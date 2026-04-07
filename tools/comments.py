"""Comment/annotation tools - read, add, reply, delete, and process @claude commands."""

import json
import datetime
import uno
from tools import tool
from uno_connection import UnoConnection


def _get_annotations(doc):
    """Get all annotation text fields from the document."""
    annotations = []
    text_fields = doc.getTextFields()
    enum = text_fields.createEnumeration()
    idx = 0
    while enum.hasMoreElements():
        field = enum.nextElement()
        if field.supportsService("com.sun.star.text.TextField.Annotation"):
            annotations.append((idx, field))
            idx += 1
    return annotations


def _get_paragraphs(doc):
    """Enumerate all paragraphs."""
    text = doc.getText()
    enum = text.createEnumeration()
    paragraphs = []
    while enum.hasMoreElements():
        para = enum.nextElement()
        if para.supportsService("com.sun.star.text.Paragraph"):
            paragraphs.append(para)
    return paragraphs


@tool(
    "get_comments",
    "List all comments (annotations) in the document. Returns author, date, text, "
    "and the text the comment is anchored to. Use filter_author to filter by author name.",
    {
        "filter_author": {
            "type": "string",
            "description": "Only return comments by this author. Optional.",
        },
    },
)
def get_comments(args):
    filter_author = args.get("filter_author", "")

    doc = UnoConnection.get().get_document()
    annotations = _get_annotations(doc)

    comments = []
    for idx, field in annotations:
        author = ""
        content = ""
        date_str = ""
        anchor_text = ""

        try:
            author = field.getPropertyValue("Author")
        except Exception:
            pass
        try:
            content = field.getPropertyValue("Content")
        except Exception:
            pass
        try:
            dt = field.getPropertyValue("DateTimeValue")
            date_str = "{:04d}-{:02d}-{:02d} {:02d}:{:02d}".format(
                dt.Year, dt.Month, dt.Day, dt.Hours, dt.Minutes
            )
        except Exception:
            pass
        try:
            anchor = field.getAnchor()
            if anchor:
                anchor_text = anchor.getString()
        except Exception:
            pass

        if filter_author and author.lower() != filter_author.lower():
            continue

        comments.append({
            "index": idx,
            "author": author,
            "date": date_str,
            "text": content,
            "anchor_text": anchor_text[:200],  # Truncate long anchors
        })

    return {
        "count": len(comments),
        "comments": comments,
    }


@tool(
    "add_comment",
    "Add a comment (annotation) anchored to a text range in the document. "
    "Specify the paragraph and character range to anchor the comment to.",
    {
        "paragraph_index": {
            "type": "integer",
            "description": "Paragraph to anchor the comment to (0-based)",
        },
        "start_pos": {
            "type": "integer",
            "description": "Start character position for the anchor. Default: 0",
        },
        "end_pos": {
            "type": "integer",
            "description": "End character position for the anchor. -1 = end of paragraph. Default: -1",
        },
        "text": {
            "type": "string",
            "description": "Comment text",
        },
        "author": {
            "type": "string",
            "description": "Comment author name. Default: 'Claude'",
        },
    },
)
def add_comment(args):
    text_content = args.get("text", "")
    if not text_content:
        return {"error": "text parameter is required"}

    para_idx = args.get("paragraph_index", 0)
    start_pos = args.get("start_pos", 0)
    end_pos = args.get("end_pos", -1)
    author = args.get("author", "Claude")

    doc = UnoConnection.get().get_document()
    paragraphs = _get_paragraphs(doc)

    if para_idx < 0 or para_idx >= len(paragraphs):
        return {"error": "Paragraph index {} out of range (0-{})".format(
            para_idx, len(paragraphs) - 1
        )}

    para = paragraphs[para_idx]
    para_text = para.getString()

    if end_pos == -1 or end_pos > len(para_text):
        end_pos = len(para_text)

    # Create the annotation
    annotation = doc.createInstance("com.sun.star.text.textfield.Annotation")
    annotation.setPropertyValue("Author", author)
    annotation.setPropertyValue("Content", text_content)

    # Set current date/time
    now = datetime.datetime.now()
    date_struct = uno.createUnoStruct("com.sun.star.util.DateTime")
    date_struct.Year = now.year
    date_struct.Month = now.month
    date_struct.Day = now.day
    date_struct.Hours = now.hour
    date_struct.Minutes = now.minute
    date_struct.Seconds = now.second
    annotation.setPropertyValue("DateTimeValue", date_struct)

    # Create cursor at the anchor position
    text_obj = doc.getText()
    cursor = text_obj.createTextCursorByRange(para.getStart())
    if end_pos > 0:
        cursor.goRight(end_pos, False)

    # Insert the annotation at the cursor position
    text_obj.insertTextContent(cursor, annotation, False)

    return {
        "success": True,
        "author": author,
        "text": text_content,
        "anchored_to_paragraph": para_idx,
    }


@tool(
    "reply_to_comment",
    "Reply to an existing comment by appending text to it. "
    "(LibreOffice 6.4 does not support threaded replies, so replies are appended.)",
    {
        "comment_index": {
            "type": "integer",
            "description": "Index of the comment to reply to (from get_comments)",
        },
        "text": {
            "type": "string",
            "description": "Reply text",
        },
        "author": {
            "type": "string",
            "description": "Reply author name. Default: 'Claude'",
        },
    },
)
def reply_to_comment(args):
    comment_idx = args.get("comment_index", 0)
    reply_text = args.get("text", "")
    author = args.get("author", "Claude")

    if not reply_text:
        return {"error": "text parameter is required"}

    doc = UnoConnection.get().get_document()
    annotations = _get_annotations(doc)

    target = None
    for idx, field in annotations:
        if idx == comment_idx:
            target = field
            break

    if target is None:
        return {"error": "Comment index {} not found".format(comment_idx)}

    # Append reply to existing content
    current_content = target.getPropertyValue("Content")
    separator = "\n\n--- Reply by {} ---\n".format(author)
    new_content = current_content + separator + reply_text
    target.setPropertyValue("Content", new_content)

    return {
        "success": True,
        "comment_index": comment_idx,
        "reply_by": author,
        "reply_text": reply_text,
    }


@tool(
    "delete_comment",
    "Delete a comment (annotation) from the document by its index.",
    {
        "comment_index": {
            "type": "integer",
            "description": "Index of the comment to delete (from get_comments)",
        },
    },
)
def delete_comment(args):
    comment_idx = args.get("comment_index", 0)

    doc = UnoConnection.get().get_document()
    annotations = _get_annotations(doc)

    target = None
    for idx, field in annotations:
        if idx == comment_idx:
            target = field
            break

    if target is None:
        return {"error": "Comment index {} not found".format(comment_idx)}

    # Get the comment text before deleting for confirmation
    comment_text = ""
    try:
        comment_text = target.getPropertyValue("Content")
    except Exception:
        pass

    # Remove the text field
    doc.getText().removeTextContent(target)

    return {
        "success": True,
        "deleted_comment_index": comment_idx,
        "deleted_text": comment_text[:100],
    }


@tool(
    "process_claude_comments",
    "Scan all comments for ones starting with '@claude' - these are commands from "
    "the user written directly in the document. Returns the commands for Claude to "
    "act on. After processing, Claude should reply to each comment with the result.",
    {},
)
def process_claude_comments(args):
    doc = UnoConnection.get().get_document()
    annotations = _get_annotations(doc)

    commands = []
    for idx, field in annotations:
        content = ""
        author = ""
        anchor_text = ""

        try:
            content = field.getPropertyValue("Content")
        except Exception:
            pass
        try:
            author = field.getPropertyValue("Author")
        except Exception:
            pass
        try:
            anchor = field.getAnchor()
            if anchor:
                anchor_text = anchor.getString()
        except Exception:
            pass

        # Check if this is a @claude command (case-insensitive)
        if content.lower().startswith("@claude"):
            command_text = content[7:].strip()  # Remove "@claude" prefix
            if command_text.startswith(":"):
                command_text = command_text[1:].strip()  # Remove optional colon

            # Skip if already processed (has a reply)
            if "--- Reply by Claude ---" in content:
                continue

            commands.append({
                "comment_index": idx,
                "author": author,
                "command": command_text,
                "anchor_text": anchor_text[:200],
            })

    return {
        "count": len(commands),
        "commands": commands,
        "instructions": (
            "Process each command and use reply_to_comment to post results. "
            "Commands reference the anchored text for context."
        ) if commands else "No pending @claude commands found.",
    }
