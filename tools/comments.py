"""Comment/annotation tools - read, add, reply, edit, delete, and process @claude commands."""

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


def _get_anchor_range(field):
    """Get the anchor start/end positions for a comment."""
    try:
        anchor = field.getAnchor()
        if anchor:
            return anchor.getString(), anchor.getStart(), anchor.getEnd()
    except Exception:
        pass
    return "", None, None


def _ranges_overlap(start1, end1, start2, end2):
    """Check if two text ranges overlap or are adjacent."""
    if start1 is None or start2 is None:
        return False
    try:
        text = start1.getText()
        # Compare positions: returns -1, 0, or 1
        s1_before_e2 = text.compareRegionStarts(start1, end2)
        s2_before_e1 = text.compareRegionStarts(start2, end1)
        # Overlapping if s1 <= e2 and s2 <= e1
        return s1_before_e2 >= 0 and s2_before_e1 >= 0
    except Exception:
        return False


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


def _field_to_comment(idx, field):
    """Extract comment data from an annotation field."""
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

    return {
        "index": idx,
        "author": author,
        "date": date_str,
        "text": content,
        "anchor_text": anchor_text[:200],
    }


@tool(
    "get_comments",
    "List all comments (annotations) in the document. Returns author, date, text, "
    "anchor text, and any replies. Replies are separate comments anchored to the "
    "same text range with 'Re ' prefix. Use filter_author to filter by author name.",
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

    # Build all comment objects
    all_comments = []
    for idx, field in annotations:
        all_comments.append((idx, field, _field_to_comment(idx, field)))

    # Identify replies (text starts with "Re ") and group them with parents
    parents = []
    reply_map = {}  # parent_idx -> [reply comments]

    for idx, field, comment in all_comments:
        if comment["text"].startswith("Re "):
            # This is a reply - find its parent by matching anchor range
            _, reply_start, reply_end = _get_anchor_range(field)
            matched_parent = None
            for pidx, pfield, pcomment in all_comments:
                if pcomment["text"].startswith("Re "):
                    continue
                _, pstart, pend = _get_anchor_range(pfield)
                if _ranges_overlap(reply_start, reply_end, pstart, pend):
                    matched_parent = pidx
                    break
            if matched_parent is not None:
                if matched_parent not in reply_map:
                    reply_map[matched_parent] = []
                reply_map[matched_parent].append(comment)
            else:
                # No parent found, treat as standalone comment
                parents.append(comment)
        else:
            parents.append(comment)

    # Attach replies to parents
    for comment in parents:
        comment["replies"] = reply_map.get(comment["index"], [])

    # Apply author filter
    if filter_author:
        parents = [c for c in parents
                   if c["author"].lower() == filter_author.lower()]

    return {
        "count": len(parents),
        "comments": parents,
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
    "Reply to an existing comment by creating a new comment anchored to the same text. "
    "The reply is a separate annotation with the original author referenced in the text.",
    {
        "comment_index": {
            "type": "integer",
            "description": "Index of the comment to reply to (from get_comments)",
        },
        "text": {
            "type": "string",
            "description": "Reply text (will be prefixed with 'Re [author]: ')",
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

    # Get original comment info
    original_author = ""
    try:
        original_author = target.getPropertyValue("Author")
    except Exception:
        pass

    # Get the anchor position of the original comment
    try:
        anchor = target.getAnchor()
        anchor_range = anchor.getEnd()
    except Exception:
        return {"error": "Could not get anchor position of comment {}".format(comment_idx)}

    # Create a new annotation as the reply
    reply_content = "Re {}: {}".format(original_author, reply_text)

    annotation = doc.createInstance("com.sun.star.text.textfield.Annotation")
    annotation.setPropertyValue("Author", author)
    annotation.setPropertyValue("Content", reply_content)

    now = datetime.datetime.now()
    date_struct = uno.createUnoStruct("com.sun.star.util.DateTime")
    date_struct.Year = now.year
    date_struct.Month = now.month
    date_struct.Day = now.day
    date_struct.Hours = now.hour
    date_struct.Minutes = now.minute
    date_struct.Seconds = now.second
    annotation.setPropertyValue("DateTimeValue", date_struct)

    # Insert at the same anchor position as the original
    text_obj = doc.getText()
    cursor = text_obj.createTextCursorByRange(anchor_range)
    text_obj.insertTextContent(cursor, annotation, False)

    return {
        "success": True,
        "comment_index": comment_idx,
        "reply_by": author,
        "reply_to": original_author,
        "reply_text": reply_content,
    }


@tool(
    "edit_comment",
    "Edit the text of an existing comment by index.",
    {
        "comment_index": {
            "type": "integer",
            "description": "Index of the comment to edit (from get_comments)",
        },
        "new_text": {
            "type": "string",
            "description": "New comment text",
        },
    },
)
def edit_comment(args):
    comment_idx = args.get("comment_index", 0)
    new_text = args.get("new_text", "")

    if not new_text:
        return {"error": "new_text parameter is required"}

    doc = UnoConnection.get().get_document()
    annotations = _get_annotations(doc)

    target = None
    for idx, field in annotations:
        if idx == comment_idx:
            target = field
            break

    if target is None:
        return {"error": "Comment index {} not found".format(comment_idx)}

    old_text = ""
    try:
        old_text = target.getPropertyValue("Content")
    except Exception:
        pass

    target.setPropertyValue("Content", new_text)

    return {
        "success": True,
        "comment_index": comment_idx,
        "old_text": old_text[:100],
        "new_text": new_text[:100],
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

    comment_text = ""
    try:
        comment_text = target.getPropertyValue("Content")
    except Exception:
        pass

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

    # Collect all reply texts to check if a command was already processed
    reply_texts = set()
    for idx, field in annotations:
        try:
            content = field.getPropertyValue("Content")
            if content.startswith("Re "):
                reply_texts.add(content)
        except Exception:
            pass

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

        if content.lower().startswith("@claude"):
            command_text = content[7:].strip()
            if command_text.startswith(":"):
                command_text = command_text[1:].strip()

            # Check if already replied to (look for "Re [author]:" in replies)
            already_processed = False
            for rt in reply_texts:
                if "Re {}:".format(author) in rt or "Re @claude" in rt.lower():
                    already_processed = True
                    break

            if already_processed:
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
