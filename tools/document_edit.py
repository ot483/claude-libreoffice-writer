"""Edit tools - insert, replace, and delete document text."""

import json
import uno
from com.sun.star.text.ControlCharacter import PARAGRAPH_BREAK

from tools import tool
from uno_connection import UnoConnection


def _get_paragraphs(doc):
    """Enumerate all paragraphs in the document body."""
    text = doc.getText()
    enum = text.createEnumeration()
    paragraphs = []
    while enum.hasMoreElements():
        para = enum.nextElement()
        if para.supportsService("com.sun.star.text.Paragraph"):
            paragraphs.append(para)
    return paragraphs


@tool(
    "insert_text",
    "Insert text into the document at a specific paragraph and character position. "
    "Set as_new_paragraph=true to insert a paragraph break before the text.",
    {
        "text": {
            "type": "string",
            "description": "Text to insert",
        },
        "paragraph_index": {
            "type": "integer",
            "description": "Target paragraph index (0-based)",
        },
        "position": {
            "type": "integer",
            "description": "Character position within the paragraph. -1 = end of paragraph. Default: -1",
        },
        "as_new_paragraph": {
            "type": "boolean",
            "description": "Insert a paragraph break before the text. Default: false",
        },
    },
)
def insert_text(args):
    text_to_insert = args.get("text", "")
    if not text_to_insert:
        return {"error": "text parameter is required"}

    para_idx = args.get("paragraph_index", 0)
    position = args.get("position", -1)
    as_new_para = args.get("as_new_paragraph", False)

    doc = UnoConnection.get().get_document()
    paragraphs = _get_paragraphs(doc)

    if para_idx < 0 or para_idx >= len(paragraphs):
        return {"error": "Paragraph index {} out of range (0-{})".format(
            para_idx, len(paragraphs) - 1
        )}

    para = paragraphs[para_idx]
    text_obj = doc.getText()

    # Create a cursor at the target position
    cursor = text_obj.createTextCursorByRange(para.getStart())
    para_text = para.getString()

    if position == -1 or position >= len(para_text):
        cursor = text_obj.createTextCursorByRange(para.getEnd())
    elif position > 0:
        cursor.goRight(position, False)

    if as_new_para:
        text_obj.insertControlCharacter(cursor, PARAGRAPH_BREAK, False)

    text_obj.insertString(cursor, text_to_insert, False)

    return {
        "success": True,
        "inserted_at": {
            "paragraph_index": para_idx,
            "position": position,
        },
        "text_length": len(text_to_insert),
    }


@tool(
    "replace_text",
    "Search and replace text in the document. Supports regex. "
    "Returns the number of replacements made.",
    {
        "search": {
            "type": "string",
            "description": "Text to find",
        },
        "replacement": {
            "type": "string",
            "description": "Replacement text",
        },
        "regex": {
            "type": "boolean",
            "description": "Use regular expression matching. Default: false",
        },
        "replace_all": {
            "type": "boolean",
            "description": "Replace all occurrences (true) or just the first (false). Default: true",
        },
    },
)
def replace_text(args):
    search = args.get("search", "")
    replacement = args.get("replacement", "")
    if not search:
        return {"error": "search parameter is required"}

    replace_all = args.get("replace_all", True)

    doc = UnoConnection.get().get_document()
    replace_desc = doc.createSearchDescriptor()
    replace_desc.SearchString = search
    replace_desc.ReplaceString = replacement
    replace_desc.SearchRegularExpression = args.get("regex", False)

    if replace_all:
        count = doc.replaceAll(replace_desc)
    else:
        # Find first and replace just that one
        found = doc.findFirst(replace_desc)
        if found:
            found.setString(replacement)
            count = 1
        else:
            count = 0

    return {
        "success": True,
        "search": search,
        "replacement": replacement,
        "replacements_made": count,
    }


@tool(
    "delete_text",
    "Delete text within a paragraph by character range. "
    "Omit end_pos to delete to end of paragraph.",
    {
        "paragraph_index": {
            "type": "integer",
            "description": "Target paragraph index (0-based)",
        },
        "start_pos": {
            "type": "integer",
            "description": "Start character position. Default: 0",
        },
        "end_pos": {
            "type": "integer",
            "description": "End character position (exclusive). -1 = end of paragraph. Default: -1",
        },
    },
)
def delete_text(args):
    para_idx = args.get("paragraph_index", 0)
    start_pos = args.get("start_pos", 0)
    end_pos = args.get("end_pos", -1)

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

    if start_pos >= end_pos:
        return {"error": "start_pos ({}) must be less than end_pos ({})".format(
            start_pos, end_pos
        )}

    deleted_text = para_text[start_pos:end_pos]

    # Create cursor over the range to delete
    text_obj = doc.getText()
    cursor = text_obj.createTextCursorByRange(para.getStart())
    if start_pos > 0:
        cursor.goRight(start_pos, False)
    cursor.goRight(end_pos - start_pos, True)  # Select the range
    cursor.setString("")  # Delete selected text

    return {
        "success": True,
        "deleted_text": deleted_text,
        "paragraph_index": para_idx,
    }


@tool(
    "insert_paragraph",
    "Insert a new paragraph at a specific position with optional paragraph style. "
    "Use after_paragraph=-1 to append at the end of the document.",
    {
        "text": {
            "type": "string",
            "description": "Paragraph text",
        },
        "after_paragraph": {
            "type": "integer",
            "description": "Insert after this paragraph index. -1 = end of document. Default: -1",
        },
        "style": {
            "type": "string",
            "description": "Paragraph style name (e.g. 'Heading 1', 'Text Body'). Optional.",
        },
    },
)
def insert_paragraph(args):
    text_to_insert = args.get("text", "")
    after_para = args.get("after_paragraph", -1)
    style = args.get("style", "")

    doc = UnoConnection.get().get_document()
    text_obj = doc.getText()
    paragraphs = _get_paragraphs(doc)

    if after_para == -1:
        # Append at end
        cursor = text_obj.createTextCursorByRange(text_obj.getEnd())
    elif after_para < 0 or after_para >= len(paragraphs):
        return {"error": "Paragraph index {} out of range (0-{})".format(
            after_para, len(paragraphs) - 1
        )}
    else:
        cursor = text_obj.createTextCursorByRange(paragraphs[after_para].getEnd())

    # Insert paragraph break then text
    text_obj.insertControlCharacter(cursor, PARAGRAPH_BREAK, False)

    if style:
        try:
            cursor.setPropertyValue("ParaStyleName", style)
        except Exception as e:
            pass  # Style might not exist - continue with default

    text_obj.insertString(cursor, text_to_insert, False)

    new_idx = after_para + 1 if after_para >= 0 else len(paragraphs)

    return {
        "success": True,
        "paragraph_index": new_idx,
        "style": style or "(default)",
    }
