"""Style tools - paragraph styles, character formatting, style listing."""

import json
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
    "set_paragraph_style",
    "Apply a named paragraph style (e.g. 'Heading 1', 'Text Body', 'List Bullet') "
    "to a paragraph. Use list_styles to see available styles.",
    {
        "paragraph_index": {
            "type": "integer",
            "description": "Target paragraph index (0-based)",
        },
        "style_name": {
            "type": "string",
            "description": "Paragraph style name (e.g. 'Heading 1', 'Text Body')",
        },
    },
)
def set_paragraph_style(args):
    para_idx = args.get("paragraph_index", 0)
    style_name = args.get("style_name", "")
    if not style_name:
        return {"error": "style_name parameter is required"}

    doc = UnoConnection.get().get_document()
    paragraphs = _get_paragraphs(doc)

    if para_idx < 0 or para_idx >= len(paragraphs):
        return {"error": "Paragraph index {} out of range (0-{})".format(
            para_idx, len(paragraphs) - 1
        )}

    # Validate style exists
    style_families = doc.getStyleFamilies()
    para_styles = style_families.getByName("ParagraphStyles")
    if not para_styles.hasByName(style_name):
        available = list(para_styles.getElementNames())[:20]
        return {
            "error": "Style '{}' not found. Some available styles: {}".format(
                style_name, ", ".join(available)
            )
        }

    para = paragraphs[para_idx]
    para.setPropertyValue("ParaStyleName", style_name)

    return {
        "success": True,
        "paragraph_index": para_idx,
        "applied_style": style_name,
    }


@tool(
    "set_character_format",
    "Apply character formatting (bold, italic, underline, font, size, color) "
    "to a text range within a paragraph. All format parameters are optional - "
    "only specified ones are applied.",
    {
        "paragraph_index": {
            "type": "integer",
            "description": "Target paragraph index (0-based)",
        },
        "start_pos": {
            "type": "integer",
            "description": "Start character position within the paragraph",
        },
        "end_pos": {
            "type": "integer",
            "description": "End character position (exclusive). -1 = end of paragraph",
        },
        "bold": {
            "type": "boolean",
            "description": "Set bold. Optional.",
        },
        "italic": {
            "type": "boolean",
            "description": "Set italic. Optional.",
        },
        "underline": {
            "type": "boolean",
            "description": "Set underline. Optional.",
        },
        "font_name": {
            "type": "string",
            "description": "Font name (e.g. 'Arial', 'Times New Roman'). Optional.",
        },
        "font_size": {
            "type": "number",
            "description": "Font size in points (e.g. 12, 14.5). Optional.",
        },
        "color": {
            "type": "string",
            "description": "Text color as hex string (e.g. '#FF0000' for red). Optional.",
        },
    },
)
def set_character_format(args):
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

    # Create cursor over the range
    text_obj = doc.getText()
    cursor = text_obj.createTextCursorByRange(para.getStart())
    if start_pos > 0:
        cursor.goRight(start_pos, False)
    cursor.goRight(end_pos - start_pos, True)  # Select the range

    applied = []

    # Bold
    if "bold" in args:
        from com.sun.star.awt.FontWeight import BOLD, NORMAL
        cursor.setPropertyValue("CharWeight", BOLD if args["bold"] else NORMAL)
        applied.append("bold={}".format(args["bold"]))

    # Italic
    if "italic" in args:
        from com.sun.star.awt.FontSlant import ITALIC, NONE
        cursor.setPropertyValue("CharPosture", ITALIC if args["italic"] else NONE)
        applied.append("italic={}".format(args["italic"]))

    # Underline
    if "underline" in args:
        from com.sun.star.awt.FontUnderline import SINGLE, NONE as UL_NONE
        cursor.setPropertyValue("CharUnderline", SINGLE if args["underline"] else UL_NONE)
        applied.append("underline={}".format(args["underline"]))

    # Font name
    if "font_name" in args:
        cursor.setPropertyValue("CharFontName", args["font_name"])
        applied.append("font={}".format(args["font_name"]))

    # Font size
    if "font_size" in args:
        cursor.setPropertyValue("CharHeight", float(args["font_size"]))
        applied.append("size={}".format(args["font_size"]))

    # Color
    if "color" in args:
        color_hex = args["color"].lstrip("#")
        color_int = int(color_hex, 16)
        cursor.setPropertyValue("CharColor", color_int)
        applied.append("color=#{}".format(color_hex))

    return {
        "success": True,
        "paragraph_index": para_idx,
        "range": "{}-{}".format(start_pos, end_pos),
        "applied": applied,
        "text_affected": para_text[start_pos:end_pos],
    }


@tool(
    "list_styles",
    "List available styles by family. Use this to discover valid style names "
    "before applying them with set_paragraph_style.",
    {
        "family": {
            "type": "string",
            "description": "Style family: 'ParagraphStyles', 'CharacterStyles', or 'PageStyles'. Default: 'ParagraphStyles'",
        },
    },
)
def list_styles(args):
    family = args.get("family", "ParagraphStyles")

    doc = UnoConnection.get().get_document()
    style_families = doc.getStyleFamilies()

    available_families = list(style_families.getElementNames())
    if family not in available_families:
        return {
            "error": "Unknown family '{}'. Available: {}".format(
                family, ", ".join(available_families)
            )
        }

    styles_collection = style_families.getByName(family)
    style_names = sorted(styles_collection.getElementNames())

    return {
        "family": family,
        "count": len(style_names),
        "styles": style_names,
    }
