"""Read tools - inspect document content, metadata, and formatting."""

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
        # Skip tables and other non-paragraph elements
        if para.supportsService("com.sun.star.text.Paragraph"):
            paragraphs.append(para)
    return paragraphs


@tool(
    "read_document",
    "Read document text. Returns paragraphs with their text and style names. "
    "Use start_para/end_para for pagination on large documents. "
    "Does NOT return formatting details - use get_paragraph_details for that.",
    {
        "start_para": {
            "type": "integer",
            "description": "First paragraph index (0-based). Default: 0",
        },
        "end_para": {
            "type": "integer",
            "description": "Last paragraph index (exclusive). Default: -1 (all)",
        },
    },
)
def read_document(args):
    doc = UnoConnection.get().get_document()
    paragraphs = _get_paragraphs(doc)
    total = len(paragraphs)

    start = args.get("start_para", 0)
    end = args.get("end_para", -1)
    if end == -1 or end > total:
        end = total

    result = []
    for i in range(start, end):
        para = paragraphs[i]
        result.append({
            "index": i,
            "text": para.getString(),
            "style": para.getPropertyValue("ParaStyleName"),
        })

    return {
        "paragraphs": result,
        "total_paragraphs": total,
        "range": "{}-{}".format(start, end),
    }


@tool(
    "get_document_info",
    "Get document metadata: title, file path, page count, paragraph count, "
    "word count, and whether it has unsaved changes.",
    {},
)
def get_document_info(args):
    doc = UnoConnection.get().get_document()
    paragraphs = _get_paragraphs(doc)

    props = doc.getDocumentProperties()

    # Word count via document statistics
    word_count = 0
    char_count = 0
    try:
        word_count = doc.getPropertyValue("WordCount")
        char_count = doc.getPropertyValue("CharacterCount")
    except Exception:
        pass

    # Page count from the controller
    page_count = 0
    try:
        controller = doc.getCurrentController()
        page_count = controller.getPropertyValue("PageCount")
    except Exception:
        pass

    return {
        "title": props.Title or "(untitled)",
        "url": doc.getURL() or "(unsaved)",
        "page_count": page_count,
        "paragraph_count": len(paragraphs),
        "word_count": word_count,
        "character_count": char_count,
        "modified": doc.isModified(),
        "description": props.Description,
        "author": props.Author,
    }


@tool(
    "search_text",
    "Search for text in the document. Returns matching text with paragraph index "
    "and position. Supports regex and case-sensitive search.",
    {
        "query": {
            "type": "string",
            "description": "Text to search for",
        },
        "regex": {
            "type": "boolean",
            "description": "Use regular expression matching. Default: false",
        },
        "case_sensitive": {
            "type": "boolean",
            "description": "Case-sensitive search. Default: false",
        },
    },
)
def search_text(args):
    query = args.get("query", "")
    if not query:
        return {"error": "query parameter is required"}

    doc = UnoConnection.get().get_document()
    search_desc = doc.createSearchDescriptor()
    search_desc.SearchString = query
    search_desc.SearchRegularExpression = args.get("regex", False)
    search_desc.SearchCaseSensitive = args.get("case_sensitive", False)

    found_all = doc.findAll(search_desc)
    matches = []

    if found_all:
        paragraphs = _get_paragraphs(doc)
        para_map = {}
        for i, p in enumerate(paragraphs):
            try:
                para_map[id(p.getText())] = i
            except Exception:
                pass

        for i in range(found_all.getCount()):
            match = found_all.getByIndex(i)
            match_text = match.getString()

            # Try to find which paragraph this match is in
            para_idx = -1
            try:
                match_text_obj = match.getText()
                # Walk paragraphs to find which one contains this range
                for pi, p in enumerate(paragraphs):
                    p_text = p.getString()
                    if match_text in p_text:
                        para_idx = pi
                        break
            except Exception:
                pass

            matches.append({
                "text": match_text,
                "paragraph_index": para_idx,
                "start": match.getStart().getPropertyValue("TextPortionType") if False else -1,
            })

    return {
        "query": query,
        "match_count": len(matches),
        "matches": matches[:100],  # Cap at 100 results
    }


@tool(
    "get_paragraph_details",
    "Get detailed formatting for a specific paragraph, including character runs "
    "with their bold, italic, font, size, and color properties.",
    {
        "paragraph_index": {
            "type": "integer",
            "description": "Paragraph index (0-based)",
        },
    },
)
def get_paragraph_details(args):
    para_idx = args.get("paragraph_index", 0)

    doc = UnoConnection.get().get_document()
    paragraphs = _get_paragraphs(doc)

    if para_idx < 0 or para_idx >= len(paragraphs):
        return {"error": "Paragraph index {} out of range (0-{})".format(
            para_idx, len(paragraphs) - 1
        )}

    para = paragraphs[para_idx]

    # Get paragraph-level properties
    para_style = para.getPropertyValue("ParaStyleName")

    # Enumerate text portions (character runs)
    runs = []
    portion_enum = para.createEnumeration()
    while portion_enum.hasMoreElements():
        portion = portion_enum.nextElement()
        portion_type = portion.getPropertyValue("TextPortionType")

        if portion_type == "Text":
            text = portion.getString()
            if not text:
                continue

            # Read character formatting
            bold = False
            italic = False
            underline = False
            font_name = ""
            font_size = 0.0
            color = ""

            try:
                weight = portion.getPropertyValue("CharWeight")
                bold = weight > 100  # BOLD = 150, NORMAL = 100
            except Exception:
                pass
            try:
                posture = portion.getPropertyValue("CharPosture")
                italic = posture.value != 0  # NONE = 0
            except Exception:
                pass
            try:
                ul = portion.getPropertyValue("CharUnderline")
                underline = ul != 0  # NONE = 0
            except Exception:
                pass
            try:
                font_name = portion.getPropertyValue("CharFontName")
            except Exception:
                pass
            try:
                font_size = portion.getPropertyValue("CharHeight")
            except Exception:
                pass
            try:
                c = portion.getPropertyValue("CharColor")
                color = "#{:06x}".format(c) if c >= 0 else ""
            except Exception:
                pass

            runs.append({
                "text": text,
                "bold": bold,
                "italic": italic,
                "underline": underline,
                "font": font_name,
                "size": font_size,
                "color": color,
            })

    return {
        "paragraph_index": para_idx,
        "style": para_style,
        "full_text": para.getString(),
        "runs": runs,
    }
