"""Navigation tools - read by section, list structure, save, undo."""

import uno

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


def _get_heading_map(doc):
    """Build a map of heading name -> (start_idx, end_idx) for all sections."""
    paragraphs = _get_paragraphs(doc)
    headings = []

    for i, para in enumerate(paragraphs):
        style = para.getPropertyValue("ParaStyleName")
        if style and ("Heading" in style or "heading" in style):
            level = 1
            # Extract level from style name like "Heading 1", "Heading 2"
            for ch in style:
                if ch.isdigit():
                    level = int(ch)
                    break
            headings.append({
                "index": i,
                "level": level,
                "title": para.getString().strip(),
                "style": style,
            })

    # Calculate end index for each heading section
    sections = []
    for j, h in enumerate(headings):
        start = h["index"]
        if j + 1 < len(headings):
            end = headings[j + 1]["index"]
        else:
            end = len(paragraphs)
        sections.append({
            "title": h["title"],
            "level": h["level"],
            "style": h["style"],
            "start_paragraph": start,
            "end_paragraph": end,
            "paragraph_count": end - start,
        })

    return sections, paragraphs


@tool(
    "list_sections",
    "List all sections (headings) in the document with their paragraph ranges. "
    "Use this to understand document structure before reading or editing specific sections.",
    {},
)
def list_sections(args):
    doc = UnoConnection.get().get_document()
    sections, paragraphs = _get_heading_map(doc)

    if not sections:
        return {
            "sections": [],
            "total_paragraphs": len(paragraphs),
            "message": "No headings found. The document may not use heading styles.",
        }

    return {
        "sections": sections,
        "total_paragraphs": len(paragraphs),
    }


@tool(
    "read_section",
    "Read a specific section by heading name. Returns all paragraphs in that section. "
    "Use list_sections first to see available section names. "
    "Matching is case-insensitive and supports partial matches.",
    {
        "heading": {
            "type": "string",
            "description": "Section heading to read (case-insensitive, partial match OK). "
                          "E.g. 'introduction', 'methods', 'results'",
        },
        "include_subsections": {
            "type": "boolean",
            "description": "Include subsections under this heading. Default: true",
        },
    },
)
def read_section(args):
    heading_query = args.get("heading", "").lower().strip()
    if not heading_query:
        return {"error": "heading parameter is required"}

    include_sub = args.get("include_subsections", True)

    doc = UnoConnection.get().get_document()
    sections, paragraphs = _get_heading_map(doc)

    if not sections:
        return {"error": "No headings found in document. Use read_document instead."}

    # Find matching section
    match = None
    for s in sections:
        if heading_query in s["title"].lower():
            match = s
            break

    if match is None:
        available = [s["title"] for s in sections]
        return {
            "error": "Section '{}' not found".format(heading_query),
            "available_sections": available,
        }

    # Determine range
    start = match["start_paragraph"]
    end = match["end_paragraph"]

    if include_sub:
        # Extend to include all subsections (higher level numbers)
        match_level = match["level"]
        for s in sections:
            if s["start_paragraph"] > start and s["level"] > match_level:
                end = max(end, s["end_paragraph"])
            elif s["start_paragraph"] > start and s["level"] <= match_level:
                break

    # Read paragraphs
    result_paragraphs = []
    for i in range(start, min(end, len(paragraphs))):
        para = paragraphs[i]
        result_paragraphs.append({
            "index": i,
            "text": para.getString(),
            "style": para.getPropertyValue("ParaStyleName"),
        })

    return {
        "section": match["title"],
        "start_paragraph": start,
        "end_paragraph": end,
        "paragraphs": result_paragraphs,
    }


@tool(
    "save_document",
    "Save the current document. If the document has been saved before, it saves to the same file. "
    "For new/untitled documents, this triggers Save As.",
    {},
)
def save_document(args):
    doc = UnoConnection.get().get_document()
    url = doc.getURL()

    if url:
        doc.store()
        return {
            "success": True,
            "message": "Document saved.",
            "url": url,
        }
    else:
        # Untitled document - use dispatch to trigger Save As dialog
        conn = UnoConnection.get()
        frame = conn.desktop.getCurrentFrame()
        dispatcher = conn.create_instance("com.sun.star.frame.DispatchHelper")
        dispatcher.executeDispatch(frame, ".uno:SaveAs", "", 0, ())
        return {
            "success": True,
            "message": "Save As dialog opened (document was untitled).",
        }


@tool(
    "undo",
    "Undo the last edit operation in the document. Can be called multiple times "
    "to undo several operations. Use this as a safety net when an edit goes wrong.",
    {
        "steps": {
            "type": "integer",
            "description": "Number of undo steps. Default: 1",
        },
    },
)
def undo(args):
    steps = args.get("steps", 1)
    doc = UnoConnection.get().get_document()
    undo_mgr = doc.getUndoManager()

    undone = 0
    descriptions = []
    for _ in range(steps):
        if undo_mgr.isUndoPossible():
            desc = undo_mgr.getCurrentUndoActionTitle()
            undo_mgr.undo()
            undone += 1
            descriptions.append(desc)
        else:
            break

    if undone == 0:
        return {"success": False, "message": "Nothing to undo."}

    return {
        "success": True,
        "steps_undone": undone,
        "descriptions": descriptions,
        "message": "Undid {} operation(s).".format(undone),
    }


@tool(
    "redo",
    "Redo the last undone operation.",
    {
        "steps": {
            "type": "integer",
            "description": "Number of redo steps. Default: 1",
        },
    },
)
def redo(args):
    steps = args.get("steps", 1)
    doc = UnoConnection.get().get_document()
    undo_mgr = doc.getUndoManager()

    redone = 0
    for _ in range(steps):
        if undo_mgr.isRedoPossible():
            undo_mgr.redo()
            redone += 1
        else:
            break

    if redone == 0:
        return {"success": False, "message": "Nothing to redo."}

    return {
        "success": True,
        "steps_redone": redone,
        "message": "Redid {} operation(s).".format(redone),
    }
