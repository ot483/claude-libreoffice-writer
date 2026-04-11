"""Selection tools - read the current text selection in LibreOffice."""

from tools import tool
from uno_connection import UnoConnection


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


def _find_paragraph_index(doc, text_range):
    """Find which paragraph a text range belongs to."""
    paragraphs = _get_paragraphs(doc)
    text = doc.getText()
    for i, para in enumerate(paragraphs):
        try:
            # Check if the range start is within this paragraph
            cmp = text.compareRegionStarts(text_range, para.getStart())
            cmp_end = text.compareRegionStarts(text_range, para.getEnd())
            if cmp >= 0 and cmp_end <= 0:
                return i
        except Exception:
            continue
    return -1


@tool(
    "get_selection",
    "Read the text currently selected (highlighted) by the user in LibreOffice Writer. "
    "Returns the selected text, its paragraph index, and character positions. "
    "Use this when the user says 'look at my selection' or 'fix the selected text'.",
    {},
)
def get_selection(args):
    conn = UnoConnection.get()
    doc = conn.get_document()
    controller = doc.getCurrentController()
    selection = controller.getSelection()

    if selection is None:
        return {"error": "Nothing is selected in the document."}

    # Handle single selection
    if selection.supportsService("com.sun.star.text.TextRange"):
        selected_text = selection.getString()
        if not selected_text:
            return {"error": "Selection is empty (cursor is placed but no text is highlighted)."}

        para_idx = _find_paragraph_index(doc, selection.getStart())

        # Find character position within paragraph
        start_pos = -1
        if para_idx >= 0:
            paragraphs = _get_paragraphs(doc)
            para_text = paragraphs[para_idx].getString()
            pos = para_text.find(selected_text)
            if pos >= 0:
                start_pos = pos

        return {
            "selected_text": selected_text,
            "paragraph_index": para_idx,
            "start_pos": start_pos,
            "end_pos": start_pos + len(selected_text) if start_pos >= 0 else -1,
            "length": len(selected_text),
        }

    # Handle multi-range selection (user held Ctrl while selecting)
    if hasattr(selection, "getCount"):
        parts = []
        for i in range(selection.getCount()):
            part = selection.getByIndex(i)
            text = part.getString()
            if text:
                para_idx = _find_paragraph_index(doc, part.getStart())
                parts.append({
                    "text": text,
                    "paragraph_index": para_idx,
                })
        if not parts:
            return {"error": "Selection is empty."}
        return {
            "selected_text": "\n".join(p["text"] for p in parts),
            "parts": parts,
            "length": sum(len(p["text"]) for p in parts),
        }

    return {"error": "Could not read selection."}
