"""Report tools - save revision reports alongside the document."""

import os
import time

from tools import tool
from uno_connection import UnoConnection


def _get_doc_dir():
    """Get the directory of the current document."""
    doc = UnoConnection.get().get_document()
    url = doc.getURL()
    if not url:
        return None
    # Convert file:///path/to/doc.odt to /path/to/doc
    import urllib.parse
    path = urllib.parse.unquote(url.replace("file://", ""))
    return os.path.dirname(path)


def _get_doc_name():
    """Get the base name of the current document without extension."""
    doc = UnoConnection.get().get_document()
    url = doc.getURL()
    if not url:
        return "untitled"
    import urllib.parse
    path = urllib.parse.unquote(url.replace("file://", ""))
    base = os.path.basename(path)
    name, _ = os.path.splitext(base)
    return name


@tool(
    "save_report",
    "Save a text report file alongside the document. Use this to save comment "
    "response reports, revision summaries, or before/after comparisons. "
    "The file is saved in the same directory as the document.",
    {
        "content": {
            "type": "string",
            "description": "The report content (plain text or markdown)",
        },
        "filename_suffix": {
            "type": "string",
            "description": "Suffix for the filename. E.g. 'comments-report' creates "
                          "'DocumentName_comments-report.md'. Default: 'report'",
        },
    },
)
def save_report(args):
    content = args.get("content", "")
    if not content:
        return {"error": "content parameter is required"}

    suffix = args.get("filename_suffix", "report")
    doc_dir = _get_doc_dir()
    doc_name = _get_doc_name()

    if doc_dir is None:
        # Unsaved document - use home directory
        doc_dir = os.path.expanduser("~")

    timestamp = time.strftime("%Y%m%d_%H%M%S")
    filename = "{}_{}.md".format(doc_name, suffix)
    filepath = os.path.join(doc_dir, filename)

    with open(filepath, "w", encoding="utf-8") as f:
        f.write(content)

    return {
        "success": True,
        "filepath": filepath,
        "filename": filename,
        "message": "Report saved to: {}".format(filepath),
    }
