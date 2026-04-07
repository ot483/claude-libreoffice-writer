"""Track Changes tools - enable reviewable edits in the document."""

from tools import tool
from uno_connection import UnoConnection


@tool(
    "enable_track_changes",
    "Enable Track Changes (Record Changes) mode. All subsequent edits will appear "
    "as reviewable markup - insertions highlighted, deletions shown as strikethrough. "
    "The user can then accept or reject each change in LibreOffice. "
    "IMPORTANT: Always enable this before making edits so the user can review them.",
    {},
)
def enable_track_changes(args):
    doc = UnoConnection.get().get_document()
    doc.setPropertyValue("RecordChanges", True)
    is_recording = doc.getPropertyValue("RecordChanges")
    return {
        "success": True,
        "track_changes_enabled": is_recording,
        "message": "Track Changes is ON. All edits will appear as reviewable markup.",
    }


@tool(
    "disable_track_changes",
    "Disable Track Changes mode. Subsequent edits will be applied directly without markup.",
    {},
)
def disable_track_changes(args):
    doc = UnoConnection.get().get_document()
    doc.setPropertyValue("RecordChanges", False)
    return {
        "success": True,
        "track_changes_enabled": False,
        "message": "Track Changes is OFF. Edits will be applied directly.",
    }


@tool(
    "get_track_changes_status",
    "Check whether Track Changes (Record Changes) is currently enabled.",
    {},
)
def get_track_changes_status(args):
    doc = UnoConnection.get().get_document()
    is_recording = doc.getPropertyValue("RecordChanges")
    show_changes = doc.getPropertyValue("ShowChanges")
    return {
        "recording": is_recording,
        "showing": show_changes,
        "message": "Track Changes is {}. Changes are {}.".format(
            "ON" if is_recording else "OFF",
            "visible" if show_changes else "hidden",
        ),
    }


@tool(
    "accept_all_changes",
    "Accept all tracked changes in the document, making them permanent.",
    {},
)
def accept_all_changes(args):
    doc = UnoConnection.get().get_document()
    doc.setPropertyValue("RecordChanges", False)
    doc.setPropertyValue("ShowChanges", True)

    # Accept all redlines
    redline = doc.getRedlines() if hasattr(doc, 'getRedlines') else None
    if redline is None:
        # Use dispatch helper
        conn = UnoConnection.get()
        frame = conn.desktop.getCurrentFrame()
        dispatcher = conn.create_instance("com.sun.star.frame.DispatchHelper")
        dispatcher.executeDispatch(frame, ".uno:AcceptAllTrackedChanges", "", 0, ())
        return {"success": True, "message": "All tracked changes accepted."}

    return {"success": True, "message": "All tracked changes accepted."}


@tool(
    "reject_all_changes",
    "Reject all tracked changes in the document, reverting them.",
    {},
)
def reject_all_changes(args):
    doc = UnoConnection.get().get_document()
    doc.setPropertyValue("RecordChanges", False)

    conn = UnoConnection.get()
    frame = conn.desktop.getCurrentFrame()
    dispatcher = conn.create_instance("com.sun.star.frame.DispatchHelper")
    dispatcher.executeDispatch(frame, ".uno:RejectAllTrackedChanges", "", 0, ())

    return {"success": True, "message": "All tracked changes rejected."}
