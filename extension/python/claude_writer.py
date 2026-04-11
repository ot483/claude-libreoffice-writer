"""Claude Writer - LibreOffice extension macro.

When the toolbar button is pressed:
1. Starts a UNO acceptor thread so external processes can connect (port 2002)
2. Launches Claude Code in a terminal
"""

import os
import sys
import time
import subprocess
import threading
import tempfile

try:
    import uno
    import unohelper
except ImportError:
    uno = None
    unohelper = None


# ---------- Configuration ----------

UNO_PORT = 2002
UNO_ACCEPT_STRING = "socket,host=localhost,port={}".format(UNO_PORT)
LOG_FILE = os.path.join(tempfile.gettempdir(), "claude_writer.log")

# ---------- UNO Acceptor ----------

_acceptor_running = False
_acceptor_lock = threading.Lock()


def _make_instance_provider(ctx):
    """Create an InstanceProvider object. Imports UNO types lazily."""
    from com.sun.star.bridge import XInstanceProvider

    class _InstanceProvider(unohelper.Base, XInstanceProvider):
        def __init__(self, ctx):
            self._ctx = ctx

        def getInstance(self, name):
            if name == "StarOffice.ComponentContext":
                return self._ctx
            raise Exception("Unknown instance requested: " + name)

    return _InstanceProvider(ctx)


def _acceptor_loop(ctx):
    """Background thread: accept UNO socket connections on the configured port."""
    global _acceptor_running
    smgr = ctx.ServiceManager

    try:
        acceptor = smgr.createInstanceWithContext(
            "com.sun.star.connection.Acceptor", ctx
        )
        bridge_factory = smgr.createInstanceWithContext(
            "com.sun.star.bridge.BridgeFactory", ctx
        )
    except Exception as e:
        _log("Failed to create acceptor services: {}".format(e))
        _acceptor_running = False
        return

    provider = _make_instance_provider(ctx)
    bridge_count = 0
    _log("UNO acceptor listening on port {}".format(UNO_PORT))

    while _acceptor_running:
        try:
            connection = acceptor.accept(UNO_ACCEPT_STRING)
            if connection is None:
                continue
            bridge_count += 1
            bridge_name = "claude_bridge_{}".format(bridge_count)
            bridge_factory.createBridge(bridge_name, "urp", connection, provider)
            _log("Bridge established: {}".format(bridge_name))
        except Exception as e:
            err = str(e)
            if "disposed" in err.lower() or "interrupt" in err.lower():
                break
            _log("Acceptor error (will retry): {}".format(err))
            time.sleep(1)


def _ensure_acceptor(ctx):
    """Start the acceptor thread if not already running."""
    global _acceptor_running
    with _acceptor_lock:
        if _acceptor_running:
            return True
        if _port_in_use(UNO_PORT):
            _acceptor_running = True
            _log("Port {} already in use - acceptor not needed".format(UNO_PORT))
            return True
        _acceptor_running = True
        t = threading.Thread(target=_acceptor_loop, args=(ctx,), daemon=True)
        t.start()
        time.sleep(0.5)
        return True


def _port_in_use(port):
    import socket
    s = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
    try:
        s.settimeout(0.5)
        s.connect(("localhost", port))
        s.close()
        return True
    except Exception:
        s.close()
        return False


# ---------- Terminal Detection ----------

def _find_terminal():
    """Find an available terminal emulator. Uses Claude Writer profile if available."""
    # Check for Claude Writer gnome-terminal profile
    gnome_profile = _get_gnome_terminal_profile()

    terminals = [
        ("gnome-terminal", ["gnome-terminal"] + (["--profile", "Claude Writer"] if gnome_profile else []) + ["--"]),
        ("xfce4-terminal", ["xfce4-terminal", "-e"]),
        ("konsole", ["konsole", "-e"]),
        ("xterm", ["xterm", "-e"]),
        ("mate-terminal", ["mate-terminal", "-e"]),
        ("lxterminal", ["lxterminal", "-e"]),
        ("terminator", ["terminator", "-e"]),
        ("tilix", ["tilix", "-e"]),
    ]
    for name, cmd in terminals:
        try:
            subprocess.run(["which", name], capture_output=True, check=True)
            return cmd
        except Exception:
            continue
    return None


def _get_gnome_terminal_profile():
    """Check if the Claude Writer gnome-terminal profile exists."""
    try:
        result = subprocess.run(
            ["dconf", "list", "/org/gnome/terminal/legacy/profiles:/"],
            capture_output=True, text=True
        )
        for line in result.stdout.strip().split("\n"):
            profile_id = line.strip().strip("/:")
            if not profile_id:
                continue
            name_result = subprocess.run(
                ["dconf", "read",
                 "/org/gnome/terminal/legacy/profiles:/:{}/visible-name".format(profile_id)],
                capture_output=True, text=True
            )
            if "Claude Writer" in name_result.stdout:
                return True
    except Exception:
        pass
    return False


# ---------- Main Entry Point ----------

def StartClaude(*args):
    """Toolbar button handler - start acceptor + launch Claude Code."""
    try:
        ctx = XSCRIPTCONTEXT.getComponentContext()
    except NameError:
        ctx = uno.getComponentContext()

    _ensure_acceptor(ctx)

    launcher = _find_launcher()
    if launcher is None:
        _show_message(
            "Claude Writer",
            "Launcher script not found.\n\n"
            "Please run install.sh again.",
        )
        return

    terminal_cmd = _find_terminal()
    if terminal_cmd is None:
        _show_message(
            "Claude Writer",
            "No terminal emulator found.\n"
            "Please open a terminal and run: claude",
        )
        return

    _launch_claude(terminal_cmd, launcher)


def _find_launcher():
    """Find the Claude Writer launcher script."""
    candidates = []

    # Next to this script
    try:
        this_dir = os.path.dirname(os.path.abspath(__file__))
        candidates.append(os.path.join(this_dir, "launcher.sh"))
        candidates.append(os.path.join(os.path.dirname(this_dir), "launcher.sh"))
    except Exception:
        pass

    # Known install location
    home = os.path.expanduser("~")
    candidates.append(os.path.join(home, ".config/libreoffice/4/user/Scripts/python/launcher.sh"))

    for p in candidates:
        if os.path.isfile(p) and os.access(p, os.X_OK):
            return p
    return None


def _launch_claude(terminal_cmd, launcher):
    """Open a terminal with the Claude Writer launcher."""
    cmd = list(terminal_cmd)
    if "gnome-terminal" in cmd[0]:
        cmd.extend(["bash", "-c", launcher])
    else:
        cmd.append(launcher)

    try:
        subprocess.Popen(cmd, start_new_session=True)
        _log("Launched Claude Writer: {}".format(" ".join(cmd)))
    except Exception as e:
        _show_message("Claude Writer", "Failed to launch terminal:\n{}".format(e))


# ---------- Helpers ----------

def _show_message(title, text):
    try:
        ctx = XSCRIPTCONTEXT.getComponentContext()
        smgr = ctx.ServiceManager
        toolkit = smgr.createInstanceWithContext("com.sun.star.awt.Toolkit", ctx)
        parent = XSCRIPTCONTEXT.getDesktop().getCurrentFrame().getContainerWindow()
        from com.sun.star.awt.MessageBoxType import MESSAGEBOX
        box = toolkit.createMessageBox(parent, MESSAGEBOX, 1, title, text)
        box.execute()
    except Exception:
        pass


def _log(msg):
    sys.stderr.write("[Claude Writer] {}\n".format(msg))
    sys.stderr.flush()
    try:
        with open(LOG_FILE, "a") as f:
            f.write("{} {}\n".format(time.strftime("%H:%M:%S"), msg))
    except Exception:
        pass


def AskClaude(*args):
    """Right-click context menu handler - send selected text to Claude."""
    try:
        ctx = XSCRIPTCONTEXT.getComponentContext()
        doc = XSCRIPTCONTEXT.getDocument()
    except NameError:
        return

    # Get selected text
    controller = doc.getCurrentController()
    selection = controller.getSelection()
    if selection is None:
        _show_message("Claude Writer", "Select some text first, then right-click > Ask Claude.")
        return

    selected_text = ""
    try:
        if hasattr(selection, "getString"):
            selected_text = selection.getString()
        elif hasattr(selection, "getCount"):
            parts = []
            for i in range(selection.getCount()):
                parts.append(selection.getByIndex(i).getString())
            selected_text = "\n".join(parts)
    except Exception:
        pass

    if not selected_text:
        _show_message("Claude Writer", "Select some text first, then right-click > Ask Claude.")
        return

    # Show input dialog asking what to do
    instruction = _input_dialog(
        "Ask Claude",
        "Selected: \"{}...\"\n\nWhat should Claude do with this text?".format(
            selected_text[:80]
        ),
    )

    if not instruction:
        return

    # Build the message for Claude
    message = 'Regarding this text: "{}"\n\nInstruction: {}'.format(
        selected_text[:500], instruction
    )

    # Send to tmux Claude session
    try:
        # Escape single quotes for tmux
        safe_message = message.replace("'", "'\\''").replace("\n", " ")
        subprocess.run(
            ["tmux", "send-keys", "-t", "claude-writer", safe_message, "Enter"],
            capture_output=True, timeout=5
        )
        _log("Sent to Claude: {}".format(instruction[:100]))
    except Exception as e:
        _show_message(
            "Claude Writer",
            "Could not send to Claude. Is the Claude Writer terminal open?\n\n"
            "Click the Claude toolbar button first, then try again.",
        )


def _input_dialog(title, message):
    """Show an input dialog and return the user's text."""
    try:
        ctx = XSCRIPTCONTEXT.getComponentContext()
        smgr = ctx.ServiceManager
        toolkit = smgr.createInstanceWithContext("com.sun.star.awt.Toolkit", ctx)
        parent = XSCRIPTCONTEXT.getDesktop().getCurrentFrame().getContainerWindow()

        # Use InputBox via dispatch or BasicIDE
        # Simpler: use the built-in input line dialog
        from com.sun.star.awt.MessageBoxType import QUERYBOX
        box = toolkit.createMessageBox(parent, QUERYBOX, 3, title, message)
        # QUERYBOX with buttons=3 gives OK/Cancel
        # But it doesn't have an input field...

        # Use XInputBoxDialog if available, otherwise fall back to simple approach
        pass
    except Exception:
        pass

    # Fall back to a simple approach using BasicIDE InputBox via dispatch
    try:
        ctx = XSCRIPTCONTEXT.getComponentContext()
        smgr = ctx.ServiceManager
        script_provider = smgr.createInstanceWithContext(
            "com.sun.star.script.provider.ScriptProviderForBasic", ctx
        )
    except Exception:
        pass

    # Simplest fallback: write to a temp file and use zenity/kdialog
    try:
        result = subprocess.run(
            ["zenity", "--entry", "--title=" + title, "--text=" + message, "--width=500"],
            capture_output=True, text=True, timeout=120
        )
        if result.returncode == 0:
            return result.stdout.strip()
    except Exception:
        pass

    try:
        result = subprocess.run(
            ["kdialog", "--inputbox", message, "--title", title],
            capture_output=True, text=True, timeout=120
        )
        if result.returncode == 0:
            return result.stdout.strip()
    except Exception:
        pass

    return None


g_exportedScripts = (StartClaude, AskClaude)
