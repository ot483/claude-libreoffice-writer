#!/usr/bin/env /usr/bin/python3
"""LibreOffice Writer MCP Server - stdio transport (JSON-RPC 2.0)."""

import sys
import json
import logging

# Set up logging to stderr (stdout is the MCP transport)
logging.basicConfig(
    stream=sys.stderr,
    level=logging.DEBUG,
    format="[LibreOffice MCP] %(levelname)s: %(message)s",
)
log = logging.getLogger("libreoffice-mcp")

# Import tool modules - each registers its tools via @tool decorator
from tools import list_tools, call_tool
from tools import document_read   # noqa: F401
from tools import document_edit   # noqa: F401
from tools import document_style  # noqa: F401
from tools import comments        # noqa: F401
from tools import track_changes   # noqa: F401
from tools import report           # noqa: F401
from tools import document_nav    # noqa: F401
from tools import tables          # noqa: F401


SERVER_INFO = {
    "name": "libreoffice-writer",
    "version": "0.1.0",
}

CAPABILITIES = {
    "tools": {},
}


def read_message():
    """Read a JSON-RPC message from stdin.

    Supports both newline-delimited JSON (Claude Code) and
    Content-Length framed messages (standard MCP).
    """
    while True:
        line = sys.stdin.buffer.readline()
        if not line:
            return None  # EOF
        line_str = line.decode("utf-8").strip()
        if not line_str:
            continue  # skip blank lines

        # Try parsing as JSON directly (newline-delimited mode)
        if line_str.startswith("{"):
            try:
                return json.loads(line_str)
            except json.JSONDecodeError:
                log.warning("Failed to parse JSON line: %s", line_str[:200])
                continue

        # Content-Length framed mode
        if line_str.lower().startswith("content-length:"):
            content_length = int(line_str.split(":", 1)[1].strip())
            # Read until blank line (end of headers)
            while True:
                header_line = sys.stdin.buffer.readline()
                if not header_line or header_line.strip() == b"":
                    break
            body = sys.stdin.buffer.read(content_length)
            if not body:
                return None
            return json.loads(body.decode("utf-8"))


def write_message(msg):
    """Write a JSON-RPC message as a newline-delimited JSON line to stdout."""
    body = json.dumps(msg)
    sys.stdout.write(body + "\n")
    sys.stdout.flush()


def handle_initialize(params):
    """Handle the initialize request."""
    # Echo back the client's protocol version for compatibility
    client_version = params.get("protocolVersion", "2024-11-05")
    return {
        "protocolVersion": client_version,
        "capabilities": CAPABILITIES,
        "serverInfo": SERVER_INFO,
    }


def handle_tools_list(params):
    """Handle tools/list request."""
    return {"tools": list_tools()}


def handle_tools_call(params):
    """Handle tools/call request."""
    name = params.get("name", "")
    arguments = params.get("arguments", {})
    log.info("Calling tool: %s with args: %s", name, json.dumps(arguments)[:200])
    return call_tool(name, arguments)


def handle_request(msg):
    """Route a JSON-RPC request to the appropriate handler."""
    method = msg.get("method", "")
    params = msg.get("params", {})
    msg_id = msg.get("id")

    handlers = {
        "initialize": handle_initialize,
        "tools/list": handle_tools_list,
        "tools/call": handle_tools_call,
    }

    # Notifications (no id) - just acknowledge
    if msg_id is None:
        log.debug("Notification: %s", method)
        return None

    handler = handlers.get(method)
    if handler is None:
        log.warning("Unknown method: %s", method)
        return {
            "jsonrpc": "2.0",
            "id": msg_id,
            "error": {
                "code": -32601,
                "message": "Method not found: {}".format(method),
            },
        }

    try:
        result = handler(params)
        return {
            "jsonrpc": "2.0",
            "id": msg_id,
            "result": result,
        }
    except Exception as e:
        log.error("Error handling %s: %s", method, str(e), exc_info=True)
        return {
            "jsonrpc": "2.0",
            "id": msg_id,
            "error": {
                "code": -32603,
                "message": str(e),
            },
        }


def main():
    log.info("LibreOffice Writer MCP server starting...")
    while True:
        msg = read_message()
        if msg is None:
            log.info("EOF on stdin - shutting down.")
            break

        log.debug("Received: %s", json.dumps(msg)[:300])
        response = handle_request(msg)

        if response is not None:
            log.debug("Sending: %s", json.dumps(response)[:300])
            write_message(response)


if __name__ == "__main__":
    main()
