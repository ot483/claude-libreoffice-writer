"""Tool registry for the LibreOffice MCP server."""

import json

# Global registry: name -> {description, schema, handler}
_tools = {}


def tool(name, description, params_schema=None):
    """Decorator to register an MCP tool.

    Usage:
        @tool("read_document", "Read document text", {"start_para": {...}, ...})
        def read_document(args):
            return {"text": "..."}
    """
    if params_schema is None:
        params_schema = {}

    def decorator(fn):
        _tools[name] = {
            "description": description,
            "schema": params_schema,
            "handler": fn,
        }
        return fn

    return decorator


def list_tools():
    """Return MCP-formatted tool list."""
    result = []
    for name, info in _tools.items():
        tool_def = {
            "name": name,
            "description": info["description"],
            "inputSchema": {
                "type": "object",
                "properties": info["schema"],
            },
        }
        result.append(tool_def)
    return result


def call_tool(name, arguments):
    """Call a registered tool by name. Returns MCP content array."""
    if name not in _tools:
        return {
            "isError": True,
            "content": [{"type": "text", "text": "Unknown tool: {}".format(name)}],
        }
    try:
        result = _tools[name]["handler"](arguments or {})
        if isinstance(result, dict) and "content" in result:
            return result
        # Auto-wrap plain dicts/lists as JSON text
        text = json.dumps(result, indent=2, default=str)
        return {"content": [{"type": "text", "text": text}]}
    except Exception as e:
        return {
            "isError": True,
            "content": [{"type": "text", "text": "Error: {}".format(str(e))}],
        }
