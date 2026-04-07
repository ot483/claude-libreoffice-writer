"""Table tools - read and inspect tables in the document."""

from tools import tool
from uno_connection import UnoConnection


@tool(
    "list_tables",
    "List all tables in the document with their names, sizes, and locations.",
    {},
)
def list_tables(args):
    doc = UnoConnection.get().get_document()
    tables = doc.getTextTables()

    if tables.getCount() == 0:
        return {"tables": [], "message": "No tables found in the document."}

    result = []
    for i in range(tables.getCount()):
        table = tables.getByIndex(i)
        result.append({
            "index": i,
            "name": table.getName(),
            "rows": table.getRows().getCount(),
            "columns": table.getColumns().getCount(),
        })

    return {"tables": result, "count": len(result)}


@tool(
    "read_table",
    "Read the contents of a specific table by index or name. "
    "Returns all cell values as a 2D array.",
    {
        "table_index": {
            "type": "integer",
            "description": "Table index (0-based). Use list_tables to find indices.",
        },
        "table_name": {
            "type": "string",
            "description": "Table name (alternative to index). E.g. 'Table1'.",
        },
    },
)
def read_table(args):
    table_idx = args.get("table_index", None)
    table_name = args.get("table_name", None)

    if table_idx is None and table_name is None:
        return {"error": "Provide either table_index or table_name"}

    doc = UnoConnection.get().get_document()
    tables = doc.getTextTables()

    table = None
    if table_name:
        try:
            table = tables.getByName(table_name)
        except Exception:
            return {"error": "Table '{}' not found".format(table_name)}
    elif table_idx is not None:
        if table_idx < 0 or table_idx >= tables.getCount():
            return {"error": "Table index {} out of range (0-{})".format(
                table_idx, tables.getCount() - 1
            )}
        table = tables.getByIndex(table_idx)

    rows = table.getRows().getCount()
    cols = table.getColumns().getCount()

    data = []
    for r in range(rows):
        row_data = []
        for c in range(cols):
            cell_name = _cell_name(r, c)
            try:
                cell = table.getCellByName(cell_name)
                if cell:
                    row_data.append(cell.getString())
                else:
                    row_data.append("")
            except Exception:
                row_data.append("")
        data.append(row_data)

    return {
        "name": table.getName(),
        "rows": rows,
        "columns": cols,
        "data": data,
    }


def _cell_name(row, col):
    """Convert 0-based row/col to table cell name like A1, B2, AA3."""
    letters = ""
    c = col
    while True:
        letters = chr(ord("A") + c % 26) + letters
        c = c // 26 - 1
        if c < 0:
            break
    return "{}{}".format(letters, row + 1)
