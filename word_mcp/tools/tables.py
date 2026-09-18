"""Live editing tools for Microsoft Word via COM automation.

These tools operate on documents that are currently open in Word,
providing real-time editing capabilities with optional tracked changes.
"""

import json
import os
import re
import sys

from word_mcp.defaults import DEFAULT_AUTHOR


# Word COM constants
WD_STORY = 6

# Word COM InsertBefore/InsertAfter limit (~32K chars).
# We use 30000 as safe margin below 2^15-1 = 32767.
_INSERT_CHUNK_SIZE = 30000


async def word_live_add_table(
    filename: str = None,
    rows: int = 2,
    cols: int = 2,
    position: str = "end",
    data: list = None,
    style: str = "Table Grid",
    autofit: str = "window",
    track_changes: bool = False,
) -> str:
    """Add a table to an open Word document.

    Args:
        filename: Document name or path.
        rows: Number of rows.
        cols: Number of columns.
        position: "start", "end", or character offset.
        data: Optional 2D list of cell data.
        style: Table style name. Default "Table Grid" (bordered).
            Use None or "" for no style.
        autofit: "window" (fit page width, default), "content" (fit cell content),
            "fixed" (fixed widths), or None for legacy behavior (no autofit).
        track_changes: Track as revision.

    Returns:
        JSON with result info.
    """
    if sys.platform != "win32":
        return json.dumps({"error": "Live editing is only available on Windows"})

    try:
        from word_mcp.word_com import get_word_app, find_document, undo_record

        app = get_word_app()
        doc = find_document(app, filename)

        if position == "start":
            rng = doc.Range(0, 0)
        elif position == "end":
            end_pos = doc.Content.End - 1
            rng = doc.Range(end_pos, end_pos)
        else:
            try:
                offset = int(position)
            except ValueError:
                return json.dumps({"error": f"Invalid position: {position}"})

            # Reject offsets that would land the new table inside an
            # existing table's range — Word would silently merge the
            # new structure into the old, breaking both.
            for t in doc.Tables:
                try:
                    ts, te = t.Range.Start, t.Range.End
                except Exception:
                    continue
                if ts <= offset <= te:
                    return json.dumps({
                        "error": (
                            f"position offset {offset} falls within an existing "
                            f"table at range [{ts}, {te}]. Choose an offset "
                            f"outside any table, or use position='end'/'start'."
                        )
                    })

            # Reject offsets immediately after an orphan cell separator
            # (residue from a prior Table.Delete with scrub disabled);
            # adding a table at such a point fuses it with the residue.
            if offset > 0:
                try:
                    probe = doc.Range(offset - 1, offset).Text or ""
                except Exception:
                    probe = ""
                if probe == "\x07":
                    return json.dumps({
                        "error": (
                            f"position offset {offset} sits immediately after "
                            f"an orphan cell separator (\\x07). Run "
                            f"word_live_modify_table operation='delete_table' "
                            f"with scrub_orphans=True (the default) on the "
                            f"prior table, or use word_live_diagnose_layout "
                            f"to locate and clean separators."
                        )
                    })

            rng = doc.Range(offset, offset)

        with undo_record(app, "MCP: Add Table"):
            prev_tracking = doc.TrackRevisions
            prev_author = app.UserName
            if track_changes:
                doc.TrackRevisions = True
                app.UserName = DEFAULT_AUTHOR

            try:
                # AutoFit behavior constants
                AUTOFIT_MAP = {
                    "window": (1, 2),   # wdWord9TableBehavior, wdAutoFitWindow
                    "content": (1, 1),  # wdWord9TableBehavior, wdAutoFitContent
                    "fixed": (0, 0),    # wdWord8TableBehavior, wdAutoFitFixed
                }

                if autofit and autofit.lower() in AUTOFIT_MAP:
                    default_behavior, autofit_behavior = AUTOFIT_MAP[autofit.lower()]
                    table = doc.Tables.Add(rng, rows, cols, default_behavior, autofit_behavior)
                else:
                    table = doc.Tables.Add(rng, rows, cols)

                # Apply table style
                if style:
                    try:
                        table.Style = doc.Styles(style)
                    except Exception:
                        pass  # Style not found; proceed without

                if data:
                    for r_idx, row_data in enumerate(data):
                        if r_idx >= rows:
                            break
                        for c_idx, cell_val in enumerate(row_data):
                            if c_idx >= cols:
                                break
                            table.Cell(r_idx + 1, c_idx + 1).Range.Text = str(cell_val)
            finally:
                if track_changes:
                    doc.TrackRevisions = prev_tracking
                    app.UserName = prev_author

        return json.dumps(
            {
                "success": True,
                "document": doc.Name,
                "rows": rows,
                "cols": cols,
                "position": position,
                "style": style or None,
                "autofit": autofit or None,
                "tracked": track_changes,
            }
        )

    except Exception as e:
        return json.dumps({"error": str(e)})


async def word_live_format_table(
    filename: str = None,
    table_index: int = -1,
    border_style: str = None,
    cell_bold: list = None,
    cell_alignment: list = None,
    column_widths: list = None,
    table_alignment: str = None,
    cell_shading: list = None,
    autofit: str = None,
) -> str:
    """Format a table in an open Word document via COM.

    Supports border removal, cell formatting, column sizing, and table alignment.
    Use table_index=-1 for the last table, 1 for the first, etc.

    Args:
        filename: Document name or path (None = active document).
        table_index: 1-based table index, or -1 for the last table.
        border_style: Border style for all edges: "none", "single", "double", "dotted",
            "dashed", "thick". "none" removes all borders.
        cell_bold: List of [row, col, bold] entries (1-indexed) to set bold on cell text.
            Example: [[1, 1, true], [1, 2, true]] bolds row 1 cells.
        cell_alignment: List of [row, col, alignment] entries. alignment: "left", "center",
            "right", "justify". Row 0 = all rows, Col 0 = all cols.
        column_widths: List of column widths in points (1-indexed order).
            Example: [200, 200] sets col 1 to 200pt, col 2 to 200pt.
        table_alignment: Table alignment on page: "left", "center", "right".
        cell_shading: List of [row, col, color_hex] entries. color_hex as "#RRGGBB".
            Row 0 = all rows. Example: [[1, 0, "#DDDDDD"]] shades entire row 1.
        autofit: "window" (fit to page width), "content" (fit to cell content),
            "fixed" (fixed column widths).

    Returns:
        JSON with result info.
    """
    if sys.platform != "win32":
        return json.dumps({"error": "Live editing is only available on Windows"})

    try:
        from word_mcp.word_com import get_word_app, find_document, undo_record

        app = get_word_app()
        doc = find_document(app, filename)

        if doc.Tables.Count == 0:
            return json.dumps({"error": "Document has no tables"})

        idx = table_index if table_index > 0 else doc.Tables.Count
        if idx < 1 or idx > doc.Tables.Count:
            return json.dumps({"error": f"Table index {table_index} out of range (1-{doc.Tables.Count})"})

        tbl = doc.Tables(idx)
        actions = []

        # Border style constants
        BORDER_STYLES = {
            "none": 0,     # wdLineStyleNone
            "single": 1,   # wdLineStyleSingle
            "double": 7,   # wdLineStyleDouble
            "dotted": 3,   # wdLineStyleDot
            "dashed": 2,   # wdLineStyleDash
            "thick": 6,    # wdLineStyleThickThinSmallGap (thick)
        }

        BORDER_IDS = [-1, -2, -3, -4, -5, -6, -7, -8]  # top, left, bottom, right, horiz, vert, etc.

        with undo_record(app, "MCP: Format Table"):
            # --- Borders ---
            if border_style is not None:
                style_val = BORDER_STYLES.get(border_style.lower())
                if style_val is None:
                    return json.dumps({"error": f"Unknown border_style: {border_style}. Use: {list(BORDER_STYLES.keys())}"})
                for bid in BORDER_IDS:
                    try:
                        tbl.Borders(bid).LineStyle = style_val
                    except Exception:
                        pass
                actions.append(f"borders={border_style}")

            # --- Autofit ---
            if autofit is not None:
                AUTOFIT = {"window": 2, "content": 1, "fixed": 0}  # wdAutoFitWindow=2, wdAutoFitContent=1, wdAutoFitFixed=0
                af_val = AUTOFIT.get(autofit.lower())
                if af_val is not None:
                    tbl.AutoFitBehavior(af_val)
                    actions.append(f"autofit={autofit}")

            # --- Table alignment ---
            if table_alignment is not None:
                ALIGN = {"left": 0, "center": 1, "right": 2}
                al_val = ALIGN.get(table_alignment.lower())
                if al_val is not None:
                    tbl.Rows.Alignment = al_val
                    actions.append(f"table_alignment={table_alignment}")

            # --- Column widths ---
            if column_widths is not None:
                for ci, width in enumerate(column_widths):
                    if ci < tbl.Columns.Count:
                        tbl.Columns(ci + 1).Width = float(width)
                actions.append(f"column_widths={column_widths}")

            # --- Cell bold ---
            if cell_bold is not None:
                for entry in cell_bold:
                    r, c, bold_val = int(entry[0]), int(entry[1]), bool(entry[2])
                    if 1 <= r <= tbl.Rows.Count and 1 <= c <= tbl.Columns.Count:
                        tbl.Cell(r, c).Range.Font.Bold = bold_val
                actions.append(f"cell_bold={len(cell_bold)} cells")

            # --- Cell alignment ---
            if cell_alignment is not None:
                PARA_ALIGN = {"left": 0, "center": 1, "right": 2, "justify": 3}
                for entry in cell_alignment:
                    r, c, align = int(entry[0]), int(entry[1]), str(entry[2]).lower()
                    al = PARA_ALIGN.get(align, 0)
                    if r == 0 and c == 0:
                        # All cells
                        for ri in range(1, tbl.Rows.Count + 1):
                            for ci in range(1, tbl.Columns.Count + 1):
                                tbl.Cell(ri, ci).Range.ParagraphFormat.Alignment = al
                    elif r == 0:
                        # Entire column
                        for ri in range(1, tbl.Rows.Count + 1):
                            tbl.Cell(ri, c).Range.ParagraphFormat.Alignment = al
                    elif c == 0:
                        # Entire row
                        for ci in range(1, tbl.Columns.Count + 1):
                            tbl.Cell(r, ci).Range.ParagraphFormat.Alignment = al
                    else:
                        if 1 <= r <= tbl.Rows.Count and 1 <= c <= tbl.Columns.Count:
                            tbl.Cell(r, c).Range.ParagraphFormat.Alignment = al
                actions.append(f"cell_alignment={len(cell_alignment)} entries")

            # --- Cell shading ---
            if cell_shading is not None:
                for entry in cell_shading:
                    r, c, color_hex = int(entry[0]), int(entry[1]), str(entry[2])
                    # Convert #RRGGBB to Word BGR integer
                    color_hex = color_hex.lstrip("#")
                    rr, gg, bb = int(color_hex[0:2], 16), int(color_hex[2:4], 16), int(color_hex[4:6], 16)
                    bgr = bb * 65536 + gg * 256 + rr

                    def shade_cell(row_i, col_i):
                        tbl.Cell(row_i, col_i).Shading.BackgroundPatternColor = bgr

                    if r == 0 and c == 0:
                        for ri in range(1, tbl.Rows.Count + 1):
                            for ci in range(1, tbl.Columns.Count + 1):
                                shade_cell(ri, ci)
                    elif r == 0:
                        for ri in range(1, tbl.Rows.Count + 1):
                            shade_cell(ri, c)
                    elif c == 0:
                        for ci in range(1, tbl.Columns.Count + 1):
                            shade_cell(r, ci)
                    else:
                        if 1 <= r <= tbl.Rows.Count and 1 <= c <= tbl.Columns.Count:
                            shade_cell(r, c)
                actions.append(f"cell_shading={len(cell_shading)} entries")

        return json.dumps(
            {
                "success": True,
                "document": doc.Name,
                "table_index": idx,
                "rows": tbl.Rows.Count,
                "cols": tbl.Columns.Count,
                "actions": actions,
            }
        )

    except Exception as e:
        return json.dumps({"error": str(e)})


async def word_live_modify_table(
    filename: str = None,
    table_index: int = 1,
    operation: str = "get_info",
    row: int = None,
    col: int = None,
    text: str = None,
    before_row: int = None,
    before_col: int = None,
    header: str = None,
    cells: list = None,
    start_row: int = None,
    start_col: int = None,
    end_row: int = None,
    end_col: int = None,
    autofit_mode: str = "content",
    accept_revisions: bool = False,
    track_changes: bool = False,
    scrub_orphans: bool = True,
) -> str:
    """[Windows only] Modify a table in an open Word document.

    Operations: get_info, set_cell, set_row, set_range, add_column, delete_column,
    add_row, delete_row, merge_cells, autofit, delete_table.
    All row/col indices are 1-based (Word COM standard).

    Args:
        filename: Document name or path (None = active document).
        table_index: 1-based table index (default 1).
        operation: One of: get_info, set_cell, set_row, set_range, add_column,
            delete_column, add_row, delete_row, merge_cells, autofit, delete_table.
        row: Row index for set_cell, set_row, delete_row.
        col: Column index for set_cell, delete_column.
        text: Text for set_cell.
        before_row: Insert row before this index (add_row). None = append at end.
        before_col: Insert column before this index (add_column). None = append at end.
        header: Header text for new column (add_column, placed in row 1).
        cells: List of cell values for set_row (1D) or set_range (2D). None values skip that cell.
            Also used for new row/column values (add_row, add_column).
        start_row: Start row for merge_cells or set_range (default 1).
        start_col: Start column for merge_cells or set_range (default 1).
        end_row: End row for merge_cells.
        end_col: End column for merge_cells.
        autofit_mode: 'content', 'window', or 'fixed' (autofit operation).
        accept_revisions: For set_cell/set_row/set_range — accept tracked changes before writing
            (prevents layered text from old revisions persisting underneath new content).
        track_changes: Track modifications as revisions.
        scrub_orphans: For delete_table — scan the deletion site for orphan
            cell-separator (\\x07) bytes and remove them. Default True.

    Returns:
        JSON with operation result.
    """
    if sys.platform != "win32":
        return json.dumps({"error": "Live editing is only available on Windows"})

    try:
        from word_mcp.word_com import get_word_app, find_document, undo_record
        from word_mcp import table_com

        app = get_word_app()
        doc = find_document(app, filename)

        # Per-call validation: re-read Tables.Count fresh in case a prior
        # MCP call (especially delete_table) reduced or zeroed the count.
        try:
            table_count = doc.Tables.Count
        except Exception as e:
            return json.dumps({
                "error": f"could not enumerate document tables: {e}"
            })

        if table_count == 0:
            return json.dumps({"error": "Document has no tables"})

        if not (1 <= table_index <= table_count):
            return json.dumps({
                "error": (
                    f"table_index {table_index} out of range. Document has "
                    f"{table_count} table(s) (valid range: 1..{table_count}). "
                    f"If a prior delete_table reduced the count, call "
                    f"word_live_get_info to refresh."
                )
            })

        table = doc.Tables(table_index)
        op = operation.lower()

        # get_info is read-only — no undo record needed
        if op == "get_info":
            result = table_com.get_info(table)
            result["document"] = doc.Name
            result["table_index"] = table_index
            return json.dumps(result, ensure_ascii=False)

        # All other operations are destructive
        with undo_record(app, "MCP: Modify Table"):
            prev_tracking = doc.TrackRevisions
            prev_author = app.UserName
            if track_changes:
                doc.TrackRevisions = True
                app.UserName = DEFAULT_AUTHOR

            try:
                if op == "set_cell":
                    if row is None or col is None or text is None:
                        return json.dumps({"error": "set_cell requires row, col, and text"})
                    result = table_com.set_cell(table, row, col, text, accept_revisions=accept_revisions)

                elif op == "set_row":
                    if row is None or not cells:
                        return json.dumps({"error": "set_row requires row and cells (list of values)"})
                    result = table_com.set_row(table, row, cells, accept_revisions=accept_revisions)

                elif op == "set_range":
                    if not cells:
                        return json.dumps({"error": "set_range requires cells (2D list of values)"})
                    result = table_com.set_range(
                        table, cells,
                        start_row=start_row or 1,
                        start_col=start_col or 1,
                        accept_revisions=accept_revisions,
                    )

                elif op == "add_column":
                    result = table_com.add_column(table, before_col, header, cells)

                elif op == "delete_column":
                    if col is None:
                        return json.dumps({"error": "delete_column requires col"})
                    result = table_com.delete_column(table, col)

                elif op == "add_row":
                    result = table_com.add_row(table, before_row, cells)

                elif op == "delete_row":
                    if row is None:
                        return json.dumps({"error": "delete_row requires row"})
                    result = table_com.delete_row(table, row)

                elif op == "merge_cells":
                    if not all(v is not None for v in [start_row, start_col, end_row, end_col]):
                        return json.dumps({"error": "merge_cells requires start_row, start_col, end_row, end_col"})
                    result = table_com.merge_cells(table, start_row, start_col, end_row, end_col)

                elif op == "autofit":
                    result = table_com.autofit(table, autofit_mode)

                elif op == "delete_table":
                    result = table_com.delete_table(table, scrub_orphans=scrub_orphans)

                else:
                    return json.dumps({
                        "error": f"Unknown operation '{op}'. Use: get_info, set_cell, set_row, set_range, "
                        "add_column, delete_column, add_row, delete_row, merge_cells, autofit, delete_table"
                    })
            finally:
                if track_changes:
                    doc.TrackRevisions = prev_tracking
                    app.UserName = prev_author

        result["success"] = True
        result["document"] = doc.Name
        result["table_index"] = table_index
        result["operation"] = op
        result["tracked"] = track_changes
        return json.dumps(result, ensure_ascii=False)

    except Exception as e:
        return json.dumps({"error": str(e)})
