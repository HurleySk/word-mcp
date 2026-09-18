"""Live editing tools for Microsoft Word via COM automation.

These tools operate on documents that are currently open in Word,
providing real-time editing capabilities with optional tracked changes.
"""

import json
from word_mcp.com_runtime import error_json
from word_mcp.live_tool import live_tool
import os
import re
import sys

from word_mcp.defaults import DEFAULT_AUTHOR


# Word COM constants
WD_STORY = 6

# Word COM InsertBefore/InsertAfter limit (~32K chars).
# We use 30000 as safe margin below 2^15-1 = 32767.
_INSERT_CHUNK_SIZE = 30000


@live_tool(mutates=True)
def word_live_insert_text(
    filename: str = None,
    text: str = "",
    position: str = "end",
    bookmark: str = None,
    track_changes: bool = False,
) -> str:
    """Insert text into an open Word document.

    Automatically chunks large text (>30K chars) to avoid Word COM limits.

    Args:
        filename: Document name or path (None = active document).
        text: Text to insert (no length limit — auto-chunked if needed).
        position: "start", "end", "cursor", or character offset as string.
        bookmark: Insert after a named bookmark (overrides position).
        track_changes: Track the insertion as a revision.

    Returns:
        JSON with result info.
    """
    if sys.platform != "win32":
        return json.dumps({"error": "Live editing is only available on Windows"})

    try:
        from word_mcp.com_runtime import get_word_app, find_document, undo_record

        app = get_word_app()
        doc = find_document(app, filename)

        # Convert literal escape sequences to actual characters.
        # MCP/JSON sends backslash-r as 2 chars; Word COM needs chr(13) for paragraph marks.
        text = text.replace("\\r\\n", "\r").replace("\\r", "\r").replace("\\n", "\r")

        # Reject control bytes (notably \x07 cell separator) — inserting
        # these outside a real table creates invalid document state that
        # subsequent Find/Replace and table operations cannot recover from.
        from word_mcp.text_safety import reject_control_chars
        try:
            reject_control_chars("text", text)
        except ValueError as e:
            return json.dumps({"error": str(e)})

        with undo_record(app, "MCP: Insert Text"):
            prev_tracking = doc.TrackRevisions
            prev_author = app.UserName
            if track_changes:
                doc.TrackRevisions = True
                app.UserName = DEFAULT_AUTHOR

            try:
                chunks = [text[i:i + _INSERT_CHUNK_SIZE]
                          for i in range(0, max(len(text), 1), _INSERT_CHUNK_SIZE)]

                if bookmark:
                    if not doc.Bookmarks.Exists(bookmark):
                        return json.dumps({"error": f"Bookmark '{bookmark}' not found"})
                    rng = doc.Bookmarks(bookmark).Range
                    for chunk in chunks:
                        rng.InsertAfter(chunk)
                        rng.Collapse(0)  # wdCollapseEnd
                elif position == "start":
                    # InsertBefore: reverse order so first chunk ends up first
                    for chunk in reversed(chunks):
                        doc.Range(0, 0).InsertBefore(chunk)
                elif position == "end":
                    for chunk in chunks:
                        end_pos = doc.Content.End - 1
                        rng = doc.Range(end_pos, end_pos)
                        rng.InsertAfter(chunk)
                elif position == "cursor":
                    for chunk in chunks:
                        app.Selection.TypeText(chunk)
                else:
                    try:
                        offset = int(position)
                    except ValueError:
                        return json.dumps(
                            {
                                "error": f"Invalid position: {position}. "
                                "Use 'start', 'end', 'cursor', or a character offset."
                            }
                        )
                    # InsertBefore at offset: reverse order so first chunk ends up at offset
                    for chunk in reversed(chunks):
                        doc.Range(offset, offset).InsertBefore(chunk)
            finally:
                if track_changes:
                    doc.TrackRevisions = prev_tracking
                    app.UserName = prev_author

        result = {
            "success": True,
            "document": doc.Name,
            "text_length": len(text),
            "position": position,
            "tracked": track_changes,
        }
        if len(chunks) > 1:
            result["chunks_used"] = len(chunks)
        return json.dumps(result)

    except Exception as e:
        return error_json(e)


@live_tool(mutates=True)
def word_live_format_text(
    filename: str = None,
    start: int = None,
    end: int = None,
    start_paragraph: int = None,
    end_paragraph: int = None,
    bold: bool = None,
    italic: bool = None,
    underline: bool = None,
    strikethrough: bool = None,
    font_name: str = None,
    font_size: float = None,
    font_color: str = None,
    highlight_color: int = None,
    style_name: str = None,
    paragraph_alignment: str = None,
    page_break_before: bool = None,
    preserve_direct_formatting: bool = False,
    track_changes: bool = False,
) -> str:
    """[Windows only] Format text in an open Word document: font, color, highlight, style, alignment, page breaks.
    Use this tool for any visual/formatting change that does NOT alter the text content itself.

    Two addressing modes (provide one):
    - start/end: Character positions (from word_live_find_text or word_live_get_page_text).
    - start_paragraph/end_paragraph: 1-indexed paragraph range (from word_live_get_text etc.).

    Args:
        filename: Document name or path (None = active document).
        start: Start character position.
        end: End character position.
        start_paragraph: First paragraph index (1-indexed). Alternative to start/end.
        end_paragraph: Last paragraph index (1-indexed, defaults to start_paragraph).
        bold: Set bold (True/False).
        italic: Set italic (True/False).
        underline: Set underline (True/False).
        strikethrough: Set strikethrough (True/False).
        font_name: Font family (e.g., "Arial", "Times New Roman").
        font_size: Font size in points (e.g., 12).
        font_color: Text color as "#RRGGBB" hex (e.g., "#FF0000" for red).
        highlight_color: Text highlight background color index.
            0 = remove highlight, 1 = black, 2 = blue, 3 = turquoise,
            4 = bright green, 5 = pink, 6 = red, 7 = yellow,
            8 = white, 9 = dark blue, 10 = teal, 11 = green,
            12 = violet, 13 = dark red, 14 = dark yellow, 15 = gray, 16 = light gray.
            Common: 7=yellow (add), 0=none (remove).
        style_name: Apply a named Word style (e.g., "Heading 1", "Normal").
        paragraph_alignment: Paragraph alignment — "left" (0), "center" (1), "right" (2), "justify" (3).
            Applies to ALL paragraphs in the selected range.
        page_break_before: Set or clear PageBreakBefore on paragraphs in range (True/False).
        preserve_direct_formatting: When True and style_name is set, saves font/size/bold/italic/
            alignment/spacing before applying the style and restores them after. Useful for changing
            a paragraph's style (e.g., Heading 5 → Normal) without losing its visual formatting.
        track_changes: Track formatting changes as revisions.

    Returns:
        JSON with result info.
    """
    if sys.platform != "win32":
        return json.dumps({"error": "Live editing is only available on Windows"})

    try:
        from word_mcp.com_runtime import get_word_app, find_document, undo_record

        app = get_word_app()
        doc = find_document(app, filename)

        # Resolve addressing mode
        if start_paragraph is not None:
            if end_paragraph is None:
                end_paragraph = start_paragraph
            total_paras = doc.Paragraphs.Count
            if start_paragraph < 1 or end_paragraph > total_paras:
                return json.dumps({
                    "error": f"Paragraph range {start_paragraph}-{end_paragraph} out of bounds (doc has {total_paras} paragraphs)"
                })
            p_start = doc.Paragraphs(start_paragraph).Range.Start
            p_end = doc.Paragraphs(end_paragraph).Range.End
            rng = doc.Range(p_start, p_end)
            range_label = f"para {start_paragraph}-{end_paragraph}"
        elif start is not None and end is not None:
            rng = doc.Range(start, end)
            range_label = f"{start}-{end}"
        else:
            return json.dumps(
                {"error": "Provide start/end character positions OR start_paragraph/end_paragraph"}
            )

        with undo_record(app, "MCP: Format Text"):
            prev_tracking = doc.TrackRevisions
            prev_author = app.UserName
            if track_changes:
                doc.TrackRevisions = True
                app.UserName = DEFAULT_AUTHOR

            try:
                # Save direct formatting before style change if requested
                saved_formats = []
                if preserve_direct_formatting and style_name is not None:
                    for para in rng.Paragraphs:
                        pr = para.Range
                        pf = para.Format
                        saved_formats.append({
                            "para": para,
                            "font_name": str(pr.Font.Name) if pr.Font.Name and pr.Font.Name != 9999999 else None,
                            "font_size": pr.Font.Size if pr.Font.Size and pr.Font.Size != 9999999 else None,
                            "bold": pr.Font.Bold if pr.Font.Bold != 9999999 else None,
                            "italic": pr.Font.Italic if pr.Font.Italic != 9999999 else None,
                            "strikethrough": pr.Font.StrikeThrough if pr.Font.StrikeThrough != 9999999 else None,
                            "alignment": pf.Alignment,
                            "space_before": pf.SpaceBefore,
                            "space_after": pf.SpaceAfter,
                            "line_spacing": pf.LineSpacing,
                            "line_spacing_rule": pf.LineSpacingRule,
                        })

                if bold is not None:
                    rng.Font.Bold = bold
                if italic is not None:
                    rng.Font.Italic = italic
                if underline is not None:
                    rng.Font.Underline = 1 if underline else 0
                if strikethrough is not None:
                    rng.Font.StrikeThrough = strikethrough
                if font_name is not None:
                    rng.Font.Name = font_name
                if font_size is not None:
                    rng.Font.Size = font_size
                if font_color is not None:
                    c = font_color.lstrip("#")
                    r, g, b = int(c[0:2], 16), int(c[2:4], 16), int(c[4:6], 16)
                    rng.Font.Color = r + (g << 8) + (b << 16)
                if highlight_color is not None:
                    rng.HighlightColorIndex = highlight_color
                if style_name is not None:
                    if preserve_direct_formatting:
                        # Apply style per-paragraph and restore formatting
                        for sf in saved_formats:
                            p = sf["para"]
                            p.Style = doc.Styles(style_name)
                            pr = p.Range
                            pf = p.Format
                            if sf["font_name"] is not None:
                                pr.Font.Name = sf["font_name"]
                            if sf["font_size"] is not None:
                                pr.Font.Size = sf["font_size"]
                            if sf["bold"] is not None:
                                pr.Font.Bold = sf["bold"]
                            if sf["italic"] is not None:
                                pr.Font.Italic = sf["italic"]
                            if sf["strikethrough"] is not None:
                                pr.Font.StrikeThrough = sf["strikethrough"]
                            pf.Alignment = sf["alignment"]
                            pf.SpaceBefore = sf["space_before"]
                            pf.SpaceAfter = sf["space_after"]
                            pf.LineSpacingRule = sf["line_spacing_rule"]
                            pf.LineSpacing = sf["line_spacing"]
                    else:
                        rng.Style = style_name
                if paragraph_alignment is not None:
                    align_map = {"left": 0, "center": 1, "right": 2, "justify": 3}
                    al = align_map.get(paragraph_alignment.lower())
                    if al is None:
                        return json.dumps({"error": f"Invalid alignment: {paragraph_alignment}. Use: left, center, right, justify"})
                    for para in rng.Paragraphs:
                        para.Format.Alignment = al
                if page_break_before is not None:
                    for para in rng.Paragraphs:
                        para.Format.PageBreakBefore = page_break_before
            finally:
                if track_changes:
                    doc.TrackRevisions = prev_tracking
                    app.UserName = prev_author

        preview = rng.Text
        if len(preview) > 50:
            preview = preview[:50] + "..."

        return json.dumps(
            {
                "success": True,
                "document": doc.Name,
                "range": range_label,
                "text_preview": preview,
                "tracked": track_changes,
            }
        )

    except Exception as e:
        return error_json(e)


@live_tool(mutates=True)
def word_live_apply_list(
    filename: str = None,
    start_paragraph: int = None,
    end_paragraph: int = None,
    list_type: str = "bullet",
    level: int = 0,
    remove: bool = False,
    continue_previous: bool = False,
    number_format: dict = None,
    number_style: dict = None,
    start_at: dict = None,
    level_map: dict = None,
    track_changes: bool = False,
) -> str:
    """[Windows only] Apply or remove bullet/numbered/multilevel list formatting on paragraphs.

    Args:
        filename: Document name or path (None = active document).
        start_paragraph: First paragraph to format (1-indexed, required).
        end_paragraph: Last paragraph to format (1-indexed, defaults to start_paragraph).
        list_type: "bullet", "number", or "multilevel" (outline numbered).
        level: Indentation level (0 = first level, 1 = second level, etc.).
            For multilevel, this sets the default list level for all paragraphs.
            Use level_map instead for per-paragraph level control.
        remove: If True, removes list formatting from the range.
        continue_previous: If True, continues numbering from a previous list above.
        number_format: (multilevel only) Dict mapping level (int) to format string.
            Example: {1: "4.%1.", 2: "(%2)", 3: "(%3)"} → "4.1.", "(a)", "(i)"
            Keys are 1-indexed levels. If not provided, defaults to {1: "%1.", 2: "%1.%2."}.
        number_style: (multilevel only) Dict mapping level (int) to numbering style string.
            Styles: "arabic" (1,2,3), "lowercase_letter" (a,b,c), "uppercase_letter" (A,B,C),
            "lowercase_roman" (i,ii,iii), "uppercase_roman" (I,II,III).
            Example: {1: "arabic", 2: "lowercase_letter", 3: "lowercase_roman"}
            If a string is given instead of dict, applies same style to all levels.
            Default: "arabic" for all levels.
        start_at: (multilevel only) Dict mapping level (int) to starting number.
            Example: {1: 5} → numbering starts at 5.
            If not provided, starts at 1.
        level_map: (multilevel only) Dict mapping paragraph index (int) to list level (int, 1-indexed).
            Example: {29: 2, 30: 2, 37: 3} → para 29 at level 2, para 37 at level 3.
            Paragraphs not in the map stay at level 1 (or the value of `level + 1`).
            Applied AFTER the list template, so the template covers the full range.
        track_changes: Track changes as revisions.

    Returns:
        JSON with result info.
    """
    if sys.platform != "win32":
        return json.dumps({"error": "Live editing is only available on Windows"})

    if start_paragraph is None:
        return json.dumps({"error": "start_paragraph is required (1-indexed)"})

    if end_paragraph is None:
        end_paragraph = start_paragraph

    try:
        from word_mcp.com_runtime import get_word_app, find_document, undo_record

        app = get_word_app()
        doc = find_document(app, filename)

        total_paras = doc.Paragraphs.Count
        if start_paragraph < 1 or end_paragraph > total_paras:
            return json.dumps({
                "error": f"Paragraph range {start_paragraph}-{end_paragraph} out of bounds (doc has {total_paras} paragraphs)"
            })

        with undo_record(app, "MCP: Apply List"):
            prev_tracking = doc.TrackRevisions
            prev_author = app.UserName
            if track_changes:
                doc.TrackRevisions = True
                app.UserName = DEFAULT_AUTHOR

            try:
                formatted = 0

                if remove:
                    for i in range(start_paragraph, end_paragraph + 1):
                        doc.Paragraphs(i).Range.ListFormat.RemoveNumbers()
                        formatted += 1
                elif list_type == "multilevel":
                    # Create custom multilevel list template (OutlineNumbered gallery)
                    lt = doc.ListTemplates.Add(OutlineNumbered=True)
                    # Normalize dict keys to int (JSON sends string keys)
                    nf = {int(k): v for k, v in (number_format or {1: "%1.", 2: "%1.%2."}).items()}
                    sa = {int(k): v for k, v in (start_at or {}).items()}
                    lm = {int(k): int(v) for k, v in (level_map or {}).items()}
                    # Map number_style string to wdListNumberStyle constant
                    style_map = {
                        "arabic": 0, "lowercase_letter": 4, "uppercase_letter": 3,
                        "lowercase_roman": 2, "uppercase_roman": 1,
                    }
                    # number_style can be a string (same for all) or dict (per-level)
                    if isinstance(number_style, dict):
                        ns_map = {int(k): style_map.get(v, 0) for k, v in number_style.items()}
                    elif isinstance(number_style, str):
                        ns_map = {lvl: style_map.get(number_style, 0) for lvl in nf}
                    else:
                        ns_map = {}
                    for lvl_num, fmt_str in nf.items():
                        lv = lt.ListLevels(int(lvl_num))
                        lv.NumberFormat = fmt_str
                        lv.NumberStyle = ns_map.get(int(lvl_num), 0)
                        lv.StartAt = sa.get(int(lvl_num), 1)
                        lv.Alignment = 0  # left
                        lv.NumberPosition = 0
                        lv.TextPosition = 28
                        lv.TabPosition = 28
                        # Do NOT set LinkedStyle — avoids Heading style side effects

                    # Apply template to the full range at once (not per-paragraph)
                    rng = doc.Range(
                        doc.Paragraphs(start_paragraph).Range.Start,
                        doc.Paragraphs(end_paragraph).Range.End,
                    )
                    rng.ListFormat.ApplyListTemplateWithLevel(
                        ListTemplate=lt,
                        ContinuePreviousList=continue_previous,
                        ApplyTo=2,  # wdListApplyToSelection
                        DefaultListBehavior=0,
                    )
                    formatted = end_paragraph - start_paragraph + 1

                    # Set per-paragraph levels from level_map
                    default_lvl = level + 1 if level > 0 else 1
                    for i in range(start_paragraph, end_paragraph + 1):
                        target_lvl = lm.get(i, default_lvl)
                        if target_lvl != 1:
                            doc.Paragraphs(i).Range.ListFormat.ListLevelNumber = target_lvl
                else:
                    # bullet or number (original logic)
                    gallery_map = {"bullet": 1, "number": 2}
                    gallery_idx = gallery_map.get(list_type, 1)
                    template = doc.Application.ListGalleries(gallery_idx).ListTemplates(1)
                    for i in range(start_paragraph, end_paragraph + 1):
                        para = doc.Paragraphs(i)
                        should_continue = (i > start_paragraph) or continue_previous
                        para.Range.ListFormat.ApplyListTemplateWithLevel(
                            ListTemplate=template,
                            ContinuePreviousList=should_continue,
                            DefaultListBehavior=1,
                        )
                        if level > 0:
                            para.Range.ListFormat.ListLevelNumber = level + 1
                        formatted += 1
            finally:
                if track_changes:
                    doc.TrackRevisions = prev_tracking
                    app.UserName = prev_author

        action = "removed" if remove else f"applied {list_type}"
        return json.dumps({
            "success": True,
            "document": doc.Name,
            "action": action,
            "paragraphs": f"{start_paragraph}-{end_paragraph}",
            "count": formatted,
            "level": level,
            "tracked": track_changes,
        })

    except Exception as e:
        return error_json(e)


@live_tool(mutates=True)
def word_live_setup_heading_numbering(
    filename: str = None,
    h1_paragraphs: list = None,
    h2_paragraphs: list = None,
    strip_manual_numbers: bool = True,
    h1_number_format: str = None,
    h2_number_format: str = None,
    font_name: str = None,
    h1_size: float = None,
    h2_size: float = None,
    bold: bool = None,
    alignment: str = None,
    font_color: str = None,
    h1_space_before: float = None,
    h1_space_after: float = None,
    h2_space_before: float = None,
    h2_space_after: float = None,
    line_spacing: float = None,
) -> str:
    """[Windows only] Set up auto-numbered headings with multilevel list (1. / 1.1).

    Creates a multilevel list template linked to Heading 1 and Heading 2 styles.
    Default formats: Level 1 = "%1." (produces "1."), Level 2 = "%1.%2" (produces "1.1").
    Custom formats supported — e.g., h1_number_format="MADDE %1 – " produces "MADDE 1 – ".

    Applies styles and numbering to the specified paragraphs, then optionally
    strips manual number prefixes. Recognizes two patterns:
    - Numeric: "1. ", "6.2. ", "10.3 " (regex: ^\\d+(\\.\\d+)*\\.?\\s+)
    - MADDE: "MADDE 6 – ", "MADDE 10 - " (regex: ^MADDE\\s+\\d+\\s*[–-]\\s*)

    If any style parameter is provided, Heading 1 and Heading 2 styles are
    customized before applying. If no style params are given, only numbering
    is applied (existing styles are preserved).

    Args:
        filename: Document name or path (None = active document).
        h1_paragraphs: List of 1-indexed paragraph numbers for Heading 1 (main sections).
        h2_paragraphs: List of 1-indexed paragraph numbers for Heading 2 (sub-sections).
        strip_manual_numbers: Remove leading number/MADDE prefix from headings (default True).
        h1_number_format: Custom Level 1 format (default "%1."). Use %1 for the number.
            Example: "MADDE %1 – " produces "MADDE 1 – ", "MADDE 2 – ", etc.
        h2_number_format: Custom Level 2 format (default "%1.%2"). Use %1 and %2.
            Example: "%1.%2." produces "1.1.", "1.2.", etc.
        font_name: Font family for both heading styles (e.g., "Cambria").
        h1_size: Font size in points for Heading 1 (e.g., 13).
        h2_size: Font size in points for Heading 2 (e.g., 11).
        bold: Set bold on both heading styles (True/False).
        alignment: Paragraph alignment — "left", "center", "right", "justify".
        font_color: Text color as "#RRGGBB" hex (e.g., "#0D0D0D").
        h1_space_before: Space before Heading 1 in points (e.g., 18).
        h1_space_after: Space after Heading 1 in points (e.g., 6).
        h2_space_before: Space before Heading 2 in points (e.g., 12).
        h2_space_after: Space after Heading 2 in points (e.g., 6).
        line_spacing: Line spacing in points for both heading styles (e.g., 13.8 for 1.15x).

    Returns:
        JSON with h1_applied, h2_applied, and stripped counts.
    """
    import re

    if sys.platform != "win32":
        return json.dumps({"error": "Live tools only on Windows"})

    if not h1_paragraphs and not h2_paragraphs:
        return json.dumps({"error": "Provide h1_paragraphs and/or h2_paragraphs"})

    try:
        from word_mcp.com_runtime import get_word_app, find_document, undo_record

        app = get_word_app()
        doc = find_document(app, filename)

        def _find_para_text(doc, text):
            """Find paragraph text in doc body, return range or None."""
            search = text[:60] if len(text) > 60 else text
            if not search:
                return None
            rng = doc.Content.Duplicate
            rng.Find.ClearFormatting()
            rng.Find.Execute(
                FindText=search, Forward=True,
                MatchCase=True, MatchWholeWord=False, Wrap=0,
            )
            return rng if rng.Find.Found else None

        with undo_record(app, "MCP: Setup Heading Numbering"):
            # --- Optionally customize heading styles ---
            has_style_params = any(p is not None for p in [
                font_name, h1_size, h2_size, bold, alignment, font_color,
                h1_space_before, h1_space_after, h2_space_before, h2_space_after,
                line_spacing,
            ])

            # These need to be defined unconditionally so the
            # _apply_direct_formatting closure (defined further down)
            # can reference them safely even when has_style_params=False.
            align_val = None
            color_int = None

            if has_style_params:
                align_map = {"left": 0, "center": 1, "right": 2, "justify": 3}
                align_val = align_map.get(alignment.lower()) if alignment else None

                if font_color:
                    c = font_color.lstrip("#")
                    r, g, b = int(c[0:2], 16), int(c[2:4], 16), int(c[4:6], 16)
                    color_int = r + (g << 8) + (b << 16)

                for style_id, size, sp_before, sp_after in [
                    (-2, h1_size, h1_space_before, h1_space_after),
                    (-3, h2_size, h2_space_before, h2_space_after),
                ]:
                    s = doc.Styles(style_id)
                    if font_name is not None:
                        s.Font.Name = font_name
                    if size is not None:
                        s.Font.Size = size
                    if bold is not None:
                        s.Font.Bold = bold
                        s.Font.Italic = False
                    if color_int is not None:
                        s.Font.Color = color_int
                    if align_val is not None:
                        s.ParagraphFormat.Alignment = align_val
                    if sp_before is not None:
                        s.ParagraphFormat.SpaceBefore = sp_before
                    if sp_after is not None:
                        s.ParagraphFormat.SpaceAfter = sp_after
                    if line_spacing is not None:
                        s.ParagraphFormat.LineSpacingRule = 5  # multiple
                        s.ParagraphFormat.LineSpacing = line_spacing
                    # H1 keeps with next (heading stays with first body para).
                    # H2 does NOT — sub-clauses are often full paragraphs;
                    # chaining keep_with_next across them breaks page layout.
                    s.ParagraphFormat.KeepWithNext = (style_id == -2)
                    s.ParagraphFormat.KeepTogether = False

            # --- Create multilevel list template ---
            lt = doc.ListTemplates.Add(OutlineNumbered=True)
            h1_fmt = h1_number_format or "%1."
            h2_fmt = h2_number_format or "%1.%2"

            # Level 1 linked to Heading 1
            lv1 = lt.ListLevels(1)
            lv1.NumberFormat = h1_fmt
            lv1.NumberStyle = 0  # wdListNumberStyleArabic
            lv1.StartAt = 1
            lv1.Alignment = 0  # left
            lv1.NumberPosition = 0
            if len(h1_fmt) > 5:
                # Long format (e.g., "MADDE %1 – ") — text follows number directly
                lv1.TextPosition = 0
                lv1.TabPosition = 0
            else:
                lv1.TextPosition = 28  # ~1cm indent for text after number
                lv1.TabPosition = 28
            lv1.LinkedStyle = "Heading 1"

            # Level 2 linked to Heading 2
            lv2 = lt.ListLevels(2)
            lv2.NumberFormat = h2_fmt
            lv2.NumberStyle = 0
            lv2.StartAt = 1
            lv2.Alignment = 0
            lv2.NumberPosition = 0
            if len(h2_fmt) > 5:
                lv2.TextPosition = 0
                lv2.TabPosition = 0
            else:
                lv2.TextPosition = 28
                lv2.TabPosition = 28
            lv2.LinkedStyle = "Heading 2"

            # --- Apply styles to paragraphs ---
            h1_applied = 0
            h2_applied = 0
            restyle_failures = []

            all_heading_paras = []
            for idx in (h1_paragraphs or []):
                all_heading_paras.append((idx, -2))  # wdStyleHeading1
            for idx in (h2_paragraphs or []):
                all_heading_paras.append((idx, -3))  # wdStyleHeading2
            all_heading_paras.sort(key=lambda x: x[0])

            target_style = {-2: doc.Styles(-2), -3: doc.Styles(-3)}
            target_label = {-2: "Heading 1", -3: "Heading 2"}

            def _apply_direct_formatting(rng, sid):
                """When has_style_params is True, mirror the heading-style
                customizations onto the range itself — this defeats any
                direct formatting inherited from a custom template style
                (e.g. "Font Style30/31") that would otherwise override the
                newly assigned Heading 1/2 style."""
                if not has_style_params:
                    return
                try:
                    if font_name is not None:
                        rng.Font.Name = font_name
                    size = h1_size if sid == -2 else h2_size
                    if size is not None:
                        rng.Font.Size = size
                    if bold is not None:
                        rng.Font.Bold = bold
                        rng.Font.Italic = False
                    if color_int is not None:
                        rng.Font.Color = color_int
                    if align_val is not None:
                        rng.ParagraphFormat.Alignment = align_val
                    sp_b = h1_space_before if sid == -2 else h2_space_before
                    sp_a = h1_space_after if sid == -2 else h2_space_after
                    if sp_b is not None:
                        rng.ParagraphFormat.SpaceBefore = sp_b
                    if sp_a is not None:
                        rng.ParagraphFormat.SpaceAfter = sp_a
                    if line_spacing is not None:
                        rng.ParagraphFormat.LineSpacingRule = 5  # multiple
                        rng.ParagraphFormat.LineSpacing = line_spacing
                except Exception:
                    pass  # best-effort

            for para_idx, style_id in all_heading_paras:
                if para_idx < 1 or para_idx > doc.Paragraphs.Count:
                    restyle_failures.append({
                        "index": para_idx, "error": "out of range"
                    })
                    continue
                para = doc.Paragraphs(para_idx)
                # Capture the style we are about to overwrite, for diagnostics.
                try:
                    old_style = para.Style.NameLocal
                except Exception:
                    old_style = None
                text = para.Range.Text.rstrip("\r\x07")
                range_len = para.Range.End - para.Range.Start
                text_len = len(text)
                inflated = (range_len > text_len + 5)

                applied_via = None
                if not inflated:
                    # Normal paragraph — direct style assignment works.
                    try:
                        para.Range.ListFormat.RemoveNumbers()
                    except Exception:
                        pass
                    try:
                        para.Range.Style = target_style[style_id]
                        _apply_direct_formatting(para.Range, style_id)
                        applied_via = "para.Range.Style"
                    except Exception as e:
                        restyle_failures.append({
                            "index": para_idx,
                            "old_style": old_style,
                            "error": f"para.Range.Style assign failed: {e}",
                        })
                        continue
                else:
                    # Inflated Range (comments/fields extend it beyond text).
                    # Use Find to locate text, then Expand to full paragraph
                    # so the paragraph mark gets the style.
                    found = _find_para_text(doc, text)
                    if not found:
                        restyle_failures.append({
                            "index": para_idx,
                            "old_style": old_style,
                            "error": "inflated range and Find could not locate text",
                        })
                        continue
                    try:
                        found.ListFormat.RemoveNumbers()
                    except Exception:
                        pass
                    try:
                        # Expand found range to full paragraph (includes \r mark)
                        found.Expand(Unit=4)  # wdParagraph
                        found.Style = target_style[style_id]
                        _apply_direct_formatting(found, style_id)
                        applied_via = "find+expand.Style"
                    except Exception as e:
                        restyle_failures.append({
                            "index": para_idx,
                            "old_style": old_style,
                            "error": f"find+expand.Style assign failed: {e}",
                        })
                        continue

                # Verify the style actually took effect; some custom
                # template styles refuse to be overwritten silently.
                try:
                    new_style = para.Style.NameLocal
                except Exception:
                    new_style = None
                expected = target_label[style_id]
                if new_style and new_style != expected:
                    restyle_failures.append({
                        "index": para_idx,
                        "old_style": old_style,
                        "post_style": new_style,
                        "expected": expected,
                        "applied_via": applied_via,
                        "error": (
                            "style assignment did not stick — paragraph "
                            "still reports a different style. Direct "
                            "formatting was applied as fallback."
                        ),
                    })

                if style_id == -2:
                    h1_applied += 1
                else:
                    h2_applied += 1

            # --- Apply list template via LinkedStyle propagation ---
            # Apply list to FIRST H1 paragraph only. Because the template
            # has LinkedStyle for Heading 1 and Heading 2, Word auto-applies
            # the correct list level to ALL paragraphs with those styles.
            # This avoids per-paragraph Range issues (inflated Range.End
            # from comments/fields/bookmarks breaks per-paragraph approach).
            if h1_paragraphs:
                first_h1 = doc.Paragraphs(sorted(h1_paragraphs)[0])
                first_h1.Range.ListFormat.ApplyListTemplateWithLevel(
                    ListTemplate=lt,
                    ContinuePreviousList=False,
                    DefaultListBehavior=1,
                )

            # --- Strip manual numbers ---
            stripped = 0
            if strip_manual_numbers:
                strip_patterns = [
                    r"^MADDE\s+\d+\s*[–\-]\s*",  # "MADDE 6 – ", "MADDE 10 - "
                    r"^\d+(\.\d+)*\.?\s+",         # "1. ", "6.2. ", "10.3 "
                ]
                for para_idx, _ in all_heading_paras:
                    if para_idx < 1 or para_idx > doc.Paragraphs.Count:
                        continue
                    para = doc.Paragraphs(para_idx)
                    text = para.Range.Text.rstrip("\r\x07")
                    for pattern in strip_patterns:
                        m = re.match(pattern, text)
                        if m:
                            prefix_len = len(m.group(0))
                            range_len = para.Range.End - para.Range.Start
                            if range_len <= len(text) + 5:
                                # Normal — para.Range.Start is reliable
                                rng = doc.Range(
                                    para.Range.Start,
                                    para.Range.Start + prefix_len,
                                )
                            else:
                                # Inflated — find text to get real position
                                found = _find_para_text(doc, text)
                                if not found:
                                    break
                                rng = doc.Range(
                                    found.Start,
                                    found.Start + prefix_len,
                                )
                            rng.Delete()
                            stripped += 1
                            break

        return json.dumps({
            "success": True,
            "document": doc.Name,
            "h1_applied": h1_applied,
            "h2_applied": h2_applied,
            "stripped": stripped,
            "restyle_failures": restyle_failures,
        })

    except Exception as e:
        return error_json(e)


@live_tool(mutates=True)
def word_live_replace_text(
    filename: str = None,
    find_text: str = "",
    replace_text: str = "",
    match_case: bool = False,
    match_whole_word: bool = False,
    use_wildcards: bool = False,
    replace_all: bool = True,
    track_changes: bool = False,
) -> str:
    """[Windows only] Find and replace text in an open Word document.

    Uses Word's native Find & Replace, which works across tracked change boundaries
    (unlike manual delete+insert). Supports Word special characters when use_wildcards=True:
    ^m (manual page break), ^t (tab), ^p (paragraph mark), ^s (non-breaking space), and Word wildcard syntax.

    Args:
        filename: Document name or path (None = active document).
        find_text: Text to find. With use_wildcards=True, supports ^m, ^t, ^p, ^s and Word wildcards.
        replace_text: Replacement text. Use "" to delete matches.
        match_case: Case-sensitive search.
        match_whole_word: Match whole words only (ignored when use_wildcards=True).
        use_wildcards: Enable Word wildcards and special characters.
        replace_all: Replace all occurrences (True) or just the first one (False).
        track_changes: Track replacements as revisions.

    Returns:
        JSON with count of replacements made.
    """
    if sys.platform != "win32":
        return json.dumps({"error": "Live editing is only available on Windows"})

    if not find_text:
        return json.dumps({"error": "find_text is required"})

    if len(find_text) > 255:
        return json.dumps({
            "error": f"find_text is {len(find_text)} chars (Word limit: 255). "
            "Break into smaller find/replace pairs."
        })
    if len(replace_text) > 255:
        return json.dumps({
            "error": f"replace_text is {len(replace_text)} chars (Word limit: 255). "
            "Break into smaller find/replace pairs."
        })

    # Reject control bytes (notably \x07 cell separator) that can corrupt
    # Find/Replace and have historically caused full-document data loss.
    from word_mcp.text_safety import reject_control_chars
    try:
        reject_control_chars("find_text", find_text)
        reject_control_chars("replace_text", replace_text)
    except ValueError as e:
        return json.dumps({"error": str(e)})

    if replace_all and track_changes:
        return json.dumps({
            "error": "replace_all=True with track_changes=True causes an infinite loop "
            "(tracked deletions stay visible to Find, triggering endless re-replacement). "
            "Use replace_all=False — each unique text only needs one replacement."
        })

    try:
        from word_mcp.com_runtime import get_word_app, find_document, undo_record

        app = get_word_app()
        doc = find_document(app, filename)

        with undo_record(app, "MCP: Replace Text"):
            prev_tracking = doc.TrackRevisions
            prev_author = app.UserName
            if track_changes:
                doc.TrackRevisions = True
                app.UserName = DEFAULT_AUTHOR
            elif replace_all and prev_tracking:
                # Issue #7: document has TrackRevisions on but caller wants
                # untracked replace_all — disable temporarily to prevent
                # infinite loop (tracked deletions stay visible to Find).
                doc.TrackRevisions = False

            try:
                count = 0
                MAX_REPLACEMENTS = 50_000  # safety ceiling
                rng = doc.Content.Duplicate
                rng.Find.ClearFormatting()

                while True:
                    found = rng.Find.Execute(
                        FindText=find_text,
                        MatchCase=match_case,
                        MatchWholeWord=match_whole_word if not use_wildcards else False,
                        MatchWildcards=use_wildcards,
                        Forward=True,
                        Wrap=0,  # wdFindStop
                    )
                    if not found:
                        break
                    # Guard: zero-length match → skip forward 1 char to avoid infinite loop
                    if rng.Start == rng.End:
                        rng.Start = rng.Start + 1
                        rng.End = doc.Content.End
                        continue
                    # Convert Word special characters to actual characters for rng.Text assignment
                    # (rng.Text doesn't interpret ^p/^t/^m like Find.Execute Replace does)
                    processed = replace_text.replace("^p", "\r").replace("^t", "\t").replace("^m", "\x0c").replace("^s", "\u00a0")
                    rng.Text = processed
                    count += 1
                    if not replace_all:
                        break
                    if count >= MAX_REPLACEMENTS:
                        break
                    rng.Collapse(0)  # wdCollapseEnd — move past replacement
            finally:
                doc.TrackRevisions = prev_tracking
                if track_changes:
                    app.UserName = prev_author

        return json.dumps({
            "success": True,
            "document": doc.Name,
            "find_text": find_text,
            "replace_text": replace_text,
            "replacements": count,
            "replace_all": replace_all,
            "tracked": track_changes,
        }, ensure_ascii=False)

    except Exception as e:
        return error_json(e)


@live_tool(mutates=True)
def word_live_insert_paragraphs(
    filename: str = None,
    paragraphs: list = None,
    target_text: str = None,
    target_paragraph_index: int = None,
    position: str = "after",
    style: str = None,
    track_changes: bool = False,
) -> str:
    """[Windows only] Insert one or more paragraphs near a target paragraph in an open Word document.

    Targets by text match or paragraph index (0-based, matching word_live_get_text output).
    Inserts all paragraphs in a single undo record.

    Args:
        filename: Document name or path (None = active document).
        paragraphs: List of paragraph texts to insert. Each string becomes one Word paragraph.
        target_text: Text to search for (first matching paragraph). Mutually exclusive with target_paragraph_index.
        target_paragraph_index: 0-based paragraph index (as returned by word_live_get_text).
        position: 'before' or 'after' the target paragraph (default 'after').
        style: Style name for inserted paragraphs. None = "Normal" (avoids inheriting heading styles).
        track_changes: Track insertions as revisions.

    Returns:
        JSON with result info including count of paragraphs inserted.
    """
    if sys.platform != "win32":
        return json.dumps({"error": "Live editing is only available on Windows"})

    if not paragraphs or not isinstance(paragraphs, list):
        return json.dumps({"error": "paragraphs must be a non-empty list of strings"})

    if target_text is None and target_paragraph_index is None:
        return json.dumps({"error": "Provide either target_text or target_paragraph_index"})

    if target_text is not None and target_paragraph_index is not None:
        return json.dumps({"error": "Provide target_text or target_paragraph_index, not both"})

    if position not in ("before", "after"):
        return json.dumps({"error": f"position must be 'before' or 'after', got '{position}'"})

    try:
        from word_mcp.com_runtime import get_word_app, find_document, undo_record

        app = get_word_app()
        doc = find_document(app, filename)

        # Find the target paragraph
        total_paras = doc.Paragraphs.Count
        target_para = None

        if target_paragraph_index is not None:
            com_index = target_paragraph_index + 1  # 0-based API → 1-based COM
            if com_index < 1 or com_index > total_paras:
                return json.dumps({
                    "error": f"target_paragraph_index {target_paragraph_index} out of range "
                    f"(0-{total_paras - 1})"
                })
            target_para = doc.Paragraphs(com_index)
        else:
            for i in range(1, total_paras + 1):
                para = doc.Paragraphs(i)
                para_text = para.Range.Text.rstrip("\r\x07")
                if target_text in para_text:
                    target_para = para
                    break
            if target_para is None:
                return json.dumps({"error": f"No paragraph found containing '{target_text}'"})

        resolved_style = style if style else "Normal"

        with undo_record(app, "MCP: Insert Paragraphs"):
            prev_tracking = doc.TrackRevisions
            prev_author = app.UserName
            if track_changes:
                doc.TrackRevisions = True
                app.UserName = DEFAULT_AUTHOR

            try:
                inserted = 0

                if position == "after":
                    rng = target_para.Range.Duplicate
                    rng.Collapse(0)  # wdCollapseEnd
                    for para_text in paragraphs:
                        rng.InsertParagraphAfter()
                        rng.Collapse(0)  # wdCollapseEnd
                        rng.InsertAfter(para_text)
                        try:
                            rng.Style = resolved_style
                        except Exception:
                            pass
                        rng.Collapse(0)  # wdCollapseEnd
                        inserted += 1
                else:  # "before"
                    for para_text in reversed(paragraphs):
                        rng = target_para.Range.Duplicate
                        rng.Collapse(1)  # wdCollapseStart
                        rng.InsertParagraphBefore()
                        rng.Collapse(1)  # wdCollapseStart
                        rng.InsertAfter(para_text)
                        try:
                            rng.Style = resolved_style
                        except Exception:
                            pass
                        inserted += 1
            finally:
                doc.TrackRevisions = prev_tracking
                if track_changes:
                    app.UserName = prev_author

        return json.dumps({
            "success": True,
            "document": doc.Name,
            "paragraphs_inserted": inserted,
            "position": position,
            "style": resolved_style,
            "tracked": track_changes,
        }, ensure_ascii=False)

    except Exception as e:
        return error_json(e)


@live_tool(mutates=True)
def word_live_delete_text(
    filename: str = None,
    start: int = None,
    end: int = None,
    track_changes: bool = False,
) -> str:
    """Delete text from an open Word document.

    Args:
        filename: Document name or path.
        start: Start character position.
        end: End character position.
        track_changes: Track deletion as a revision.

    Returns:
        JSON with deleted text info.
    """
    if sys.platform != "win32":
        return json.dumps({"error": "Live editing is only available on Windows"})

    if start is None or end is None:
        return json.dumps(
            {"error": "Both 'start' and 'end' character positions are required"}
        )

    try:
        from word_mcp.com_runtime import get_word_app, find_document, undo_record

        app = get_word_app()
        doc = find_document(app, filename)
        rng = doc.Range(start, end)
        deleted_text = rng.Text

        with undo_record(app, "MCP: Delete Text"):
            prev_tracking = doc.TrackRevisions
            prev_author = app.UserName
            if track_changes:
                doc.TrackRevisions = True
                app.UserName = DEFAULT_AUTHOR

            try:
                # Delete any table objects within the range first
                # (rng.Delete only removes text, leaving ghost table structure)
                for i in range(doc.Tables.Count, 0, -1):
                    tbl = doc.Tables(i)
                    if tbl.Range.Start >= start and tbl.Range.End <= end:
                        tbl.Delete()
                # Delete remaining text in the range
                rng = doc.Range(start, min(end, doc.Content.End))
                if rng.Start < rng.End:
                    rng.Delete()
            finally:
                if track_changes:
                    doc.TrackRevisions = prev_tracking
                    app.UserName = prev_author

        preview = deleted_text
        if len(preview) > 100:
            preview = preview[:100] + "..."

        return json.dumps(
            {
                "success": True,
                "document": doc.Name,
                "deleted_text": preview,
                "range": f"{start}-{end}",
                "tracked": track_changes,
            }
        )

    except Exception as e:
        return error_json(e)


@live_tool(mutates=True)
def word_live_undo(
    filename: str = None,
    times: int = 1,
) -> str:
    """[Windows only] Undo the last N operations in an open Word document.

    Each MCP destructive tool call is grouped as a single undo entry (e.g.,
    "MCP: Insert Text"). Calling undo(times=1) reverts the last MCP operation;
    undo(times=3) reverts the last three.

    Args:
        filename: Document name or path (None = active document).
        times: Number of undo steps (default 1).

    Returns:
        JSON with success status and number of undone steps.
    """
    if sys.platform != "win32":
        return json.dumps({"error": "Live editing is only available on Windows"})

    if times < 1:
        return json.dumps({"error": "times must be >= 1"})

    try:
        from word_mcp.com_runtime import get_word_app, find_document

        app = get_word_app()
        doc = find_document(app, filename)

        result = doc.Undo(times)

        return json.dumps({
            "success": bool(result),
            "document": doc.Name,
            "times_requested": times,
            "undo_result": bool(result),
        })

    except Exception as e:
        return error_json(e)


@live_tool(mutates=True)
def word_live_save(
    filename: str = None,
    save_as: str = None,
) -> str:
    """Save an open Word document.

    Saves the document. Optionally saves to a new path with save_as.

    Args:
        filename: Document name or path (None = active document).
        save_as: Optional new file path to save as. If omitted, saves in place.

    Returns:
        JSON with save result.
    """
    if sys.platform != "win32":
        return json.dumps({"error": "Live editing is only available on Windows"})

    try:
        from word_mcp.com_runtime import get_word_app, find_document

        app = get_word_app()
        doc = find_document(app, filename)

        if save_as:
            save_path = os.path.abspath(save_as)
            # Determine format from extension
            ext = os.path.splitext(save_path)[1].lower()
            format_map = {
                ".docx": 16,  # wdFormatXMLDocument
                ".doc": 0,    # wdFormatDocument
                ".pdf": 17,   # wdFormatPDF
                ".rtf": 6,    # wdFormatRTF
                ".txt": 2,    # wdFormatText
            }
            file_format = format_map.get(ext, 16)
            doc.SaveAs2(save_path, FileFormat=file_format)
            return json.dumps({
                "success": True,
                "document": doc.Name,
                "saved_as": save_path,
                "format": ext,
            }, ensure_ascii=False)
        else:
            doc.Save()
            return json.dumps({
                "success": True,
                "document": doc.Name,
                "path": doc.FullName,
            }, ensure_ascii=False)

    except Exception as e:
        return error_json(e)


@live_tool(mutates=True)
def word_live_toggle_track_changes(
    filename: str = None,
    enable: bool = None,
) -> str:
    """Toggle or set track changes mode on an open Word document.

    If enable is omitted, toggles the current state.

    Args:
        filename: Document name or path (None = active document).
        enable: True to enable, False to disable, None to toggle.

    Returns:
        JSON with the new track changes state.
    """
    if sys.platform != "win32":
        return json.dumps({"error": "Live editing is only available on Windows"})

    try:
        from word_mcp.com_runtime import get_word_app, find_document

        app = get_word_app()
        doc = find_document(app, filename)

        previous = bool(doc.TrackRevisions)
        if enable is None:
            doc.TrackRevisions = not previous
        else:
            doc.TrackRevisions = enable

        return json.dumps({
            "success": True,
            "document": doc.Name,
            "previous_state": previous,
            "track_changes": bool(doc.TrackRevisions),
        })

    except Exception as e:
        return error_json(e)
