import os
import sys

os.environ.setdefault("FASTMCP_LOG_LEVEL", "WARNING")

from fastmcp import FastMCP
from mcp.types import ToolAnnotations

from word_mcp.defaults import DEFAULT_AUTHOR
from word_mcp.tools import edit, layout, read, references, screen, tables

mcp = FastMCP("word-mcp")


def register_tools():
    @mcp.tool(
        annotations=ToolAnnotations(
            title="Word Screen Capture",
            readOnlyHint=True,
        ),
    )
    async def word_screen_capture(filename: str = None, output_path: str = None):
        """[Windows only] Capture a screenshot of a Word document window.
        Returns the path to the saved PNG image. Requires Word to be running."""
        return await screen.word_screen_capture(filename, output_path)

    @mcp.tool(
        annotations=ToolAnnotations(
            title="Word Live Insert Text",
            destructiveHint=True,
        ),
    )
    async def word_live_insert_text(
        filename: str = None,
        text: str = "",
        position: str = "end",
        bookmark: str = None,
        track_changes: bool = False,
    ):
        """[Windows only] Insert text into a Word document that is open in Word.
        Position: 'start', 'end', 'cursor', or character offset. Requires Word running."""
        return await edit.word_live_insert_text(
            filename, text, position, bookmark, track_changes
        )

    @mcp.tool(
        annotations=ToolAnnotations(
            title="Word Live Format Text",
            destructiveHint=True,
        ),
        description=edit.word_live_format_text.__doc__,
    )
    async def word_live_format_text(
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
    ):
        return await edit.word_live_format_text(
            filename, start, end, start_paragraph, end_paragraph,
            bold, italic, underline, strikethrough,
            font_name, font_size, font_color, highlight_color,
            style_name, paragraph_alignment, page_break_before,
            preserve_direct_formatting, track_changes,
        )

    @mcp.tool(
        annotations=ToolAnnotations(
            title="Word Live Replace Text",
            destructiveHint=True,
        ),
        description=edit.word_live_replace_text.__doc__,
    )
    async def word_live_replace_text(
        filename: str = None,
        find_text: str = "",
        replace_text: str = "",
        match_case: bool = False,
        match_whole_word: bool = False,
        use_wildcards: bool = False,
        replace_all: bool = True,
        track_changes: bool = False,
    ):
        return await edit.word_live_replace_text(
            filename, find_text, replace_text, match_case,
            match_whole_word, use_wildcards, replace_all, track_changes,
        )

    @mcp.tool(
        annotations=ToolAnnotations(
            title="Word Live Insert Paragraphs",
            destructiveHint=True,
        ),
        description=edit.word_live_insert_paragraphs.__doc__,
    )
    async def word_live_insert_paragraphs(
        filename: str = None,
        paragraphs: list = None,
        target_text: str = None,
        target_paragraph_index: int = None,
        position: str = "after",
        style: str = None,
        track_changes: bool = False,
    ):
        return await edit.word_live_insert_paragraphs(
            filename, paragraphs, target_text, target_paragraph_index,
            position, style, track_changes,
        )

    @mcp.tool(
        annotations=ToolAnnotations(
            title="Word Live Add Table",
            destructiveHint=True,
        ),
    )
    async def word_live_add_table(
        filename: str = None,
        rows: int = 2,
        cols: int = 2,
        position: str = "end",
        data: list = None,
        style: str = "Table Grid",
        autofit: str = "window",
        track_changes: bool = False,
    ):
        """[Windows only] Add a table to a Word document open in Word.
        Optionally provide data as 2D list. Default style is 'Table Grid' with
        autofit to window width. Set style=None for no style, autofit=None for
        legacy fixed behavior. Requires Word running."""
        return await tables.word_live_add_table(
            filename, rows, cols, position, data, style, autofit, track_changes
        )

    @mcp.tool(
        annotations=ToolAnnotations(
            title="Word Live Format Table",
            destructiveHint=True,
        ),
        description=tables.word_live_format_table.__doc__,
    )
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
    ):
        return await tables.word_live_format_table(
            filename, table_index, border_style, cell_bold, cell_alignment,
            column_widths, table_alignment, cell_shading, autofit
        )

    @mcp.tool(
        annotations=ToolAnnotations(
            title="Word Live Modify Table",
            destructiveHint=True,
        ),
        description=tables.word_live_modify_table.__doc__,
    )
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
    ):
        return await tables.word_live_modify_table(
            filename, table_index, operation, row, col, text,
            before_row, before_col, header, cells,
            start_row, start_col, end_row, end_col,
            autofit_mode, accept_revisions, track_changes,
        )

    @mcp.tool(
        annotations=ToolAnnotations(
            title="Word Live Delete Text",
            destructiveHint=True,
        ),
    )
    async def word_live_delete_text(
        filename: str = None,
        start: int = None,
        end: int = None,
        track_changes: bool = False,
    ):
        """[Windows only] Delete text from a Word document open in Word.
        Specify start/end character positions. Requires Word running."""
        return await edit.word_live_delete_text(
            filename, start, end, track_changes
        )

    @mcp.tool(
        annotations=ToolAnnotations(
            title="Word Live Apply List",
            destructiveHint=True,
        ),
        description=edit.word_live_apply_list.__doc__,
    )
    async def word_live_apply_list(
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
    ):
        return await edit.word_live_apply_list(
            filename, start_paragraph, end_paragraph, list_type,
            level, remove, continue_previous, number_format,
            number_style, start_at, level_map, track_changes,
        )

    @mcp.tool(
        annotations=ToolAnnotations(
            title="Word Live Setup Heading Numbering",
            destructiveHint=True,
        ),
        description=edit.word_live_setup_heading_numbering.__doc__,
    )
    async def word_live_setup_heading_numbering(
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
    ):
        return await edit.word_live_setup_heading_numbering(
            filename, h1_paragraphs, h2_paragraphs, strip_manual_numbers,
            h1_number_format, h2_number_format,
            font_name, h1_size, h2_size, bold, alignment, font_color,
            h1_space_before, h1_space_after, h2_space_before, h2_space_after,
            line_spacing,
        )

    @mcp.tool(
        annotations=ToolAnnotations(
            title="Word Live Get Text",
            readOnlyHint=True,
        ),
    )
    async def word_live_get_text(filename: str = None):
        """[Windows only] Get all text from a Word document open in Word, paragraph by paragraph. For large documents (200+ paragraphs), automatically returns only the first 3 pages — use word_live_get_page_text to read specific pages."""
        return await read.word_live_get_text(filename)

    @mcp.tool(
        annotations=ToolAnnotations(
            title="Word Live Take Snapshot",
            readOnlyHint=True,
        ),
    )
    async def word_live_take_snapshot(filename: str = None):
        """[Windows only] Store a snapshot of the current document text for later diffing without returning the full text. Use word_live_get_diff afterwards to see what changed."""
        return await read.word_live_take_snapshot(filename)

    @mcp.tool(
        annotations=ToolAnnotations(
            title="Word Live Get Diff",
            readOnlyHint=True,
        ),
    )
    async def word_live_get_diff(filename: str = None):
        """[Windows only] Return only paragraphs that changed since the last snapshot. Compares current document against snapshot from word_live_take_snapshot. Returns added, modified, deleted paragraphs. Automatically updates snapshot after diffing."""
        return await read.word_live_get_diff(filename)

    @mcp.tool(
        annotations=ToolAnnotations(
            title="Word Live Snapshot Status",
            readOnlyHint=True,
        ),
    )
    async def word_live_snapshot_status(filename: str = None):
        """[Windows only] Check whether a snapshot exists for the document and how old it is. Returns has_snapshot, age_seconds, and paragraph_count."""
        return await read.word_live_snapshot_status(filename)

    @mcp.tool(
        annotations=ToolAnnotations(
            title="Word Live Get Paragraph Format",
            readOnlyHint=True,
        ),
        description=read.word_live_get_paragraph_format.__doc__,
    )
    async def word_live_get_paragraph_format(
        filename: str = None,
        start_paragraph: int = None,
        end_paragraph: int = None,
        include_runs: bool = False,
    ):
        return await read.word_live_get_paragraph_format(
            filename, start_paragraph, end_paragraph, include_runs,
        )

    @mcp.tool(
        annotations=ToolAnnotations(
            title="Word Live Get Info",
            readOnlyHint=True,
        ),
    )
    async def word_live_get_info(filename: str = None):
        """[Windows only] Get document info (pages, words, sections, etc.) from a Word document open in Word. Requires Word running."""
        return await read.word_live_get_info(filename)

    @mcp.tool(
        annotations=ToolAnnotations(
            title="Word Live Set Core Properties",
            destructiveHint=True,
        ),
        description=read.word_live_set_core_properties.__doc__,
    )
    async def word_live_set_core_properties(
        filename: str = None,
        title: str = None,
        subject: str = None,
        author: str = None,
        keywords: str = None,
        comments: str = None,
        category: str = None,
        manager: str = None,
        company: str = None,
        last_author: str = None,
    ):
        return await read.word_live_set_core_properties(
            filename=filename,
            title=title,
            subject=subject,
            author=author,
            keywords=keywords,
            comments=comments,
            category=category,
            manager=manager,
            company=company,
            last_author=last_author,
        )

    @mcp.tool(
        annotations=ToolAnnotations(
            title="Word Live List Open",
            readOnlyHint=True,
        ),
    )
    async def word_live_list_open():
        """[Windows only] List all documents currently open in Word with name, path, pages, and saved status."""
        return await read.word_live_list_open()

    @mcp.tool(
        annotations=ToolAnnotations(
            title="Word Live Find Text",
            readOnlyHint=True,
        ),
    )
    async def word_live_find_text(
        filename: str = None,
        search_text: str = "",
        match_case: bool = False,
        whole_word: bool = False,
        use_wildcards: bool = False,
        context_chars: int = 60,
        max_results: int = 50,
    ):
        """[Windows only] Find text in a Word document open in Word. Returns positions and context.
        With use_wildcards=True, supports ^m (page break), ^t (tab), ^p (paragraph mark) and Word wildcards.
        context_chars controls how many characters of surrounding context to return (default 60). Requires Word running."""
        return await read.word_live_find_text(
            filename, search_text, match_case, whole_word,
            use_wildcards, context_chars, max_results,
        )

    @mcp.tool(
        annotations=ToolAnnotations(
            title="Word Live Get Comments",
            readOnlyHint=True,
        ),
    )
    async def word_live_get_comments(filename: str = None):
        """[Windows only] Get all comments from a Word document open in Word. Requires Word running."""
        return await read.word_live_get_comments(filename)

    @mcp.tool(
        annotations=ToolAnnotations(
            title="Word Live Add Comment",
            destructiveHint=True,
        ),
    )
    async def word_live_add_comment(
        filename: str = None,
        start: int = None,
        end: int = None,
        paragraph_index: int = None,
        text: str = "",
        author: str = DEFAULT_AUTHOR,
    ):
        """[Windows only] Add a comment to a Word document open in Word.
        Specify start/end character positions or paragraph_index (1-indexed). Requires Word running."""
        return await read.word_live_add_comment(
            filename, start, end, paragraph_index, text, author
        )

    @mcp.tool(
        annotations=ToolAnnotations(
            title="Word Live Reply to Comment",
            destructiveHint=True,
        ),
    )
    async def word_live_reply_to_comment(
        filename: str = None,
        comment_index: int = None,
        text: str = "",
        author: str = DEFAULT_AUTHOR,
    ):
        """[Windows only] Reply to an existing comment in a Word document open in Word.
        Adds a threaded reply. Use word_live_get_comments to find the comment_index.
        Requires Word 2016+ running."""
        return await read.word_live_reply_to_comment(
            filename, comment_index, text, author
        )

    @mcp.tool(
        annotations=ToolAnnotations(
            title="Word Live Resolve Comment",
            destructiveHint=True,
        ),
    )
    async def word_live_resolve_comment(
        filename: str = None,
        comment_index: int = None,
        resolve: bool = True,
    ):
        """[Windows only] Resolve or unresolve a comment in a Word document open in Word.
        Sets the comment's Done property. Use word_live_get_comments to find comment_index.
        Requires Word 2016+ running."""
        return await read.word_live_resolve_comment(
            filename, comment_index, resolve
        )

    @mcp.tool(
        annotations=ToolAnnotations(
            title="Word Live Delete Comment",
            destructiveHint=True,
        ),
    )
    async def word_live_delete_comment(
        filename: str = None,
        comment_index: int = None,
    ):
        """[Windows only] Delete a comment from a Word document open in Word.
        Permanently removes the comment. Use word_live_get_comments to find comment_index.
        Requires Word running."""
        return await read.word_live_delete_comment(
            filename, comment_index
        )

    @mcp.tool(
        annotations=ToolAnnotations(
            title="Word Live List Revisions",
            readOnlyHint=True,
        ),
    )
    async def word_live_list_revisions(filename: str = None):
        """[Windows only] List all tracked changes (revisions) in a Word document open in Word. Requires Word running."""
        return await read.word_live_list_revisions(filename)

    @mcp.tool(
        annotations=ToolAnnotations(
            title="Word Live Accept Revisions",
            destructiveHint=True,
        ),
    )
    async def word_live_accept_revisions(
        filename: str = None,
        author: str = None,
        revision_ids: list[int] = None,
    ):
        """[Windows only] Accept tracked changes in a Word document open in Word.
        Filter by author or specific revision IDs. Requires Word running."""
        return await read.word_live_accept_revisions(
            filename, author, revision_ids
        )

    @mcp.tool(
        annotations=ToolAnnotations(
            title="Word Live Reject Revisions",
            destructiveHint=True,
        ),
    )
    async def word_live_reject_revisions(
        filename: str = None,
        author: str = None,
        revision_ids: list[int] = None,
    ):
        """[Windows only] Reject tracked changes in a Word document open in Word.
        Filter by author or specific revision IDs. Requires Word running."""
        return await read.word_live_reject_revisions(
            filename, author, revision_ids
        )

    @mcp.tool(
        annotations=ToolAnnotations(
            title="Word Live Get Page Text",
            readOnlyHint=True,
        ),
        description=read.word_live_get_page_text.__doc__,
    )
    async def word_live_get_page_text(
        filename: str = None,
        page: int = 1,
        end_page: int = None,
    ):
        return await read.word_live_get_page_text(
            filename, page, end_page,
        )

    @mcp.tool(
        annotations=ToolAnnotations(
            title="Word Live Get Undo History",
            readOnlyHint=True,
        ),
    )
    async def word_live_get_undo_history(filename: str = None):
        """[Windows only] Get the undo stack from a Word document open in Word.
        Shows MCP tool operations as named entries. Requires Word running."""
        return await read.word_live_get_undo_history(filename)

    @mcp.tool(
        annotations=ToolAnnotations(
            title="Word Live Undo",
            destructiveHint=True,
        ),
    )
    async def word_live_undo(
        filename: str = None,
        times: int = 1,
    ):
        """[Windows only] Undo the last N operations in a Word document open in Word.
        Each MCP tool call is one undo entry. Requires Word running."""
        return await edit.word_live_undo(filename, times)

    @mcp.tool(
        annotations=ToolAnnotations(
            title="Word Live Save",
            destructiveHint=True,
        ),
    )
    async def word_live_save(
        filename: str = None,
        save_as: str = None,
    ):
        """[Windows only] Save a Word document open in Word.
        Optionally save to a new path with save_as. Requires Word running."""
        return await edit.word_live_save(filename, save_as)

    @mcp.tool(
        annotations=ToolAnnotations(
            title="Word Live Toggle Track Changes",
            destructiveHint=True,
        ),
    )
    async def word_live_toggle_track_changes(
        filename: str = None,
        enable: bool = None,
    ):
        """[Windows only] Toggle or set Track Changes mode on a Word document.
        If enable is omitted, toggles current state. Requires Word running."""
        return await edit.word_live_toggle_track_changes(filename, enable)

    @mcp.tool(
        annotations=ToolAnnotations(
            title="Word Live Insert Image",
            destructiveHint=True,
        ),
        description=references.word_live_insert_image.__doc__,
    )
    async def word_live_insert_image(
        filename: str = None,
        image_path: str = "",
        paragraph_index: int = None,
        position: str = "end",
        width_inches: float = None,
        height_inches: float = None,
        width_pt: float = None,
        height_pt: float = None,
        alignment: str = None,
        wrapping: str = None,
        border_style: str = None,
        border_width_pt: float = None,
        border_color: str = None,
        link_to_file: bool = False,
    ):
        return await references.word_live_insert_image(
            filename, image_path, paragraph_index, position,
            width_inches, height_inches, width_pt, height_pt,
            alignment, wrapping, border_style, border_width_pt,
            border_color, link_to_file
        )

    @mcp.tool(
        annotations=ToolAnnotations(
            title="Word Live Insert Cross Reference",
            destructiveHint=True,
        ),
    )
    async def word_live_insert_cross_reference(
        filename: str = None,
        ref_type: str = "heading",
        ref_item: int = 1,
        ref_kind: str = "text",
        insert_position: str = "end",
        paragraph_index: int = None,
        insert_as_hyperlink: bool = True,
    ):
        """[Windows only] Insert a cross-reference to a heading, bookmark, figure, table, etc.
        First use word_live_list_cross_reference_items to discover available targets.
        ref_type: heading, bookmark, figure, table, equation, footnote, endnote.
        ref_kind: text, number, number_no_context, page, above_below.
        Requires Word running."""
        return await references.word_live_insert_cross_reference(
            filename, ref_type, ref_item, ref_kind,
            insert_position, paragraph_index, insert_as_hyperlink
        )

    @mcp.tool(
        annotations=ToolAnnotations(
            title="Word Live List Cross Reference Items",
            readOnlyHint=True,
        ),
    )
    async def word_live_list_cross_reference_items(
        filename: str = None,
        ref_type: str = "heading",
    ):
        """[Windows only] List available cross-reference targets in a Word document.
        Returns items that can be referenced with word_live_insert_cross_reference.
        ref_type: heading, bookmark, figure, table, equation, footnote, endnote.
        Requires Word running."""
        return await references.word_live_list_cross_reference_items(filename, ref_type)

    @mcp.tool(
        annotations=ToolAnnotations(
            title="Word Live Insert Equation",
            destructiveHint=True,
        ),
    )
    async def word_live_insert_equation(
        filename: str = None,
        equation: str = "",
        paragraph_index: int = None,
        position: str = "end",
        display_mode: bool = False,
    ):
        """[Windows only] Insert a mathematical equation into a Word document using UnicodeMath syntax.
        Examples: "x^2 + y^2 = z^2", "(a+b)/(c+d)" (fraction), "\\sqrt(x^2+y^2)" (root),
        "\\alpha + \\beta" (Greek), "\\int_0^\\infty e^(-x^2) dx" (integral),
        "\\sum_(i=1)^n i^2" (summation), "\\matrix(a&b@c&d)" (matrix).
        display_mode=True centers the equation on its own line.
        Requires Word running."""
        return await references.word_live_insert_equation(
            filename, equation, paragraph_index, position, display_mode
        )

    @mcp.tool(
        annotations=ToolAnnotations(
            title="Word Live Diagnose Layout",
            readOnlyHint=True,
        ),
        description=read.word_live_diagnose_layout.__doc__,
    )
    async def word_live_diagnose_layout(filename: str = None):
        return await read.word_live_diagnose_layout(filename)

    @mcp.tool(
        annotations=ToolAnnotations(
            title="Word Live Set Page Layout",
            destructiveHint=True,
        ),
    )
    async def word_live_set_page_layout(
        filename: str = None,
        section_index: int = 1,
        orientation: str = None,
        page_width_inches: float = None,
        page_height_inches: float = None,
        margin_top_inches: float = None,
        margin_bottom_inches: float = None,
        margin_left_inches: float = None,
        margin_right_inches: float = None,
    ):
        """[Windows only] Set page layout (orientation, size, margins) for a section in a Word document open in Word. Requires Word running."""
        return await layout.word_live_set_page_layout(
            filename, section_index, orientation,
            page_width_inches, page_height_inches,
            margin_top_inches, margin_bottom_inches,
            margin_left_inches, margin_right_inches,
        )

    @mcp.tool(
        annotations=ToolAnnotations(
            title="Word Live Add Header/Footer",
            destructiveHint=True,
        ),
    )
    async def word_live_add_header_footer(
        filename: str = None,
        section_index: int = 1,
        header_text: str = None,
        footer_text: str = None,
        header_alignment: str = "center",
        footer_alignment: str = "center",
    ):
        """[Windows only] Add header and/or footer to a section in a Word document open in Word. Requires Word running."""
        return await layout.word_live_add_header_footer(
            filename, section_index, header_text, footer_text,
            header_alignment, footer_alignment,
        )

    @mcp.tool(
        annotations=ToolAnnotations(
            title="Word Live Add Page Numbers",
            destructiveHint=True,
        ),
    )
    async def word_live_add_page_numbers(
        filename: str = None,
        section_index: int = 1,
        position: str = "footer",
        alignment: str = "center",
        prefix: str = "",
        suffix: str = "",
        include_total: bool = False,
    ):
        """[Windows only] Add page numbers to header or footer in a Word document open in Word. Requires Word running."""
        return await layout.word_live_add_page_numbers(
            filename, section_index, position, alignment,
            prefix, suffix, include_total,
        )

    @mcp.tool(
        annotations=ToolAnnotations(
            title="Word Live Add Section Break",
            destructiveHint=True,
        ),
    )
    async def word_live_add_section_break(
        filename: str = None,
        break_type: str = "new_page",
    ):
        """[Windows only] Add a section break (new_page, continuous, even_page, odd_page) to a Word document open in Word. Requires Word running."""
        return await layout.word_live_add_section_break(
            filename, break_type,
        )

    @mcp.tool(
        annotations=ToolAnnotations(
            title="Word Live Set Paragraph Spacing",
            destructiveHint=True,
        ),
    )
    async def word_live_set_paragraph_spacing(
        filename: str = None,
        paragraph_index: int = None,
        start_paragraph: int = None,
        end_paragraph: int = None,
        space_before_pt: float = None,
        space_after_pt: float = None,
        line_spacing: float = None,
        line_spacing_rule: str = None,
        keep_with_next: bool = None,
        keep_together: bool = None,
        alignment: str = None,
    ):
        """[Windows only] Set paragraph spacing and layout properties in a Word document open in Word. Paragraphs are 1-indexed. Requires Word running."""
        return await layout.word_live_set_paragraph_spacing(
            filename, paragraph_index, start_paragraph, end_paragraph,
            space_before_pt, space_after_pt, line_spacing, line_spacing_rule,
            keep_with_next, keep_together, alignment,
        )

    @mcp.tool(
        annotations=ToolAnnotations(
            title="Word Live Add Bookmark",
            destructiveHint=True,
        ),
    )
    async def word_live_add_bookmark(
        filename: str = None,
        paragraph_index: int = 1,
        bookmark_name: str = "",
    ):
        """[Windows only] Add a named bookmark at a paragraph in a Word document open in Word.
        Paragraph is 1-indexed. Requires Word running."""
        return await layout.word_live_add_bookmark(
            filename, paragraph_index, bookmark_name,
        )

    @mcp.tool(
        annotations=ToolAnnotations(
            title="Word Live Add Watermark",
            destructiveHint=True,
        ),
    )
    async def word_live_add_watermark(
        filename: str = None,
        text: str = "TASLAK",
        font_size: int = 72,
        font_color: str = "C0C0C0",
        rotation: int = -45,
        section_index: int = 1,
    ):
        """[Windows only] Add a diagonal text watermark to a Word document open in Word. Requires Word running."""
        return await layout.word_live_add_watermark(
            filename, text, font_size, font_color, rotation, section_index,
        )


def run():
    if sys.platform != "win32":
        print("word-mcp only runs on Windows", file=sys.stderr)
        raise SystemExit(1)
    register_tools()
    mcp.run(transport="stdio", show_banner=False)
