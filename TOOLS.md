# Tool Reference

The 45 tools provided by word-mcp. They operate on documents currently open in Word on Windows through COM automation. Each editing call is a single Ctrl+Z entry in Word's undo stack.

<details open>
<summary><b>Editing</b></summary>

| Tool | Description |
|------|-------------|
| `word_live_insert_text` | Insert text at a position (with optional tracked changes) |
| `word_live_delete_text` | Delete a character range |
| `word_live_replace_text` | Find & replace via COM — works across tracked change boundaries; supports wildcards including `^s` (non-breaking space) |
| `word_live_insert_paragraphs` | Insert multiple paragraphs near a target (by text or index) in a single undo record |
| `word_live_format_text` | Format text (bold, italic, font, highlight, paragraph alignment, page break before) |
| `word_live_add_table` | Insert a table |
| `word_live_format_table` | Format an existing table |
| `word_live_apply_list` | Apply bullet, numbered, or multilevel list formatting |
| `word_live_setup_heading_numbering` | Auto-numbered headings (1. / 1.1) with configurable style |
| `word_live_modify_table` | Modify table structure: get info, set cell, set row, set range, add/delete rows/columns, merge cells, autofit, or delete table |
| `word_live_save` | Save document in place or save-as to a new path (docx, pdf, rtf, txt) |
| `word_live_toggle_track_changes` | Toggle or explicitly set track changes mode on/off |
| `word_live_insert_image` | Insert an image with sizing, alignment, wrapping, and optional border |
| `word_live_insert_cross_reference` | Insert a live cross-reference to headings, bookmarks, figures, tables, equations, footnotes, or endnotes |
| `word_live_insert_equation` | Insert a mathematical equation using UnicodeMath syntax |

</details>

<details open>
<summary><b>Reading</b></summary>

| Tool | Description |
|------|-------------|
| `word_live_list_open` | List all documents currently open in Word with name, path, pages, and saved status |
| `word_live_get_text` | Get all text paragraph by paragraph |
| `word_live_take_snapshot` | Store paragraph baseline for efficient change detection |
| `word_live_get_diff` | Compare current document against snapshot — returns only changed paragraphs |
| `word_live_snapshot_status` | Check snapshot existence and age |
| `word_live_get_page_text` | Get text from specific page(s) with char offsets for chaining |
| `word_live_get_paragraph_format` | Inspect paragraph formatting (font, spacing, alignment, list info, per-run detail) |
| `word_live_get_info` | Get document metadata (pages, words, sections) |
| `word_live_find_text` | Find text with context; supports wildcards |
| `word_live_get_undo_history` | List undo stack entries |
| `word_live_list_cross_reference_items` | List available cross-reference targets (headings, bookmarks, figures, tables) with indices |
| `word_live_diagnose_layout` | Scan for layout problems (keep_with_next chains, style misuse, break issues) |

</details>

<details open>
<summary><b>Comments & Revisions</b></summary>

| Tool | Description |
|------|-------------|
| `word_live_get_comments` | Get all comments |
| `word_live_add_comment` | Add a comment anchored to text |
| `word_live_list_revisions` | List tracked changes |
| `word_live_reply_to_comment` | Add a threaded reply to an existing comment (Word 2016+) |
| `word_live_resolve_comment` | Mark a comment as resolved or unresolve it (Word 2016+) |
| `word_live_delete_comment` | Permanently delete a comment from the document |
| `word_live_accept_revisions` | Accept tracked changes (all or by author/type) |
| `word_live_reject_revisions` | Reject tracked changes (all or by author/type) |

</details>

<details open>
<summary><b>Layout</b></summary>

| Tool | Description |
|------|-------------|
| `word_live_set_page_layout` | Set orientation, size, and margins |
| `word_live_add_header_footer` | Add header/footer text |
| `word_live_add_page_numbers` | Add page numbers |
| `word_live_add_section_break` | Add section break |
| `word_live_set_paragraph_spacing` | Set paragraph spacing |
| `word_live_add_bookmark` | Add a named bookmark |
| `word_live_add_watermark` | Add a text watermark |
| `word_live_set_core_properties` | Set built-in document properties (title, subject, author, keywords, and others) |

</details>

<details open>
<summary><b>Undo & Screen Capture</b></summary>

| Tool | Description |
|------|-------------|
| `word_live_undo` | Undo last N operations (each tool call = one undo entry) |
| `word_screen_capture` | Screenshot of the Word window |

</details>
