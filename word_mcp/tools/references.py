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
) -> str:
    """Insert an image into an open Word document.

    The image can be placed at a specific paragraph, at the start or end,
    or at a character offset position.

    Args:
        filename: Document name or path (None = active document).
        image_path: Full path to the image file (PNG, JPG, BMP, etc.).
        paragraph_index: 1-indexed paragraph to insert before (image goes before the paragraph).
        position: "start", "end", or character offset as string. Only used if paragraph_index is None.
        width_inches: Optional width in inches (aspect ratio maintained if only one dimension given).
        height_inches: Optional height in inches.
        width_pt: Optional width in points (1 inch = 72 pt). Overrides width_inches if both given.
        height_pt: Optional height in points. Overrides height_inches if both given.
        alignment: Paragraph alignment for the image: "left", "center", "right". Default: unchanged.
        wrapping: Text wrapping style: "inline" (default), "square", "tight", "behind",
            "infront", "topbottom". Non-inline converts to a floating Shape.
        border_style: Border style around the image: "single", "double", "dotted", "dashed",
            "thick", "none". Default: no border.
        border_width_pt: Border line width in points (e.g. 1.0, 2.0). Default: 1.0.
        border_color: Border color as "#RRGGBB" hex string. Default: black (#000000).
        link_to_file: If True, links to the file instead of embedding it.

    Returns:
        JSON with image insertion result.
    """
    if sys.platform != "win32":
        return json.dumps({"error": "Live editing is only available on Windows"})

    if not image_path:
        return json.dumps({"error": "image_path is required"})

    abs_path = os.path.abspath(image_path)
    if not os.path.isfile(abs_path):
        return json.dumps({"error": f"Image file not found: {abs_path}"})

    try:
        from word_mcp.word_com import get_word_app, find_document, undo_record

        app = get_word_app()
        doc = find_document(app, filename)

        # Determine insertion range
        if paragraph_index is not None:
            if paragraph_index < 1 or paragraph_index > doc.Paragraphs.Count:
                return json.dumps({
                    "error": f"paragraph_index {paragraph_index} out of range (1-{doc.Paragraphs.Count})"
                })
            rng = doc.Paragraphs(paragraph_index).Range
            rng.Collapse(1)  # wdCollapseStart
        elif position == "start":
            rng = doc.Range(0, 0)
        elif position == "end":
            rng = doc.Range()
            rng.Collapse(0)  # wdCollapseEnd
        else:
            try:
                offset = int(position)
                rng = doc.Range(offset, offset)
            except (ValueError, TypeError):
                rng = doc.Range()
                rng.Collapse(0)

        # Resolve final size in points (pt params override inches params)
        final_w = None
        final_h = None
        if width_pt is not None:
            final_w = float(width_pt)
        elif width_inches is not None:
            final_w = float(width_inches) * 72.0
        if height_pt is not None:
            final_h = float(height_pt)
        elif height_inches is not None:
            final_h = float(height_inches) * 72.0

        # Wrapping style constants (wdWrapType)
        WRAP_STYLES = {
            "inline": None,       # keep as InlineShape
            "square": 0,          # wdWrapSquare
            "tight": 1,           # wdWrapTight
            "behind": 3,          # wdWrapBehind
            "infront": 4,         # wdWrapFront
            "topbottom": 2,       # wdWrapTopBottom
        }
        wrap_val = None
        if wrapping is not None:
            wrap_val = WRAP_STYLES.get(wrapping.lower())
            if wrapping.lower() != "inline" and wrap_val is None:
                return json.dumps({"error": f"Unknown wrapping: {wrapping}. Use: {list(WRAP_STYLES.keys())}"})

        # Border style constants
        BORDER_STYLES = {
            "none": 0,     # wdLineStyleNone
            "single": 1,   # wdLineStyleSingle
            "double": 7,   # wdLineStyleDouble
            "dotted": 3,   # wdLineStyleDot
            "dashed": 2,   # wdLineStyleDash
            "thick": 6,    # wdLineStyleThickThinSmallGap
        }

        # Alignment map
        ALIGN_MAP = {"left": 0, "center": 1, "right": 2}

        with undo_record(app, "MCP: Insert Image"):
            inline_shape = rng.InlineShapes.AddPicture(
                FileName=abs_path,
                LinkToFile=link_to_file,
                SaveWithDocument=not link_to_file,
            )

            # Resize if requested (preserves aspect ratio if only one dimension given)
            if final_w is not None and final_h is not None:
                inline_shape.Width = final_w
                inline_shape.Height = final_h
            elif final_w is not None:
                original_ratio = inline_shape.Height / inline_shape.Width
                inline_shape.Width = final_w
                inline_shape.Height = final_w * original_ratio
            elif final_h is not None:
                original_ratio = inline_shape.Width / inline_shape.Height
                inline_shape.Height = final_h
                inline_shape.Width = final_h * original_ratio

            result_width = inline_shape.Width
            result_height = inline_shape.Height
            result_wrapping = "inline"

            # Convert to floating Shape for non-inline wrapping
            if wrap_val is not None:
                float_shape = inline_shape.ConvertToShape()
                float_shape.WrapFormat.Type = wrap_val
                result_wrapping = wrapping.lower()
                result_width = float_shape.Width
                result_height = float_shape.Height

                # Apply border to floating shape
                if border_style is not None:
                    b_style = BORDER_STYLES.get(border_style.lower())
                    if b_style is None:
                        return json.dumps({"error": f"Unknown border_style: {border_style}. Use: {list(BORDER_STYLES.keys())}"})
                    b_width = float(border_width_pt) if border_width_pt else 1.0
                    # Parse border color
                    b_color = 0  # black
                    if border_color:
                        bc = border_color.lstrip("#")
                        rr, gg, bb = int(bc[0:2], 16), int(bc[2:4], 16), int(bc[4:6], 16)
                        b_color = bb * 65536 + gg * 256 + rr  # Word BGR
                    line = float_shape.Line
                    if b_style == 0:  # none
                        line.Visible = False
                    else:
                        line.Visible = True
                        DASH_MAP = {"single": 1, "double": 1, "dotted": 3, "dashed": 4, "thick": 1}
                        line.DashStyle = DASH_MAP.get(border_style.lower(), 1)
                        line.Weight = b_width
                        line.ForeColor.RGB = b_color
                        if border_style.lower() == "double":
                            line.Style = 3  # msoLineThinThin

                # Apply alignment for floating shape using relative positioning
                if alignment is not None:
                    al = alignment.lower()
                    if al in ALIGN_MAP:
                        # Use margin-relative positioning
                        float_shape.RelativeHorizontalPosition = 0  # wdRelativeHorizontalPositionMargin
                        float_shape.RelativeVerticalPosition = 2    # wdRelativeVerticalPositionParagraph
                        page_setup = doc.PageSetup
                        text_width = page_setup.PageWidth - page_setup.LeftMargin - page_setup.RightMargin
                        if al == "left":
                            float_shape.Left = 0
                        elif al == "right":
                            float_shape.Left = max(0, text_width - float_shape.Width)
                        else:  # center
                            float_shape.Left = max(0, (text_width - float_shape.Width) / 2)
            else:
                # Inline shape: apply border via inline shape borders
                if border_style is not None:
                    b_style = BORDER_STYLES.get(border_style.lower())
                    if b_style is None:
                        return json.dumps({"error": f"Unknown border_style: {border_style}. Use: {list(BORDER_STYLES.keys())}"})
                    b_width = float(border_width_pt) if border_width_pt else 1.0
                    b_color = 0  # black
                    if border_color:
                        bc = border_color.lstrip("#")
                        rr, gg, bb = int(bc[0:2], 16), int(bc[2:4], 16), int(bc[4:6], 16)
                        b_color = bb * 65536 + gg * 256 + rr
                    # Apply to all 4 borders of inline shape
                    for bid in [-1, -2, -3, -4]:  # top, left, bottom, right
                        try:
                            border = inline_shape.Borders(bid)
                            border.LineStyle = b_style
                            if b_style != 0:
                                border.LineWidth = b_width
                                border.Color = b_color
                        except Exception:
                            pass

                # Apply alignment for inline shape (set paragraph alignment)
                if alignment is not None:
                    al = ALIGN_MAP.get(alignment.lower())
                    if al is not None:
                        inline_shape.Range.ParagraphFormat.Alignment = al

        return json.dumps({
            "success": True,
            "document": doc.Name,
            "image": os.path.basename(abs_path),
            "width_pt": result_width,
            "height_pt": result_height,
            "alignment": alignment or "unchanged",
            "wrapping": result_wrapping,
            "border": border_style or "none",
            "linked": link_to_file,
        }, ensure_ascii=False)

    except Exception as e:
        return json.dumps({"error": str(e)})


async def word_live_insert_cross_reference(
    filename: str = None,
    ref_type: str = "heading",
    ref_item: int = 1,
    ref_kind: str = "text",
    insert_position: str = "end",
    paragraph_index: int = None,
    insert_as_hyperlink: bool = True,
) -> str:
    """Insert a cross-reference to a heading, bookmark, figure, or table.

    Cross-references are live fields that update automatically (e.g., "see Section 2.1").

    Args:
        filename: Document name or path (None = active document).
        ref_type: Type of item to reference: "heading", "bookmark", "figure",
                  "table", "equation", "footnote", "endnote".
        ref_item: 1-indexed item number within that reference type.
        ref_kind: What to display: "text" (full text), "number" (label+number),
                  "number_no_context" (just number), "page" (page number),
                  "above_below" ("above" or "below").
        insert_position: "start", "end", or character offset. Used if paragraph_index is None.
        paragraph_index: Insert at the start of this 1-indexed paragraph.
        insert_as_hyperlink: If True, the reference is a clickable hyperlink.

    Returns:
        JSON with cross-reference result.
    """
    if sys.platform != "win32":
        return json.dumps({"error": "Live editing is only available on Windows"})

    # Map ref_type to Word constants (wdRefType)
    ref_type_map = {
        "heading": 1,        # wdRefTypeHeading
        "bookmark": 2,       # wdRefTypeBookmark
        "footnote": 3,       # wdRefTypeFootnote
        "endnote": 4,        # wdRefTypeEndnote
        "figure": 10,        # wdRefTypeFigure (SEQ Figure)
        "table": 11,         # wdRefTypeTable (SEQ Table)
        "equation": 12,      # wdRefTypeEquation
    }

    # Map ref_kind to Word constants (wdReferenceKind)
    ref_kind_map = {
        "text": 0,                 # wdContentText
        "number": 1,               # wdNumberFullContext
        "number_no_context": 2,    # wdNumberNoContext
        "number_relative": 3,      # wdNumberRelativeContext
        "page": 7,                 # wdPageNumber
        "above_below": 6,          # wdAboveBelow
    }

    ref_type_lower = ref_type.lower()
    if ref_type_lower not in ref_type_map:
        return json.dumps({
            "error": f"Invalid ref_type '{ref_type}'. Use: {', '.join(ref_type_map.keys())}"
        })

    ref_kind_lower = ref_kind.lower()
    if ref_kind_lower not in ref_kind_map:
        return json.dumps({
            "error": f"Invalid ref_kind '{ref_kind}'. Use: {', '.join(ref_kind_map.keys())}"
        })

    try:
        from word_mcp.word_com import get_word_app, find_document, undo_record

        app = get_word_app()
        doc = find_document(app, filename)

        # Move selection to insertion point
        if paragraph_index is not None:
            if paragraph_index < 1 or paragraph_index > doc.Paragraphs.Count:
                return json.dumps({
                    "error": f"paragraph_index {paragraph_index} out of range (1-{doc.Paragraphs.Count})"
                })
            rng = doc.Paragraphs(paragraph_index).Range
            rng.Collapse(1)  # wdCollapseStart
        elif insert_position == "start":
            rng = doc.Range(0, 0)
        elif insert_position == "end":
            rng = doc.Range()
            rng.Collapse(0)  # wdCollapseEnd
        else:
            try:
                offset = int(insert_position)
                rng = doc.Range(offset, offset)
            except (ValueError, TypeError):
                rng = doc.Range()
                rng.Collapse(0)

        rng.Select()

        with undo_record(app, "MCP: Insert Cross Reference"):
            app.Selection.InsertCrossReference(
                ReferenceType=ref_type_map[ref_type_lower],
                ReferenceKind=ref_kind_map[ref_kind_lower],
                ReferenceItem=ref_item,
                InsertAsHyperlink=insert_as_hyperlink,
            )

        return json.dumps({
            "success": True,
            "document": doc.Name,
            "ref_type": ref_type,
            "ref_item": ref_item,
            "ref_kind": ref_kind,
            "as_hyperlink": insert_as_hyperlink,
        }, ensure_ascii=False)

    except Exception as e:
        return json.dumps({"error": str(e)})


async def word_live_list_cross_reference_items(
    filename: str = None,
    ref_type: str = "heading",
) -> str:
    """List all available cross-reference targets of a given type.

    Use this to discover which headings, bookmarks, figures, etc. can be
    referenced, and their 1-based index for use with word_live_insert_cross_reference.

    Args:
        filename: Document name or path (None = active document).
        ref_type: Type to list: "heading", "bookmark", "figure", "table", "equation",
                  "footnote", "endnote".

    Returns:
        JSON with list of referenceable items and their indices.
    """
    if sys.platform != "win32":
        return json.dumps({"error": "Live editing is only available on Windows"})

    valid_types = {"heading", "bookmark", "footnote", "endnote", "figure", "table", "equation"}
    ref_type_lower = ref_type.lower()
    if ref_type_lower not in valid_types:
        return json.dumps({
            "error": f"Invalid ref_type '{ref_type}'. Use: {', '.join(sorted(valid_types))}"
        })

    try:
        from word_mcp.word_com import get_word_app, find_document

        app = get_word_app()
        doc = find_document(app, filename)

        result = []

        if ref_type_lower == "heading":
            idx = 1
            for i in range(1, doc.Paragraphs.Count + 1):
                p = doc.Paragraphs(i)
                style_name = p.Style.NameLocal
                if style_name.startswith("Heading"):
                    text = p.Range.Text.strip()
                    if text:
                        result.append({
                            "index": idx,
                            "text": text,
                            "style": style_name,
                            "paragraph": i,
                        })
                        idx += 1

        elif ref_type_lower == "bookmark":
            for i in range(1, doc.Bookmarks.Count + 1):
                bm = doc.Bookmarks(i)
                text = bm.Range.Text.strip()[:100] if bm.Range else ""
                result.append({
                    "index": i,
                    "name": bm.Name,
                    "text": text,
                })

        elif ref_type_lower == "footnote":
            for i in range(1, doc.Footnotes.Count + 1):
                fn = doc.Footnotes(i)
                text = fn.Range.Text.strip()[:100]
                result.append({
                    "index": i,
                    "text": text,
                })

        elif ref_type_lower == "endnote":
            for i in range(1, doc.Endnotes.Count + 1):
                en = doc.Endnotes(i)
                text = en.Range.Text.strip()[:100]
                result.append({
                    "index": i,
                    "text": text,
                })

        elif ref_type_lower in ("figure", "table", "equation"):
            # Scan for captioned items (SEQ fields)
            seq_label = {"figure": "Figure", "table": "Table", "equation": "Equation"}[ref_type_lower]
            idx = 1
            for i in range(1, doc.Paragraphs.Count + 1):
                p = doc.Paragraphs(i)
                text = p.Range.Text.strip()
                if text.startswith(seq_label):
                    result.append({
                        "index": idx,
                        "text": text[:100],
                        "paragraph": i,
                    })
                    idx += 1

        return json.dumps({
            "success": True,
            "document": doc.Name,
            "ref_type": ref_type,
            "items": result,
            "count": len(result),
        }, ensure_ascii=False)

    except Exception as e:
        return json.dumps({"error": str(e)})


async def word_live_insert_equation(
    filename: str = None,
    equation: str = "",
    paragraph_index: int = None,
    position: str = "end",
    display_mode: bool = False,
) -> str:
    """Insert a mathematical equation into a Word document using UnicodeMath syntax.

    LaTeX-like commands (e.g. \\int, \\sum, \\alpha) are automatically converted to
    Unicode math symbols before insertion, ensuring proper rendering.

    Args:
        filename: Document name (uses active document if None).
        equation: Equation text in UnicodeMath syntax. Examples:
            Simple: "x^2 + y^2 = z^2", "E = mc^2"
            Fractions: "(a+b)/(c+d)"
            Square root: "\\sqrt(x^2+y^2)"
            Greek letters: "\\alpha + \\beta = \\gamma"
            Integrals: "\\int_0^\\infty e^(-x^2) dx"
            Summation: "\\sum_(i=1)^n i^2"
            Matrix: "\\matrix(a&b@c&d)"
            Taylor series: "f(x) = \\sum_(n=0)^\\infty (f^((n))(a))/(n!) (x-a)^n"
        paragraph_index: Insert after this paragraph (1-based). None = use position.
        position: "start" or "end" of document. Ignored if paragraph_index given.
        display_mode: If True, equation is centered on its own line (display style).
            If False, equation is inline with surrounding text.

    Returns:
        JSON with success status and equation details.
    """
    # LaTeX-like command to Unicode math symbol mapping.
    # Word's COM OMaths.Add + BuildUp doesn't process autocorrect entries,
    # so we must pre-convert commands like \int, \sum to their Unicode equivalents.
    UNICODE_MATH = {
        # Greek lowercase
        r"\alpha": "\u03B1", r"\beta": "\u03B2", r"\gamma": "\u03B3",
        r"\delta": "\u03B4", r"\epsilon": "\u03B5", r"\varepsilon": "\u03B5",
        r"\zeta": "\u03B6", r"\eta": "\u03B7", r"\theta": "\u03B8",
        r"\vartheta": "\u03D1", r"\iota": "\u03B9", r"\kappa": "\u03BA",
        r"\lambda": "\u03BB", r"\mu": "\u03BC", r"\nu": "\u03BD",
        r"\xi": "\u03BE", r"\pi": "\u03C0", r"\rho": "\u03C1",
        r"\sigma": "\u03C3", r"\varsigma": "\u03C2", r"\tau": "\u03C4",
        r"\upsilon": "\u03C5", r"\phi": "\u03C6", r"\varphi": "\u03D5",
        r"\chi": "\u03C7", r"\psi": "\u03C8", r"\omega": "\u03C9",
        # Greek uppercase
        r"\Gamma": "\u0393", r"\Delta": "\u0394", r"\Theta": "\u0398",
        r"\Lambda": "\u039B", r"\Xi": "\u039E", r"\Pi": "\u03A0",
        r"\Sigma": "\u03A3", r"\Upsilon": "\u03A5", r"\Phi": "\u03A6",
        r"\Psi": "\u03A8", r"\Omega": "\u03A9",
        # Operators / big operators
        r"\int": "\u222B", r"\iint": "\u222C", r"\iiint": "\u222D",
        r"\oint": "\u222E", r"\sum": "\u2211", r"\prod": "\u220F",
        r"\coprod": "\u2210",
        # Roots and radicals
        r"\sqrt": "\u221A", r"\cbrt": "\u221B",
        # Calculus / analysis
        r"\partial": "\u2202", r"\nabla": "\u2207",
        r"\infty": "\u221E",
        # Logic / set theory
        r"\forall": "\u2200", r"\exists": "\u2203", r"\nexists": "\u2204",
        r"\in": "\u2208", r"\notin": "\u2209",
        r"\subset": "\u2282", r"\supset": "\u2283",
        r"\subseteq": "\u2286", r"\supseteq": "\u2287",
        r"\cup": "\u222A", r"\cap": "\u2229",
        r"\emptyset": "\u2205",
        r"\neg": "\u00AC", r"\land": "\u2227", r"\lor": "\u2228",
        # Arithmetic / relations
        r"\pm": "\u00B1", r"\mp": "\u2213",
        r"\times": "\u00D7", r"\div": "\u00F7", r"\cdot": "\u22C5",
        r"\leq": "\u2264", r"\geq": "\u2265", r"\neq": "\u2260",
        r"\approx": "\u2248", r"\equiv": "\u2261", r"\cong": "\u2245",
        r"\sim": "\u223C", r"\propto": "\u221D",
        r"\ll": "\u226A", r"\gg": "\u226B",
        # Arrows
        r"\rightarrow": "\u2192", r"\leftarrow": "\u2190",
        r"\leftrightarrow": "\u2194",
        r"\Rightarrow": "\u21D2", r"\Leftarrow": "\u21D0",
        r"\Leftrightarrow": "\u21D4",
        r"\uparrow": "\u2191", r"\downarrow": "\u2193",
        r"\mapsto": "\u21A6",
        # Dots
        r"\cdots": "\u22EF", r"\ldots": "\u2026", r"\vdots": "\u22EE",
        r"\ddots": "\u22F1",
        # Miscellaneous
        r"\angle": "\u2220", r"\degree": "\u00B0",
        r"\star": "\u22C6", r"\circ": "\u2218",
        r"\bullet": "\u2022", r"\diamond": "\u22C4",
        r"\triangle": "\u25B3",
        r"\hbar": "\u210F", r"\ell": "\u2113",
        r"\Re": "\u211C", r"\Im": "\u2124",
        r"\aleph": "\u2135",
        # Matrix (Word UnicodeMath uses ■ for matrix)
        r"\matrix": "\u25A0", r"\pmatrix": "\u25A0",
        # Function names (these stay as text but without backslash)
        r"\lim": "lim", r"\sin": "sin", r"\cos": "cos", r"\tan": "tan",
        r"\sec": "sec", r"\csc": "csc", r"\cot": "cot",
        r"\arcsin": "arcsin", r"\arccos": "arccos", r"\arctan": "arctan",
        r"\sinh": "sinh", r"\cosh": "cosh", r"\tanh": "tanh",
        r"\log": "log", r"\ln": "ln", r"\exp": "exp",
        r"\det": "det", r"\dim": "dim", r"\ker": "ker",
        r"\min": "min", r"\max": "max", r"\inf": "inf", r"\sup": "sup",
        r"\gcd": "gcd", r"\arg": "arg", r"\mod": "mod",
    }
    if sys.platform != "win32":
        return json.dumps({"error": "Live editing is only available on Windows"})

    try:
        from word_mcp.word_com import get_word_app, find_document, undo_record

        app = get_word_app()
        doc = find_document(app, filename)

        if not equation or not equation.strip():
            return json.dumps({"error": "equation text is required"})

        with undo_record(app, "MCP: Insert Equation"):
            # Determine insertion range
            if paragraph_index is not None:
                if paragraph_index < 1 or paragraph_index > doc.Paragraphs.Count:
                    return json.dumps({
                        "error": f"paragraph_index {paragraph_index} out of range (1-{doc.Paragraphs.Count})"
                    })
                rng = doc.Paragraphs(paragraph_index).Range
                rng.Collapse(0)  # After the paragraph
                rng.InsertParagraphAfter()
                rng.Collapse(0)
            elif position == "start":
                rng = doc.Paragraphs(1).Range
                rng.Collapse(1)  # Before first paragraph
                rng.InsertParagraphBefore()
                rng = doc.Paragraphs(1).Range
                rng.Collapse(1)
            else:  # "end"
                rng = doc.Content
                rng.Collapse(0)  # After last content
                rng.InsertParagraphAfter()
                rng.Collapse(0)

            # Convert LaTeX-like commands to Unicode math symbols.
            # Sort by length descending so longer matches take priority
            # (e.g. \iint before \int, \infty before \in).
            # Use negative lookahead (?![a-zA-Z]) to avoid partial matches.
            _commands = sorted(UNICODE_MATH.keys(), key=len, reverse=True)
            _pattern = '|'.join(re.escape(c) for c in _commands)
            _pattern = f'({_pattern})(?![a-zA-Z])'
            eq_text = re.sub(_pattern, lambda m: UNICODE_MATH[m.group(1)], equation)

            # Insert the converted equation text
            rng.Text = eq_text

            # Convert to OMath
            doc.OMaths.Add(rng)
            omath = doc.OMaths(doc.OMaths.Count)

            # Set display mode (centered on own line) vs inline
            if display_mode:
                omath.Type = 1  # wdOMathDisplay
            else:
                omath.Type = 0  # wdOMathInline

            # Build up the equation (render UnicodeMath to formatted equation)
            omath.BuildUp()

        return json.dumps({
            "success": True,
            "document": doc.Name,
            "equation": equation,
            "display_mode": display_mode,
            "omath_count": doc.OMaths.Count,
        }, ensure_ascii=False)

    except Exception as e:
        return json.dumps({"error": str(e)})
