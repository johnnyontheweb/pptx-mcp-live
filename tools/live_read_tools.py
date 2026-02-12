"""
Live read-only tools using PowerPoint COM automation (Windows only).
These work even when the file is locked/open in PowerPoint.
"""
import json
import sys
from typing import Optional


def register_live_read_tools(app):
    """Register 7 read-only COM tools with the FastMCP app."""

    # ------------------------------------------------------------------
    # 1. ppt_live_get_info
    # ------------------------------------------------------------------
    @app.tool()
    async def ppt_live_get_info(filename: Optional[str] = None) -> str:
        """[Windows only] Get metadata for an open presentation: name, path, slide count, saved state, dimensions, built-in properties.

        Args:
            filename: Presentation filename (basename). None = active presentation.
        """
        if sys.platform != "win32":
            return json.dumps({"error": "Live tools are only available on Windows"})
        try:
            from core.ppt_com import get_ppt_app, find_presentation
            ppt = get_ppt_app()
            pres = find_presentation(ppt, filename)

            props = {}
            try:
                bp = pres.BuiltInDocumentProperties
                for name in ("Title", "Subject", "Author", "Keywords", "Comments",
                             "Last Author", "Company", "Category"):
                    try:
                        props[name] = str(bp(name).Value)
                    except Exception:
                        pass
            except Exception:
                pass

            return json.dumps({
                "name": pres.Name,
                "full_name": pres.FullName,
                "path": pres.Path,
                "slide_count": pres.Slides.Count,
                "saved": bool(pres.Saved),
                "slide_width_pt": float(pres.PageSetup.SlideWidth),
                "slide_height_pt": float(pres.PageSetup.SlideHeight),
                "properties": props,
            }, ensure_ascii=False)
        except Exception as e:
            return json.dumps({"error": str(e)})

    # ------------------------------------------------------------------
    # 2. ppt_live_get_text
    # ------------------------------------------------------------------
    @app.tool()
    async def ppt_live_get_text(filename: Optional[str] = None) -> str:
        """[Windows only] Extract all text and speaker notes from every slide.

        Args:
            filename: Presentation filename (basename). None = active presentation.
        """
        if sys.platform != "win32":
            return json.dumps({"error": "Live tools are only available on Windows"})
        try:
            from core.ppt_com import get_ppt_app, find_presentation
            ppt = get_ppt_app()
            pres = find_presentation(ppt, filename)

            slides = []
            for si in range(1, pres.Slides.Count + 1):
                slide = pres.Slides(si)
                texts = []
                for sh_i in range(1, slide.Shapes.Count + 1):
                    shape = slide.Shapes(sh_i)
                    if shape.HasTextFrame:
                        txt = shape.TextFrame.TextRange.Text
                        if txt.strip():
                            texts.append({"shape": shape.Name, "text": txt})

                notes = ""
                try:
                    notes = slide.NotesPage.Shapes.Placeholders(2).TextFrame.TextRange.Text
                except Exception:
                    pass

                slides.append({
                    "slide_index": si,
                    "texts": texts,
                    "notes": notes,
                })

            return json.dumps({"slide_count": pres.Slides.Count, "slides": slides},
                              ensure_ascii=False)
        except Exception as e:
            return json.dumps({"error": str(e)})

    # ------------------------------------------------------------------
    # 3. ppt_live_get_slide_info
    # ------------------------------------------------------------------
    @app.tool()
    async def ppt_live_get_slide_info(slide_index: int, filename: Optional[str] = None) -> str:
        """[Windows only] Get detailed shape information for a specific slide.

        Args:
            slide_index: 1-based slide index.
            filename: Presentation filename (basename). None = active presentation.
        """
        if sys.platform != "win32":
            return json.dumps({"error": "Live tools are only available on Windows"})
        try:
            from core.ppt_com import get_ppt_app, find_presentation
            ppt = get_ppt_app()
            pres = find_presentation(ppt, filename)

            if slide_index < 1 or slide_index > pres.Slides.Count:
                return json.dumps({"error": f"slide_index {slide_index} out of range (1-{pres.Slides.Count})"})

            slide = pres.Slides(slide_index)
            shapes = []
            for i in range(1, slide.Shapes.Count + 1):
                s = slide.Shapes(i)
                info = {
                    "index": i,
                    "name": s.Name,
                    "type": int(s.Type),
                    "left": float(s.Left),
                    "top": float(s.Top),
                    "width": float(s.Width),
                    "height": float(s.Height),
                }
                if s.HasTextFrame:
                    info["text"] = s.TextFrame.TextRange.Text
                if s.HasTable:
                    info["has_table"] = True
                    info["table_rows"] = s.Table.Rows.Count
                    info["table_columns"] = s.Table.Columns.Count
                shapes.append(info)

            return json.dumps({
                "slide_index": slide_index,
                "shape_count": slide.Shapes.Count,
                "shapes": shapes,
            }, ensure_ascii=False)
        except Exception as e:
            return json.dumps({"error": str(e)})

    # ------------------------------------------------------------------
    # 4. ppt_live_find_text
    # ------------------------------------------------------------------
    @app.tool()
    async def ppt_live_find_text(
        search_text: str,
        filename: Optional[str] = None,
        match_case: bool = False,
    ) -> str:
        """[Windows only] Search for text across all slides. Returns slide/shape/context for each match.

        Args:
            search_text: Text to search for.
            filename: Presentation filename (basename). None = active presentation.
            match_case: Case-sensitive search.
        """
        if sys.platform != "win32":
            return json.dumps({"error": "Live tools are only available on Windows"})
        try:
            from core.ppt_com import get_ppt_app, find_presentation
            ppt = get_ppt_app()
            pres = find_presentation(ppt, filename)

            results = []
            for si in range(1, pres.Slides.Count + 1):
                slide = pres.Slides(si)
                for sh_i in range(1, slide.Shapes.Count + 1):
                    shape = slide.Shapes(sh_i)
                    if not shape.HasTextFrame:
                        continue
                    full = shape.TextFrame.TextRange.Text
                    compare_full = full if match_case else full.lower()
                    compare_search = search_text if match_case else search_text.lower()
                    if compare_search in compare_full:
                        results.append({
                            "slide_index": si,
                            "shape_index": sh_i,
                            "shape_name": shape.Name,
                            "context": full[:200],
                        })

            return json.dumps({"search_text": search_text, "match_count": len(results),
                               "results": results}, ensure_ascii=False)
        except Exception as e:
            return json.dumps({"error": str(e)})

    # ------------------------------------------------------------------
    # 5. ppt_live_get_notes
    # ------------------------------------------------------------------
    @app.tool()
    async def ppt_live_get_notes(
        filename: Optional[str] = None,
        slide_index: Optional[int] = None,
    ) -> str:
        """[Windows only] Get speaker notes. If slide_index is None, returns notes for all slides.

        Args:
            filename: Presentation filename (basename). None = active presentation.
            slide_index: 1-based slide index. None = all slides.
        """
        if sys.platform != "win32":
            return json.dumps({"error": "Live tools are only available on Windows"})
        try:
            from core.ppt_com import get_ppt_app, find_presentation
            ppt = get_ppt_app()
            pres = find_presentation(ppt, filename)

            def _get_notes(slide):
                try:
                    return slide.NotesPage.Shapes.Placeholders(2).TextFrame.TextRange.Text
                except Exception:
                    return ""

            if slide_index is not None:
                if slide_index < 1 or slide_index > pres.Slides.Count:
                    return json.dumps({"error": f"slide_index {slide_index} out of range (1-{pres.Slides.Count})"})
                return json.dumps({"slide_index": slide_index,
                                   "notes": _get_notes(pres.Slides(slide_index))},
                                  ensure_ascii=False)

            notes = []
            for si in range(1, pres.Slides.Count + 1):
                n = _get_notes(pres.Slides(si))
                if n.strip():
                    notes.append({"slide_index": si, "notes": n})
            return json.dumps({"slides_with_notes": len(notes), "notes": notes},
                              ensure_ascii=False)
        except Exception as e:
            return json.dumps({"error": str(e)})

    # ------------------------------------------------------------------
    # 6. ppt_live_list_comments
    # ------------------------------------------------------------------
    @app.tool()
    async def ppt_live_list_comments(filename: Optional[str] = None) -> str:
        """[Windows only] List all comments across all slides.

        Args:
            filename: Presentation filename (basename). None = active presentation.
        """
        if sys.platform != "win32":
            return json.dumps({"error": "Live tools are only available on Windows"})
        try:
            from core.ppt_com import get_ppt_app, find_presentation
            ppt = get_ppt_app()
            pres = find_presentation(ppt, filename)

            comments = []
            for si in range(1, pres.Slides.Count + 1):
                slide = pres.Slides(si)
                try:
                    for ci in range(1, slide.Comments.Count + 1):
                        c = slide.Comments(ci)
                        comments.append({
                            "slide_index": si,
                            "comment_index": ci,
                            "author": c.Author,
                            "author_initials": c.AuthorInitials,
                            "text": c.Text,
                            "date_time": str(c.DateTime),
                        })
                except Exception:
                    pass

            return json.dumps({"comment_count": len(comments), "comments": comments},
                              ensure_ascii=False)
        except Exception as e:
            return json.dumps({"error": str(e)})

    # ------------------------------------------------------------------
    # 7. ppt_live_list_open
    # ------------------------------------------------------------------
    @app.tool()
    async def ppt_live_list_open() -> str:
        """[Windows only] List all presentations currently open in PowerPoint."""
        if sys.platform != "win32":
            return json.dumps({"error": "Live tools are only available on Windows"})
        try:
            from core.ppt_com import get_ppt_app
            ppt = get_ppt_app()

            presos = []
            for i in range(1, ppt.Presentations.Count + 1):
                p = ppt.Presentations(i)
                presos.append({
                    "index": i,
                    "name": p.Name,
                    "full_name": p.FullName,
                    "slide_count": p.Slides.Count,
                    "saved": bool(p.Saved),
                })

            active_name = ""
            try:
                active_name = ppt.ActivePresentation.Name
            except Exception:
                pass

            return json.dumps({
                "count": len(presos),
                "active": active_name,
                "presentations": presos,
            }, ensure_ascii=False)
        except Exception as e:
            return json.dumps({"error": str(e)})
