"""
Live editing tools using PowerPoint COM automation (Windows only).
These work even when the file is locked/open in PowerPoint.
"""
import json
import sys
from typing import Optional


def register_live_edit_tools(app):
    """Register 8 edit tools with the FastMCP app."""

    # ------------------------------------------------------------------
    # 8. ppt_live_insert_text
    # ------------------------------------------------------------------
    @app.tool()
    async def ppt_live_insert_text(
        slide_index: int,
        text: str,
        filename: Optional[str] = None,
        shape_index: Optional[int] = None,
        shape_name: Optional[str] = None,
        position: str = "end",
    ) -> str:
        """[Windows only] Insert text into a shape on a slide.

        Args:
            slide_index: 1-based slide index.
            text: Text to insert.
            filename: Presentation filename (basename). None = active presentation.
            shape_index: 1-based shape index. Provide either this or shape_name.
            shape_name: Shape name. Provide either this or shape_index.
            position: Where to insert: "end" (append), "start" (prepend), or "replace" (overwrite).
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

            # Find shape
            shape = None
            if shape_name:
                for i in range(1, slide.Shapes.Count + 1):
                    if slide.Shapes(i).Name == shape_name:
                        shape = slide.Shapes(i)
                        break
                if shape is None:
                    names = [slide.Shapes(i).Name for i in range(1, slide.Shapes.Count + 1)]
                    return json.dumps({"error": f"Shape '{shape_name}' not found. Available: {names}"})
            elif shape_index:
                if shape_index < 1 or shape_index > slide.Shapes.Count:
                    return json.dumps({"error": f"shape_index {shape_index} out of range (1-{slide.Shapes.Count})"})
                shape = slide.Shapes(shape_index)
            else:
                return json.dumps({"error": "Provide either shape_index or shape_name"})

            if not shape.HasTextFrame:
                return json.dumps({"error": f"Shape '{shape.Name}' has no text frame"})

            tr = shape.TextFrame.TextRange
            if position == "replace":
                tr.Text = text
            elif position == "start":
                tr.InsertBefore(text)
            else:  # end
                tr.InsertAfter(text)

            return json.dumps({
                "slide_index": slide_index,
                "shape_name": shape.Name,
                "position": position,
                "new_text": shape.TextFrame.TextRange.Text[:200],
            }, ensure_ascii=False)
        except Exception as e:
            return json.dumps({"error": str(e)})

    # ------------------------------------------------------------------
    # 9. ppt_live_replace_text
    # ------------------------------------------------------------------
    @app.tool()
    async def ppt_live_replace_text(
        old_text: str,
        new_text: str,
        filename: Optional[str] = None,
        slide_index: Optional[int] = None,
        match_case: bool = False,
    ) -> str:
        """[Windows only] Find and replace text across slides.

        Args:
            old_text: Text to find.
            new_text: Replacement text.
            filename: Presentation filename (basename). None = active presentation.
            slide_index: 1-based slide index. None = all slides.
            match_case: Case-sensitive matching.
        """
        if sys.platform != "win32":
            return json.dumps({"error": "Live tools are only available on Windows"})
        try:
            from core.ppt_com import get_ppt_app, find_presentation
            ppt = get_ppt_app()
            pres = find_presentation(ppt, filename)

            replacements = 0
            start = slide_index if slide_index else 1
            end = (slide_index if slide_index else pres.Slides.Count) + 1

            for si in range(start, end):
                if si < 1 or si > pres.Slides.Count:
                    continue
                slide = pres.Slides(si)
                for sh_i in range(1, slide.Shapes.Count + 1):
                    shape = slide.Shapes(sh_i)
                    if not shape.HasTextFrame:
                        continue
                    full = shape.TextFrame.TextRange.Text
                    if match_case:
                        if old_text in full:
                            shape.TextFrame.TextRange.Text = full.replace(old_text, new_text)
                            replacements += full.count(old_text)
                    else:
                        lower_full = full.lower()
                        lower_old = old_text.lower()
                        if lower_old in lower_full:
                            # Case-insensitive replace
                            import re
                            new_val = re.sub(re.escape(old_text), new_text, full, flags=re.IGNORECASE)
                            count = len(re.findall(re.escape(old_text), full, re.IGNORECASE))
                            shape.TextFrame.TextRange.Text = new_val
                            replacements += count

            return json.dumps({
                "old_text": old_text,
                "new_text": new_text,
                "replacements": replacements,
            }, ensure_ascii=False)
        except Exception as e:
            return json.dumps({"error": str(e)})

    # ------------------------------------------------------------------
    # 10. ppt_live_add_slide
    # ------------------------------------------------------------------
    @app.tool()
    async def ppt_live_add_slide(
        filename: Optional[str] = None,
        layout_index: int = 12,
        position: Optional[int] = None,
        title: Optional[str] = None,
    ) -> str:
        """[Windows only] Add a new slide.

        Args:
            filename: Presentation filename (basename). None = active presentation.
            layout_index: PowerPoint layout enum (1=Title, 2=Title+Content, 11=TitleOnly, 12=Blank).
            position: 1-based insert position. None = append at end.
            title: Optional title text to set on the new slide.
        """
        if sys.platform != "win32":
            return json.dumps({"error": "Live tools are only available on Windows"})
        try:
            from core.ppt_com import get_ppt_app, find_presentation
            ppt = get_ppt_app()
            pres = find_presentation(ppt, filename)

            pos = position if position else pres.Slides.Count + 1
            new_slide = pres.Slides.Add(pos, layout_index)

            if title and new_slide.Shapes.HasTitle:
                new_slide.Shapes.Title.TextFrame.TextRange.Text = title

            return json.dumps({
                "slide_index": new_slide.SlideIndex,
                "layout_index": layout_index,
                "title": title,
                "total_slides": pres.Slides.Count,
            }, ensure_ascii=False)
        except Exception as e:
            return json.dumps({"error": str(e)})

    # ------------------------------------------------------------------
    # 11. ppt_live_delete_slide
    # ------------------------------------------------------------------
    @app.tool()
    async def ppt_live_delete_slide(
        slide_index: int,
        filename: Optional[str] = None,
    ) -> str:
        """[Windows only] Delete a slide.

        Args:
            slide_index: 1-based slide index to delete.
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

            pres.Slides(slide_index).Delete()

            return json.dumps({
                "deleted_slide_index": slide_index,
                "remaining_slides": pres.Slides.Count,
            }, ensure_ascii=False)
        except Exception as e:
            return json.dumps({"error": str(e)})

    # ------------------------------------------------------------------
    # 12. ppt_live_set_notes
    # ------------------------------------------------------------------
    @app.tool()
    async def ppt_live_set_notes(
        slide_index: int,
        notes_text: str,
        filename: Optional[str] = None,
    ) -> str:
        """[Windows only] Set speaker notes for a slide.

        Args:
            slide_index: 1-based slide index.
            notes_text: Notes text to set.
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
            slide.NotesPage.Shapes.Placeholders(2).TextFrame.TextRange.Text = notes_text

            return json.dumps({
                "slide_index": slide_index,
                "notes_set": True,
            }, ensure_ascii=False)
        except Exception as e:
            return json.dumps({"error": str(e)})

    # ------------------------------------------------------------------
    # 13. ppt_live_add_image
    # ------------------------------------------------------------------
    @app.tool()
    async def ppt_live_add_image(
        slide_index: int,
        image_path: str,
        filename: Optional[str] = None,
        left: float = 0,
        top: float = 0,
        width: Optional[float] = None,
        height: Optional[float] = None,
    ) -> str:
        """[Windows only] Add an image to a slide.

        Args:
            slide_index: 1-based slide index.
            image_path: Absolute path to the image file.
            filename: Presentation filename (basename). None = active presentation.
            left: Left position in points.
            top: Top position in points.
            width: Width in points. None = original size.
            height: Height in points. None = original size.
        """
        if sys.platform != "win32":
            return json.dumps({"error": "Live tools are only available on Windows"})
        try:
            import os
            from core.ppt_com import get_ppt_app, find_presentation
            ppt = get_ppt_app()
            pres = find_presentation(ppt, filename)

            if slide_index < 1 or slide_index > pres.Slides.Count:
                return json.dumps({"error": f"slide_index {slide_index} out of range (1-{pres.Slides.Count})"})

            if not os.path.exists(image_path):
                return json.dumps({"error": f"Image file not found: {image_path}"})

            slide = pres.Slides(slide_index)

            # AddPicture(FileName, LinkToFile, SaveWithDocument, Left, Top, Width, Height)
            w = width if width else -1
            h = height if height else -1
            shape = slide.Shapes.AddPicture(
                os.path.abspath(image_path), False, True, left, top, w, h
            )

            return json.dumps({
                "slide_index": slide_index,
                "shape_name": shape.Name,
                "left": float(shape.Left),
                "top": float(shape.Top),
                "width": float(shape.Width),
                "height": float(shape.Height),
            }, ensure_ascii=False)
        except Exception as e:
            return json.dumps({"error": str(e)})

    # ------------------------------------------------------------------
    # 14. ppt_live_add_comment
    # ------------------------------------------------------------------
    @app.tool()
    async def ppt_live_add_comment(
        slide_index: int,
        text: str,
        filename: Optional[str] = None,
        author: str = "Av. Y\u00fcce Karapazar",
        author_initials: str = "YK",
        left: float = 0,
        top: float = 0,
    ) -> str:
        """[Windows only] Add a comment to a slide.

        Args:
            slide_index: 1-based slide index.
            text: Comment text.
            filename: Presentation filename (basename). None = active presentation.
            author: Comment author name.
            author_initials: Author initials.
            left: Comment position left (points).
            top: Comment position top (points).
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
            comment = slide.Comments.Add(left, top, author, author_initials, text)

            return json.dumps({
                "slide_index": slide_index,
                "comment_index": slide.Comments.Count,
                "author": author,
                "text": text,
            }, ensure_ascii=False)
        except Exception as e:
            return json.dumps({"error": str(e)})

    # ------------------------------------------------------------------
    # 15. ppt_live_delete_comment
    # ------------------------------------------------------------------
    @app.tool()
    async def ppt_live_delete_comment(
        slide_index: int,
        comment_index: int,
        filename: Optional[str] = None,
    ) -> str:
        """[Windows only] Delete a comment from a slide.

        Args:
            slide_index: 1-based slide index.
            comment_index: 1-based comment index on that slide.
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
            if comment_index < 1 or comment_index > slide.Comments.Count:
                return json.dumps({"error": f"comment_index {comment_index} out of range (1-{slide.Comments.Count})"})

            slide.Comments(comment_index).Delete()

            return json.dumps({
                "slide_index": slide_index,
                "deleted_comment_index": comment_index,
                "remaining_comments": slide.Comments.Count,
            }, ensure_ascii=False)
        except Exception as e:
            return json.dumps({"error": str(e)})
