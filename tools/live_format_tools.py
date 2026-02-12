"""
Live formatting tools using PowerPoint COM automation (Windows only).
These provide capabilities that python-pptx CANNOT do:
  - Picture manipulation (crop, brightness, contrast, color, transparency)
  - Shape effects (shadow, glow, reflection, soft edge, 3D, fill, line)
  - Shape arrange (group, ungroup, align, distribute, z-order, flip, duplicate)
  - Slide transitions
  - Headers/Footers
  - Section management
  - Table row/column add/delete/merge
"""
import json
import sys
from typing import Optional, List, Dict


def _hex_to_rgb_int(hex_color: str) -> int:
    """Convert '#RRGGBB' to COM RGB integer (BGR order for VBA)."""
    h = hex_color.lstrip("#")
    r, g, b = int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16)
    return r + (g << 8) + (b << 16)


def _get_shape(slide, shape_index: int):
    """Get a shape by 1-based index with validation."""
    if shape_index < 1 or shape_index > slide.Shapes.Count:
        raise ValueError(f"shape_index {shape_index} out of range (1-{slide.Shapes.Count})")
    return slide.Shapes(shape_index)


def register_live_format_tools(app):
    """Register 7 format tools with the FastMCP app."""

    # ------------------------------------------------------------------
    # 16. ppt_live_format_picture
    # ------------------------------------------------------------------
    @app.tool()
    async def ppt_live_format_picture(
        slide_index: int,
        shape_index: int,
        filename: Optional[str] = None,
        brightness: Optional[float] = None,
        contrast: Optional[float] = None,
        color_type: Optional[str] = None,
        crop_left: Optional[float] = None,
        crop_right: Optional[float] = None,
        crop_top: Optional[float] = None,
        crop_bottom: Optional[float] = None,
        transparent_background: Optional[bool] = None,
        transparency_color: Optional[str] = None,
    ) -> str:
        """[Windows only] Format a picture shape: brightness, contrast, color type, crop, transparency.
        (python-pptx CANNOT do any of these.)

        Args:
            slide_index: 1-based slide index.
            shape_index: 1-based shape index (must be a picture).
            filename: Presentation filename. None = active.
            brightness: -1.0 to 1.0 (0 = normal).
            contrast: -1.0 to 1.0 (0 = normal).
            color_type: "auto", "grayscale", "washout", "blackwhite".
            crop_left: Crop from left in points.
            crop_right: Crop from right in points.
            crop_top: Crop from top in points.
            crop_bottom: Crop from bottom in points.
            transparent_background: Set transparent background.
            transparency_color: Transparency color as "#RRGGBB".
        """
        if sys.platform != "win32":
            return json.dumps({"error": "Live tools are only available on Windows"})
        try:
            from core.ppt_com import get_ppt_app, find_presentation
            ppt = get_ppt_app()
            pres = find_presentation(ppt, filename)

            if slide_index < 1 or slide_index > pres.Slides.Count:
                return json.dumps({"error": f"slide_index {slide_index} out of range"})

            slide = pres.Slides(slide_index)
            shape = _get_shape(slide, shape_index)
            pf = shape.PictureFormat

            applied = []

            if brightness is not None:
                pf.Brightness = brightness
                applied.append(f"brightness={brightness}")
            if contrast is not None:
                pf.Contrast = contrast
                applied.append(f"contrast={contrast}")
            if color_type is not None:
                ct_map = {"auto": 1, "grayscale": 3, "washout": 4, "blackwhite": 5}
                ct = ct_map.get(color_type.lower())
                if ct is None:
                    return json.dumps({"error": f"Invalid color_type '{color_type}'. Use: auto, grayscale, washout, blackwhite"})
                pf.ColorType = ct
                applied.append(f"color_type={color_type}")
            if crop_left is not None:
                pf.CropLeft = crop_left
                applied.append(f"crop_left={crop_left}")
            if crop_right is not None:
                pf.CropRight = crop_right
                applied.append(f"crop_right={crop_right}")
            if crop_top is not None:
                pf.CropTop = crop_top
                applied.append(f"crop_top={crop_top}")
            if crop_bottom is not None:
                pf.CropBottom = crop_bottom
                applied.append(f"crop_bottom={crop_bottom}")
            if transparent_background is not None:
                # msoTrue = -1, msoFalse = 0
                pf.TransparentBackground = -1 if transparent_background else 0
                applied.append(f"transparent_background={transparent_background}")
            if transparency_color is not None:
                pf.TransparencyColor = _hex_to_rgb_int(transparency_color)
                applied.append(f"transparency_color={transparency_color}")

            return json.dumps({
                "slide_index": slide_index,
                "shape_index": shape_index,
                "shape_name": shape.Name,
                "applied": applied,
            }, ensure_ascii=False)
        except Exception as e:
            return json.dumps({"error": str(e)})

    # ------------------------------------------------------------------
    # 17. ppt_live_format_shape
    # ------------------------------------------------------------------
    @app.tool()
    async def ppt_live_format_shape(
        slide_index: int,
        shape_index: int,
        filename: Optional[str] = None,
        shadow: Optional[dict] = None,
        glow: Optional[dict] = None,
        reflection: Optional[dict] = None,
        soft_edge: Optional[dict] = None,
        fill: Optional[dict] = None,
        line: Optional[dict] = None,
        three_d: Optional[dict] = None,
    ) -> str:
        """[Windows only] Apply visual effects to a shape: shadow, glow, reflection, soft edge, 3D, fill, line.
        (python-pptx has minimal shadow only; rest is IMPOSSIBLE.)

        Args:
            slide_index: 1-based slide index.
            shape_index: 1-based shape index.
            filename: Presentation filename. None = active.
            shadow: {visible, blur, offset_x, offset_y, color("#RRGGBB"), transparency(0-1)}
            glow: {radius, color("#RRGGBB"), transparency(0-1)}
            reflection: {type(0-9 preset index)}
            soft_edge: {radius(points)}
            fill: {type("solid"/"none"), color("#RRGGBB"), transparency(0-1)}
            line: {color("#RRGGBB"), weight(points), dash_style("solid"/"dash"/"dot"/"dash_dot")}
            three_d: {rotation_x, rotation_y, rotation_z, bevel_top(int), depth, material(int), lighting(int)}
        """
        if sys.platform != "win32":
            return json.dumps({"error": "Live tools are only available on Windows"})
        try:
            from core.ppt_com import get_ppt_app, find_presentation
            ppt = get_ppt_app()
            pres = find_presentation(ppt, filename)

            if slide_index < 1 or slide_index > pres.Slides.Count:
                return json.dumps({"error": f"slide_index {slide_index} out of range"})

            slide = pres.Slides(slide_index)
            shape = _get_shape(slide, shape_index)
            applied = []

            # Shadow
            if shadow is not None:
                s = shape.Shadow
                if shadow.get("visible") is not None:
                    s.Visible = -1 if shadow["visible"] else 0  # msoTrue/msoFalse
                if shadow.get("blur") is not None:
                    s.Blur = shadow["blur"]
                if shadow.get("offset_x") is not None:
                    s.OffsetX = shadow["offset_x"]
                if shadow.get("offset_y") is not None:
                    s.OffsetY = shadow["offset_y"]
                if shadow.get("color") is not None:
                    s.ForeColor.RGB = _hex_to_rgb_int(shadow["color"])
                if shadow.get("transparency") is not None:
                    s.Transparency = shadow["transparency"]
                applied.append("shadow")

            # Glow
            if glow is not None:
                g = shape.Glow
                if glow.get("radius") is not None:
                    g.Radius = glow["radius"]
                if glow.get("color") is not None:
                    g.Color.RGB = _hex_to_rgb_int(glow["color"])
                if glow.get("transparency") is not None:
                    g.Color.Transparency = glow["transparency"]
                applied.append("glow")

            # Reflection
            if reflection is not None:
                if reflection.get("type") is not None:
                    shape.Reflection.Type = reflection["type"]
                applied.append("reflection")

            # Soft Edge
            if soft_edge is not None:
                if soft_edge.get("radius") is not None:
                    shape.SoftEdge.Radius = soft_edge["radius"]
                applied.append("soft_edge")

            # Fill
            if fill is not None:
                f = shape.Fill
                fill_type = fill.get("type", "solid")
                if fill_type == "none":
                    f.Background()
                else:
                    f.Solid()
                    if fill.get("color") is not None:
                        f.ForeColor.RGB = _hex_to_rgb_int(fill["color"])
                    if fill.get("transparency") is not None:
                        f.Transparency = fill["transparency"]
                applied.append("fill")

            # Line
            if line is not None:
                ln = shape.Line
                if line.get("color") is not None:
                    ln.ForeColor.RGB = _hex_to_rgb_int(line["color"])
                if line.get("weight") is not None:
                    ln.Weight = line["weight"]
                if line.get("dash_style") is not None:
                    ds_map = {"solid": 1, "dash": 4, "dot": 3, "dash_dot": 5}
                    ds = ds_map.get(line["dash_style"].lower(), 1)
                    ln.DashStyle = ds
                applied.append("line")

            # 3D
            if three_d is not None:
                td = shape.ThreeD
                if three_d.get("rotation_x") is not None:
                    td.RotationX = three_d["rotation_x"]
                if three_d.get("rotation_y") is not None:
                    td.RotationY = three_d["rotation_y"]
                if three_d.get("rotation_z") is not None:
                    td.RotationZ = three_d["rotation_z"]
                if three_d.get("bevel_top") is not None:
                    td.BevelTopType = three_d["bevel_top"]
                if three_d.get("depth") is not None:
                    td.Depth = three_d["depth"]
                if three_d.get("material") is not None:
                    td.PresetMaterial = three_d["material"]
                if three_d.get("lighting") is not None:
                    td.PresetLightingDirection = three_d["lighting"]
                applied.append("three_d")

            return json.dumps({
                "slide_index": slide_index,
                "shape_index": shape_index,
                "shape_name": shape.Name,
                "effects_applied": applied,
            }, ensure_ascii=False)
        except Exception as e:
            return json.dumps({"error": str(e)})

    # ------------------------------------------------------------------
    # 18. ppt_live_arrange_shapes
    # ------------------------------------------------------------------
    @app.tool()
    async def ppt_live_arrange_shapes(
        slide_index: int,
        shape_indices: list,
        operation: str,
        filename: Optional[str] = None,
        relative_to_slide: bool = True,
    ) -> str:
        """[Windows only] Arrange shapes: group, ungroup, align, distribute, z-order, flip, duplicate.
        (python-pptx CANNOT do any of these.)

        Args:
            slide_index: 1-based slide index.
            shape_indices: List of 1-based shape indices to operate on.
            operation: One of: group, ungroup, align_left, align_center, align_right,
                       align_top, align_middle, align_bottom, distribute_h, distribute_v,
                       bring_front, send_back, bring_forward, send_backward, duplicate, flip_h, flip_v.
            filename: Presentation filename. None = active.
            relative_to_slide: For align/distribute, relative to slide (True) or to each other (False).
        """
        if sys.platform != "win32":
            return json.dumps({"error": "Live tools are only available on Windows"})
        try:
            from core.ppt_com import get_ppt_app, find_presentation
            ppt = get_ppt_app()
            pres = find_presentation(ppt, filename)

            if slide_index < 1 or slide_index > pres.Slides.Count:
                return json.dumps({"error": f"slide_index {slide_index} out of range"})

            slide = pres.Slides(slide_index)

            # Validate indices
            for idx in shape_indices:
                if idx < 1 or idx > slide.Shapes.Count:
                    return json.dumps({"error": f"shape_index {idx} out of range (1-{slide.Shapes.Count})"})

            # Get shape names for Range
            shape_names = [slide.Shapes(i).Name for i in shape_indices]
            rel = -1 if relative_to_slide else 0  # msoTrue / msoFalse

            op = operation.lower()
            result_info = {"operation": op, "shape_indices": shape_indices}

            if op == "group":
                if len(shape_indices) < 2:
                    return json.dumps({"error": "Need at least 2 shapes to group"})
                sr = slide.Shapes.Range(shape_names)
                grp = sr.Group()
                result_info["group_name"] = grp.Name

            elif op == "ungroup":
                shape = slide.Shapes(shape_indices[0])
                shape.Ungroup()

            elif op.startswith("align_"):
                # msoAlignLefts=0, Centers=1, Rights=2, Tops=3, Middles=4, Bottoms=5
                align_map = {
                    "align_left": 0, "align_center": 1, "align_right": 2,
                    "align_top": 3, "align_middle": 4, "align_bottom": 5,
                }
                if op not in align_map:
                    return json.dumps({"error": f"Unknown align operation: {op}"})
                sr = slide.Shapes.Range(shape_names)
                sr.Align(align_map[op], rel)

            elif op == "distribute_h":
                sr = slide.Shapes.Range(shape_names)
                sr.Distribute(0, rel)  # msoDistributeHorizontally=0

            elif op == "distribute_v":
                sr = slide.Shapes.Range(shape_names)
                sr.Distribute(1, rel)  # msoDistributeVertically=1

            elif op == "bring_front":
                for idx in shape_indices:
                    slide.Shapes(idx).ZOrder(0)  # msoBringToFront
            elif op == "send_back":
                for idx in shape_indices:
                    slide.Shapes(idx).ZOrder(1)  # msoSendToBack
            elif op == "bring_forward":
                for idx in shape_indices:
                    slide.Shapes(idx).ZOrder(2)  # msoBringForward
            elif op == "send_backward":
                for idx in shape_indices:
                    slide.Shapes(idx).ZOrder(3)  # msoSendBackward

            elif op == "duplicate":
                for idx in shape_indices:
                    slide.Shapes(idx).Duplicate()

            elif op == "flip_h":
                for idx in shape_indices:
                    slide.Shapes(idx).Flip(0)  # msoFlipHorizontal
            elif op == "flip_v":
                for idx in shape_indices:
                    slide.Shapes(idx).Flip(1)  # msoFlipVertical

            else:
                return json.dumps({"error": f"Unknown operation: {op}"})

            return json.dumps(result_info, ensure_ascii=False)
        except Exception as e:
            return json.dumps({"error": str(e)})

    # ------------------------------------------------------------------
    # 19. ppt_live_set_transition
    # ------------------------------------------------------------------
    @app.tool()
    async def ppt_live_set_transition(
        filename: Optional[str] = None,
        slide_index: Optional[int] = None,
        effect: str = "fade",
        duration: Optional[float] = None,
        advance_on_click: Optional[bool] = None,
        advance_on_time: Optional[bool] = None,
        advance_time: Optional[float] = None,
    ) -> str:
        """[Windows only] Set slide transition effect and timing.
        (python-pptx CANNOT do this — the existing tool is a stub.)

        Args:
            filename: Presentation filename. None = active.
            slide_index: 1-based slide index. None = apply to ALL slides.
            effect: Transition type: fade, push, wipe, split, reveal, cut, random, none.
            duration: Transition duration in seconds.
            advance_on_click: Advance on mouse click.
            advance_on_time: Advance automatically after time.
            advance_time: Auto-advance time in seconds.
        """
        if sys.platform != "win32":
            return json.dumps({"error": "Live tools are only available on Windows"})
        try:
            from core.ppt_com import get_ppt_app, find_presentation
            ppt = get_ppt_app()
            pres = find_presentation(ppt, filename)

            # Effect mapping (ppEntryEffect constants)
            effect_map = {
                "none": 0,        # ppEffectNone
                "fade": 3844,     # ppEffectFade
                "push": 3852,     # ppEffectPushDown
                "wipe": 3856,     # ppEffectWipeRight
                "split": 3848,    # ppEffectSplitHorizontalOut
                "reveal": 3845,   # ppEffectFadeSmoothly
                "cut": 257,       # ppEffectCut
                "random": 513,    # ppEffectRandom
            }
            eff = effect_map.get(effect.lower())
            if eff is None:
                return json.dumps({"error": f"Unknown effect '{effect}'. Use: {list(effect_map.keys())}"})

            start = slide_index if slide_index else 1
            end = (slide_index if slide_index else pres.Slides.Count) + 1
            applied_count = 0

            for si in range(start, end):
                if si < 1 or si > pres.Slides.Count:
                    continue
                t = pres.Slides(si).SlideShowTransition
                t.EntryEffect = eff
                if duration is not None:
                    t.Duration = duration
                if advance_on_click is not None:
                    t.AdvanceOnClick = -1 if advance_on_click else 0
                if advance_on_time is not None:
                    t.AdvanceOnTime = -1 if advance_on_time else 0
                if advance_time is not None:
                    t.AdvanceTime = advance_time
                applied_count += 1

            return json.dumps({
                "effect": effect,
                "slides_affected": applied_count,
                "duration": duration,
            }, ensure_ascii=False)
        except Exception as e:
            return json.dumps({"error": str(e)})

    # ------------------------------------------------------------------
    # 20. ppt_live_set_header_footer
    # ------------------------------------------------------------------
    @app.tool()
    async def ppt_live_set_header_footer(
        filename: Optional[str] = None,
        slide_index: Optional[int] = None,
        slide_number: Optional[bool] = None,
        date_time: Optional[bool] = None,
        date_auto_update: Optional[bool] = None,
        date_format: Optional[int] = None,
        date_fixed_text: Optional[str] = None,
        footer_visible: Optional[bool] = None,
        footer_text: Optional[str] = None,
    ) -> str:
        """[Windows only] Set slide headers/footers: slide numbers, date/time, footer text.
        (python-pptx CANNOT do this.)

        Args:
            filename: Presentation filename. None = active.
            slide_index: 1-based slide index. None = apply to ALL slides.
            slide_number: Show slide number.
            date_time: Show date/time.
            date_auto_update: Use auto-updating date format (True) or fixed text (False).
            date_format: Date format index (e.g., 1=M/d/yyyy). Only when date_auto_update=True.
            date_fixed_text: Fixed date text. Only when date_auto_update=False.
            footer_visible: Show footer.
            footer_text: Footer text.
        """
        if sys.platform != "win32":
            return json.dumps({"error": "Live tools are only available on Windows"})
        try:
            from core.ppt_com import get_ppt_app, find_presentation
            ppt = get_ppt_app()
            pres = find_presentation(ppt, filename)

            start = slide_index if slide_index else 1
            end = (slide_index if slide_index else pres.Slides.Count) + 1
            applied_count = 0

            for si in range(start, end):
                if si < 1 or si > pres.Slides.Count:
                    continue
                hf = pres.Slides(si).HeadersFooters

                if slide_number is not None:
                    hf.SlideNumber.Visible = -1 if slide_number else 0
                if date_time is not None:
                    hf.DateAndTime.Visible = -1 if date_time else 0
                if date_auto_update is not None:
                    hf.DateAndTime.UseFormat = -1 if date_auto_update else 0
                if date_format is not None:
                    hf.DateAndTime.Format = date_format
                if date_fixed_text is not None:
                    hf.DateAndTime.Text = date_fixed_text
                if footer_visible is not None:
                    hf.Footer.Visible = -1 if footer_visible else 0
                if footer_text is not None:
                    hf.Footer.Text = footer_text
                applied_count += 1

            return json.dumps({
                "slides_affected": applied_count,
                "settings": {
                    "slide_number": slide_number,
                    "date_time": date_time,
                    "footer_visible": footer_visible,
                    "footer_text": footer_text,
                },
            }, ensure_ascii=False)
        except Exception as e:
            return json.dumps({"error": str(e)})

    # ------------------------------------------------------------------
    # 21. ppt_live_manage_sections
    # ------------------------------------------------------------------
    @app.tool()
    async def ppt_live_manage_sections(
        operation: str,
        filename: Optional[str] = None,
        section_index: Optional[int] = None,
        section_name: Optional[str] = None,
        slide_index: Optional[int] = None,
        delete_slides: bool = False,
    ) -> str:
        """[Windows only] Manage presentation sections: list, add, rename, delete.
        (python-pptx CANNOT do this.)

        Args:
            operation: "list", "add", "rename", or "delete".
            filename: Presentation filename. None = active.
            section_index: 1-based section index (for rename/delete).
            section_name: Section name (for add/rename).
            slide_index: 1-based slide index where new section starts (for add).
            delete_slides: Also delete slides in the section (for delete).
        """
        if sys.platform != "win32":
            return json.dumps({"error": "Live tools are only available on Windows"})
        try:
            from core.ppt_com import get_ppt_app, find_presentation
            ppt = get_ppt_app()
            pres = find_presentation(ppt, filename)
            sp = pres.SectionProperties

            op = operation.lower()

            if op == "list":
                sections = []
                for i in range(1, sp.Count + 1):
                    sections.append({
                        "index": i,
                        "name": sp.Name(i),
                        "first_slide": sp.FirstSlide(i),
                        "slide_count": sp.SlidesCount(i),
                    })
                return json.dumps({"section_count": sp.Count, "sections": sections},
                                  ensure_ascii=False)

            elif op == "add":
                if not section_name:
                    return json.dumps({"error": "section_name is required for add"})
                si = slide_index if slide_index else 1
                idx = sp.AddSection(si, section_name)
                return json.dumps({"added_section_index": idx, "name": section_name,
                                   "starts_at_slide": si}, ensure_ascii=False)

            elif op == "rename":
                if section_index is None or not section_name:
                    return json.dumps({"error": "section_index and section_name required for rename"})
                if section_index < 1 or section_index > sp.Count:
                    return json.dumps({"error": f"section_index {section_index} out of range (1-{sp.Count})"})
                old_name = sp.Name(section_index)
                sp.Rename(section_index, section_name)
                return json.dumps({"section_index": section_index,
                                   "old_name": old_name, "new_name": section_name},
                                  ensure_ascii=False)

            elif op == "delete":
                if section_index is None:
                    return json.dumps({"error": "section_index required for delete"})
                if section_index < 1 or section_index > sp.Count:
                    return json.dumps({"error": f"section_index {section_index} out of range (1-{sp.Count})"})
                name = sp.Name(section_index)
                sp.Delete(section_index, delete_slides)
                return json.dumps({"deleted_section": name,
                                   "slides_deleted": delete_slides,
                                   "remaining_sections": sp.Count},
                                  ensure_ascii=False)

            else:
                return json.dumps({"error": f"Unknown operation '{op}'. Use: list, add, rename, delete"})

        except Exception as e:
            return json.dumps({"error": str(e)})

    # ------------------------------------------------------------------
    # 22. ppt_live_modify_table
    # ------------------------------------------------------------------
    @app.tool()
    async def ppt_live_modify_table(
        slide_index: int,
        shape_index: int,
        operation: str,
        filename: Optional[str] = None,
        row_index: Optional[int] = None,
        col_index: Optional[int] = None,
        merge_to_row: Optional[int] = None,
        merge_to_col: Optional[int] = None,
    ) -> str:
        """[Windows only] Modify a table: add/delete rows/columns, merge cells.
        (python-pptx can only create tables, not add/delete rows/columns.)

        Args:
            slide_index: 1-based slide index.
            shape_index: 1-based shape index (must have a table).
            operation: "add_row", "delete_row", "add_column", "delete_column", "merge_cells".
            filename: Presentation filename. None = active.
            row_index: 1-based row index (for add_row: insert before; for delete_row: which row).
            col_index: 1-based column index (for add_column: insert before; for delete_column: which column).
            merge_to_row: 1-based target row for merge_cells.
            merge_to_col: 1-based target column for merge_cells.
        """
        if sys.platform != "win32":
            return json.dumps({"error": "Live tools are only available on Windows"})
        try:
            from core.ppt_com import get_ppt_app, find_presentation
            ppt = get_ppt_app()
            pres = find_presentation(ppt, filename)

            if slide_index < 1 or slide_index > pres.Slides.Count:
                return json.dumps({"error": f"slide_index {slide_index} out of range"})

            slide = pres.Slides(slide_index)
            shape = _get_shape(slide, shape_index)

            if not shape.HasTable:
                return json.dumps({"error": f"Shape {shape_index} ('{shape.Name}') is not a table"})

            table = shape.Table
            op = operation.lower()

            if op == "add_row":
                ri = row_index if row_index else table.Rows.Count + 1
                table.Rows.Add(ri)
                return json.dumps({"operation": "add_row", "row_index": ri,
                                   "total_rows": table.Rows.Count}, ensure_ascii=False)

            elif op == "delete_row":
                if not row_index:
                    return json.dumps({"error": "row_index required for delete_row"})
                if row_index < 1 or row_index > table.Rows.Count:
                    return json.dumps({"error": f"row_index {row_index} out of range (1-{table.Rows.Count})"})
                table.Rows(row_index).Delete()
                return json.dumps({"operation": "delete_row", "deleted_row": row_index,
                                   "remaining_rows": table.Rows.Count}, ensure_ascii=False)

            elif op == "add_column":
                ci = col_index if col_index else table.Columns.Count + 1
                table.Columns.Add(ci)
                return json.dumps({"operation": "add_column", "col_index": ci,
                                   "total_columns": table.Columns.Count}, ensure_ascii=False)

            elif op == "delete_column":
                if not col_index:
                    return json.dumps({"error": "col_index required for delete_column"})
                if col_index < 1 or col_index > table.Columns.Count:
                    return json.dumps({"error": f"col_index {col_index} out of range (1-{table.Columns.Count})"})
                table.Columns(col_index).Delete()
                return json.dumps({"operation": "delete_column", "deleted_col": col_index,
                                   "remaining_columns": table.Columns.Count}, ensure_ascii=False)

            elif op == "merge_cells":
                if not all([row_index, col_index, merge_to_row, merge_to_col]):
                    return json.dumps({"error": "row_index, col_index, merge_to_row, merge_to_col all required"})
                table.Cell(row_index, col_index).Merge(table.Cell(merge_to_row, merge_to_col))
                return json.dumps({"operation": "merge_cells",
                                   "from": [row_index, col_index],
                                   "to": [merge_to_row, merge_to_col]}, ensure_ascii=False)

            else:
                return json.dumps({"error": f"Unknown operation '{op}'. Use: add_row, delete_row, add_column, delete_column, merge_cells"})

        except Exception as e:
            return json.dumps({"error": str(e)})
