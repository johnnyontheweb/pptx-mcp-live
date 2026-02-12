"""
Live export tools using PowerPoint COM automation (Windows only).
Export to PDF, slide images, and shape images — python-pptx CANNOT do any of these.
"""
import json
import os
import sys
import tempfile
from typing import Optional


def register_live_export_tools(app):
    """Register 3 export tools with the FastMCP app."""

    # ------------------------------------------------------------------
    # 23. ppt_live_export_pdf
    # ------------------------------------------------------------------
    @app.tool()
    async def ppt_live_export_pdf(
        filename: Optional[str] = None,
        output_path: Optional[str] = None,
    ) -> str:
        """[Windows only] Export presentation to PDF.
        (python-pptx CANNOT export PDF.)

        Args:
            filename: Presentation filename. None = active.
            output_path: Output PDF path. None = auto temp path.
        """
        if sys.platform != "win32":
            return json.dumps({"error": "Live tools are only available on Windows"})
        try:
            from core.ppt_com import get_ppt_app, find_presentation
            ppt = get_ppt_app()
            pres = find_presentation(ppt, filename)

            if not output_path:
                export_dir = os.path.join(tempfile.gettempdir(), "ppt_mcp_exports")
                os.makedirs(export_dir, exist_ok=True)
                base = pres.Name.rsplit(".", 1)[0] if "." in pres.Name else pres.Name
                output_path = os.path.join(export_dir, f"{base}.pdf")

            # Ensure absolute path
            output_path = os.path.abspath(output_path)

            # ppFixedFormatTypePDF = 2, ppFixedFormatIntentScreen = 1
            # PrintRange=None is required: makepy defaults to 0 which can't
            # marshal to VT_DISPATCH (COM object), causing E_FAIL.
            pres.ExportAsFixedFormat(
                Path=output_path,
                FixedFormatType=2,
                Intent=1,
                PrintRange=None,
            )

            size = os.path.getsize(output_path) if os.path.exists(output_path) else 0
            return json.dumps({
                "path": output_path,
                "size_bytes": size,
                "slide_count": pres.Slides.Count,
            }, ensure_ascii=False)
        except Exception as e:
            return json.dumps({"error": str(e)})

    # ------------------------------------------------------------------
    # 24. ppt_live_export_slide_image
    # ------------------------------------------------------------------
    @app.tool()
    async def ppt_live_export_slide_image(
        slide_index: int,
        filename: Optional[str] = None,
        output_path: Optional[str] = None,
        format: str = "PNG",
        width: int = 1920,
    ) -> str:
        """[Windows only] Export a single slide as an image (PNG/JPG).
        (python-pptx CANNOT export images.)

        Args:
            slide_index: 1-based slide index.
            filename: Presentation filename. None = active.
            output_path: Output image path. None = auto temp path.
            format: Image format: "PNG" or "JPG".
            width: Image width in pixels (height auto-calculated from aspect ratio).
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
            fmt = format.upper()
            if fmt not in ("PNG", "JPG"):
                return json.dumps({"error": "format must be 'PNG' or 'JPG'"})

            # Calculate height from aspect ratio
            sw = pres.PageSetup.SlideWidth
            sh = pres.PageSetup.SlideHeight
            height = int(width * sh / sw) if sw > 0 else int(width * 9 / 16)

            if not output_path:
                export_dir = os.path.join(tempfile.gettempdir(), "ppt_mcp_exports")
                os.makedirs(export_dir, exist_ok=True)
                ext = "png" if fmt == "PNG" else "jpg"
                output_path = os.path.join(export_dir, f"slide_{slide_index}.{ext}")

            output_path = os.path.abspath(output_path)
            slide.Export(output_path, fmt, width, height)

            size = os.path.getsize(output_path) if os.path.exists(output_path) else 0
            return json.dumps({
                "path": output_path,
                "size_bytes": size,
                "width": width,
                "height": height,
                "format": fmt,
            }, ensure_ascii=False)
        except Exception as e:
            return json.dumps({"error": str(e)})

    # ------------------------------------------------------------------
    # 25. ppt_live_export_shape_image
    # ------------------------------------------------------------------
    @app.tool()
    async def ppt_live_export_shape_image(
        slide_index: int,
        shape_index: int,
        filename: Optional[str] = None,
        output_path: Optional[str] = None,
        format: str = "PNG",
    ) -> str:
        """[Windows only] Export a single shape as an image.
        (python-pptx CANNOT export shapes.)

        Args:
            slide_index: 1-based slide index.
            shape_index: 1-based shape index.
            filename: Presentation filename. None = active.
            output_path: Output image path. None = auto temp path.
            format: Image format: "PNG" or "JPG".
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
            if shape_index < 1 or shape_index > slide.Shapes.Count:
                return json.dumps({"error": f"shape_index {shape_index} out of range (1-{slide.Shapes.Count})"})

            shape = slide.Shapes(shape_index)
            fmt = format.upper()
            if fmt not in ("PNG", "JPG"):
                return json.dumps({"error": "format must be 'PNG' or 'JPG'"})

            # ppShapeFormatPNG = 2, ppShapeFormatJPG = 1
            fmt_enum = 2 if fmt == "PNG" else 1

            if not output_path:
                export_dir = os.path.join(tempfile.gettempdir(), "ppt_mcp_exports")
                os.makedirs(export_dir, exist_ok=True)
                ext = "png" if fmt == "PNG" else "jpg"
                output_path = os.path.join(export_dir, f"shape_s{slide_index}_{shape_index}.{ext}")

            output_path = os.path.abspath(output_path)
            shape.Export(output_path, fmt_enum)

            size = os.path.getsize(output_path) if os.path.exists(output_path) else 0
            return json.dumps({
                "path": output_path,
                "size_bytes": size,
                "shape_name": shape.Name,
                "format": fmt,
            }, ensure_ascii=False)
        except Exception as e:
            return json.dumps({"error": str(e)})
