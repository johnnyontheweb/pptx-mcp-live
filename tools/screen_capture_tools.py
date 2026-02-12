"""
Screen capture tool for PowerPoint window (Windows only).
Uses PrintWindow API via pywin32 + Pillow.
"""
import json
import os
import sys
import tempfile
from typing import Optional


def _capture_window_to_png(hwnd: int) -> bytes:
    """Capture a window by HWND and return PNG bytes."""
    import win32gui
    import win32ui
    from ctypes import windll
    from PIL import Image
    import io

    rect = win32gui.GetWindowRect(hwnd)
    w = rect[2] - rect[0]
    h = rect[3] - rect[1]
    if w <= 0 or h <= 0:
        raise RuntimeError(f"Invalid window dimensions: {w}x{h}")

    wDC = win32gui.GetWindowDC(hwnd)
    dcObj = win32ui.CreateDCFromHandle(wDC)
    cDC = dcObj.CreateCompatibleDC()
    bmp = win32ui.CreateBitmap()
    bmp.CreateCompatibleBitmap(dcObj, w, h)
    cDC.SelectObject(bmp)

    result = windll.user32.PrintWindow(hwnd, cDC.GetSafeHdc(), 2)
    if not result:
        windll.user32.PrintWindow(hwnd, cDC.GetSafeHdc(), 0)

    bmpinfo = bmp.GetInfo()
    bmpstr = bmp.GetBitmapBits(True)
    img = Image.frombuffer(
        "RGB",
        (bmpinfo["bmWidth"], bmpinfo["bmHeight"]),
        bmpstr, "raw", "BGRX", 0, 1,
    )

    dcObj.DeleteDC()
    cDC.DeleteDC()
    win32gui.ReleaseDC(hwnd, wDC)
    win32gui.DeleteObject(bmp.GetHandle())

    buf = io.BytesIO()
    img.save(buf, "PNG")
    return buf.getvalue()


def register_screen_capture_tools(app):
    """Register 1 screen capture tool with the FastMCP app."""

    @app.tool()
    async def ppt_screen_capture(
        filename: Optional[str] = None,
        output_path: Optional[str] = None,
    ) -> str:
        """[Windows only] Take a screenshot of the PowerPoint window.

        Args:
            filename: Presentation filename to activate its window. None = active window.
            output_path: Where to save the PNG. None = auto temp path.
        """
        if sys.platform != "win32":
            return json.dumps({"error": "Screen capture is only available on Windows"})
        try:
            from core.ppt_com import get_ppt_app, find_presentation
            import win32gui

            ppt = get_ppt_app()

            # Activate the right window
            if filename:
                pres = find_presentation(ppt, filename)
                try:
                    pres.Windows(1).Activate()
                except Exception:
                    pass

            # Get HWND
            hwnd = None
            try:
                hwnd = int(ppt.ActiveWindow.HWND)
            except Exception:
                pass

            if not hwnd:
                hwnd = win32gui.FindWindow("PPTFrameClass", None)
            if not hwnd:
                return json.dumps({"error": "Could not find PowerPoint window handle"})

            # Capture
            png_bytes = _capture_window_to_png(hwnd)

            # Save
            if not output_path:
                capture_dir = os.path.join(tempfile.gettempdir(), "ppt_mcp_captures")
                os.makedirs(capture_dir, exist_ok=True)
                safe_name = "ppt_capture"
                try:
                    safe_name = f"ppt_capture_{ppt.ActivePresentation.Name.rsplit('.', 1)[0]}"
                except Exception:
                    pass
                output_path = os.path.join(capture_dir, f"{safe_name}.png")

            with open(output_path, "wb") as f:
                f.write(png_bytes)

            return json.dumps({
                "path": output_path,
                "size_bytes": len(png_bytes),
            }, ensure_ascii=False)
        except Exception as e:
            return json.dumps({"error": str(e)})
