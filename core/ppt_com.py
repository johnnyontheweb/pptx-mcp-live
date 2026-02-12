"""
PowerPoint COM automation core — Windows only.
Provides get_ppt_app() and find_presentation() for all ppt_live_* tools.
"""
import os
import sys


def get_ppt_app():
    """Get a running PowerPoint.Application COM object.

    Raises RuntimeError if not on Windows or PowerPoint is not running.
    """
    if sys.platform != "win32":
        raise RuntimeError("PowerPoint COM automation is only available on Windows")
    import win32com.client
    try:
        return win32com.client.GetActiveObject("PowerPoint.Application")
    except Exception:
        raise RuntimeError("PowerPoint is not running. Please open PowerPoint first.")


def find_presentation(ppt_app, filename=None):
    """Find an open presentation by filename (basename match).

    If *filename* is None/empty, returns the ActivePresentation.
    Raises ValueError when no presentations are open or the target is not found.
    """
    if ppt_app.Presentations.Count == 0:
        raise ValueError("No presentations are open in PowerPoint")
    if not filename:
        return ppt_app.ActivePresentation
    target = os.path.basename(filename).lower()
    for i in range(1, ppt_app.Presentations.Count + 1):
        pres = ppt_app.Presentations(i)
        if pres.Name.lower() == target:
            return pres
    open_list = [ppt_app.Presentations(i).Name
                 for i in range(1, ppt_app.Presentations.Count + 1)]
    raise ValueError(f"'{filename}' is not open in PowerPoint. Open files: {open_list}")
