"""
Path utilities for PowerPoint MCP Server.
Provides locked-file detection hook for python-pptx.
"""
import os


def install_pptx_path_hook():
    """Monkey-patch PhysPkgReader to detect locked files and give helpful error messages."""
    from pptx.opc.phys_pkg import PhysPkgReader
    from pptx.opc.exceptions import PackageNotFoundError

    if getattr(PhysPkgReader.__new__, "_locked_file_hooked", False):
        return

    _orig = PhysPkgReader.__new__

    def _patched(cls, pkg_file, *a, **kw):
        if isinstance(pkg_file, str) and not os.path.isdir(pkg_file):
            if not os.path.exists(pkg_file):
                raise PackageNotFoundError(f"Not found: '{pkg_file}'")
            try:
                with open(pkg_file, "rb"):
                    pass
            except PermissionError:
                raise PackageNotFoundError(
                    f"File is locked (open in PowerPoint): '{pkg_file}'. "
                    "Use ppt_live_* tools instead."
                )
        return _orig(cls, pkg_file, *a, **kw)

    _patched._locked_file_hooked = True
    PhysPkgReader.__new__ = _patched
