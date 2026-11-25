# -*- coding: utf-8 -*-
"""Open PPTX files in a real PowerPoint instance and report repair failures.

This catches PowerPoint-specific problems (notably stale relationship targets
in copied media shapes) that OOXML schema validators do not detect.

Usage:
  python check_powerpoint_open.py <deck.pptx> [more.pptx ...]

Requires pythonnet. PowerPoint must be installed. The application is made
visible because PowerPoint refuses to run fully hidden; files are opened
read-only and closed immediately.
"""
import sys, glob, os

PPT_PAT = r'C:/Windows/assembly/GAC_MSIL/Microsoft.Office.Interop.PowerPoint/**/*.dll'
OFFICE_PAT = r'C:/Windows/assembly/GAC_MSIL/office/**/OFFICE.DLL'


def first(pattern):
    hits = [f for f in glob.glob(pattern, recursive=True) if os.path.isfile(f)]
    return hits[0] if hits else None


def main():
    try:
        import clr
    except ImportError:
        print('pythonnet is required: pip install pythonnet')
        return 2
    ppt = first(PPT_PAT)
    office = first(OFFICE_PAT)
    if not ppt or not office:
        print('PowerPoint interop assemblies not found')
        return 2
    clr.AddReference(ppt)
    clr.AddReference(office)
    from Microsoft.Office.Interop.PowerPoint import Application
    from Microsoft.Office.Core import MsoTriState

    app = Application()
    app.Visible = MsoTriState.msoTrue
    failed = 0
    try:
        for path in sys.argv[1:]:
            try:
                pres = app.Presentations.Open(
                    os.path.abspath(path),
                    MsoTriState.msoTrue,   # ReadOnly
                    MsoTriState.msoFalse,  # Untitled
                    MsoTriState.msoFalse,  # WithWindow
                )
                print(f'OK   {path} ({pres.Slides.Count} slides)')
                pres.Close()
            except Exception as e:
                failed += 1
                print(f'FAIL {path}: {str(e).splitlines()[0][:160]}')
    finally:
        try:
            app.Quit()
        except Exception:
            pass
    return 1 if failed else 0


if __name__ == '__main__':
    raise SystemExit(main())
