# -*- coding: utf-8 -*-
"""Strict OOXML validation using the DocumentFormat.OpenXml assemblies shipped
with Microsoft Office. Catches schema issues that zip/XML well-formedness
checks (and python-pptx) miss.

Usage:
  python validate_openxml.py <deck.pptx> [more.pptx ...]

Requires pythonnet (pip install pythonnet). The Office assembly is discovered
automatically under C:\\Program Files\\Microsoft Office.
"""
import sys, glob, os

DLL_PATTERNS = [
    r'C:/Program Files/Microsoft Office/root/Office16/ADDINS/**/DocumentFormat.OpenXml.dll',
    r'C:/Program Files/Microsoft Office/root/vfs/**/DocumentFormat.OpenXml.dll',
    r'C:/Program Files (x86)/Microsoft Office/**/DocumentFormat.OpenXml.dll',
]


def find_dll():
    for pat in DLL_PATTERNS:
        for f in glob.glob(pat, recursive=True):
            if os.path.isfile(f):
                return f
    return None


def main():
    try:
        import clr
    except ImportError:
        print('pythonnet is required: pip install pythonnet')
        return 2
    dll = find_dll()
    if not dll:
        print('DocumentFormat.OpenXml.dll not found under Microsoft Office')
        return 2
    clr.AddReference(dll)
    from DocumentFormat.OpenXml.Validation import OpenXmlValidator
    from DocumentFormat.OpenXml.Packaging import PresentationDocument

    failed = 0
    for path in sys.argv[1:]:
        doc = PresentationDocument.Open(os.path.abspath(path), False)
        errs = list(OpenXmlValidator().Validate(doc))
        doc.Close()
        print(f'{path}: {len(errs)} OOXML validation error(s)')
        for e in errs[:50]:
            try:
                xp = e.Path.XPath
            except Exception:
                xp = ''
            print(f'  [{e.ErrorType}] {e.Description} @ {xp}')
        if errs:
            failed += 1
    return 1 if failed else 0


if __name__ == '__main__':
    raise SystemExit(main())
