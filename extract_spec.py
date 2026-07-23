#!/usr/bin/env python3
"""Thin wrapper so `python extract_spec.py ...` works from a repo clone
without installing the package. Installed users get the `xlsb-extract-spec`
console script instead (see xlsb_reader/_spec_extractor.py for the implementation).
"""

from xlsb_reader._spec_extractor import main

if __name__ == "__main__":
    main()
