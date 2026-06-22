#!/usr/bin/env python3
"""Read slide_antibiotics.pdf and extract text content."""
# /// script
# requires-python = ">=3.11"
# dependencies = ["pdfplumber>=0.11"]
# ///

import pdfplumber
import sys

pdf_path = sys.argv[1] if len(sys.argv) > 1 else "/project/data/slide_antibiotics.pdf"

with pdfplumber.open(pdf_path) as pdf:
    print(f"Total pages: {len(pdf.pages)}")
    for i, page in enumerate(pdf.pages, 1):
        text = page.extract_text()
        if text and text.strip():
            print(f"\n=== Page {i} ===")
            print(text)
