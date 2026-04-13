Windows local document-to-Markdown converter

Files in this folder:
- convert_to_md.py
- pdf_to_md_ai.py
- requirements.txt
- setup_windows.ps1

Install:
1. Install Python 3.12+
2. Install Node.js
3. Run:
   powershell -ExecutionPolicy Bypass -File .\setup_windows.ps1

Run:
1. Convert all supported files under current folder:
   py .\convert_to_md.py "." --skip-unsupported

2. Convert specific files:
   py .\convert_to_md.py ".\sample.pdf" ".\sample.hwp" ".\sample.docx"

Supported:
- pdf
- hwp
- docx / xlsx / pptx
- doc / xls / ppt (requires Microsoft Office installed)
- html / csv / json / xml / txt and more through MarkItDown

Notes:
- hwp conversion requires @ohah/hwpjs
- pdf conversion uses the bundled pdf_to_md_ai.py
- output is written to .\output_md by default
