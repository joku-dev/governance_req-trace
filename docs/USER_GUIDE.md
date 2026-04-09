# User Guide

## Install
```bash
pip install -r requirements.txt
```

## Run
GUI mode:

Windows:
```powershell
python src\devsecops_requirements_extractor.py
```

Linux/macOS:
```bash
python3 src/devsecops_requirements_extractor.py
```

CLI mode (explicit files, no GUI required):
```bash
python3 src/devsecops_requirements_extractor.py ./input1.docx ./input2.docx -o ./output.xlsx
```

Launcher scripts:
- Windows: `run_extractor.bat`
- Linux/macOS: `./run_extractor.sh`

## Usage steps
1. Select one or more Word documents, or pass files as CLI arguments.
2. Select a save target for the generated Excel workbook (or pass `-o`).
3. Open the workbook.
4. Review extracted requirements and hyperlinks.

## Expected output
The workbook contains:
- a master requirement list
- source excerpts
- a topic cross-reference sheet
- a document inventory sheet
- a README sheet

## Platform and format note
- `.docx` is supported on Windows, Linux, and macOS.
- `.doc` is supported only on Windows with Microsoft Word installed.

## Operational note
This is a heuristic extractor. Do not use the raw output as an authoritative compliance baseline without human review.
