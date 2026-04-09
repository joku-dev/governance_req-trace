# User Guide

## Install
```powershell
pip install -r requirements.txt
```

## Run
```powershell
python src\devsecops_requirements_extractor.py
```

Or double-click `run_extractor.bat`.

## Usage steps
1. Select one or more Word documents.
2. Select a save target for the generated Excel workbook.
3. Open the workbook.
4. Review extracted requirements and hyperlinks.

## Expected output
The workbook contains:
- a master requirement list
- source excerpts
- a topic cross-reference sheet
- a document inventory sheet
- a README sheet

## Operational note
This is a heuristic extractor. Do not use the raw output as an authoritative compliance baseline without human review.
