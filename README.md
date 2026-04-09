# DevSecOps Requirements Extractor

Windows-compatible utility to extract requirement-like statements from selected Word documents and generate a structured Excel workbook with working hyperlinks.

## Features
- Multi-file selection dialog for `.docx` and `.doc`
- Heuristic extraction of requirements using modal verbs such as `SHALL`, `MUST`, `SHALL NOT`, `MUST NOT`, `SHOULD`, and `MAY`
- Output workbook with the following sheets:
  - `README`
  - `Requirements_Master`
  - `Source_Excerpts`
  - `Cross_Reference_Map`
  - `Documents`
- Working hyperlinks between workbook sheets
- Working file hyperlinks from Excel back to the original source documents

## Repository structure
```text
.
├── .gitignore
├── LICENSE
├── README.md
├── requirements.txt
├── run_extractor.bat
├── src/
│   └── devsecops_requirements_extractor.py
├── docs/
│   ├── ARCHITECTURE.md
│   └── USER_GUIDE.md
├── examples/
│   └── output_placeholder.txt
└── tests/
    └── README.md
```

## Prerequisites
- Windows 10/11
- Python 3.10+
- Microsoft Word installed

Install dependencies:
```powershell
pip install -r requirements.txt
```

## Start
### Option 1: PowerShell / CMD
```powershell
python src\devsecops_requirements_extractor.py
```

### Option 2: Double-click under Windows
Double-click:
```text
run_extractor.bat
```

## What the tool does
1. Opens a file picker to select Word source documents
2. Reads document content via Word COM automation
3. Detects requirement-like statements heuristically
4. Groups entries into cross-reference topics
5. Generates an Excel workbook with internal links and source file links

## Limitations
This is a first-pass extractor, not a fully semantic compliance parser. Every extracted requirement should be reviewed before governance, audit, or certification use.

## Recommended next evolution
- PDF support
- direct write-back into an existing corporate workbook template
- configurable taxonomy for `XREF_Group_ID`
- stronger classification logic for policy vs. control vs. guidance
- metadata enrichment for owners, control families, and evidence mappings
