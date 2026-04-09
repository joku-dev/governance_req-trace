# DevSecOps Requirements Extractor

Cross-platform utility to extract requirement-like statements from selected Word documents and generate a structured Excel workbook with working hyperlinks.

## Features
- Multi-file selection dialog for `.docx` and `.doc`
- CLI mode for headless execution on Linux/macOS/CI
- Heuristic extraction of requirements using modal verbs such as `SHALL`, `MUST`, `SHALL NOT`, `MUST NOT`, `SHOULD`, and `MAY`
- Stronger classification into `Policy`, `Control`, and `Guidance` with confidence scoring
- Metadata enrichment for `Owner_Role`, `Control_Family`, and `Evidence_Mapping`
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
├── run_extractor.sh
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
- Python 3.10+
- Microsoft Word installed only if you need to parse legacy `.doc` files

Platform support:
- `.docx`: Windows, Linux, macOS
- `.doc`: Windows only (requires Microsoft Word + `pywin32`)

Install dependencies:
```bash
pip install -r requirements.txt
```

## Start
### Option 1: GUI
Windows (PowerShell/CMD):
```powershell
python src\devsecops_requirements_extractor.py
```

Linux/macOS (Terminal):
```bash
python3 src/devsecops_requirements_extractor.py
```

### Option 2: CLI with explicit files (recommended for Linux/macOS servers)
```bash
python3 src/devsecops_requirements_extractor.py ./docs/input1.docx ./docs/input2.docx -o ./output.xlsx
```

### Option 3: Launcher scripts
Windows (double-click):
```text
run_extractor.bat
```

Linux/macOS:
```bash
./run_extractor.sh ./docs/input1.docx -o ./output.xlsx
```

## What the tool does
1. Opens a file picker to select Word source documents
2. Reads `.docx` using `python-docx` (cross-platform) and `.doc` via Word COM (Windows only)
3. Detects requirement-like statements heuristically
4. Groups entries into cross-reference topics
5. Generates an Excel workbook with internal links and source file links

## Limitations
This is a first-pass extractor, not a fully semantic compliance parser. Every extracted requirement should be reviewed before governance, audit, or certification use. Legacy `.doc` files are not supported on Linux/macOS.

## Recommended next evolution
- PDF support
- direct write-back into an existing corporate workbook template
- configurable taxonomy for `XREF_Group_ID`
