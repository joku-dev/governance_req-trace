# Architecture

## Overview
The tool follows a straightforward desktop/CLI batch-processing pattern:

1. **User interaction layer**
   - GUI file picker and save dialog (if tkinter is available)
   - CLI argument mode for non-GUI environments

2. **Document ingestion**
   - `python-docx` parses `.docx` on all platforms
   - Microsoft Word COM automation parses `.doc` and can also parse `.docx` on Windows
   - paragraphs are normalized into processable text units

3. **Requirement extraction**
   - heuristic detection based on modal verbs and requirement phrasing
   - scoring-based classification into `Policy`, `Control`, and `Guidance`
   - confidence assignment (`High`, `Medium`, `Low`)

4. **Cross-reference enrichment**
   - topic mapping based on keyword rules
   - metadata enrichment for control family, owner role, and evidence mapping
   - linkage between requirements, excerpts, and topic groups

5. **Workbook generation**
   - `openpyxl` creates the workbook
   - internal hyperlinks connect workbook entities
   - file hyperlinks point back to original source documents

## Design rationale
Python was selected over pure VBA because document parsing and workbook generation are more maintainable, extensible, and testable in Python for this use case.
