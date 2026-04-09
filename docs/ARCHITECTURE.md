# Architecture

## Overview
The tool follows a straightforward desktop batch-processing pattern:

1. **User interaction layer**
   - Windows file dialog for selecting source documents
   - save dialog for choosing the target Excel workbook

2. **Document ingestion**
   - Microsoft Word COM automation reads `.doc` and `.docx`
   - paragraphs are normalized into processable text units

3. **Requirement extraction**
   - heuristic detection based on modal verbs and requirement phrasing
   - classification into requirement strength categories

4. **Cross-reference enrichment**
   - topic mapping based on keyword rules
   - linkage between requirements, excerpts, and topic groups

5. **Workbook generation**
   - `openpyxl` creates the workbook
   - internal hyperlinks connect workbook entities
   - file hyperlinks point back to original source documents

## Design rationale
Python was selected over pure VBA because document parsing and workbook generation are more maintainable, extensible, and testable in Python for this use case.
