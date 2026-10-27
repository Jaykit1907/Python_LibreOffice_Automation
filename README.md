# Excel Report Automation

## Folder structure
- `input/Andhra Pradesh.xlsx`: source workbook (each worksheet is a district)
- `input/Template.docx`: Word template
- `input/Nellore.pdf`: visual reference
- `main.py`: complete single-file automation
- `output/`: generated Word and PDF reports

## Install Python packages
```bash
python -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install -r requirements.txt
```

Install LibreOffice separately for PDF conversion.

## Run
From the project root:
```bash
python main.py
```
