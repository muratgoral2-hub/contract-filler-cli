# Contract Filler CLI

This project is a **Python-based automation tool** that fills in Word contract templates (`.docx`) using client data from Excel/CSV/JSON files.  
It can also export the final document to PDF and add a company logo.

## Features
- Read client data from:
  - Excel (`.xlsx`)
  - CSV
  - JSON / JSONL
- Replace placeholders in Word templates (`{name}`, `{surname}`, `{company}`, etc.)
- Fill both paragraphs and tables in Word documents
- Export to PDF automatically
- Add a logo on the first page of the PDF

## Example Usage
```bash
python contractfillercli.py   --template contract_template.docx   --data client.xlsx   --out output_contracts   --logo logo.png
```

## Project Structure
```
contract_filler_github_demo/
│── contractfillercli.py       # Main Python script
│── contract_template.docx     # Contract template with placeholders
│── client.xlsx                # Example client data
│── logo.png                   # Company logo
│── requirements.txt           # Dependencies
│── README.md                  # Documentation
```

## Requirements
Install all dependencies with:
```bash
pip install -r requirements.txt
```

## Example Placeholders
Inside the contract template, placeholders should look like:
```
This contract is signed by {name} {surname}, representative of {company}.
```

These placeholders will be replaced with the corresponding values from `client.xlsx`.
