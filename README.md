# Mabel Letter Generator

This tool automates the generation of client letters by merging data exported from LawLogix with a Word template. For each row in the Excel file it produces a personalized letter in DOCX format, replacing placeholder tags (`<caps_full_name>`, `<full_name>`, `<rm>`, `<date>`) in the template.

## Directory Structure

mabel-letter-generator/
├── fill_letters.py
├── in/
│   ├── Letter_Template.docx
│   └── Letter_Excel.xlsx
├── out/
├── .gitignore
└── README.md

## Usage

1. **Prepare your inputs**  
   - Copy your blank Word template into `in/Letter_Template.docx`. Ensure it contains the placeholder tags:  
     - `<caps_full_name>`  
     - `<full_name>`  
     - `<rm>`  
     - `<date>`  
2. **Export your client data from LawLogix to Excel**
     - name it `in/Letter_Excel.xlsx`
     - verify it has the columns exactly as shown above.

4. **Install dependencies**  
   ```bash
   python3 -m venv .venv
   source .venv/bin/activate
   pip install pandas python-docx
