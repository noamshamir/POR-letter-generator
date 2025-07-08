# Mabel Letter Generator

This tool automates the generation of client letters by merging data exported from LawLogix with a Word template. For each row in the Excel file it produces a personalized letter in DOCX format, replacing placeholder tags in the template.

## Usage

1. **Prepare your inputs**  
   - Create a word document here: `in/Letter_Template.docx`
   - Add your specific template template. Ensure it contains the placeholder tags:  
     - `<caps_full_name>`  
     - `<full_name>`  
     - `<rm>`  
     - `<date>`
       
2. **Export your client data from LawLogix to Excel**
     - name it `in/Letter_Excel.xlsx`
     - verify it has the columns exactly as shown above.

4. **Install dependencies**
   - In the python script, the file names are relative to the repo. Make sure these are correct for you.
   ```bash
   python3 -m venv .venv
   source .venv/bin/activate
   pip install pandas python-docx
   
