# excel-auto-fill

This Python script processes and formats Excel files. It extracts data from a given Excel file, inserts the extracted data into a template, applies styles and formatting, locks certain cells for protection, and saves the modified file. The script also handles both `.xls` and `.xlsx` file formats.

Note: The script is currently designed to work with a specific input file and template. If different files are used, the code may need modifications to ensure compatibility.

## Features

- Extracts data from a specified input Excel file.
- Handles `.xls` and `.xlsx` file formats.
- Uses a template file for output formatting.
- Protects specific cells and columns with a password.
- Automatically opens the output file after processing.

## Requirements

- Python 3.x
- `openpyxl` (for working with Excel files)
- `pandas` (for reading and writing data)
- `xlrd` (for reading `.xls` files)

You can install the necessary packages with:

```bash
pip install openpyxl pandas xlrd
```

## How to Use the Program

### Windows
1. Double-click on the `excel-auto-fill.py` file to run the program.
2. The program will prompt you to enter the file name.
3. The program will automatically process the input file, generate the output, and open it in Excel.
### macOS
1. Open **Terminal**.
2. Navigate to the directory where the `excel-auto-fill.py` file is located using the `cd` command:
   ```bash
   cd /path/to/your/program
3. Run the program using the following command: 
   ```bash
   python3 excel-auto-fill.py
   ```
4. The program will prompt you to enter the file name.
5. The program will automatically process the input file, generate the output, and open it in Excel.

## Future Upgrades
- Support more file formats: Extend the program to support additional file formats such as CSV, JSON, or even Google Sheets.
- Add logging features: Implement logging to track events and errors in the program.
- Enable users to select different template styles dynamically through the program, such as choosing font, borders, colors, and other formatting options.
- Use Outlook API: Integrate with the Outlook API to automatically create a draft email containing the generated output file as an attachment.
- Automatically remove protection from password-protected .xls and .xlsx files
- Optional: Allow users to create a draft email with the Outlook API and choose the contact category from their Outlook contacts (e.g., client, supplier, etc.).
