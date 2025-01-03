# Bulk Document Generator - VBA Project

This is a VBA project for generating bulk documents by using a template and populating it with data from a spreadsheet.

## Files
1. **Data Sheet**: `datasheet.xlsx`
   - The source data for generating the documents. This Excel sheet contains the information that will replace placeholders in the Word template.
   
2. **Template**: `sampleTemplate.docx`
   - A sample Word template where placeholders are defined. The placeholders will be replaced by the data from the Excel sheet.

3. **VBA Code**: `BulkDocumentGenerator.bas`
   - This is the VBA code file that handles the bulk document generation. It takes the data from the Excel file and uses the template to create new documents.

4. **Generated Sample Documents**:
   - After running the code, three sample documents will be generated based on the data in `datasheet.xlsx` and the structure in `sampleTemplate.docx`. These files demonstrate how the final output will look.

## How It Works
The VBA code takes data from the `datasheet.xlsx`, opens the template document `sampleTemplate.docx`, and replaces the placeholders in the Word template with the data in the Excel file. The resulting documents are saved with new names based on the data from the Excel sheet.

1. **Data Sheet**: The Excel sheet must have columns with relevant data, such as names, dates, project names, etc. Each row contains information for generating one document.
   
2. **Template**: In the Word template, placeholders are in the format `<Placeholder>`. The VBA code will search for these placeholders and replace them with actual data from the Excel file.

3. **Output**: After the macro is executed, new Word files are saved in the working directory with unique names based on the content of the Excel sheet.

## How to Use

1. **Prepare Data**: Make sure `datasheet.xlsx` is populated with the relevant information. The first row contains the headers (placeholders to match in the template).

2. **Prepare Template**: Edit `sampleTemplate.docx` to include placeholders in the format `<Placeholder>`. Ensure the placeholders match the column headers in the Excel sheet.

3. **Run the VBA Code**:
   - Open the Excel workbook that contains the data sheet.
   - Press `Alt + F11` to open the VBA editor.
   - Import the `BulkDocumentGenerator.bas` code.
   - Run the macro from the VBA editor.

4. **Output Files**: The VBA code will generate documents based on the template and save them in the project folder.

## Example Output

- `GeneratedDocument_JohnDoe_1.docx`
- `GeneratedDocument_JaneSmith_2.docx`
- `GeneratedDocument_MichaelJohnson_3.docx`

These files are based on the data in `datasheet.xlsx`.

## Requirements

- Microsoft Excel
- Microsoft Word

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.
