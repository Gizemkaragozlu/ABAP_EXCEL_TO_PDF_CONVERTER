# ABAP Excel to PDF Converter

## Overview

This ABAP program converts an Excel file to a PDF format using OLE automation in SAP. It provides a class `ZCL_EXCEL_PDF_CONVERTER` with methods to perform the conversion. The class uses OLE (Object Linking and Embedding) to interact with Microsoft Excel, manipulate the spreadsheet, and export it as a PDF.

## Prerequisites

- SAP system with ABAP support
- Microsoft Excel installed on the system where this code runs
- Access to the Excel file and a destination path for the PDF

## How It Works

1. **Initialize Excel Application**: The program creates an instance of the Excel application.
2. **Open Workbook**: It opens the specified Excel file.
3. **Adjust Formatting**: Sets up page layout and formatting options.
4. **Export as PDF**: Converts the Excel file to PDF format and saves it to the specified destination.
5. **Cleanup**: Closes the workbook without saving changes and quits the Excel application.

## Parameters

- `p_source`: The path to the source Excel file (e.g., `C:\path\to\file.xlsx`).
- `p_dest`: The path where the PDF will be saved (e.g., `C:\path\to\output.pdf`).

## Usage

To use this program, include the following ABAP code in your selection screen or report:

```abap
PARAMETERS: p_source TYPE string LOWER CASE. " Path to source Excel file
PARAMETERS: p_dest   TYPE string LOWER CASE. " Path to output PDF file

START-OF-SELECTION.
  ZCL_EXCEL_PDF_CONVERTER=>get_instance( )->convert(
    EXPORTING
      source      = p_source
      destination = p_dest ).
```

## Class Definition

### `ZCL_EXCEL_PDF_CONVERTER`

#### Methods

- **`get_instance`**: Returns a singleton instance of the class.
- **`convert`**: Converts the specified Excel file to PDF.

### Usage Example

```abap
DATA(lo_converter) = ZCL_EXCEL_PDF_CONVERTER=>get_instance( ).
lo_converter->convert(
  EXPORTING
    source      = 'C:\path\to\source.xlsx'
    destination = 'C:\path\to\output.pdf' ).
```

## Notes

- Ensure that the path to the Excel file and the destination PDF file are accessible and writable by the SAP system.
- This code uses OLE automation, which requires Microsoft Excel to be installed and configured properly on the machine where the SAP system is running.

## Troubleshooting

- **Excel Application Not Visible**: Ensure that the Excel application is properly installed and not blocked by security settings.
- **Path Issues**: Verify that the paths provided are correct and accessible.

## License

This code is provided "as-is" without warranties or guarantees. Use it at your own risk. 

## Contact

For issues or questions, please contact [Your Contact Information].


Feel free to customize the README to better suit your needs or add any additional information that might be useful.
