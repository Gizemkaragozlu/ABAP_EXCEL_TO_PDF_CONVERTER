CLASS zcl_excel_pdf_converter DEFINITION
  PUBLIC
  FINAL
  CREATE PUBLIC .
  PUBLIC SECTION.
    CLASS-METHODS:
      get_instance
        RETURNING
          VALUE(ro_instance) TYPE REF TO zcl_excel_pdf_converter.
    METHODS:
      convert
        IMPORTING
          source      TYPE string
          destination TYPE string.
  PRIVATE SECTION.
    CLASS-DATA: instance TYPE REF TO zcl_excel_pdf_converter.
ENDCLASS.



CLASS ZCL_EXCEL_PDF_CONVERTER IMPLEMENTATION.


  METHOD convert.
* Include necessary for OLE
    INCLUDE ole2incl.
* OLE object holders
    DATA: lo_excel_application TYPE ole2_object, " Excel application object
          lo_workbooks         TYPE ole2_object, " Collection of workbooks
          lo_workbook          TYPE ole2_object, " Single workbook object
          lo_worksheet         TYPE ole2_object, " Worksheet object
          lo_cells             TYPE ole2_object, " Cells object
          lo_columns           TYPE ole2_object. " Columns object
    " Create Excel application
    CREATE OBJECT lo_excel_application 'EXCEL.APPLICATION'.
    " Make Excel application visible
    SET PROPERTY OF lo_excel_application 'Visible' = 1.
    " Get the collection of workbooks
    CALL METHOD OF lo_excel_application 'Workbooks' = lo_workbooks.
    " Open the specified workbook
    CALL METHOD OF lo_workbooks 'Open'
      EXPORTING
        #1 = source.
    " Get the first workbook
    CALL METHOD OF lo_workbooks 'Item' = lo_workbook
    EXPORTING
      #1 = 1.
    " Get the first worksheet
    CALL METHOD OF lo_workbook 'Worksheets' = lo_worksheet
    EXPORTING
      #1 = 1.
    " Get cells
    CALL METHOD OF lo_worksheet 'Cells' = lo_cells.
    " Select columns and auto-fit their widths
    CALL METHOD OF lo_cells 'EntireColumn' = lo_columns.
    CALL METHOD OF lo_columns 'AutoFit'.
*.CenterHeader = ""                " Center Header = ""
*.RightHeader = ""                 " Right Header = ""
*.LeftFooter = ""                  " Left Footer = ""
*.CenterFooter = ""                " Center Footer = ""
*.RightFooter = ""                 " Right Footer = ""
*.PrintHeadings = False            " Print Headings = False
*.PrintGridlines = False           " Print Gridlines = False
*.CenterHorizontally = False       " Center Horizontally = False
*.CenterVertically = False         " Center Vertically = False
*.Draft = False                    " Draft Mode = False
*.BlackAndWhite = False            " Black and White Print = False
*.Zoom = False                     " Zoom = False
*.FitToPagesWide = 1               " Fit to Width of Pages = 1
*.FitToPagesTall = 1               " Fit to Height of Pages = 1
    " Configure page layout settings
    CALL METHOD OF lo_excel_application 'ExecuteExcel4Macro'
      EXPORTING
        #1 = 'PAGE.SETUP(,,,,,,,,,,,,{1,0})'.
    " Export the workbook as a PDF
    CALL METHOD OF lo_workbook 'ExportAsFixedFormat'
      EXPORTING
        #1 = 0        " xlTypePDF
        #2 = destination. " Path for the output PDF
    " Close the workbook without saving
    CALL METHOD OF lo_workbook 'Close'
      EXPORTING
        #1 = 0.       " Save changes = False
    " Make the Excel application visible
    SET PROPERTY OF lo_excel_application 'Visible' = 1.
    " Quit the Excel application
    CALL METHOD OF lo_excel_application 'Quit'.
    " Release the OLE object
    FREE OBJECT lo_excel_application.
  ENDMETHOD.


  METHOD get_instance.
    IF instance IS NOT BOUND.
      instance = NEW #( ).
    ENDIF.
    ro_instance = instance.
  ENDMETHOD.
ENDCLASS.

