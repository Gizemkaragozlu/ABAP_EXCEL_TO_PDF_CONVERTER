CLASS  zcl_excel_pdf_converter DEFINITION
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
* OLE için gerekli include
    INCLUDE ole2incl.
* OLE nesne tutucuları
    DATA: lo_excel_application TYPE ole2_object, " Excel uygulama nesnesi
          lo_workbooks         TYPE ole2_object, " Çalışma kitapları koleksiyonu
          lo_workbook          TYPE ole2_object, " Tek bir çalışma kitabı
          lo_worksheet         TYPE ole2_object, " Çalışma sayfası nesnesi
          lo_cells             TYPE ole2_object, " Hücre nesnesi
          lo_columns           TYPE ole2_object. " Sütun nesnesi
    " Excel uygulamasını oluştur
    CREATE OBJECT lo_excel_application 'EXCEL.APPLICATION'.
    " Excel uygulamasını görünür yap
    SET PROPERTY OF lo_excel_application 'Visible' = 1.
    " Çalışma kitapları koleksiyonunu al
    CALL METHOD OF lo_excel_application 'Workbooks' = lo_workbooks.
    " Belirtilen çalışma kitabını aç
    CALL METHOD OF lo_workbooks 'Open'
      EXPORTING
        #1 = source.
    " İlk çalışma kitabını al
    CALL METHOD OF lo_workbooks 'Item' = lo_workbook
    EXPORTING
      #1 = 1.
    " İlk çalışma sayfasını al
    CALL METHOD OF lo_workbook 'Worksheets' = lo_worksheet
    EXPORTING
      #1 = 1.
    " Hücreleri al
    CALL METHOD OF lo_worksheet 'Cells' = lo_cells.
    " Sütunları seç ve genişliklerini otomatik ayarla
    CALL METHOD OF lo_cells 'EntireColumn' = lo_columns.
    CALL METHOD OF lo_columns 'AutoFit'.
*.CenterHeader = ""                " Üst Merkez Başlığı = ""
*.RightHeader = ""                 " Üst Sağ Başlık = ""
*.LeftFooter = ""                  " Alt Sol Altbilgi = ""
*.CenterFooter = ""                " Alt Merkez Altbilgi = ""
*.RightFooter = ""                 " Alt Sağ Altbilgi = ""
*.PrintHeadings = False            " Başlıkları Yazdır = Yanlış
*.PrintGridlines = False           " Izgara Çizgilerini Yazdır = Yanlış
*.CenterHorizontally = False       " Yatayda Merkezle = Yanlış
*.CenterVertically = False         " Düşeyde Merkezle = Yanlış
*.Draft = False                    " Taslak Modu = Yanlış
*.BlackAndWhite = False            " Siyah-Beyaz Yazdır = Yanlış
*.Zoom = False                     " Yakınlaştırma = Yanlış
*.FitToPagesWide = 1               " Sayfaların Genişliğine Sığdır = 1
*.FitToPagesTall = 1               " Sayfaların Uzunluğuna Sığdır = 1
    " Sayfa düzeni ayarlarını yapılandır
    CALL METHOD OF lo_excel_application 'ExecuteExcel4Macro'
      EXPORTING
        #1 = 'PAGE.SETUP(,,,,,,,,,,,,{1,0})'.
    " Çalışma kitabını PDF formatında dışa aktar
    CALL METHOD OF lo_workbook 'ExportAsFixedFormat'
      EXPORTING
        #1 = 0        " xlTypePDF
        #2 = destination. " Çıktı PDF dosyasının yolu
    " Çalışma kitabını kaydetmeden kapat
    CALL METHOD OF lo_workbook 'Close'
      EXPORTING
        #1 = 0.       " Kaydetme = False
    " Excel uygulamasını görünür yap
    SET PROPERTY OF lo_excel_application 'Visible' = 1.
    " Excel uygulamasını kapat
    CALL METHOD OF lo_excel_application 'Quit'.
    " OLE nesnesini serbest bırak
    FREE OBJECT lo_excel_application.
  ENDMETHOD.


  METHOD get_instance.
    IF instance IS NOT BOUND.
      instance = NEW #( ).
    ENDIF.
    ro_instance = instance.
  ENDMETHOD.
ENDCLASS.
