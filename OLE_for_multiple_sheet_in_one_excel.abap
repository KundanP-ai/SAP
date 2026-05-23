*----------------------------------------------------------------------*
* FORM: f_download_multiple
* PURPOSE: Creates an Excel workbook with multiple sheets (PIM, PAYINFO, PAYDTL)
*          and populates them with data from different internal tables
* AUTHOR: SAP Developer
* DATE: 2023
*----------------------------------------------------------------------*

FORM f_download_multiple .

  DATA: lt_210 TYPE STANDARD TABLE OF zusxx_str_arn_equo1_210,
        lt_220 TYPE STANDARD TABLE OF zusxx_str_arn_equo1_220,
        lt_202 TYPE STANDARD TABLE OF zusxx_str_arn_equo1_202.

  CREATE OBJECT gv_application 'excel.application'.

  SET PROPERTY OF gv_application 'Visible' = 0.
  CALL METHOD OF gv_application 'Workbooks' = gv_books.
  CALL METHOD OF gv_books 'Add' = gv_book.
  CALL METHOD OF gv_book 'WORKSHEETS' = gv_sheet.
  CALL METHOD OF gv_sheet 'Add'.
  CALL METHOD OF gv_sheet 'Add'.

  MOVE gt_202 TO lt_202.
  MOVE gt_210 TO lt_210.
  MOVE gt_220 TO lt_220.

  PERFORM f_sheet1 TABLES lt_202.
  PERFORM f_sheet2 TABLES lt_210.
  PERFORM f_sheet3 TABLES lt_220.

  CALL METHOD OF gv_book 'SaveAs'
    EXPORTING
      #1 = p_xl
      #2 = 1.

  CALL METHOD OF gv_book 'SAVE'.
  CALL METHOD OF gv_book 'CLOSE'.
  SET PROPERTY OF gv_book 'visible' = 0.

  CALL METHOD OF gv_books 'Application.Quit()'.

  FREE OBJECT:gv_name,gv_cell,gv_column,gv_sheet,gv_book,gv_books,gv_application.

ENDFORM.

*----------------------------------------------------------------------*
* FORM: f_fill_cell
* PURPOSE: Fills a single cell in the Excel sheet with a value
* PARAMETERS:
*   lv_row - Row number
*   lv_col - Column number
*   lv_val - Value to fill
*----------------------------------------------------------------------*

FORM f_fill_cell USING lv_row lv_col lv_val.
  CALL METHOD OF gv_sheet 'cells' = gv_cell NO FLUSH EXPORTING #1 = lv_row #2 = lv_col.
  SET PROPERTY OF gv_cell 'value' = lv_val.
  FREE OBJECT gv_cell NO FLUSH.
ENDFORM.

*----------------------------------------------------------------------*
* FORM: f_sheet1
* PURPOSE: Creates and populates the 'PIM' worksheet with data from
*          the ZUSXX_STR_ARN_EQUO1_202 structure
* PARAMETERS:
*   lv_itab202 - Internal table containing data for PIM sheet
*----------------------------------------------------------------------*

FORM f_sheet1 TABLES lv_itab202 STRUCTURE zusxx_str_arn_equo1_202.
  gv_name = 'PIM'.
  gv_no = gv_no + 1.
  CALL METHOD OF gv_book 'worksheets' = gv_sheet NO FLUSH EXPORTING #1 = gv_no.
  SET PROPERTY OF gv_sheet 'Name' = gv_name NO FLUSH.
  PERFORM f_fill_sheet1 TABLES lv_itab202 USING gv_no gv_name.
  CALL METHOD OF gv_sheet 'Columns' = gv_column.
  FREE OBJECT gv_sheet.
  CALL METHOD OF gv_column 'Autofit'.
  FREE OBJECT gv_column.
ENDFORM.

*----------------------------------------------------------------------*
* FORM: f_fill_sheet1
* PURPOSE: Populates the PIM sheet with field headers and data rows
* PARAMETERS:
*   lt_itab1 - Internal table with PIM data
*   gv_no   - Sheet number
*   gv_name - Sheet name
*----------------------------------------------------------------------*

FORM f_fill_sheet1
TABLES lt_itab1 STRUCTURE zusxx_str_arn_equo1_202
USING gv_no gv_name.

  DATA: lv_index  TYPE i,
        lt_struct TYPE TABLE OF dfies WITH HEADER LINE,
        lt_fields TYPE STANDARD TABLE OF string WITH HEADER LINE.

  CALL FUNCTION 'DDIF_FIELDINFO_GET'
    EXPORTING
      tabname        = 'ZUSXX_STR_ARN_EQUO1_202'
      langu          = sy-langu
    TABLES
      dfies_tab      = lt_struct
    EXCEPTIONS
      not_found      = 1
      internal_error = 2
      OTHERS         = 3.

  lv_index = 1.
  gv_row = 1.
  LOOP AT lt_struct.
    PERFORM f_fill_cell USING gv_row lv_index lt_struct-fieldtext.
    lv_index = lv_index + 1.
    lt_fields = lt_struct-fieldname.
    APPEND lt_fields.
  ENDLOOP.
  CLEAR:  lv_index.

  LOOP AT lt_itab1 INTO DATA(ls_itab1).
    gv_row = gv_row + 1.
    lv_index = 1.
    LOOP AT lt_fields INTO DATA(ls_field).
      ASSIGN COMPONENT ls_field OF STRUCTURE ls_itab1 TO FIELD-SYMBOL(<field_value>).
      IF sy-subrc = 0.
        PERFORM f_fill_cell USING gv_row lv_index <field_value>.
      ENDIF.
      lv_index = lv_index + 1.
    ENDLOOP.
  ENDLOOP.

  UNASSIGN:  <field_value>.
  CLEAR:  lv_index, gv_row.
  REFRESH:  lt_fields,lt_itab1.
ENDFORM.

*----------------------------------------------------------------------*
* FORM: f_sheet2
* PURPOSE: Creates and populates the 'PAYINFO' worksheet with data from
*          the ZUSXX_STR_ARN_EQUO1_210 structure
* PARAMETERS:
*   lv_itab210 - Internal table containing data for PAYINFO sheet
*----------------------------------------------------------------------*

FORM f_sheet2 TABLES lv_itab210 STRUCTURE zusxx_str_arn_equo1_210.
  gv_name = 'PAYINFO'.
  gv_no = gv_no + 1.
  CALL METHOD OF gv_book 'worksheets' = gv_sheet NO FLUSH EXPORTING #1 = gv_no.
  SET PROPERTY OF gv_sheet 'Name' = gv_name NO FLUSH.
  PERFORM f_fill_sheet2 TABLES lv_itab210 USING gv_no gv_name.
  CALL METHOD OF gv_sheet 'Columns' = gv_column.
  FREE OBJECT gv_sheet.
  CALL METHOD OF gv_column 'Autofit'.
  FREE OBJECT gv_column.
ENDFORM.

*----------------------------------------------------------------------*
* FORM: f_fill_sheet2
* PURPOSE: Populates the PAYINFO sheet with field headers and data rows
* PARAMETERS:
*   lt_itab1 - Internal table with PAYINFO data
*   gv_no   - Sheet number
*   gv_name - Sheet name
*----------------------------------------------------------------------*

FORM f_fill_sheet2
TABLES lt_itab1 STRUCTURE zusxx_str_arn_equo1_210
USING gv_no gv_name.

  DATA: lv_index  TYPE i,
        lt_struct TYPE TABLE OF dfies WITH HEADER LINE,
        lt_fields TYPE STANDARD TABLE OF string WITH HEADER LINE.

  CALL FUNCTION 'DDIF_FIELDINFO_GET'
    EXPORTING
      tabname        = 'ZUSXX_STR_ARN_EQUO1_210'
      langu          = sy-langu
    TABLES
      dfies_tab      = lt_struct
    EXCEPTIONS
      not_found      = 1
      internal_error = 2
      OTHERS         = 3.

  lv_index = 1.
  gv_row = 1.
  LOOP AT lt_struct.
    PERFORM f_fill_cell USING gv_row lv_index lt_struct-fieldtext.
    lv_index = lv_index + 1.
    lt_fields = lt_struct-fieldname.
    APPEND lt_fields.
  ENDLOOP.
  CLEAR:  lv_index.

  LOOP AT lt_itab1 INTO DATA(ls_itab1).
    gv_row = gv_row + 1.
    lv_index = 1.
    LOOP AT lt_fields INTO DATA(ls_field).
      ASSIGN COMPONENT ls_field OF STRUCTURE ls_itab1 TO FIELD-SYMBOL(<field_value>).
      IF sy-subrc = 0.
        PERFORM f_fill_cell USING gv_row lv_index <field_value>.
      ENDIF.
      lv_index = lv_index + 1.
    ENDLOOP.
  ENDLOOP.

  UNASSIGN:  <field_value>.
  CLEAR:  lv_index, gv_row.
  REFRESH:  lt_fields,lt_itab1.
ENDFORM.

*----------------------------------------------------------------------*
* FORM: f_sheet3
* PURPOSE: Creates and populates the 'PAYDTL' worksheet with data from
*          the ZUSXX_STR_ARN_EQUO1_220 structure
* PARAMETERS:
*   lv_itab220 - Internal table containing data for PAYDTL sheet
*----------------------------------------------------------------------*

FORM f_sheet3 TABLES lv_itab220 STRUCTURE zusxx_str_arn_equo1_220.
  gv_name = 'PAYDTL'.
  gv_no = gv_no + 1.
  CALL METHOD OF gv_book 'worksheets' = gv_sheet NO FLUSH EXPORTING #1 = gv_no.
  SET PROPERTY OF gv_sheet 'Name' = gv_name NO FLUSH.
  PERFORM f_fill_sheet3 TABLES lv_itab220 USING gv_no gv_name.
  CALL METHOD OF gv_sheet 'Columns' = gv_column.
  FREE OBJECT gv_sheet.
  CALL METHOD OF gv_column 'Autofit'.
  FREE OBJECT gv_column.
ENDFORM.

*----------------------------------------------------------------------*
* FORM: f_fill_sheet3
* PURPOSE: Populates the PAYDTL sheet with field headers and data rows
* PARAMETERS:
*   lt_itab1 - Internal table with PAYDTL data
*   gv_no   - Sheet number
*   gv_name - Sheet name
*----------------------------------------------------------------------*

FORM f_fill_sheet3
TABLES lt_itab1 STRUCTURE zusxx_str_arn_equo1_220
USING gv_no gv_name.

  DATA: lv_index  TYPE i,
        lt_struct TYPE TABLE OF dfies WITH HEADER LINE,
        lt_fields TYPE STANDARD TABLE OF string WITH HEADER LINE.

  CALL FUNCTION 'DDIF_FIELDINFO_GET'
    EXPORTING
      tabname        = 'ZUSXX_STR_ARN_EQUO1_220'
      langu          = sy-langu
    TABLES
      dfies_tab      = lt_struct
    EXCEPTIONS
      not_found      = 1
      internal_error = 2
      OTHERS         = 3.

  lv_index = 1.
  gv_row = 1.
  LOOP AT lt_struct.
    PERFORM f_fill_cell USING gv_row lv_index lt_struct-fieldtext.
    lv_index = lv_index + 1.
    lt_fields = lt_struct-fieldname.
    APPEND lt_fields.
  ENDLOOP.
  CLEAR:  lv_index.

  LOOP AT lt_itab1 INTO DATA(ls_itab1).
    gv_row = gv_row + 1.
    lv_index = 1.
    LOOP AT lt_fields INTO DATA(ls_field).
      ASSIGN COMPONENT ls_field OF STRUCTURE ls_itab1 TO FIELD-SYMBOL(<field_value>).
      IF sy-subrc = 0.
        PERFORM f_fill_cell USING gv_row lv_index <field_value>.
      ENDIF.
      lv_index = lv_index + 1.
    ENDLOOP.
  ENDLOOP.

  UNASSIGN:  <field_value>.
  CLEAR:  lv_index, gv_row.
  REFRESH:  lt_fields,lt_itab1.
ENDFORM.
