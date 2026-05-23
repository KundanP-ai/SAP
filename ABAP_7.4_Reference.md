# ABAP 7.4 Reference Guide

## Table of Contents
1. [CDS Annotations](#cds-annotations)
2. [Append Operations](#append-operations)
3. [For Loop Patterns](#for-loop-patterns)
4. [Value Operations](#value-operations)
5. [Select Statements](#select-statements)
6. [Read Table Operations](#read-table-operations)
7. [Move Values](#move-values)
8. [Contains and String Operations](#contains-and-string-operations)
9. [Concatenation](#concatenation)
10. [Loop with Field Symbols](#loop-with-field-symbols)
11. [OData Methods](#odata-methods)
12. [Message Handling](#message-handling)
13. [Advanced Operations](#advanced-operations)

---

## CDS Annotations

### Foreign Key Association with Value Help
```abap
@ObjectModel.foreignKey.association: '_storage'
@Consumption.valueHelpDefinition: [
  {
    entity: { name: 'I_Storage_VH', element: 'LGORT' }
  }
]
```

### SPROXY Documentation
```
Note: SPROXY object documentation can be accessed via SPROXY transaction code,
not from SE24 transaction.
```

---

## Append Operations

### Append to Table with VALUE Constructor
```abap
APPEND VALUE #( col1 = sy-index col2 = sy-index ** 2 ) TO wa_tab.
```

### Append String with Variable Interpolation
```abap
APPEND |{ gi_qals-stat01 } { '-' } { gi_qave-vdatum DATE = USER }| TO gv_titleval.
```

---

## For Loop Patterns

### Simple For Loop
```abap
DATA(lt_stocktype) = VALUE #( FOR ls_line IN lt_config ( ls_line-zto ) ).
```

### For Loop with CORRESPONDING
```abap
itab_target = CORRESPONDING #( itab_source MAPPING field2 = field1 field4 = field3 ).
```

### For Loop with WHERE and CORRESPONDING
```abap
DATA(lt_outputyps) = VALUE efg_tab_ranges( 
  FOR ls_data IN lt_data 
  WHERE ( variable_name = k_pu_outpty )
  ( CORRESPONDING #( ls_data MAPPING option = opti ) ) 
).
```

### For Loop with Table Type
```abap
DATA(lt_lang) = VALUE fkk_rt_sprsl( 
  FOR lwa_config IN lt_ztbc 
  WHERE ( variable_name = 'LANGU' )
  ( 
    sign = 'I' 
    option = 'EQ' 
    low = lwa_config-low 
  ) 
).
```

### For Loop with Sort Table
```abap
DATA(lt_sort) = VALUE abap_sortorder_tab( 
  FOR lwa_sort IN it_order 
  ( 
    name = lwa_sort-property
    descending = COND #( 
      WHEN lwa_sort-order EQ TEXT-001 
      THEN abap_true 
      ELSE abap_false 
    ) 
  ) 
).

SORT et_entityset BY (lt_sort).
```

### For Loop with Complex Condition
```abap
ot_chem = VALUE ztt_vccont_chem( 
  FOR lwa_val_char IN lt_val_char 
  WHERE ( charact = 'ZFM_VACUUM_MEDIA' )
  ( 
    chemname = lwa_val_char-value_char 
    chemsym = lwa_val_char-value_neutral
    special_inst = VALUE #( 
      lt_val_char[ 
        charact = 'ZFM_VACUUM_HAZARD' 
        value_neutral = lwa_val_char-value_neutral 
      ]-value_char 
    ) 
  ) 
).
```

### For Loop with IN Condition
```abap
ot_equnr = VALUE dfps_tb_lm_equnr( 
  FOR lwa_line1 IN lt_eqplist
  WHERE ( equicatgry IN lt_eqcat ) 
  ( equnr = lwa_line1-equipment ) 
).
```

---

## Value Operations

### VALUE with Conditional WHEN-ELSE
```abap
i_symbols = VALUE #( 
  ( 
    key = '&EMAIL_TO&'
    value = COND #( 
      WHEN i_rectype EQ 'QC' THEN |Dear Quality Control|
      WHEN i_rectype EQ 'LC' THEN |Dear Line Control| 
    ) 
  )
  ( 
    key = '&HEADER_FIELD&'
    value = COND #( 
      WHEN sy-cprog EQ g_cprog THEN |Expiry Date: { i_mseg-vfdat DATE = USER }|
      ELSE |Quality Issue: { i_mseg-ablad }| 
    ) 
  )
).
```

---

## Select Statements

### SELECT SINGLE into Multiple Variables
```abap
SELECT SINGLE herst, typbz 
FROM equi 
INTO (@DATA(l_herst), @DATA(typbz))
WHERE equnr = @iv_equnr.
```

---

## Read Table Operations

### Read Table by Index
```abap
g_faxno = VALUE #( lt_fax[ 1 ]-fax_no OPTIONAL ).
```

### Read Table by Variable Condition
```abap
DATA(lv_attach) = lt_configuration[ 
  variable_name = 'EMAIL_ATTACH' 
  add_search = ud_data-vcode 
]-low.

gv_adobe = VALUE #( 
  lt_ztbc[ 
    variable_name = 'ADOBE_ATTACH' 
    sign = 'I' 
    opti = 'EQ' 
  ]-low 
  OPTIONAL 
).
```

### Read Table with Conditional Value
```abap
os_fetch-failure = COND #( 
  WHEN VALUE #( 
    lt_val_char[ charact = 'ZFM_VACUUM_RUN_TO_FAIL' ]-value_char 
    OPTIONAL 
  ) EQ k_yes 
  THEN abap_true 
  ELSE abap_false 
).
```

### Read Table Work Area (WA)
```abap
DATA(ls_flight) = it_flights [ carrid = 'SQ' connid = '0026' ].
```

---

## Move Values

### Move to Work Area with VALUE
```abap
gwa_zqmheader = VALUE #( 
  method = TEXT-011 
  l_spec = TEXT-012 
  u_spec = TEXT-013 
).
```

### Get Line Index
```abap
DATA(lv_idx) = line_index( gi_text[ table_line = 'Defect(s):' ] ).
```

### Check if Line Exists
```abap
IF line_exists( lt_ztqmr[ werks = lwa_mseg-werks ] ).
  DATA(l_iretu) = abap_true.
ENDIF.
```

### Get Number of Lines
```abap
DATA : l_lines TYPE i.
l_lines = lines( lt_lines ).
```

---

## Contains and String Operations

### Check Contains Pattern
```abap
IF is_mat_comp-code CP '*E0013*'.
  " Do something
ENDIF.
```

### Remove Leading Zeros
```abap
DATA(l_equnr) = |{ is_print-equnr ALPHA = IN }|.
```

---

## Concatenation

### Concatenate with Space
```abap
DATA(l_flname) = |{ is_print-tplnr } | & | | & |{ is_print-pltxt }|.

g_endline = | { TEXT-018 } { ' ' } { gwa_qals-prueflos } |.
```

### Concatenate Lines of Table
```abap
CONCATENATE LINES OF lt_pump_type INTO lwa_final-pump_type SEPARATED BY ','.
```

---

## Loop with Field Symbols

### Loop with Field Symbol Assignment
```abap
LOOP AT gi_files ASSIGNING FIELD-SYMBOL(<fs_files>) 
  WHERE file = TEXT-224.
  
  IF <fs_files> IS ASSIGNED AND <fs_files>-file = TEXT-224.
    " Do something
  ENDIF.
  
ENDLOOP.
```

---

## OData Methods

### Get Filter Information
```abap
DATA lo_filter TYPE REF TO /iwbep/if_mgw_req_filter.
DATA lt_filter_select_options TYPE /iwbep/t_mgw_select_option.
DATA lv_filter_str TYPE string.

lo_filter = io_tech_request_context->get_filter( ).
lt_filter_select_options = lo_filter->get_filter_select_options( ).
lv_filter_str = lo_filter->get_filter_string( ).
```

---

## Message Handling

### Display Message from System Variables
```abap
MESSAGE ID sy-msgid TYPE sy-msgty NUMBER sy-msgno
  WITH sy-msgv1 sy-msgv2 sy-msgv3 sy-msgv4.
```

### Add Message to Container and Raise Exception
```abap
" Instantiate Message Container
DATA : lo_msg TYPE REF TO /iwbep/if_message_container.

CALL METHOD me->/iwbep/if_mgw_conv_srv_runtime~get_message_container
  RECEIVING
    ro_message_container = lo_msg.

CALL METHOD lo_msg->add_message
  EXPORTING
    iv_msg_type   = /iwbep/cl_cos_logger=>error
    iv_msg_id     = 'ZFM_VACUUM'
    iv_msg_number = '003'.

" Raising Exception
RAISE EXCEPTION TYPE /iwbep/cx_mgw_tech_exception
  EXPORTING
    message_container = lo_msg.
```

### Log Message in Application Log
```abap
" Log message in the application log
me->/iwbep/if_sb_dpc_comm_services~log_message(
  EXPORTING
    iv_msg_type   = 'E'
    iv_msg_id     = 'ZFM_VACUUM'
    iv_msg_number = 003 ).

" Raise Exception
RAISE EXCEPTION TYPE /iwbep/cx_mgw_tech_exception
  EXPORTING
    textid = /iwbep/cx_mgw_tech_exception=>internal_error.
```

---

## Advanced Operations

### REDUCE with Conditional Concatenation
```abap
lwa_final-pump_type = REDUCE zfm_particularity( 
  INIT l_string TYPE string 
  FOR lwa_pump_type IN lt_pump_type
  NEXT l_string = COND zfm_particularity( 
    WHEN l_string IS INITIAL 
    THEN lwa_pump_type-pump_type
    ELSE l_string && ',' && lwa_pump_type-pump_type 
  ) 
).
```

### Concatenate Table Lines with Separator
```abap
CONCATENATE LINES OF lt_pump_type INTO lwa_final-pump_type SEPARATED BY ','.
```

---

## Quick Reference Tips

- Use `VALUE #()` constructor for efficient table building
- Use `CORRESPONDING #()` to map fields between structures
- Use field symbols (`<fs_variable>`) for loop performance
- Use inline comments for complex conditional logic
- Always use `OPTIONAL` when reading tables to avoid exceptions
- Use string interpolation `|{ variable }|` for flexible string formatting
