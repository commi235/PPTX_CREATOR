CREATE OR REPLACE PACKAGE PPTX_CREATOR_PKG
  AUTHID CURRENT_USER
AS 
  TYPE t_single_slide IS RECORD
    ( slide_num NUMBER
    , slide_id NUMBER
    , relation_id NUMBER
    , content_type VARCHAR2(200) := 'application/vnd.openxmlformats-officedocument.presentationml.slide+xml'
    , relation_type VARCHAR2(200) := 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide'
    , slide_data CLOB
    );

  TYPE t_all_slides IS TABLE OF t_single_slide INDEX BY PLS_INTEGER;

  TYPE t_replace_values_tab IS TABLE OF t_str_array INDEX BY PLS_INTEGER;

  FUNCTION convert_template ( p_template IN BLOB
                            , p_replace_patterns IN t_str_array
                            , p_replace_values IN t_replace_values_tab 
                            )
    RETURN BLOB;

END PPTX_CREATOR_PKG;
/