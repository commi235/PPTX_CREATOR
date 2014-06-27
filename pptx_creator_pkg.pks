CREATE OR REPLACE PACKAGE PPTX_CREATOR_PKG
  AUTHID CURRENT_USER
AS 

  TYPE t_replace_values_tab IS TABLE OF t_str_array INDEX BY PLS_INTEGER;

  FUNCTION convert_template ( p_template IN BLOB
                            , p_replace_patterns IN t_str_array
                            , p_replace_values IN t_replace_values_tab 
                            )
    RETURN BLOB;

END PPTX_CREATOR_PKG;
/
