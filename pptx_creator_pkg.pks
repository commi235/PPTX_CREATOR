CREATE OR REPLACE PACKAGE PPTX_CREATOR_PKG
  AUTHID CURRENT_USER
AS 
/**
* Package provides functions to generate PPTX files based on a template.
* @headcom
*/

  c_sub_character CONSTANT VARCHAR2(1) := '#';

  TYPE t_vc_value_row IS TABLE OF VARCHAR2(4000);

  /**
  * Simple two-dimensional array with values,
  * which should replace the substitution strings in the template.
  */
  TYPE t_vc_value_tab IS TABLE OF t_vc_value_row INDEX BY PLS_INTEGER;
  
  TYPE t_replace_value IS RECORD ( varchar_value VARCHAR2(4000)
                                 , number_value NUMBER
                                 , date_value DATE
                                 , format_mask VARCHAR2(100)
                                 );
  
  TYPE t_replace_value_row IS TABLE OF t_replace_value INDEX BY VARCHAR2(30);

  /**
  * Array of the values indexed by the substitution string.
  */
  TYPE t_replace_value_tab IS TABLE OF t_replace_value_row INDEX BY PLS_INTEGER;

  FUNCTION convert_template ( p_template IN BLOB
                            , p_replace_patterns IN t_vc_value_row
                            , p_replace_values IN t_vc_value_tab 
                            )
    RETURN BLOB;
    
  FUNCTION convert_template ( p_template IN BLOB
                            , p_replace_name_value IN t_replace_value_tab
                            , p_substitution_pattern IN VARCHAR2 DEFAULT c_sub_character
                            )
    RETURN BLOB;
    
END PPTX_CREATOR_PKG;
/
