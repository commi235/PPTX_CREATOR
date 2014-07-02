CREATE OR REPLACE PACKAGE PPTX_CREATOR_PKG
  AUTHID CURRENT_USER
AS 
/**
* Package provides functions to generate PPTX files based on a template.
* @headcom
*/

  /**
  * Simple two-dimensional array with values,
  * which should replace the substitution strings in the template.
  */
  TYPE t_replace_values_tab IS TABLE OF t_str_array INDEX BY PLS_INTEGER;

  /**
  * Array of the values indexed by the substitution string.
  */
  TYPE t_name_value_tab IS TABLE OF t_str_array INDEX BY VARCHAR2(30);

  FUNCTION convert_template ( p_template IN BLOB
                            , p_replace_patterns IN t_str_array
                            , p_replace_values IN t_replace_values_tab 
                            )
    RETURN BLOB;

END PPTX_CREATOR_PKG;
/
