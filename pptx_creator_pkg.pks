CREATE OR REPLACE PACKAGE PPTX_CREATOR_PKG
  AUTHID CURRENT_USER
AS 
/**
* Package provides functions to generate PPTX files based on a template.
* @headcom
*/

  /** Default value for substitution string */
  c_enclose_character CONSTANT VARCHAR2(1) := '#';

  /**
  * Simple one-dimensional array of varchar.
  */
  TYPE t_vc_value_row IS TABLE OF VARCHAR2(4000);

  /**
  * Simple two-dimensional array with values,
  * which should replace the substitution strings in the template.
  */
  TYPE t_vc_value_tab IS TABLE OF t_vc_value_row INDEX BY PLS_INTEGER;
  
  /**
  * Complex type to hold the value for a single replacement.
  * Only one of the *_value parameters should be set.
  * Order of evaluation is: varchar_value -> number_value -> date_value -> clob_value
  * If number or date is used, format_mask needs to be filled also.
  * @param varchar_value Substitution value as VARCHAR2
  * @param number_value Substitution value as NUMBER (format_mask mandatory)
  * @param date_value Substitution value as DATE (format_mask mandatory)
  * @param clob_value Substitution value as CLOB (use with caution)
  */
  TYPE t_replace_value IS RECORD ( varchar_value VARCHAR2(32767)
                                 , number_value NUMBER
                                 , date_value DATE
                                 , clob_value CLOB
                                 , format_mask VARCHAR2(100)
                                 );

  /**
  * Array of the values indexed by the substitution string.
  * This acts as the value list for a single slide.
  */  
  TYPE t_replace_value_row IS TABLE OF t_replace_value INDEX BY VARCHAR2(32767);

  /**
  * Array of named value lists.
  * This acts as a table holding all slide lists.
  */
  TYPE t_replace_value_tab IS TABLE OF t_replace_value_row INDEX BY PLS_INTEGER;

  /**
  * Converts the given template into a slide-deck.
  * This is the simpler interface just taking varchar values.
  * @param p_template The PPTX file which should serve as the template.
  * @param p_replace_patterns One-dimensional array of the substitution strings to be replaced.
  * @param p_replace_values Two-dimensional array with the values. (Every row must be ordered same as p_replace_patterns)
  * @return New PPTX file with substitution strings replaced and template slide duplicated as needed.
  */
  FUNCTION convert_template ( p_template IN BLOB
                            , p_replace_patterns IN t_vc_value_row
                            , p_replace_values IN t_vc_value_tab 
                            )
    RETURN BLOB;
    
  /**
  * Converts the given template into a slide-deck.
  * This overload uses a more sophisticated input, which can deal with varchar, number and date values.
  * @param p_template The PPTX file which should serve as the template.
  * @param p_replace_name_value Two-dimensional array holding all replacement data.
  * @param p_enclose_char The character surrounding substitution strings.
  * @return New PPTX file with substitution strings replaced and template slide duplicated as needed.
  */
  FUNCTION convert_template ( p_template IN BLOB
                            , p_replace_name_value IN t_replace_value_tab
                            , p_enclose_char IN VARCHAR2 DEFAULT c_enclose_character
                            )
    RETURN BLOB;

  /**
  * Converts the given template into a slide-deck.
  * This overload allows to pass in a ref cursor
  * which is parsed and the column aliases are used as substitution strings.
  * @param p_template The PPTX file which should serve as the template.
  * @param p_cursor A ref cursor defining the data to use.
  * @param p_enclose_char The character surrounding substitution strings.
  */
  FUNCTION convert_template ( p_template IN BLOB
                            , p_cursor IN OUT NOCOPY sys_refcursor
                            , p_enclose_char IN VARCHAR2 DEFAULT c_enclose_character
                            )
    RETURN BLOB;

END PPTX_CREATOR_PKG;
/
