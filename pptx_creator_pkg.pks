CREATE OR REPLACE PACKAGE PPTX_CREATOR_PKG
  AUTHID CURRENT_USER
AS 
/**
* Package provides functions to generate PPTX files based on a template.
* @headcom
*/

  $IF $$apex_installed IS NULL $THEN
    $error q'[Set CCFLAG apex_installed to either 0 or 1.
    e.g. ALTER SESSION SET plsql_ccflags = 'apex_installed:1';
    ]' $end
  $END

  /**
  * Returns the current version of this package.
  * @return Current version in format major.minor.patch
  */
  FUNCTION get_version
    RETURN VARCHAR2;


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
  * Array used to map column position by column alias.
  */
  TYPE t_pi_by_vc IS TABLE OF pls_integer INDEX BY VARCHAR2(32767);
  
  /**
  * Array used to map column alias to column position.
  */
  TYPE t_vc_by_pi IS TABLE OF VARCHAR2(32767) INDEX BY PLS_INTEGER;
  
  /**
  * Used to hold column position to column name mapping.
  */
  TYPE t_col_info_rec IS RECORD ( col_position NUMBER, col_name VARCHAR2(4000) );
  
  /**
  * Array of column position to column name mappings.
  */
  TYPE t_col_info_tab IS TABLE OF t_col_info_rec;
  
  /**
  * Used to hold substitution variable name to column alias mapping.
  */
  TYPE t_col_map_rec IS record ( sub_name VARCHAR2(4000), map_col VARCHAR2(4000) );
  
  /**
  * Array of substitution variable to column alias mapping.
  */
  TYPE t_col_map_tab IS TABLE OF t_col_map_rec;
  
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
  * Extracts the substitution variables from the given template.
  * @param  p_template     The PPTX file used as template.
  * @param  p_enclose_char The enclosing character to identify the substitution variables.
  * @return Array of the substitution variables without enclosing character
  */
  FUNCTION get_substitutions( p_template IN BLOB
                            , p_enclose_char IN VARCHAR2 DEFAULT c_enclose_character
                            )
    RETURN t_vc_value_row;

  /**
  * Extracts the substitution variables from given template.
  * This overload is a pipelined table function for easy use within SQL.
  * @param  p_template     The PPTX file used as template.
  * @param  p_enclose_char The enclosing character to identify the substitution variables.
  * @return Array of the substitution variables without enclosing character
  */
  FUNCTION get_substitutions_tf( p_template IN BLOB
                               , p_enclose_char IN VARCHAR2 DEFAULT c_enclose_character
                               )
    RETURN t_vc_value_row PIPELINED;

  /**
  * Returns the SQL column names and their position in given SQL statement
  * @param  p_sql The SQL statement to parse.
  * @return Array of column number indexed by column name
  */
  FUNCTION get_sql_cols( p_sql IN VARCHAR2 )
    RETURN t_pi_by_vc;

  /**
  * Returns the column positions and names from given SQL statement.
  * Useful for creating LOVs.
  * @param  p_sql The SQL statement to parse.
  * @return Table with columns COL_POSITION and COL_NAME
  */
  FUNCTION get_sql_col_info( p_sql IN VARCHAR2 )
    RETURN t_col_info_tab PIPELINED;

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

-- APEX specific, not included if no APEX installed
$IF $$apex_installed = 1 $THEN

  /**
  * Converts a mapping string into the internal column substitution mapping format.
  * Input string needs to be of following format:
  * SUBSTITUTION1,COL_POSITION1:SUBSTITUTION2,COL_POSITION2:SUBSTITUTION3,COL_POSITION3
  * COL_POSITION needs to be whole numbers
  * @param  p_mapping String to parse
  * @return Array of substitution variable names indexed by column position
  */
  FUNCTION convert_col_sub_map( p_mapping IN VARCHAR2 )
    RETURN t_vc_by_pi;

  /**
  * Exracts substitution variables from template, columns from query.
  * Provides a table with substitution variable and candidate columns,
  * candidate columns are defaulted by comparing column alias and substitution variable.
  * @param p_template     The template file to be used.
  * @param p_enclose_char Character with which substitution variables are enclosed.
  * @param p_sql          SQL query from which to derive the columns.
  * @return               Table with substitution variable and column mapping.
  */
  FUNCTION get_sub_map( p_template IN BLOB
                      , p_enclose_char IN VARCHAR2 DEFAULT c_enclose_character
                      , p_sql IN VARCHAR2
                      )
    RETURN t_col_map_tab;
  
  /**
  * Extracts substitution variables from template, columns from query.
  * Provides a table with substitution variable and candidate columns,
  * candidate columns are defaulted by comparing column alias and substitution variable.
  * Can only be used from SQL as this is a pipelined table function.
  * @param p_template_tab Table name where templates are stored.
  * @param p_blob_col     Column in which the file is stored.
  * @param p_search_col   Column used to pinpoint selected template.
  * @param p_search_value Search value to pinpoint selected template.
  * @param p_enclose_char Character with which substitution variables are enclosed.
  * @param p_sql          SQL query from which to derive the columns.
  * @return               Table with substitution variable and column mapping.
  */
  FUNCTION get_sub_map_tf( p_template_tab IN VARCHAR2
                         , p_blob_col IN VARCHAR2
                         , p_search_col IN VARCHAR2
                         , p_search_value IN VARCHAR2
                         , p_enclose_char IN VARCHAR2 DEFAULT c_enclose_character
                         , p_sql IN VARCHAR2
                         )
    RETURN t_col_map_tab pipelined;

  /**
  * Extracts substitution variables from template and columns from query derived from interactive report.
  * Provides a table with substitution variable and candidate columns,
  * candidate columns are defaulted by comparing column alias and substitution variable.
  * Can only be used from SQL as this is a pipelined table function.
  * @param p_template_tab Table name where templates are stored.
  * @param p_blob_col     Column in which the file is stored.
  * @param p_search_col   Column used to pinpoint selected template.
  * @param p_search_value Search value to pinpoint selected template.
  * @param p_enclose_char Character with which substitution variables are enclosed.
  * @param p_app_id       Application ID of interactive report.
  * @param p_app_page_id  Page ID of interactive report.
  * @param p_ir_region_id Region ID of interactive report. If omitted or NULL region ID is derived.
  * @return               Table with substitution variable and column mapping.
  */
  FUNCTION get_sub_map_tf( p_template_tab IN VARCHAR2
                         , p_blob_col IN VARCHAR2
                         , p_search_col IN VARCHAR2
                         , p_search_value IN VARCHAR2
                         , p_enclose_char IN VARCHAR2 DEFAULT c_enclose_character
                         , p_app_id IN NUMBER DEFAULT nv('APP_ID')
                         , p_app_page_id IN NUMBER DEFAULT nv('APP_PAGE_ID')
                         , p_ir_region_id IN NUMBER DEFAULT NULL
                         )
    RETURN t_col_map_tab pipelined;

  /**
  * Converts the given template into a slide-deck.
  * This overload allows to pass in an APEX IR definition and a column to variable mapping
  * @param p_template     The PPTX file which should serve as the template.
  * @param p_report       The IR definition as derived with APEX_IR.GET_REPORT.
  * @param p_col_sub_map  Array containing the column to variable mapping.
  * @param p_enclose_char The character surrounding substitution strings.
  * @return               Generated slide deck.
  */
  FUNCTION convert_template( p_template IN BLOB
                           , p_report IN apex_ir.t_report
                           , p_col_sub_map IN t_vc_by_pi
                           , p_enclose_char IN VARCHAR2 DEFAULT c_enclose_character
                           )
    RETURN BLOB;
  
  /**
  * Converts the given template into a slide-deck.
  * This overload allows to pass in an APEX IR definition and a column to variable mapping
  * @param p_template       The PPTX file which should serve as the template.
  * @param p_report         The IR definition as derived with APEX_IR.GET_REPORT.
  * @param p_col_sub_map_vc String containing the column to variable mappings.
  * @param p_enclose_char   The character surrounding substitution strings.
  * @return                 Generated slide deck.
  */
  FUNCTION convert_template( p_template IN BLOB
                           , p_report IN apex_ir.t_report
                           , p_col_sub_map_vc IN VARCHAR2
                           , p_enclose_char IN VARCHAR2 DEFAULT c_enclose_character
                           )
    RETURN BLOB;

  /**
  * Converts the given template into a slide-deck.
  * This overload allows to pass in application id, page id, IR region id
  * and a column to variable mapping.
  * @param p_template       The PPTX file which should serve as the template.
  * @param p_application_id Application ID the IR belongs to. (Defaults to APP_ID)
  * @param p_app_page_id    Page ID of the IR region. (Defaults to APP_PAGE_ID)
  * @param p_ir_region_id   Region ID of IR. (Derived if NULL and only 1 IR on specified page)
  * @param p_col_sub_map_vc String containing the column to variable mappings.
  * @param p_enclose_char   The character surrounding substitution strings.
  * @return                 Generated slide deck.
  */    
  FUNCTION convert_template( p_template IN BLOB
                           , p_application_id IN NUMBER DEFAULT nv('APP_ID')
                           , p_app_page_id IN NUMBER DEFAULT nv('APP_PAGE_ID')
                           , p_ir_region_id IN NUMBER DEFAULT NULL
                           , p_col_sub_map IN t_vc_by_pi
                           , p_enclose_char IN VARCHAR2 DEFAULT c_enclose_character
                           )
    RETURN BLOB;

  /**
  * Converts the given template into a slide-deck.
  * This overload allows to pass in application id, page id, IR region id
  * and a column to variable mapping
  * @param p_template       The PPTX file which should serve as the template.
  * @param p_application_id Application ID the IR belongs to. (Defaults to APP_ID)
  * @param p_app_page_id    Page ID of the IR region. (Defaults to APP_PAGE_ID)
  * @param p_ir_region_id   Region ID of IR. (Derived if NULL and only 1 IR on specified page)
  * @param p_col_sub_map_vc String containing the column to variable mappings.
  * @param p_enclose_char   The character surrounding substitution strings.
  * @return                 Generated slide deck.
  */    
  FUNCTION convert_template( p_template IN BLOB
                           , p_application_id IN NUMBER DEFAULT nv('APP_ID')
                           , p_app_page_id IN NUMBER DEFAULT nv('APP_PAGE_ID')
                           , p_ir_region_id IN NUMBER DEFAULT NULL
                           , p_col_sub_map_vc IN VARCHAR2
                           , p_enclose_char IN VARCHAR2 DEFAULT c_enclose_character
                           )
    RETURN BLOB;
  
  /**
  * Converts the given template into a slide-deck and immediately pushes the file to the browser.
  * @param p_template       The PPTX file which should serve as the template.
  * @param p_application_id Application ID the IR belongs to. (Defaults to APP_ID)
  * @param p_app_page_id    Page ID of the IR region. (Defaults to APP_PAGE_ID)
  * @param p_ir_region_id   Region ID of IR. (Derived if NULL and only 1 IR on specified page)
  * @param p_col_sub_map_vc String containing the column to variable mappings.
  * @param p_enclose_char   The character surrounding substitution strings.
  * @param p_filename       Filename for download, current date and extension are always appended. If NULL application ID and page ID are used.
  */
  PROCEDURE download( p_template IN BLOB
                    , p_application_id IN NUMBER DEFAULT nv('APP_ID')
                    , p_app_page_id IN NUMBER DEFAULT nv('APP_PAGE_ID')
                    , p_ir_region_id IN NUMBER DEFAULT NULL
                    , p_col_sub_map_vc IN VARCHAR2
                    , p_enclose_char IN VARCHAR2 DEFAULT c_enclose_character
                    , p_filename IN VARCHAR2
                    )
  ;
$END

END PPTX_CREATOR_PKG;
/
