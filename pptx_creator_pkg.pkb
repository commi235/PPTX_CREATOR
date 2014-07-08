create or replace PACKAGE BODY PPTX_CREATOR_PKG
AS

  /*
    Types
  */
  /**
  * Record which holds all data about a single slide
  * @param slide_num
  * @param slide_id
  * @param relation_id
  * @param slide_data
  * @param notes_data
  */
  TYPE t_single_slide IS RECORD
    ( slide_num NUMBER
    , slide_id NUMBER
    , relation_id NUMBER
    , slide_data CLOB
    , notes_data CLOB
    );

  /**
  * Array of slide details
  */
  TYPE t_all_slides IS TABLE OF t_single_slide INDEX BY PLS_INTEGER;

  /**
  * Record to keep all relevant templates together
  * @param slide The template slide with substitution patterns
  * @param slide_relations Template slide relations
  * @param notes The template notes with substitution patterns
  * @param notes_relations Template notes relations
  */
  TYPE t_templates IS RECORD
    ( slide CLOB
    , slide_relations CLOB
    , notes CLOB
    , notes_relations CLOB
    );

  TYPE t_col_values IS RECORD
    ( varchar_values dbms_sql.varchar2a
    , number_values dbms_sql.number_table
    , date_values dbms_sql.date_table
    , clob_values dbms_sql.clob_table
    );
    
  TYPE t_sql_rows IS TABLE OF t_col_values INDEX BY VARCHAR2(32767);

  /*
    Constants
  */
  
  /* Static filenames for top level files */
  c_content_types_fname CONSTANT VARCHAR(19 CHAR) := '[Content_Types].xml';
  c_presentation_fname CONSTANT VARCHAR2(20 CHAR) := 'ppt/presentation.xml';
  c_pres_rel_fname CONSTANT VARCHAR2(31 CHAR) := 'ppt/_rels/presentation.xml.rels';
  
  /* Filenames for templates */
  c_template_slide CONSTANT VARCHAR2(21 CHAR) := 'ppt/slides/slide1.xml';
  c_template_slide_rel CONSTANT VARCHAR2(32 CHAR) := 'ppt/slides/_rels/slide1.xml.rels';
  c_template_notes CONSTANT VARCHAR2(31 CHAR) := 'ppt/notesSlides/notesSlide1.xml';
  c_template_notes_rel CONSTANT VARCHAR2(42 CHAR) := 'ppt/notesSlides/_rels/notesSlide1.xml.rels';
  
  /* Directories for slides and notes */
  c_slide_dname CONSTANT VARCHAR2(11 CHAR) := 'ppt/slides/';
  c_slide_rel_dname CONSTANT VARCHAR2(17 CHAR) := 'ppt/slides/_rels/';
  c_notes_dname CONSTANT VARCHAR2(16 CHAR) := 'ppt/notesSlides/';
  c_notes_rel_dname CONSTANT VARCHAR2(22 CHAR) := 'ppt/notesSlides/_rels/';
  
  /* Template filenames for slide and notes */
  c_slide_fname_pattern CONSTANT VARCHAR2(14 CHAR) := 'slide#NUM#.xml';
  c_notes_fname_pattern CONSTANT VARCHAR2(20 CHAR) := 'notesSlide#NUM#.xml';
  c_relation_suffix CONSTANT VARCHAR2(5 CHAR) := '.rels';

  c_slide_content_type CONSTANT VARCHAR2(70 CHAR) := 'application/vnd.openxmlformats-officedocument.presentationml.slide+xml';
  c_sliderel_content_type CONSTANT VARCHAR2(73 CHAR) := 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide';
  c_notes_content_type CONSTANT VARCHAR2(75 CHAR) := 'application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml';
  c_presentation_slide CONSTANT VARCHAR2(45 CHAR) := '<p:sldId id="#SLIDE_ID#" r:id="rId#REL_ID#"/>';

  c_slide_num_offset CONSTANT NUMBER := 0;
  c_bulk_size CONSTANT PLS_INTEGER := 100;

  /*
    Globals
  */
  g_base_file BLOB;
  g_file_list zip_util_pkg.t_file_list;

  g_slide_id_offset NUMBER := 255;
  g_slide_relation_id_offset NUMBER := 6;

  g_templates t_templates;
  g_slides t_all_slides;

  g_replace_value_tab t_replace_value_tab;
  g_enclose_character VARCHAR2(1);

  /*
    Private API
  */

  FUNCTION convert_replace( p_replace_patterns IN t_vc_value_row
                          , p_replace_values IN t_vc_value_tab
                          )
    RETURN t_replace_value_tab
  AS
    l_returnvalue t_replace_value_tab;
  BEGIN
    FOR i IN 1..p_replace_values.count LOOP
      FOR j IN 1..p_replace_patterns.count LOOP
        l_returnvalue(i)(p_replace_patterns(j)).varchar_value := p_replace_values(i)(j);
      END LOOP;
    END LOOP;
    RETURN l_returnvalue;
  END convert_replace;
  
  FUNCTION convert_replace( p_sql_result IN t_sql_rows )
    RETURN t_replace_value_tab
  AS
    l_returnvalue t_replace_value_tab;
    l_cur_col VARCHAR2(32767);
  BEGIN
    l_cur_col := p_sql_result.first();
    WHILE l_cur_col IS NOT NULL LOOP
      IF p_sql_result(l_cur_col).varchar_values.count > 0 THEN
        FOR i IN 1..p_sql_result(l_cur_col).varchar_values.count LOOP
          l_returnvalue(i)(l_cur_col).varchar_value := p_sql_result(l_cur_col).varchar_values(i);
        END LOOP;
      ELSIF p_sql_result(l_cur_col).number_values.count > 0 THEN
        FOR i IN 1..p_sql_result(l_cur_col).number_values.count LOOP
          l_returnvalue(i)(l_cur_col).number_value := p_sql_result(l_cur_col).number_values(i);
        END LOOP;
      ELSIF p_sql_result(l_cur_col).date_values.count > 0 THEN
        FOR i IN 1..p_sql_result(l_cur_col).date_values.count LOOP
          l_returnvalue(i)(l_cur_col).date_value := p_sql_result(l_cur_col).date_values(i);
        END LOOP;
      ELSIF p_sql_result(l_cur_col).clob_values.count > 0 THEN
        FOR i IN 1..p_sql_result(l_cur_col).clob_values.count LOOP
          l_returnvalue(i)(l_cur_col).clob_value := p_sql_result(l_cur_col).clob_values(i);
        END LOOP;      
      END IF;
      l_cur_col := p_sql_result.next(l_cur_col);
    END LOOP;
    RETURN l_returnvalue;
  END convert_replace;

  FUNCTION convert_value ( p_value IN t_replace_value )
    RETURN VARCHAR2
  AS
    l_returnvalue VARCHAR2(32767);
  BEGIN
    IF p_value.varchar_value IS NOT NULL THEN
      l_returnvalue := p_value.varchar_value;
    ELSIF p_value.number_value IS NOT NULL THEN
      IF p_value.format_mask IS NOT NULL THEN
        l_returnvalue := to_char( p_value.number_value, p_value.format_mask );
      ELSE
        l_returnvalue := to_char( p_value.number_value);
      END IF;
    ELSIF p_value.date_value IS NOT NULL THEN
      IF p_value.format_mask IS NOT NULL THEN
        l_returnvalue := to_char( p_value.date_value, p_value.format_mask );
      ELSE
        l_returnvalue := to_char( p_value.date_value);
      END IF;
    ELSIF p_value.clob_value IS NOT NULL THEN
      l_returnvalue := dbms_lob.substr(p_value.clob_value);
    END IF;
    RETURN l_returnvalue;
  END convert_value;

  FUNCTION replace_substitution( p_original IN CLOB
                               , p_row_num IN PLS_INTEGER
                               )
    RETURN CLOB
  AS
    l_returnvalue CLOB := p_original;
    l_current_index VARCHAR2(30);
  BEGIN
    l_current_index := g_replace_value_tab(p_row_num).FIRST;
    WHILE l_current_index IS NOT NULL LOOP
      l_returnvalue := REPLACE( l_returnvalue
                              , g_enclose_character || l_current_index || g_enclose_character
                              , convert_value(g_replace_value_tab(p_row_num)(l_current_index))
                              );
      l_current_index := g_replace_value_tab(p_row_num).NEXT(l_current_index);
    END LOOP;
    RETURN l_returnvalue;
  END replace_substitution;

  PROCEDURE set_offsets
  AS
    l_data XMLTYPE;
    l_clob CLOB;
    l_domdoc dbms_xmldom.DOMDocument;
    l_root_node dbms_xmldom.DOMNode;
    l_rels dbms_xmldom.DOMNodeList;
    l_act_node dbms_xmldom.DOMNode;
    l_act_element dbms_xmldom.DOMElement;
    l_attrs dbms_xmldom.DOMNamedNodeMap;
    l_attr VARCHAR2(32676);
    l_match_left NUMBER;
    l_match_right NUMBER;
  BEGIN
    -- presentation to get slide Ids
    l_clob := zip_util_pkg.get_file_clob( g_base_file, c_presentation_fname );
    l_match_left := INSTR(l_clob, 'sldId');
    l_match_left := INSTR(l_clob, '"', l_match_left);
    l_match_right := INSTR(l_clob, '"', l_match_left + 1);
    l_attr := SUBSTR(l_clob, l_match_left + 1, l_match_right - l_match_left - 1);
    IF l_attr IS NOT NULL THEN
      g_slide_id_offset := GREATEST(g_slide_id_offset, to_number(l_attr));
    END IF;
    
    -- presentation rels to get rId
    l_data := XMLTYPE(zip_util_pkg.get_file_clob( g_base_file, c_pres_rel_fname ));
    l_domdoc := dbms_xmldom.newDOMDocument(l_data);
    l_root_node := dbms_xmldom.makeNode(DBMS_XMLDOM.getDocumentElement(l_domdoc));
    l_rels := dbms_xmldom.getChildNodes(l_root_node);
    FOR i IN 1..dbms_xmldom.getLength(l_rels)
    LOOP
      l_act_element := dbms_xmldom.makeElement(dbms_xmldom.item(l_rels, i));
      l_attr := dbms_xmldom.getAttribute(l_act_element, 'Id');
      IF l_attr IS NOT NULL THEN
        g_slide_relation_id_offset := GREATEST(g_slide_relation_id_offset, to_number(substr(l_attr, 4)));
      END IF;
    END LOOP;
  END set_offsets;

  PROCEDURE set_template_slide
  AS
  BEGIN
    g_templates.slide := zip_util_pkg.get_file_clob( g_base_file, c_template_slide );
    g_templates.slide_relations := zip_util_pkg.get_file_clob( g_base_file, c_template_slide_rel );
    g_templates.notes := zip_util_pkg.get_file_clob( g_base_file, c_template_notes );
    g_templates.notes_relations := zip_util_pkg.get_file_clob( g_base_file, c_template_notes_rel );
  END set_template_slide;
  
  FUNCTION process_slide_relation( p_slide_num IN NUMBER )
    RETURN CLOB
  AS
  BEGIN
    RETURN REPLACE( g_templates.slide_relations, 'notesSlide1', 'notesSlide' || to_char(p_slide_num) );
  END process_slide_relation;
  
  FUNCTION process_notes_relation( p_slide_num IN NUMBER )
    RETURN CLOB
  AS
  BEGIN
    RETURN REPLACE( g_templates.notes_relations, 'slide1', 'slide' || to_char(p_slide_num) );
  END process_notes_relation;  

  -- Insert override tags into content types for each slide
  PROCEDURE update_content_types( p_file IN OUT NOCOPY BLOB )
  AS
    l_data XMLTYPE;
    l_domdoc dbms_xmldom.DOMDocument;
    l_root_node dbms_xmldom.DOMNode;
    l_slide_element dbms_xmldom.DOMElement;
    l_slide_node dbms_xmldom.DOMNode;
    l_notes_element dbms_xmldom.DOMElement;
    l_notes_node dbms_xmldom.DOMNode;
  BEGIN
    l_data := XMLTYPE(zip_util_pkg.get_file_clob(g_base_file, c_content_types_fname));
    l_domdoc := dbms_xmldom.newDOMDocument(l_data);
    l_root_node := dbms_xmldom.makeNode(DBMS_XMLDOM.getDocumentElement(l_domdoc));
    FOR i IN 2..g_slides.COUNT
    LOOP
      l_slide_element := dbms_xmldom.createElement(l_domdoc, 'Override');
      dbms_xmldom.setAttribute(l_slide_element, 'PartName', '/' || c_slide_dname || REPLACE(c_slide_fname_pattern, '#NUM#', to_char(g_slides(i).slide_num)) );
      dbms_xmldom.setAttribute(l_slide_element, 'ContentType', c_slide_content_type);
      l_slide_node := dbms_xmldom.appendChild(l_root_node, dbms_xmldom.makeNode(l_slide_element));
      IF g_templates.notes IS NOT NULL THEN
        l_notes_element := dbms_xmldom.createElement(l_domdoc, 'Override');
        dbms_xmldom.setAttribute(l_notes_element, 'PartName', '/' || c_notes_dname || REPLACE(c_notes_fname_pattern, '#NUM#', to_char(g_slides(i).slide_num)) );
        dbms_xmldom.setAttribute(l_notes_element, 'ContentType', c_notes_content_type);        
        l_notes_node := dbms_xmldom.appendChild(l_root_node, dbms_xmldom.makeNode(l_notes_element));
      END IF;
    END LOOP;
    dbms_xmldom.freeDocument(l_domdoc);
    zip_util_pkg.add_file(p_file, c_content_types_fname, REPLACE(l_data.getClobVal, 'ISO-8859-15', 'UTF-8'));
  END update_content_types;

  -- Insert slide id tag into slide id list of presentation for each slide
  PROCEDURE update_presentation( p_file IN OUT NOCOPY BLOB )
  AS
    l_data XMLTYPE;
    l_domdoc dbms_xmldom.DOMDocument;
    l_root_node dbms_xmldom.DOMNode;
    l_slide_list dbms_xmldom.DOMNode;
    l_element dbms_xmldom.DOMElement;
    l_slide_node dbms_xmldom.DOMNode;
    l_id_attr dbms_xmldom.DOMAttr;
    l_rid_attr dbms_xmldom.DOMAttr;
  BEGIN
    l_data := XMLTYPE(zip_util_pkg.get_file_clob(g_base_file, c_presentation_fname));
    l_domdoc := dbms_xmldom.newDOMDocument(l_data);
    l_root_node := dbms_xmldom.makeNode(DBMS_XMLDOM.getDocumentElement(l_domdoc));
    l_slide_list := dbms_xmldom.item(dbms_xmldom.getElementsByTagName(l_domdoc, 'sldIdLst'), 0);
    FOR i IN 2..g_slides.COUNT
    LOOP
      l_element := dbms_xmldom.createElement(l_domdoc, 'p:sldId');
      l_id_attr := dbms_xmldom.createAttribute(l_domdoc, 'id');
      l_rid_attr := dbms_xmldom.createAttribute(l_domdoc, 'r:id');
      dbms_xmldom.setValue(l_rid_attr, 'rId' || to_char(g_slides(i).relation_id));
      dbms_xmldom.setValue(l_id_attr, to_char(g_slides(i).slide_id));
      l_id_attr := dbms_xmldom.setAttributeNode(l_element, l_id_attr);
      l_rid_attr := dbms_xmldom.setAttributeNode(l_element, l_rid_attr);      
      l_slide_node := dbms_xmldom.appendChild(l_slide_list, dbms_xmldom.makeNode(l_element));
    END LOOP;
    dbms_xmldom.freeDocument(l_domdoc);
    zip_util_pkg.add_file(p_file, c_presentation_fname, REPLACE(l_data.getClobVal, 'ISO-8859-15', 'UTF-8'));
  END update_presentation;
  
  -- Insert relationship tag into presenation relations for every slide
  PROCEDURE update_pres_rel( p_file IN OUT NOCOPY BLOB )
  AS
    l_data XMLTYPE;
    l_domdoc dbms_xmldom.DOMDocument;
    l_root_node dbms_xmldom.DOMNode;
    l_element dbms_xmldom.DOMElement;
    l_slide_node dbms_xmldom.DOMNode;
  BEGIN
    l_data := XMLTYPE(zip_util_pkg.get_file_clob(g_base_file, c_pres_rel_fname));
    l_domdoc := dbms_xmldom.newDOMDocument(l_data);
    l_root_node := dbms_xmldom.makeNode(DBMS_XMLDOM.getDocumentElement(l_domdoc));
    FOR i IN 2..g_slides.COUNT
    LOOP
      l_element := dbms_xmldom.createElement(l_domdoc, 'Relationship');
      dbms_xmldom.setAttribute(l_element, 'Id', 'rId' || to_char(g_slides(i).relation_id));
      dbms_xmldom.setAttribute(l_element, 'Type', c_sliderel_content_type);
      dbms_xmldom.setAttribute(l_element, 'Target', 'slides/' || REPLACE(c_slide_fname_pattern, '#NUM#', to_char(g_slides(i).slide_num)));
      l_slide_node := dbms_xmldom.appendChild(l_root_node, dbms_xmldom.makeNode(l_element));
    END LOOP;
    dbms_xmldom.freeDocument(l_domdoc);
    zip_util_pkg.add_file(p_file, c_pres_rel_fname, REPLACE(l_data.getClobVal, 'ISO-8859-15', 'UTF-8'));
  END update_pres_rel;
  

  /*
    Public API
  */
  
  FUNCTION convert_template ( p_template IN BLOB
                            , p_replace_patterns IN t_vc_value_row
                            , p_replace_values IN t_vc_value_tab 
                            )
    RETURN BLOB
  AS
  BEGIN
    RETURN convert_template( p_template => p_template
                           , p_replace_name_value => convert_replace( p_replace_patterns, p_replace_values )
                           , p_enclose_char => NULL
                           );
  END convert_template;

  FUNCTION convert_template ( p_template IN BLOB
                            , p_replace_name_value IN t_replace_value_tab
                            , p_enclose_char IN VARCHAR2 DEFAULT c_enclose_character
                            )
    RETURN BLOB
  AS
    l_current_slide t_single_slide;
    l_returnvalue BLOB;
    l_current_file BLOB;
  BEGIN
    /* Init globals */
    g_base_file := p_template;
    g_file_list := zip_util_pkg.get_file_list (g_base_file);
    set_template_slide;
    set_offsets;
    g_replace_value_tab := p_replace_name_value;
    g_enclose_character := p_enclose_char;
    
    IF g_replace_value_tab.count = 0 THEN
      RETURN NULL;
    END IF;
    
    -- Create all slides and notes based on template and add to file
    FOR i IN 1..g_replace_value_tab.COUNT
    LOOP
      l_current_slide.slide_num := c_slide_num_offset + i;
      l_current_slide.slide_id := g_slide_id_offset + i;
      l_current_slide.relation_id := g_slide_relation_id_offset + i;
      
      /* Slide and Relations (always)*/      
      l_current_slide.slide_data := replace_substitution( p_original => g_templates.slide
                                                        , p_row_num => i
                                                        );
      zip_util_pkg.add_file ( l_returnvalue
                            , c_slide_dname || REPLACE(c_slide_fname_pattern, '#NUM#', to_char(l_current_slide.slide_num))
                            , l_current_slide.slide_data
                            );
      zip_util_pkg.add_file ( l_returnvalue
                            , c_slide_rel_dname || REPLACE(c_slide_fname_pattern, '#NUM#', to_char(l_current_slide.slide_num)) || c_relation_suffix
                            , process_slide_relation( l_current_slide.slide_num )
                            );
      /* Notes and Relations (only if notes exist) */
      IF g_templates.notes IS NOT NULL THEN
        l_current_slide.notes_data := replace_substitution( p_original => g_templates.notes
                                                          , p_row_num => i
                                                          );
        zip_util_pkg.add_file ( l_returnvalue
                              , c_notes_dname || REPLACE(c_notes_fname_pattern, '#NUM#', to_char(l_current_slide.slide_num))
                              , l_current_slide.notes_data
                              );
        zip_util_pkg.add_file ( l_returnvalue
                              , c_notes_rel_dname || REPLACE(c_notes_fname_pattern, '#NUM#', to_char(l_current_slide.slide_num)) || c_relation_suffix
                              , process_notes_relation( l_current_slide.slide_num )
                              );        
      END IF;
      g_slides(i) := l_current_slide;
    END LOOP;
    
    -- slides prepared, loop through file content and adapt where needed
    FOR i IN 1..g_file_list.COUNT
    LOOP
      IF g_file_list(i) = c_content_types_fname
      THEN
        update_content_types(l_returnvalue); --update content types file with overrides for slides
      ELSIF g_file_list(i) = c_presentation_fname
      THEN
        update_presentation(l_returnvalue); -- update presentation file with slide refs
      ELSIF g_file_list(i) = c_pres_rel_fname
      THEN
        update_pres_rel(l_returnvalue); --update presentation rels with slide infos
      ELSIF g_file_list(i) NOT LIKE c_slide_dname || '%' AND g_file_list(i) NOT LIKE c_notes_dname || '%'
      THEN
        zip_util_pkg.add_file( l_returnvalue, g_file_list(i), zip_util_pkg.get_file( g_base_file, g_file_list(i)));
      END IF;
    END LOOP;
    zip_util_pkg.finish_zip (l_returnvalue);
    RETURN l_returnvalue;
  END convert_template;

  FUNCTION convert_template ( p_template IN BLOB
                            , p_cursor IN OUT NOCOPY sys_refcursor
                            , p_enclose_char IN VARCHAR2 DEFAULT c_enclose_character
                            )
    RETURN BLOB
  AS
    l_cursor_id PLS_INTEGER;
    l_desc_tab dbms_sql.desc_tab2;
    l_col_cnt PLS_INTEGER;
    l_row_cnt PLS_INTEGER;
    l_sql_result t_sql_rows;
  BEGIN
    l_cursor_id := dbms_sql.to_cursor_number( p_cursor );
    dbms_sql.describe_columns2( l_cursor_id, l_col_cnt, l_desc_tab );
    
    -- Prepare the arrays
    FOR i IN 1..l_col_cnt LOOP
      CASE
        WHEN l_desc_tab( i ).col_type IN ( 2, 100, 101 ) THEN
          dbms_sql.define_array( l_cursor_id, i, l_sql_result(l_desc_tab(i).col_name).number_values, c_bulk_size, 1 );
        WHEN l_desc_tab( i ).col_type IN ( 12, 178, 179, 180, 181 , 231 ) THEN
          dbms_sql.define_array( l_cursor_id, i, l_sql_result(l_desc_tab(i).col_name).date_values, c_bulk_size, 1 );
        WHEN l_desc_tab( i ).col_type IN ( 1, 8, 9, 96 ) THEN
          dbms_sql.define_array( l_cursor_id, i, l_sql_result(l_desc_tab(i).col_name).varchar_values, c_bulk_size, 1 );
        WHEN l_desc_tab( i ).col_type = 112 THEN
          dbms_sql.define_array( l_cursor_id, i, l_sql_result(l_desc_tab(i).col_name).clob_values, c_bulk_size, 1 );
        ELSE
          NULL;
      END CASE;      
    END LOOP;
    
    -- Execute and fill arrays
    l_row_cnt := dbms_sql.EXECUTE( l_cursor_id );
    LOOP
      l_row_cnt := dbms_sql.fetch_rows( l_cursor_id );
      IF l_row_cnt > 0 THEN
        FOR i IN 1..l_col_cnt LOOP
          CASE
            WHEN l_desc_tab( i ).col_type IN ( 2, 100, 101 ) THEN
              dbms_sql.COLUMN_VALUE( l_cursor_id, i, l_sql_result(l_desc_tab(i).col_name).number_values);
            WHEN l_desc_tab( i ).col_type IN ( 12, 178, 179, 180, 181 , 231 ) THEN
              dbms_sql.COLUMN_VALUE( l_cursor_id, i, l_sql_result(l_desc_tab(i).col_name).date_values);
            WHEN l_desc_tab( i ).col_type IN ( 1, 8, 9, 96 ) THEN
              dbms_sql.COLUMN_VALUE( l_cursor_id, i, l_sql_result(l_desc_tab(i).col_name).varchar_values);
            WHEN l_desc_tab( i ).col_type = 112 THEN
              dbms_sql.COLUMN_VALUE( l_cursor_id, i, l_sql_result(l_desc_tab(i).col_name).clob_values);
            ELSE
              NULL;
          END CASE;          
        END LOOP;
      END IF;
      EXIT WHEN l_row_cnt != c_bulk_size;
    END LOOP;
    dbms_sql.close_cursor(l_cursor_id);
    RETURN convert_template( p_template => p_template
                           , p_replace_name_value => convert_replace( l_sql_result )
                           , p_enclose_char => p_enclose_char
                           );
    EXCEPTION
      WHEN OTHERS THEN
        IF dbms_sql.is_open(l_cursor_id) THEN
          dbms_sql.close_cursor(l_cursor_id);
        END IF;
        RETURN NULL;
  END convert_template;

END PPTX_CREATOR_PKG;
/
