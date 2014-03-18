create or replace PACKAGE BODY PPTX_CREATOR_PKG
AS

  /*
    Constants
  */
  c_content_types_fname CONSTANT VARCHAR(200) := '[Content_Types].xml';
  c_presentation_fname CONSTANT VARCHAR2(200) := 'ppt/presentation.xml';
  c_pres_rel_fname CONSTANT VARCHAR2(200) := 'ppt/_rels/presentation.xml.rels';
  
  c_template_slide CONSTANT VARCHAR2(200) := 'ppt/slides/slide1.xml';
  c_template_slide_rel CONSTANT VARCHAR2(200) := 'ppt/slides/_rels/slide1.xml.rels';
  c_template_note CONSTANT VARCHAR2(200) := 'ppt/notesSlides/notesSlide1.xml';
  c_template_note_rel CONSTANT VARCHAR2(200) := 'ppt/notesSlides/_rels/notesSlide1.xml.rels';
  
  c_slide_dname CONSTANT VARCHAR2(200) := 'ppt/slides/';
  c_slide_rel_dname CONSTANT VARCHAR2(200) := 'ppt/slides/_rels/';
  
  c_slide_pattern CONSTANT VARCHAR2(200) := 'slide#NUM#.xml';
  c_relation_suffix CONSTANT VARCHAR2(200) := '.rels';

  c_note_content_type CONSTANT VARCHAR2(200) := 'application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml';
  c_override_pattern CONSTANT VARCHAR2(200) := '<Override PartName="/ppt/slides/slide#NUM#.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>';
  c_presentation_slide CONSTANT VARCHAR2(200) := '<p:sldId id="#SLIDE_ID#" r:id="rId#REL_ID#"/>';
  c_pres_rel_pattern CONSTANT VARCHAR2(200) := '<Relationship Id="rId#REL_ID#" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide#NUM#.xml"/>';

  c_slide_num_offset CONSTANT NUMBER := 0;

  /*
    Globals
  */
  g_base_file BLOB;
  g_file_list zip_util_pkg.t_file_list;

  g_slide_id_offset NUMBER := 255;
  g_slide_relation_id_offset NUMBER := 6;

  g_template_slide CLOB;
  g_template_slide_rel BLOB;
  g_slides t_all_slides;
  

  /*
    Private API
  */

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
    l_clob := sql_util_pkg.blob_to_clob(zip_util_pkg.get_file( g_base_file, c_presentation_fname ));
    l_match_left := INSTR(l_clob, 'sldId');
    l_match_left := INSTR(l_clob, '"', l_match_left);
    l_match_right := INSTR(l_clob, '"', l_match_left + 1);
    l_attr := SUBSTR(l_clob, l_match_left + 1, l_match_right - l_match_left - 1);
    IF l_attr IS NOT NULL THEN
      g_slide_id_offset := GREATEST(g_slide_id_offset, to_number(l_attr));
    END IF;
    
    -- presentation rels to get rId
    l_data := XMLTYPE(sql_util_pkg.blob_to_clob(zip_util_pkg.get_file( g_base_file, c_pres_rel_fname )));
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
    g_template_slide := sql_util_pkg.blob_to_clob( zip_util_pkg.get_file( g_base_file, c_template_slide ) );
    g_template_slide_rel := zip_util_pkg.get_file( g_base_file, c_template_slide_rel );
  END set_template_slide;
  
  PROCEDURE process_slides
  AS
  BEGIN
    NULL;
  END process_slides;

  -- Insert override tags into content types for each slide
  FUNCTION update_content_types (p_data IN BLOB)
  RETURN BLOB
  AS
    l_data XMLTYPE;
    l_domdoc dbms_xmldom.DOMDocument;
    l_root_node dbms_xmldom.DOMNode;
    l_element dbms_xmldom.DOMElement;
    l_slide_node dbms_xmldom.DOMNode;
  BEGIN
    l_data := XMLTYPE(sql_util_pkg.blob_to_clob(p_data));
    l_domdoc := dbms_xmldom.newDOMDocument(l_data);
    l_root_node := dbms_xmldom.makeNode(DBMS_XMLDOM.getDocumentElement(l_domdoc));
    FOR i IN 2..g_slides.COUNT
    LOOP
      l_element := dbms_xmldom.createElement(l_domdoc, 'Override');
      dbms_xmldom.setAttribute(l_element, 'PartName', '/' || c_slide_dname || REPLACE(c_slide_pattern, '#NUM#', to_char(g_slides(i).slide_num)) );
      dbms_xmldom.setAttribute(l_element, 'ContentType', g_slides(i).content_type);
      l_slide_node := dbms_xmldom.appendChild(l_root_node, dbms_xmldom.makeNode(l_element));
    END LOOP;
    dbms_xmldom.freeDocument(l_domdoc);
    RETURN sql_util_pkg.clob_to_blob(REPLACE(l_data.getClobVal, 'ISO-8859-15', 'UTF-8'));
  END update_content_types;

  -- Insert slide id tag into slide id list of presentation for each slide
  FUNCTION update_presentation (p_data IN BLOB)
  RETURN BLOB
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
    l_data := XMLTYPE(sql_util_pkg.blob_to_clob(p_data));
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
    RETURN sql_util_pkg.clob_to_blob(REPLACE(l_data.getClobVal, 'ISO-8859-15', 'UTF-8'));
  END update_presentation;
  
  -- Insert relationship tag into presenation relations for every slide
  FUNCTION update_pres_rel (p_data IN BLOB)
  RETURN BLOB
  AS
    l_data XMLTYPE;
    l_domdoc dbms_xmldom.DOMDocument;
    l_root_node dbms_xmldom.DOMNode;
    l_element dbms_xmldom.DOMElement;
    l_slide_node dbms_xmldom.DOMNode;
  BEGIN
    l_data := XMLTYPE(sql_util_pkg.blob_to_clob(p_data));
    l_domdoc := dbms_xmldom.newDOMDocument(l_data);
    l_root_node := dbms_xmldom.makeNode(DBMS_XMLDOM.getDocumentElement(l_domdoc));
    FOR i IN 2..g_slides.COUNT
    LOOP
      l_element := dbms_xmldom.createElement(l_domdoc, 'Relationship');
      dbms_xmldom.setAttribute(l_element, 'Id', 'rId' || to_char(g_slides(i).relation_id));
      dbms_xmldom.setAttribute(l_element, 'Type', g_slides(i).relation_type);
      dbms_xmldom.setAttribute(l_element, 'Target', 'slides/' || REPLACE(c_slide_pattern, '#NUM#', to_char(g_slides(i).slide_num)));
      l_slide_node := dbms_xmldom.appendChild(l_root_node, dbms_xmldom.makeNode(l_element));
    END LOOP;
    dbms_xmldom.freeDocument(l_domdoc);
    RETURN sql_util_pkg.clob_to_blob(REPLACE(l_data.getClobVal, 'ISO-8859-15', 'UTF-8'));
  END update_pres_rel;
  

  /*
    Public API
  */
  
  FUNCTION convert_template ( p_template IN BLOB
                            , p_replace_patterns IN t_str_array
                            , p_replace_values IN t_replace_values_tab 
                            )
    RETURN BLOB
  AS
    l_retval BLOB;
    l_cur_blob BLOB;
    l_current_slide t_single_slide;
  BEGIN
    g_base_file := p_template;
    g_file_list := zip_util_pkg.get_file_list (g_base_file);
    
    set_template_slide;
    set_offsets;
    
    -- create all slides based on template and add to new file
    FOR i IN 1..p_replace_values.COUNT
    LOOP
      l_current_slide.slide_num := c_slide_num_offset + i;
      l_current_slide.slide_id := g_slide_id_offset + i;
      l_current_slide.relation_id := g_slide_relation_id_offset + i;
      l_current_slide.slide_data := string_util_pkg.multi_replace ( g_template_slide, p_replace_patterns, p_replace_values(i));
      g_slides(i) := l_current_slide;
      l_cur_blob := sql_util_pkg.clob_to_blob (l_current_slide.slide_data);
      zip_util_pkg.add_file (l_retval, c_slide_dname || REPLACE(c_slide_pattern, '#NUM#', to_char(l_current_slide.slide_num)), l_cur_blob);
      zip_util_pkg.add_file (l_retval, c_slide_rel_dname || REPLACE(c_slide_pattern, '#NUM#', to_char(l_current_slide.slide_num)) || c_relation_suffix, g_template_slide_rel);
    END LOOP;
    
    -- slides prepared, loop through file content and adapt where needed
    FOR i IN 1..g_file_list.COUNT
    LOOP
      l_cur_blob := zip_util_pkg.get_file (p_template, g_file_list(i));
      IF g_file_list(i) = c_content_types_fname
      THEN
        l_cur_blob := update_content_types(l_cur_blob); --update content types file with overrides for slides
      ELSIF g_file_list(i) = c_presentation_fname
      THEN
        l_cur_blob := update_presentation(l_cur_blob); -- update presentation file with slide refs
      ELSIF g_file_list(i) = c_pres_rel_fname
      THEN
        l_cur_blob := update_pres_rel(l_cur_blob); --update presentation rels with slide infos
      END IF;
      
      IF g_file_list(i) NOT LIKE c_slide_dname || '%'  -- ignore all files in ppt/slides folder, those were generated above
      THEN
        zip_util_pkg.add_file (l_retval, g_file_list(i), l_cur_blob);
      END IF;      
    END LOOP;
    zip_util_pkg.finish_zip (l_retval);
    RETURN l_retval;
  END convert_template;

END PPTX_CREATOR_PKG;
/