REM Install PPTX Creator
SET define '^'
define HASAPX = '0'
column apx new_val HASAPX

SELECT '1' AS apx
  FROM all_registry_banners
 WHERE banner LIKE 'Oracle Application Express%'
   AND ROWNUM = 1
;

ALTER SESSION SET plsql_ccflags = 'apex_installed:^HASAPX';

@@pptx_creator_pkg.pks
@@pptx_creator_pkg.pkb

show errors