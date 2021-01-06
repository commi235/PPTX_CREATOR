PPTX_CREATOR
============
Create PPTX files from a template replacing substitution strings.  

INSTALLATION
------------
Simply run the provided "install_all.sql" file if you want everything installed at once.  
If you have the referenced libraries already you can run "install_main.sql" to install the main package.  
If you just want to install the ZIP_UTIL_PKG then run "/lib/zip_util/install.sql".  
You can also install the packages manually,
but then you need to specify if APEX is installed using before installing
the PPTX_CREATOR_PKG package.  
Before compiling the package run
```sql
ALTER SESSION SET plsql_ccflags='apex_installed:1';
```
to enable APEX-specific functionality or run
```sql
ALTER SESSION SET plsql_ccflags='apex_installed:0';
```
to disable them.

DEPENDENCIES
------------
1. ZIP_UTIL_PKG (Included)
2. DBMS_XMLDOM (Oracle XDB)
3. XMLTYPE (Oracle XDB)
4. Oracle Application Express (not required)  
   The install script checks if APEX is installed and enables additional
   functions which only work together with APEX.

HOW TO USE
----------
