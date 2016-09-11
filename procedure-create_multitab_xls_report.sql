create or replace procedure create_multitab_xls_report
(result_file varchar2,v_dir  varchar2, v_id number )
as

/****************************
PURPOSE: CREATE MULTI PAGE XLS REPORT 
WITH SPECIFICATION GIVEN FOR FORMATS AND LAYOUT BASED ON DATATYPE I.E. CHARACTERS, NUMBER, DATE 
USING DATA GATHERED FROM SELECT QUERIES STORED IN TABLE MULTITAB_WORKSHEET

PREREQUISITES:
1. USER ACCESS PERMISSION ON THE DIRECTORY WHERE FILE WILL BE CREATED
2. USER MUST CREATE TABLE MULTITAB_WORKSHEET 
TO PROVIDE SIMPLE METADATA INFORMATION ABOUT REPORT.
EXAMPLE: WORKBOOK NAME, WORKSHEET NAME, SELECT QUERY TO DISPLAY RESULT IN WORKSHEET, FILE_NUMBER

DESC multitab_worksheet;
Name                           Null     Type                                                                                                                                                                                          
------------------------------ -------- ----------------------------------------
WORKBOOK_NAME                           VARCHAR2(50)                                                                                                                                                                                  
WORKSHEET_NUMBER                        NUMBER                                                                                                                                                                                        
WORKSHEET_NAME                          VARCHAR2(50)                                                                                                                                                                                  
WORKSHEET_QUERY                         VARCHAR2(1000)                                                                                                                                                                                
FILE_NUMBER                             NUMBER   

3. USER ACCESS TO CREATE ASSOCIATIVE ARRAY

PARAMETERS:
RESULTFILE: 
NAME OF THE EXCEL FILE

V_DIR: 
NAME OF THE DIRECTORY WHERE FILE WILL STORED.
NOTE: USER ACCESS IS REQUIRED ON THE DIRECTORY BEFORE EXECUTION OF THE PROCEDURE

V_ID:
UNIQUE FILE_NUMBER STORED IN TABLE MULTITAB_WORKSHEET FOR EACH EXCEL REPORT

AUTHOR: SINGHINDU87
VERSION: 1.0
CONCEPTS USED: COLLECTIONS, REF_CURSORS, DBMS_SQL, UTL_FILE
****************************/

type type_worksheet is table of multitab_worksheet%rowtype index by binary_integer;
t_worksheet type_worksheet;
rfc_worksheet sys_refcursor;
v_writefile utl_file.file_type;
c number;
d number;
rec_tab dbms_sql.desc_tab;
col_cnt integer;
v_v_val varchar2 (4000); 
v_n_val number;
v_d_val date;
v_ret number;
begin
select * bulk collect into t_worksheet from multitab_worksheet where file_number=v_id ; --fetch all the worksheet for unqiue excel report driven by a file_number

dbms_output.put_line ('......WORKBOOK BUILDING......');
dbms_output.put_line ( '......WORKBOOK NAME......' || t_worksheet (1).workbook_name);
v_writefile :=utl_file.fopen (upper (v_dir),t_worksheet (1).workbook_name,'w',32767);
utl_file.put_line (v_writefile, '<?xml version="1.0"  encoding="ISO-8859-9"?>');
utl_file.put_line (v_writefile,'<ss:Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet"  xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel" xmlns:html="http://www.w3.org/2001/XMLSchema">');
utl_file.put_line (v_writefile,'<ss:Styles> <ss:Style ss:ID="Header"> <ss:Font ss:background-color="grey" ss:Bold="1"/> </ss:Style> <ss:Style ss:ID="OracleDate"> <ss:NumberFormat ss:Format="dd/mm/yyyy\ hh:mm:ss"/> </ss:Style> </ss:Styles>');
for i in 1 .. t_worksheet.count
loop
utl_file.put_line (v_writefile,'<ss:Worksheet ss:Name="' || t_worksheet (i).worksheet_name || '">');
utl_file.put_line (v_writefile, '<ss:Table>');
c := dbms_sql.open_cursor;
dbms_sql.parse (c, t_worksheet (i).worksheet_query, dbms_sql.native);
d := dbms_sql.execute (c);
dbms_sql.describe_columns (c, col_cnt, rec_tab);
for j in 1..col_cnt
loop
utl_file.put_line (v_writefile, '<ss:Column ss:Width="80"/>');
case rec_tab(j).col_type
when 1 then
dbms_sql.define_column(c,j,v_v_val,4000);
when 2 then
dbms_sql.define_column(c,j,v_n_val);
when 12 then
dbms_sql.define_column(c,j,v_d_val);
else
dbms_sql.define_column(c,j,v_v_val,4000);
end case;
end loop;
utl_file.put_line(v_writefile,'<ss:Row ss:StyleID="Header">');
-- Output the header
for j in 1..col_cnt
loop
utl_file.put_line(v_writefile,'<ss:Cell>');
utl_file.put_line(v_writefile,'<ss:Data ss:Type="String">'||rec_tab(j).col_name||'</ss:Data>');
utl_file.put_line(v_writefile,'</ss:Cell>');
end loop;
utl_file.put_line(v_writefile,'</ss:Row>');
-- Output the data
loop
v_ret := dbms_sql.fetch_rows(c);
exit
when v_ret = 0;
utl_file.put_line(v_writefile,'<ss:Row>');
for j in 1..col_cnt
loop
case rec_tab(j).col_type
when 1 then
dbms_sql.column_value(c,j,v_v_val);
utl_file.put_line(v_writefile,'<ss:Cell>');
utl_file.put_line(v_writefile,'<ss:Data ss:Type="String">'||v_v_val||'</ss:Data>');
utl_file.put_line(v_writefile,'</ss:Cell>');
when 2 then
dbms_sql.column_value(c,j,v_n_val);
utl_file.put_line(v_writefile,'<ss:Cell>');
utl_file.put_line(v_writefile,'<ss:Data ss:Type="Number">'||to_char(v_n_val)||'</ss:Data>');
utl_file.put_line(v_writefile,'</ss:Cell>');
when 12 then
dbms_sql.column_value(c,j,v_d_val);
utl_file.put_line(v_writefile,'<ss:Cell ss:StyleID="OracleDate">');
utl_file.put_line(v_writefile,'<ss:Data ss:Type="DateTime">'||to_char(v_d_val,'YYYY-MM-DD"T"HH24:MI:SS')||'</ss:Data>');
utl_file.put_line(v_writefile,'</ss:Cell>');
else
dbms_sql.column_value(c,j,v_v_val);
utl_file.put_line(v_writefile,'<ss:Cell>');
utl_file.put_line(v_writefile,'<ss:Data ss:Type="String">'||v_v_val||'</ss:Data>');
utl_file.put_line(v_writefile,'</ss:Cell>');
end case;
end loop;
utl_file.put_line(v_writefile,'</ss:Row>');
end loop;
dbms_sql.close_cursor(c);
utl_file.put_line (v_writefile, '</ss:Table>');
utl_file.put_line (v_writefile, '</ss:Worksheet>');
end loop;
utl_file.put_line (v_writefile, '</ss:Workbook>');
utl_file.fclose (v_writefile);
utl_file.fclose(v_writefile);
exception
when others then
raise_application_error(-20001,'An error was encountered - '||sqlcode||' -ERROR- '||sqlerrm);
end create_multitab_xls_report;
