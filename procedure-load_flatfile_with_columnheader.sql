create or replace PROCEDURE load_flatfile_with_columnheader
(
v_filename   VARCHAR2,
v_directory  VARCHAR2,
v_delimiter  VARCHAR2,
v_table_name VARCHAR2 )
AS
/****************************
PURPOSE: USING ANY FLATFILE (STORED IN DIRECTORY)
CREATE AND INSERT DATA INTO DATABASE USING PL/SQL

PARAMETERS:
V_FILENAME: 
NAME OF THE FLATFILE I.E. CSV , XLS ETC

V_DIRECTORY: 
NAME OF THE DIRECTORY WHERE FILE IS STORED.
NOTE: USER ACCESS IS REQUIRED ON THE DIRECTORY BEFORE EXECUTION OF THE PROCEDURE

V_DELIMITER:
MENTION THE DELIMITER USED IN FLATFILE

V_TABLE_NAME:
NAME OF THE TABLE TO BE CREATED IN DATABASE

OTHER NOTES:
IMPORTANT:
CREATE NESTED TABLE TYPE USING FOLLOWING COMMAND.USER ACCESS IS REQUIRED TO CREATE TYPE.
create or replace type t_str_col as table of varchar2(4000);

TABLE CREATED WILL HAVE DATATYPE VARCHAR2(100) FOR ALL COLUMNS.
USER CAN TWEAK THE CODE IF DIFFERENT DATATYPE IS REQUIRED.

AUTHOR: SINGHINDU87
VERSION: 1.0
CONCEPTS USED: UTL_FILE, NESTED TABLES

****************************/

f_handler utl_file.file_type;
f_header     VARCHAR2(2000):=NULL;
f_line       VARCHAR2(2000):=NULL;
f_counter    NUMBER        :=0;
header_count NUMBER        :=0;
l_index      NUMBER        :=0;
f_tab t_str_col            :=t_str_col();
f_sql               VARCHAR2(1000)       :=NULL;
f_sub_sql           VARCHAR2(1000)       :=NULL;
f_table             NUMBER               :=0;
table_already_exist EXCEPTION;
BEGIN
SELECT COUNT(*)
INTO f_table
FROM all_tables
WHERE table_name=upper(v_table_name);
IF f_table      =0 THEN
f_handler    :=utl_file.fopen( v_directory,v_filename,'r');
f_counter    :=1;
IF utl_file.is_open(f_handler) THEN
LOOP
BEGIN
utl_file.get_line(f_handler, f_line);

/*** create header**/ --step can be ignored if table structure is already created
IF f_counter =1 THEN
f_header  := f_line;
SELECT (LENGTH (REPLACE(f_header,v_delimiter,'  ')) -LENGTH(f_header))+1
INTO header_count
FROM dual;

dbms_output.put_line('delimiter selected: '|| v_delimiter || ' no of columns: ' || header_count);
/** get column name into array f_tab(i) **/
l_index :=1;
f_header:=f_header|| v_delimiter;
dbms_output.put_line('f_header '|| f_header );
f_sql := 'create table '||v_table_name||'( ';
FOR i IN 1..header_count
LOOP
EXIT
WHEN instr(f_header, v_delimiter,l_index)=0;
f_tab.EXTEND;
f_tab(i):=SUBSTR(f_header,l_index,instr(f_header, v_delimiter,l_index)-l_index);
l_index :=instr(f_header, v_delimiter,l_index)                        +1;
f_sub_sql := f_sub_sql ||f_tab(i)||' varchar2(100),'; --create column with datatype varchar2(100);
dbms_output.put_line('f_sub_sql: '|| f_sub_sql );
END LOOP;
f_sub_sql:=SUBSTR(f_sub_sql,1,LENGTH(f_sub_sql)-1);
f_sql    := f_sql||f_sub_sql ||')';
dbms_output.put_line('f_sql: '|| f_sql );
EXECUTE IMMEDIATE f_sql;
/**end creating header **/
ELSE
/***start inserting data **/
f_sql:=null;
f_sub_sql:=null;
l_index :=1;
f_line:=f_line|| v_delimiter;
dbms_output.put_line('f_line: '|| f_line );
f_tab.delete;
f_sql:='insert into ' || v_table_name ||' values(';
FOR i IN 1..header_count
LOOP
EXIT
WHEN instr(f_line, v_delimiter,l_index)=0;
f_tab.EXTEND;
f_tab(i):=SUBSTR(f_line,l_index,instr(f_line, v_delimiter,l_index)-l_index);
l_index :=instr(f_line, v_delimiter,l_index)                        +1;
f_sub_sql := f_sub_sql||''''||f_tab(i)||''',';
dbms_output.put_line('f_sub_sql: '|| f_sub_sql );
END LOOP;
f_sub_sql:=substr(f_sub_sql,0,length(f_sub_sql)-1); -- to remove extra delimiter
f_sql    := f_sql||f_sub_sql ||')';
dbms_output.put_line('f_sql: '|| f_sql );
EXECUTE IMMEDIATE f_sql;
COMMIT;
END IF;

f_counter:=f_counter+1;
END;
END LOOP;
END IF;
utl_file.fclose(f_handler);
ELSE
RAISE table_already_exist;
END IF;
EXCEPTION
WHEN no_data_found THEN
raise_application_error(-20002,'An error was encountered - '||SQLCODE||' -ERROR- '||SQLERRM);
WHEN table_already_exist THEN
raise_application_error(-20003,'An error was encountered - '||SQLCODE||' -ERROR- '||SQLERRM);
WHEN OTHERS THEN
raise_application_error(-20001,'An error was encountered - '||SQLCODE||' -ERROR- '||SQLERRM);
END;
