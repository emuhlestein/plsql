Custom Solution (SQL)
If you are using SQL*Plus, a simple solution is to avoid PL/SQL altogether and concatenate all the column values together with dividing commas. The output from this type of query can be spooled out to a file.

SET LINESIZE 1000 TRIMSPOOL ON PAGESIZE 0 FEEDBACK OFF

SPOOL c:\oracle\extract\emp.csv

SELECT empno || ',' ||
       ename || ',' ||
       job || ',' ||
       mgr || ',' ||
       TO_CHAR(hiredate,'DD-MON-YYYY') AS hiredate || ',' ||
       sal || ',' ||
       comm || ',' ||
       deptno
FROM   emp
ORDER BY ename;

SPOOL OFF

SET PAGESIZE 14
Custom Solution (PL/SQL)
Define a directory object which points to an existing filesystem directory on the server. We must grant the necessary access privilege on the directory object to the user who will perform the extract.

CONNECT / AS SYSDBA
CREATE OR REPLACE DIRECTORY EXTRACT_DIR AS 'c:\oracle\extract';
GRANT READ, WRITE ON DIRECTORY EXTRACT_DIR TO SCOTT;
GRANT EXECUTE ON UTL_FILE TO SCOTT;
Next we create the extract procedure.

CONNECT scott/tiger

CREATE OR REPLACE PROCEDURE EMP_CSV AS
  CURSOR c_data IS
    SELECT empno,
           ename,
           job,
           mgr,
           TO_CHAR(hiredate,'DD-MON-YYYY') AS hiredate,
           sal,
           comm,
           deptno
    FROM   emp
    ORDER BY ename;
    
  v_file  UTL_FILE.FILE_TYPE;
BEGIN
  v_file := UTL_FILE.FOPEN(location     => 'EXTRACT_DIR',
                           filename     => 'emp_csv.txt',
                           open_mode    => 'w',
                           max_linesize => 32767);
  FOR cur_rec IN c_data LOOP
    UTL_FILE.PUT_LINE(v_file,
                      cur_rec.empno    || ',' ||
                      cur_rec.ename    || ',' ||
                      cur_rec.job      || ',' ||
                      cur_rec.mgr      || ',' ||
                      cur_rec.hiredate || ',' ||
                      cur_rec.empno    || ',' ||
                      cur_rec.sal      || ',' ||
                      cur_rec.comm     || ',' ||
                      cur_rec.deptno);
  END LOOP;
  UTL_FILE.FCLOSE(v_file);
  
EXCEPTION
  WHEN OTHERS THEN
    UTL_FILE.FCLOSE(v_file);
    RAISE;
END;
/
We are now able to perform the extract as follows.

EXEC EMP_CSV;
In the event of a problem the current exception handler gives no useful information. In order to report something more useful we could alter the exception handler as follows.

EXCEPTION
  WHEN UTL_FILE.INVALID_PATH THEN
    UTL_FILE.FCLOSE(v_file);
    RAISE_APPLICATION_ERROR(-20000, 'File location is invalid.');
    
  WHEN UTL_FILE.INVALID_MODE THEN
    UTL_FILE.FCLOSE(v_file);
    RAISE_APPLICATION_ERROR(-20001, 'The open_mode parameter in FOPEN is invalid.');

  WHEN UTL_FILE.INVALID_FILEHANDLE THEN
    UTL_FILE.FCLOSE(v_file);
    RAISE_APPLICATION_ERROR(-20002, 'File handle is invalid.');

  WHEN UTL_FILE.INVALID_OPERATION THEN
    UTL_FILE.FCLOSE(v_file);
    RAISE_APPLICATION_ERROR(-20003, 'File could not be opened or operated on as requested.');

  WHEN UTL_FILE.READ_ERROR THEN
    UTL_FILE.FCLOSE(v_file);
    RAISE_APPLICATION_ERROR(-20004, 'Operating system error occurred during the read operation.');

  WHEN UTL_FILE.WRITE_ERROR THEN
    UTL_FILE.FCLOSE(v_file);
    RAISE_APPLICATION_ERROR(-20005, 'Operating system error occurred during the write operation.');

  WHEN UTL_FILE.INTERNAL_ERROR THEN
    UTL_FILE.FCLOSE(v_file);
    RAISE_APPLICATION_ERROR(-20006, 'Unspecified PL/SQL error.');

  WHEN UTL_FILE.CHARSETMISMATCH THEN
    UTL_FILE.FCLOSE(v_file);
    RAISE_APPLICATION_ERROR(-20007, 'A file is opened using FOPEN_NCHAR, but later I/O ' ||
                                    'operations use nonchar functions such as PUTF or GET_LINE.');

  WHEN UTL_FILE.FILE_OPEN THEN
    UTL_FILE.FCLOSE(v_file);
    RAISE_APPLICATION_ERROR(-20008, 'The requested operation failed because the file is open.');

  WHEN UTL_FILE.INVALID_MAXLINESIZE THEN
    UTL_FILE.FCLOSE(v_file);
    RAISE_APPLICATION_ERROR(-20009, 'The MAX_LINESIZE value for FOPEN() is invalid; it should ' || 
                                    'be within the range 1 to 32767.');

  WHEN UTL_FILE.INVALID_FILENAME THEN
    UTL_FILE.FCLOSE(v_file);
    RAISE_APPLICATION_ERROR(-20010, 'The filename parameter is invalid.');

  WHEN UTL_FILE.ACCESS_DENIED THEN
    UTL_FILE.FCLOSE(v_file);
    RAISE_APPLICATION_ERROR(-20011, 'Permission to access to the file location is denied.');

  WHEN UTL_FILE.INVALID_OFFSET THEN
    UTL_FILE.FCLOSE(v_file);
    RAISE_APPLICATION_ERROR(-20012, 'The ABSOLUTE_OFFSET parameter for FSEEK() is invalid; ' ||
                                    'it should be greater than 0 and less than the total ' ||
                                    'number of bytes in the file.');

  WHEN UTL_FILE.DELETE_FAILED THEN
    UTL_FILE.FCLOSE(v_file);
    RAISE_APPLICATION_ERROR(-20013, 'The requested file delete operation failed.');

  WHEN UTL_FILE.RENAME_FAILED THEN
    UTL_FILE.FCLOSE(v_file);
    RAISE_APPLICATION_ERROR(-20014, 'The requested file rename operation failed.');

  WHEN OTHERS THEN
    UTL_FILE.FCLOSE(v_file);
    RAISE;
END;
Generic Solution
For a simple generic solution, load the csv.sql package into the SCOTT schema.

CONNECT scott/tiger
@csv.sql
Using the directory object created previously, execute the following commands.

-- Query.
EXEC csv.generate('EXTRACT_DIR', 'emp.csv', p_query => 'SELECT * FROM emp');

-- REF CURSOR.
DECLARE
  l_refcursor  SYS_REFCURSOR;
BEGIN
  OPEN l_refcursor FOR
    SELECT * FROM emp;

  csv.generate_rc('DBA_DIR','generate.csv', l_refcursor);
END;
/
