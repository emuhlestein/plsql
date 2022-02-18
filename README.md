PL/SQL code to create a CSV file

Yes, excel is quite happy with CSV files (comma separated values). Here is a simple stored procedure that shows this at work:

create or replace function dump_csv(
  p_query in varchar2,
  p_separator in varchar2 default ',',
  p_dir in varchar2 ,
  p_filename in varchar2 )  return number is
  
  l_output utl_file.file_type;
  l_theCursor integer default dbms_sql.open_cursor;
  l_columnValue varchar2(2000);
  l_status integer;
  l_colCnt number default 0;
  l_separator varchar2(10) default '';
  l_cnt number default 0;
begin
  l_output := utl_file.fopen( p_dir, p_filename, 'w' );

  dbms_sql.parse( l_theCursor, p_query, dbms_sql.native );

  for i in 1 .. 255 loop
  begin
    dbms_sql.define_column( l_theCursor, i, l_columnValue, 2000 );
    l_colCnt := i;
    exception
      when others then
      if ( sqlcode = -1007 ) then
        exit;
      else
        raise;
      end if;
  end;
  end loop;

  dbms_sql.define_column( l_theCursor, 1, l_columnValue, 2000 );

  l_status := dbms_sql.execute(l_theCursor);

  loop
  exit when ( dbms_sql.fetch_rows(l_theCursor) <= 0 );
  l_separator := '';
  for i in 1 .. l_colCnt loop
    dbms_sql.column_value( l_theCursor, i, l_columnValue );
    utl_file.put( l_output, l_separator || l_columnValue );
    l_separator := p_separator;
  end loop;
  utl_file.new_line( l_output );
  l_cnt := l_cnt+1;
end loop;
dbms_sql.close_cursor(l_theCursor);

utl_file.fclose( l_output );
return l_cnt;
end dump_csv;
/
