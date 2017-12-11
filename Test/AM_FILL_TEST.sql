create or replace FUNCTION am_fill_test
( p_workbook BLOB := null
, p_options varchar2:=''
) return BLOB
as
  l_query varchar2(2000);
  l_cursor am_fill.ref_cursor;
  l_bxlsx BLOB;
  l_rows number := 5;
begin
  if p_workbook is null then
/* init by build-in workbook */
    am_fill.init(am_fill.new_workbook(), p_options);
  else
/* init by user-defined workbook */
    am_fill.init(p_workbook, p_options);
  end if;
/* One table per hidden Sheet0 */
  l_query := 'SELECT rownum rn, o.* FROM ALL_OBJECTS o 
     WHERE OBJECT_TYPE IN (''FUNCTION'',''PROCEDURE'',''PACKAGE'')';
--  l_query := l_query||' and rownum <=200';
  am_fill.in_field('query B1','sheet0!A1');
  am_fill.in_table(l_query, 'B1', 'h');
  am_fill.in_sheet('sheet0','One table per hidden Sheet0'); 

/* Three tables per Sheet1 */
  l_query := 'SELECT rownum rn, o.object_name,o.created,o.timestamp from all_objects o where rownum <= 5';
  am_fill.in_field('query B1 hi','Sheet1!A1');
  am_fill.in_table(l_query, 'B1', 'hi');

  am_fill.in_field('fields B2:E2 i','A2');
  for c in (select rownum rn, o.* from all_objects o where rownum <= 5)
  loop
    am_fill.in_field(c.rn, 'B2', 'i');
    am_fill.in_field(c.OBJECT_NAME, 'C2', 'i');
    am_fill.in_field(c.CREATED, 'E2', 'i');
    am_fill.in_field(c.TIMESTAMP, 'D2', 'i');
  end loop;

  am_fill.in_field('cursor B4 i','A4');
  open l_cursor for l_query;
  am_fill.in_table(l_cursor, 'B4', 'hi');
  
  am_fill.in_field('Success?','H3');
  am_fill.in_sheet('Sheet1', 'Three tables per Sheet1', 'b');
/* Save filled workbook */
  am_fill.finish(l_bxlsx);
  return l_bxlsx;
exception
  when others then
    am_fill.init;
    dbms_output.put_line ( DBMS_UTILITY.FORMAT_ERROR_STACK() );
    dbms_output.put_line ( DBMS_UTILITY.FORMAT_ERROR_BACKTRACE() );
  return null;
END;
