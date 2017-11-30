create or replace FUNCTION am_fill_test(p_workbook BLOB := null) return BLOB
as
  l_query varchar2(2000);
  l_cursor am_fill.ref_cursor;
  l_bxlsx BLOB;
  l_rows number := 5;
begin
  if p_workbook is null then
/* init by build-in workbook */
    am_fill.init(am_fill.new_workbook(),'e');
  else
/* init by user-defined workbook */
    am_fill.init(p_workbook, 'e');
  end if;
/* One table per sheet */
  l_query := 'SELECT rownum rn, o.* FROM ALL_OBJECTS o 
     WHERE OBJECT_TYPE IN (''FUNCTION'',''PROCEDURE'',''PACKAGE'')';
  am_fill.in_table(l_query, 'B1', 'h');
  am_fill.in_sheet('Sheet1','One table per sheet'); 
/* Three tables per sheet */
  am_fill.in_table(l_query||' and rownum <= '||l_rows, 'B1', 'hi');
  for c in (select rownum r, o.* from ALL_OBJECTS o where rownum <= l_rows)
  loop
    am_fill.in_field(c.r, 'B2', 'i');
    am_fill.in_field(c.OBJECT_NAME, 'C2', 'i');
    am_fill.in_field(c.OBJECT_ID, 'D2', 'i');
    am_fill.in_field(c.CREATED, 'E2', 'i');
  end loop;
  open l_cursor for SELECT OBJECT_NAME, OBJECT_TYPE FROM ALL_OBJECTS
    where rownum <= l_rows;
  am_fill.in_table(l_cursor, 'B4', 'hi');
/* Save and hide source sheet */
  am_fill.in_sheet('Sheet1', 'Three tables per sheet', 'bh');
--  am_fill.in_sheet('Sheet0','Empty sheet'); --???
/* Save filled workbook */
  am_fill.finish(l_bxlsx);
  return l_bxlsx;
  exception
  when others then
    am_fill.init;
    dbms_output.put_line ( DBMS_UTILITY.FORMAT_ERROR_STACK() );
    dbms_output.put_line ( DBMS_UTILITY.FORMAT_ERROR_BACKTRACE() );
    raise;
end;
