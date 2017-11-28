create or replace FUNCTION am_fill_test(p_workbook BLOB := null) return BLOB
as
  l_query varchar2(2000);
  l_cursor am_fill.ref_cursor;
  l_bxlsx BLOB;
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
     WHERE OBJECT_TYPE IN (''FUNCTION'',''PROCEDURE'',''PACKAGE'') and rownum <= 10';
  am_fill.in_table(l_query, 'A1', 'h');
  am_fill.in_sheet('Sheet1','One table per sheet'); 
/* Three tables per sheet */
  for c in (select rownum r, o.* from USER_OBJECTS o)
  loop
    am_fill.in_field(c.r, 'A2', 'i');
    am_fill.in_field(c.OBJECT_NAME, 'B2', 'i');
    am_fill.in_field(c.OBJECT_ID, 'C2', 'i');
    am_fill.in_field(c.CREATED, 'D2', 'i');
  end loop;
  am_fill.in_table(l_query, 'A1', 'hi');
  open l_cursor for SELECT OBJECT_NAME, OBJECT_TYPE FROM ALL_OBJECTS where rownum < 300;
  am_fill.in_table(l_cursor, 'A4', 'h');
/* Save and hide source sheet */
  am_fill.in_sheet('Sheet1', 'Three tables per sheet', 'bh');
  am_fill.in_sheet('Sheet0','Empty sheet'); --???
/* Save filled workbook */
  am_fill.finish(l_bxlsx);
  return l_bxlsx;
end;
