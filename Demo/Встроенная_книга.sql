Function am_fill_demo1 return BLOB
as
  l_query varchar2(2000);
  l_cursor am_fill.ref_cursor;
  l_bxlsx BLOB;
begin
/* инициализация встроенной книгой, стили дат замещаются */
  am_fill.init(am_fill.new_workbook(),'ed');
/* одна таблица на листе, запросы без завершающей ; */
  l_query := 'SELECT rownum rn, a.* FROM ALL_OBJECTS a 
     WHERE OBJECT_TYPE IN (''FUNCTION'',''PROCEDURE'',''PACKAGE'')';
  am_fill.in_table(l_query, 'A1', 'h');
  am_fill.in_sheet(‘Sheet1’,'Одна таблица на листе');
/* две таблицы на листе */
  am_fill.in_table(l_query, 'A1', 'hi');
  open l_cursor for  'SELECT OBJECT_NAME, OBJECT_TYPE FROM USER_OBJECTS';
  am_fill.in_table(l_cursor, 'A4', 'h');
/* сохраним с новым именем и скроем исходный лист */
  am_fill.in_sheet('Sheet1', 'Две таблицы на листе', 'h');
/* сформируем xlsx */
  am_fill.finish(l_bxlsx);
  return l_bxlsx;
end;
