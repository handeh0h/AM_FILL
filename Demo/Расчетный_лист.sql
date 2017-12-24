Function am_fill_demo
( p_xlsx BLOB  -- шаблон листка: Расчетный_лист.xlsx
) return BLOB
as
  type tp_names is table of varchar2(100);
  type tp_sums is table of number(10,2);
-- Для вывода переменной части листка необходима таблица (матрица)
-- значений, которая для простоты представлена набором векторов.
-- Расчетная часть опущена.
  l_inames tp_names := tp_names('Оклад','Премия', 'Отпуск','всего'); 
  l_isums  tp_sums := tp_sums(60000, 120000, 40000, 220000); 
  l_idays  tp_names := tp_names('32 дн','I кв','20 дн');
  l_rnames tp_names := tp_names('НДФЛ','','','всего удержано');
  l_rsums  tp_sums := tp_sums(28600, null, null, 28600); 
  l_xlsx BLOB;
begin
-- инициализация шаблоном с разрешением exception по ошибке имени
  am_fill.init(p_xlsx,'e');
-- заполнение именованных областей титульной части 
  am_fill.in_field(sysdate, 'Расчетная_дата');
  am_fill.in_field('Привалов Александр Иванович', 'ФИО_сотрудника');
  am_fill.in_field('заведующий', 'Должность'); 
  am_fill.in_field('вычислительный центр', 'Подразделение'); 
  am_fill.in_field(l_isums(1), 'Оклад'); 
-- заполнение переменной части в режиме последовательной вставки строк 
  for k in 1..3 loop
-- колонки (поля) предпочтительно именовать
    am_fill.in_field(l_inames(k), 'Нач_Вид', 'i'); 
    am_fill.in_field(l_isums(k),  'C8', 'i'); 
    am_fill.in_field(l_idays(k),  'D8', 'i'); 
    am_fill.in_field(l_rnames(k), 'E8', 'i');
    am_fill.in_field(l_rsums(k),  'G8', 'i');
  end loop;
-- заполнение итоговой части
  am_fill.in_field(l_isums(4), 'C9');
  am_fill.in_field(l_rsums(4), 'G9');
  am_fill.in_field(l_isums(4)-l_rsums(4), 'К_выплате');
-- формирование документа
  am_fill.finish(l_xlsx);
  return l_xlsx;
end;
