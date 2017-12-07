create or replace PACKAGE "AM_FILL" as  
/*   
*    Using the Microsoft XLSX files as a templates for data output
*
*    Require AS_ZIP package designed by Anton Scheffer  
*    Download from: http://technology.amis.nl/wp-content/uploads/2010/06/as_zip7.txt  
*  
*    Based on the code of packages designed by Anton Scheffer  
*        https://technology.amis.nl/wp-content/uploads/2011/02/as_xlsx11.txt  
*        https://technology.amis.nl/wp-content/uploads/2013/01/as_read_xlsx9.txt  
*   
*    Author: miktim@mail.ru, Petrozavodsk State University   
*    Date: 2013-06-24   
*    Updated:   
*      2017-12-.. refactoring in progress...  
*                 active sheet detection, sheet names, 
*                 in_table, in_sheet, new_workbook, address,...  
*      2017-03-28 fixed bug: 200 rows 'limitation' (IN_TABLE),   
*                 options added (IN_SHEET),   
*      2016-06-08 fixed bug: calc new rId (IN_SHEET)   
*      2016-06-07 fixed bugs: get sheet xml name, sheets without merges (align_loc)   
*                 Thanks github.com/Zulus88   
*      2016-02-01 free xmlDocument objects   
*      2015-03-18 procedure in_sheet added 
*      2015-02-27 support Oracle non UTF-8 charSets.   
******************************************************************************   
* Copyright (C) 2011 - 2013 Anton Scheffer (as_xlsx, as_read_xlsx, as_zip)  
*               2013 - 2017 MikTim   
* License: MIT   
******************************************************************************   
*/  
version constant varchar2(10):='71200';  
/* 
  Exception messages: 
    #WORKBOOK! unknown structure  
    #SHEET!... 
    #REF!... unknown sheet/named range
*/ 
data_error EXCEPTION; 
PRAGMA EXCEPTION_INIT(data_error, -20711); 
/*    
   INIT: Initialize package by xlsx template   
   p_options:   
     e - enable exception on #REF!, otherwise ignore filling
     d - replace the common style for DATEs by user-defined or workbook default
*/         
Procedure init  
( p_xtemplate BLOB   -- xlsx template BLOB
, p_options varchar2:=''     
);  
/* INIT: Clear internal structures */   
Procedure init;   
/*    
   IN_FIELD: Fill in cell. Destination sheet becomes active (current).

   p_address: relative A1 style cell address. 
     You can not enclose in single quotes sheet names without !$
   Exаmples:
     am_fill.in_field(123, 'First sheet!a12'); 
     am_fill.in_field('string', 'Named_Area!A12);
     am_fill.in_field(sysdate, 'Named_Area');
     am_fill.in_field(null, 'A11'); 
      
   p_options:   
     i - row insert mode (sequentially on every call), default - overwrite   
         WARNING: cut vertical merges
*/   
Procedure in_field   
( p_value date   
, p_address varchar2   
, p_options varchar2:='');   
Procedure in_field   
( p_value number   
, p_address varchar2   
, p_options varchar2:='');   
Procedure in_field   
( p_value varchar2   
, p_address varchar2   
, p_options varchar2:='');   
/*    
   IN_TABLE: Fill in table   
   p_table: ref_cursor or sql query text (without trailing semicolon)  
   p_address: cell address for first field (see IN_FIELD) 
   p_options:   
     h - print headings (field names)   
     i - rows insert mode. WARNING: cut vertical merges.
*/   
Type ref_cursor is REF CURSOR;  
  
Procedure in_table   
( p_table in out ref_cursor   
, p_address varchar2    
, p_options varchar2 := '');  
  
Procedure in_table   
( p_table CLOB   
, p_address varchar2  
, p_options varchar2 := '');  
  
/*   
   IN_SHEET: Save FILLED sheet with new name AFTER source sheet,   
     clears data from source sheet, new sheet becomes active and visible.   
   WARNING:
     all data, except numbers, strings, dates, will be REMOVED from the new sheet
   p_newsheet_name cannot contain /\*[]:?'
   p_options:   
     h - hide source sheet   
     b - insert BEFORE source sheet   
*/   
Procedure in_sheet   
( p_sheet_name varchar2      -- filled sheet name 
, p_newsheet_name varchar2   -- new sheet name (max 31 char)   
, p_options varchar2:=''     -- options   
);   

/*   
   FINISH: Generate workbook. Save FILLED sheets, clear internal structures.   
   WARNING: 
     all data, except numbers, strings, dates, will be REMOVED from saved sheets  
*/   
Procedure finish   
( p_xfile in out nocopy BLOB -- filled in xlsx returns   
);

/*
   ADDRESS: returns relative A1 style cell address or null on #REF!.  
   p_address - cell address (see IN_FIELD).
   p_options: 
     o - align left-unaligned outside merge 
     l - align to left cell in merge
     t - align to top cell in merge
*/
Function address   
( p_row pls_integer          -- from 1     
, p_col pls_integer          -- from 1
, p_address varchar2 
, p_options varchar2 := ''   --  o l t 
) return varchar2;

/*   
   NEW_WORKBOOK: returns workbook with two sheets:    
     'Sheet1' - visible.   
     'Sheet0' - hidden. A1 cell formatted as date (YYYY-MM-DD)  
*/   
Function new_workbook return BLOB;   
  
end;  
/
create or replace PACKAGE BODY  "AM_FILL" is  
 
c_errcode constant number := -20711; 
 
c_sheeturl constant varchar2(200) := 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'; 
c_sheetns  constant varchar2(200) := 'xmlns="'||c_sheeturl||'"'; 
c_rurl   constant varchar2(200) := 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'; 
c_rns    constant varchar2(200) := 'xmlns:r="'||c_rurl||'"'; 
c_relsns constant varchar2(200) := 'xmlns="http://schemas.openxmlformats.org/package/2006/relationships"';  
   
type tp_one_range is record  
  ( range_name varchar2(200)  -- sheet or range name  
  , isSheet boolean := false  
  , sht_nr  pls_integer := null  -- sheet number  
  , row_nr  pls_integer := null  -- upper left row  
  , col_nr  pls_integer := null  -- upper left col
  );  
type tp_all_ranges is table of tp_one_range index by pls_integer;  
--  
type tp_loc is record  
  ( sht_nr pls_integer  := null
  , row_nr pls_integer  := null  
  , row_off pls_integer := 0  --  v_size-1 for merges, inserted row index  
  , col_nr pls_integer  := null  
  , col_off pls_integer := 0  --  h_size-1 for merges  
  );  
--  
type tp_one_field is record  
  ( type char(1)  -- N,D,S  
  , value number  -- for strings (S) index tp_str_ind  
  );  
-- field access: (sht)(row)(row ins)(col)  
type tp_fcols is table of tp_one_field index by pls_integer;  
type tp_frows_of is table of tp_fcols index by pls_integer;  
type tp_frows is table of tp_frows_of index by pls_integer;  
type tp_fsheets is table of tp_frows index by pls_integer;  
--  
type tp_strings is table of pls_integer index by varchar2(32767);  
type tp_str_ind is table of varchar2(32767) index by pls_integer;  
--  
type tp_one_cell_style is record  
  ( type char(1)  
  , style pls_integer  
  );  
type tp_styles is table of tp_one_cell_style index by pls_integer;  
type tp_rows is table of tp_styles index by pls_integer;  
type tp_one_row_style is record  
  ( height number  --:=12.1  
  , style pls_integer  
  );  
type tp_row_styles is table of tp_one_row_style index by pls_integer;  
type tp_sheet is record  
  ( cells tp_rows  
  , rows tp_row_styles  
  );  
type tp_one_merge is record  
  ( row_off pls_integer  -- v_size-1  
  , col_off pls_integer  -- h_size-1  
  );  
type tp_col_merge is table of tp_one_merge  index by pls_integer;  
type tp_row_merge is table of tp_col_merge index by pls_integer;  
type tp_merges is table of tp_row_merge index by pls_integer;  
type tp_context is record  
  ( template BLOB  
  , options varchar2(10)  
  , insmode boolean := false       -- insert/overwrite  
  , date1904 boolean := true       -- workbook date format  
  , date_style pls_integer := null -- default date style  
  , current_sht pls_integer := 1  
  , strings tp_strings  
  , str_ind tp_str_ind  
  , str_cnt_of pls_integer  
  , ranges tp_strings         -- sheet or range names in upper shift
  , ran_ind tp_all_ranges  
  , merges tp_merges  
  , fields tp_fsheets  
  , styles tp_sheet  
  );  
context tp_context;  
-- 
Procedure debug(str1 varchar2,str2 varchar2:=null,str3 varchar2:=null,str4 varchar2:=null,str5 varchar2:=null) 
is 
begin 
  dbms_output.put_line(trim(str1||' '||str2||' '||str3||' '||str4||' '||str5)); 
end; 
--  
Procedure clear_styles  
is  
  r pls_integer;  
begin  
  r := context.styles.cells.first();  
  while r is not null  
  loop  
    context.styles.cells( r ).delete();  
    r := context.styles.cells.next( r );  
  end loop;  
  context.styles.cells.delete();  
  context.styles.rows.delete();  
end;  
--  
Procedure clear_sheet_fields(p_s pls_integer)  
as  
  r pls_integer;  
  ro pls_integer;  
begin  
  if not context.fields.exists(p_s) then return; end if; 
  r := context.fields(p_s).first();  
  while r is not null loop  
    ro := context.fields(p_s)(r).first();  
    while ro is not null loop  
      context.fields( p_s )( r )( ro ).delete();  
      ro := context.fields( p_s )( r ).next(ro);  
    end loop;  
    context.fields( p_s )( r ).delete();  
    r := context.fields(p_s).next(r);  
  end loop;  
  context.fields( p_s ).delete();  
end;      
--  
Procedure clear_fields  
is  
  s pls_integer;  
begin  
  s := context.fields.first();  
  while s is not null loop  
    clear_sheet_fields(s);  
    s := context.fields.next( s );  
  end loop;  
  context.fields.delete();  
end;  
--  
Procedure clear_merges  
is  
  s pls_integer;  
  r pls_integer;  
begin    
  s:=context.merges.first();  
  while s is not null  
  loop  
    r := context.merges(s).first();  
    while r is not null  
    loop  
      context.merges(s)( r ).delete();  
      r := context.merges(s).next( r );  
    end loop;  
    s:=context.merges.next(s);  
  end loop;  
  context.merges.delete();  
end;  
--  
Procedure clear_context  
is  
  s pls_integer;  
  r pls_integer;  
begin  
  context.strings.delete;  
  context.str_ind.delete;  
  context.ranges.delete;  
  context.ran_ind.delete;  
  context.date1904 := true;  
  context.date_style := null;  
  context.current_sht := 1;  
  clear_styles;  
  clear_fields;  
  clear_merges;  
end;  
--  
Procedure add_style  
  ( p_style pls_integer  
  , p_type char := 'G' -- general  
  , p_row pls_integer  
  , p_col pls_integer  
  )  
is  
begin  
  context.styles.cells( p_row )( p_col ).type := p_type;  
  context.styles.cells( p_row )( p_col ).style := p_style;  
end;  
--  
Function get_cell_style(r pls_integer, c pls_integer, p_field tp_one_field) return varchar2  
is  
t_style tp_one_cell_style;  
t_s varchar2(50):='';  
t_s_ind pls_integer;  
begin  
  begin  
    t_style := context.styles.cells(r)(c);  
  exception when no_data_found then null;  
  end;  
  if nvl(t_style.type,'-') <> p_field.type  
  then  
    t_style.type:=p_field.type;  
    if (t_style.style is null and t_style.type = 'D'
        and instr(context.options,'d') > 0 and context.date_style is not null) then  
      t_style.style:=context.date_style;  
    end if;  
  end if;  
  t_s := case  
       when t_style.type='N' and p_field.value is not null then ' t="n" '  
       when t_style.type='S' and p_field.value is not null then ' t="s" '  
       when t_style.type='D' and p_field.value is not null then ' t="n" '  
       else '' end  
       ||case when t_style.style is null then ''  
            else ' s="'||t_style.style||'"' end;  
  return t_s;  
end;  
--  
function add_string( p_string varchar2 )  
return pls_integer  
is  
  t_cnt pls_integer;  
begin  
  if p_string is null then return null; end if;  
  if context.strings.exists( p_string )  
  then  
    t_cnt := context.strings( p_string );  
  else  
    t_cnt := nvl(context.str_cnt_of,0)+context.strings.count();    
    context.str_ind( t_cnt ) := nvl( p_string, '' );  
    context.strings( nvl( p_string, '' ) ) := t_cnt;  
  end if;  
  return t_cnt;  
end;  
--  
function blob2node( p_blob blob )  
return dbms_xmldom.domnode  
is  
begin  
  if p_blob is null or dbms_lob.getlength( p_blob ) = 0  
  then  
    return null;  
  end if;  
  return dbms_xmldom.makenode( dbms_xmldom.getdocumentelement( dbms_xmldom.newdomdocument( xmltype( p_blob, nls_charset_id( 'AL32UTF8' ) ) ) ) );  
end;  
--*** in-zip replace needeed!!
procedure replace1file  
  ( p_zipped_blob in out blob  
  , p_name varchar2  
  , p_content blob  
  )  
is  
  t_blob blob;  
  zip_files as_zip.file_list;  
begin  
  if p_zipped_blob is null  
  then  
    dbms_lob.createtemporary( p_zipped_blob, true );  
  end if;  
  zip_files  := as_zip.get_file_list(p_zipped_blob);  
  for i in zip_files.first() .. zip_files.last  
  loop  
     begin     
       if zip_files(i) <> p_name then  
         as_zip.add1file(t_blob  
                  , zip_files( i )  
                  , as_zip.get_file(p_zipped_blob,zip_files( i ))  
                  );  
        end if;  
        exception           -- zip entry is empty folder  
           when others then  null;  
      end;  
  end loop;  
  as_zip.add1file(t_blob, p_name, p_content);  
  as_zip.finish_zip(t_blob);  
  dbms_lob.trim(p_zipped_blob,0);  
  dbms_lob.append(p_zipped_blob,t_blob);  
end;  
--  
procedure replace1xml  
( p_msFile in out nocopy blob  
, p_filePath varchar2  
, p_xml xmlType  
)  
is  
  t_blob BLOB;  
  t_xml xmltype;  
begin  
  t_blob := p_xml.getBlobVal(nls_charset_id('AL32UTF8'),4,0);  
  replace1file( p_msFile, p_filePath, t_blob );  
  if dbms_lob.istemporary(t_blob)=1 then dbms_lob.freetemporary(t_blob); end if;  
end;  
--  
Function get1xml  
( p_msfile in out nocopy blob  
, p_filename varchar2  
) return xmlType  
is  
  t_btmp blob;  
  t_xml xmltype;  
begin  
  t_btmp := as_zip.get_file(p_msfile,p_filename);  
  if t_btmp is null or dbms_lob.getlength( t_btmp ) = 0  
  then  
    return null;  
  end if;  
  t_xml := xmltype( t_btmp, nls_charset_id( 'AL32UTF8' ) );  
  if dbms_lob.istemporary(t_btmp) = 1 then dbms_lob.freetemporary(t_btmp); end if;  
  return t_xml;  
end;  
--  
function col2alfan( p_col pls_integer )  
return varchar2  
is  
begin  
  return  
    case  
      when p_col > 702 then  
        chr( 64 + trunc( ( p_col - 27 ) / 676 ) )  
          || chr( 65 + mod( trunc( ( p_col - 1 ) / 26 ) - 1, 26 ) )  
          || chr( 65 + mod( p_col - 1, 26 ) )  
      when p_col > 26  then  
        chr( 64 + trunc( ( p_col - 1 ) / 26 ) ) || chr( 65 + mod( p_col - 1, 26 ) )  
      else chr( 64 + p_col )  
    end;  
end;  
--  
Function cell2alfan
( p_row pls_integer
, p_col pls_integer
, p_range_name varchar2:=null
, p_opt varchar2:='' -- '1' absolute style
) return varchar2  
is
  t_$ varchar2(1) := case when instr(p_opt,'1') > 0 then '' end;
  t_alfan varchar2(200) := p_range_name;
begin
  if t_alfan is not null then
    if regexp_instr(t_alfan,'^[[:alnum:]_]*$') = 0 then
      t_alfan := ''''||t_alfan||'''';
    end if;
    if p_row is null and p_col is null then return t_alfan; end if;
    t_alfan := t_alfan||'!';
  end if;
  return t_alfan||t_$||col2alfan(p_col)||t_$||p_row;  
end;  
--    
function alfan2col( p_col varchar2 )  
return pls_integer  
is  
begin  
  return ascii( substr( p_col, -1 ) ) - 64  
       + nvl( ( ascii( substr( p_col, -2, 1 ) ) - 64 ) * 26, 0 )  
       + nvl( ( ascii( substr( p_col, -3, 1 ) ) - 64 ) * 676, 0 );  
end;  
--
Function range2loc(p_ind pls_integer) return tp_loc
is
  t_loc tp_loc;
  t_ran tp_one_range := context.ran_ind(p_ind);
begin
  t_loc.sht_nr := t_ran.sht_nr; 
  t_loc.row_nr := t_ran.row_nr; 
  t_loc.col_nr := t_ran.col_nr; 
  return t_loc;
end;
--
Function addr2loc
( p_address varchar2
, raise_ref boolean := instr(context.options,'e') > 0
) return tp_loc
is
  t_loc tp_loc;
  t_range_name varchar2(100);
  t_addrp varchar2(200):= -- 
    '^(\''[^'']+\''|[^\!]+)(\!([a-zA-Z0-9\:\$]+))?$'; --'1,3 
  t_rangep varchar2(200):=
    '^(\$?([a-zA-Z]{1,3})\$?([1-9][0-9]*))(:(\$?([a-zA-Z]{1,3})\$?([1-9][0-9]*)))?$'; --2,3 6,7
  t_s varchar2(200); -- sheet/range name
  t_r varchar2(200); -- cell range (a1:b2)
begin
  if p_address is null or regexp_instr(p_address,t_addrp) = 0
  then raise data_error; end if;
  t_s:=regexp_replace(p_address,t_addrp,'\1'); -- sheet name
  t_r:=regexp_replace(p_address,t_addrp,'\3'); -- cell range
  if (t_r is not null and regexp_instr(t_r, t_rangep) = 0) -- range parsing error
  then raise data_error; end if; 
  t_range_name:= replace(t_s,''''); -- set unquoted range name
  if context.ranges.exists(upper(t_range_name)) then
    t_loc := range2loc(context.ranges(upper(t_range_name)));
  else
    if t_r is null and regexp_instr(t_s,t_rangep) > 0 then
      t_r := t_s; -- assume sheet name is cell addr
      t_loc := range2loc(context.current_sht);
    else raise data_error; end if;
  end if;
  t_loc.col_nr := 
    t_loc.col_nr+nvl(alfan2col(regexp_replace(t_r,t_rangep,'\2')),1)-1;
  t_loc.row_nr := 
    t_loc.row_nr+nvl(to_number(regexp_replace(t_r,t_rangep,'\3')),1)-1;
  t_loc.col_off := 
    nvl(alfan2col(regexp_replace(t_r,t_rangep,'\6')),0);
  t_loc.row_off := 
    nvl(to_number(regexp_replace(t_r,t_rangep,'\7')),0);
--debug(t_loc.sht_nr,t_loc.row_nr,t_loc.row_off,t_loc.col_nr,t_loc.col_off);  
  return t_loc;
exception
  when data_error then
    if raise_ref then 
      raise_application_error(c_errcode,'#REF!: '||nvl(p_address,'null'));
    else
      t_loc.sht_nr := null;
      return t_loc;
    end if;
end;
--
Function loc2addr(p_loc tp_loc, p_opt varchar2) return varchar2
is
  t_addr varchar2(500);
begin
  if not (p_loc.row_nr > 0 and p_loc.col_nr > 0) then 
    return null;
  end if;
  if p_loc.sht_nr > 0 then
    t_addr := context.ran_ind(p_loc.sht_nr).range_name;
  end if;
  t_addr := cell2alfan(p_loc.row_nr,p_loc.col_nr,t_addr,p_opt);
  if p_loc.col_off > 0 or p_loc.row_off > 0 then
    t_addr := t_addr||':'||cell2alfan(p_loc.row_off+1,p_loc.col_off+1,null,p_opt);
  end if;
  return t_addr;
end;
--  
Procedure add_range  
( p_range_name varchar2  
, p_range_def varchar2  
)  
is  
  t_cnt pls_integer;  
  t_loc tp_loc;
begin  
  if not context.ranges.exists( upper(p_range_name) )  
  then  
    t_cnt := context.ranges.count()+1;    
    context.ran_ind( t_cnt ).range_name := p_range_name;  
    context.ranges( upper(nvl( p_range_name, '' ) )) := t_cnt;  
    if p_range_def is not null then  
      t_loc := addr2loc(p_range_def);  
--debug(p_range_name);  
      context.ran_ind( t_cnt ).sht_nr := t_loc.sht_nr;  
      context.ran_ind( t_cnt ).row_nr := t_loc.row_nr;  
      context.ran_ind( t_cnt ).col_nr := t_loc.col_nr;  
    else  
      context.ran_ind( t_cnt ).sht_nr := t_cnt;  
      context.ran_ind( t_cnt ).isSheet := true;
      context.ran_ind( t_cnt ).row_nr := 1;
      context.ran_ind( t_cnt ).col_nr := 1;
    end if;          
  end if;  
end;  
--  
Function merge_exists
( p_sht_nr pls_integer
, p_row_nr pls_integer
, p_col_nr pls_integer
) return boolean  
is  
begin  
  return context.merges(p_sht_nr)(p_row_nr).exists(p_col_nr);  
exception when no_data_found then return false;  
end;  
 -- Return merge as tp_loc  
Function merge2loc  
( p_sht_nr pls_integer  
, p_row_nr pls_integer  
, p_col_nr pls_integer  
) return tp_loc  
is  
  t_loc tp_loc;  
begin  
  t_loc.sht_nr := p_sht_nr;  
  t_loc.row_nr := p_row_nr;  
  t_loc.col_nr := p_col_nr;  
  t_loc.row_off := context.merges(p_sht_nr)(p_row_nr)(p_col_nr).row_off;  
  t_loc.col_off := context.merges(p_sht_nr)(p_row_nr)(p_col_nr).col_off;  
  return t_loc;  
end;  
-- Return merge that contain loc  
Function get_merge( p_loc tp_loc ) return tp_loc  
is  
  t_loc tp_loc:=p_loc;  
  r pls_integer;  
  c pls_integer;  
begin  
  if context.merges.exists(p_loc.sht_nr) then  
    r := context.merges(p_loc.sht_nr).first();  
    while r is not null and r <= p_loc.row_nr loop  
      c := context.merges(p_loc.sht_nr)(r).first();  
      while c is not null and c <= p_loc.col_nr loop  
        t_loc := merge2loc(p_loc.sht_nr, r, c);  
        if (p_loc.row_nr between r and r+t_loc.row_off)  
          and (p_loc.col_nr between c and c+t_loc.col_off)  
        then  
          return t_loc;  
        end if;  
        c := context.merges(p_loc.sht_nr)(r).next(c);  
      end loop;  
      r := context.merges(p_loc.sht_nr).next(r);  
    end loop;  
  end if;  
  t_loc := p_loc;  
  t_loc.row_off := 0;  
  t_loc.col_off := 0;  
  return t_loc;  
end;  
-- Cut vertical merge  
Procedure cut_merge  
  ( p_mloc tp_loc  -- merge loc  
  , p_cloc tp_loc  -- cell loc  
  )  
as  
  t_dc pls_integer;  
  t_merge tp_one_merge;  
begin  
  t_merge.col_off := p_mloc.col_off;  
  if p_mloc.row_off = 0 then return; end if;  
  t_merge.row_off := 0;  
  context.merges(p_mloc.sht_nr)(p_mloc.row_nr)(p_mloc.col_nr) := t_merge;  
end;  
-- Align location to merge  
Function align_loc  
  ( p_loc tp_loc  
  , t_insert boolean := context.insmode  
  ) return tp_loc  
is  
  t_mloc tp_loc;  
  t_loc tp_loc := p_loc;  
begin  
  t_mloc := get_merge(p_loc);  
  t_loc.col_nr := t_mloc.col_nr;  
--debug(t_mloc.row_off);  
  if t_insert then cut_merge(t_mloc, t_loc);  
  else t_loc.row_nr := t_mloc.row_nr; end if;  
  return t_loc;  
end;  
--  
Function field_exists(p_loc tp_loc)  
return boolean  
is  
begin  
  return context.fields(p_loc.sht_nr)(p_loc.row_nr)(p_loc.row_off).exists(p_loc.col_nr);  
exception   
  when others then return false;  
end;  
--  
Function field_exists(s pls_integer, r pls_integer, ro pls_integer:=0, c pls_integer:=0)  
return boolean  
is  
begin  
  return context.fields(s)(r)(0).exists(c);  
exception  
  when others then return false;  
end;  
---  
Function next_row  
  ( p_loc tp_loc  
  , p_insert boolean:=context.insmode  
  ) return tp_loc  
is  
  t_loc tp_loc := p_loc;  
  t_mloc tp_loc;  
begin  
  if p_insert then  
    t_loc.row_off := t_loc.row_off + 1;  
  else  
    t_loc.row_nr := t_loc.row_nr + 1;  
--    t_mloc := get_merge(p_loc);  
--    t_loc.row_nr := t_mloc.row_nr + t_mloc.row_off + 1;  
  end if;  
  return t_loc;  
end;  
--  
Function next_col  
( p_loc tp_loc  
, p_insert boolean  
) return tp_loc  
is  
  t_loc tp_loc := p_loc;  
  t_mloc tp_loc := get_merge( p_loc );
begin  
  t_loc.col_nr := t_mloc.col_nr + t_mloc.col_off + 1;  
  return t_loc;  
end;  
--  
Procedure add_value  
( p_value number  
, p_type char  
, p_loc tp_loc  
)  
is  
  t_field tp_one_field;  
  t_loc tp_loc := p_loc;  
begin  
  context.current_sht:=p_loc.sht_nr;  
  t_field.type := p_type;  
  t_field.value := p_value;  
  t_loc := align_loc(p_loc, context.insmode); -- real alignment   
--debug(t_loc.sht_nr,t_loc.row_nr,t_loc.row_off,t_loc.col_nr,t_loc.col_off);  
  context.fields( t_loc.sht_nr )( t_loc.row_nr )( t_loc.row_off )( t_loc.col_nr ) := t_field;  
end;  
--  
Procedure add_field  
( p_value date  
, p_loc tp_loc  
)  
is  
begin  
  add_value( p_value - case when context.date1904 then to_date('01-01-1904','DD-MM-YYYY')  
          else to_date('01-01-1900','DD-MM-YYYY') end + 2  
     ,'D', p_loc);  
end;  
--  
Procedure add_field  
( p_value number  
, p_loc tp_loc  
)  
is  
begin  
  add_value(p_value, 'N', p_loc);  
end;  
--  
Procedure add_field  
( p_value varchar2  
, p_loc tp_loc  
)  
is  
begin  
  add_value(add_string(p_value), 'S', p_loc);  
end;  
--  
-- Init next inserted row  
Procedure set_cell_off(p_loc in out tp_loc)  
is  
begin  
  while field_exists(p_loc)  
  loop  
    p_loc.row_off := p_loc.row_off + 1;  
  end loop;  
end;  
--  
Procedure in_field(p_value number, p_address varchar2, p_options varchar2:='')  
is  
  t_loc tp_loc;  
begin  
  context.insmode := instr(p_options,'i') > 0;  
  t_loc := addr2loc(p_address);  
  if t_loc.sht_nr is null then return; end if;  
  t_loc := align_loc(t_loc);  
  if context.insmode then set_cell_off(t_loc); end if;  
  add_field( p_value, t_loc);  
end;  
--  
Procedure in_field(p_value date, p_address varchar2, p_options varchar2:='')  
is  
  t_loc tp_loc;  
begin  
  context.insmode := instr(p_options,'i') > 0;  
  t_loc := addr2loc(p_address);  
  if t_loc.sht_nr is null then return; end if;  
  t_loc := align_loc(t_loc);  
  if context.insmode then set_cell_off(t_loc); end if;  
  add_field( p_value, t_loc );  
end;  
--  
Procedure in_field(p_value varchar2, p_address varchar2, p_options varchar2:='')  
is  
  t_loc tp_loc;  
begin  
  context.insmode := instr(p_options,'i') > 0;  
  t_loc := addr2loc(p_address);  
  if t_loc.sht_nr is null then return; end if;  
  t_loc := align_loc(t_loc);  
  if context.insmode then set_cell_off(t_loc); end if;  
  add_field( p_value, t_loc );  
end;  
-- 
Procedure in_table(p_table CLOB, p_address varchar2, p_options varchar2:='') 
as 
  l_cursor ref_cursor; 
begin 
-- Open REF CURSOR variable: 
  OPEN l_cursor FOR p_table; 
  in_table(l_cursor, p_address, p_options); 
end; 
-- 
Procedure in_table(p_table in out ref_cursor, p_address varchar2, p_options varchar2:='')  
as  
  t_header boolean := instr(p_options, 'h') > 0;  
  t_insert boolean := instr(p_options, 'i') > 0;  
  t_loc tp_loc;  
  t_cloc tp_loc;  
  t_rloc tp_loc;  
  t_c integer;  
  t_col_cnt integer;  
  t_desc_tab dbms_sql.desc_tab2;  
  d_tab dbms_sql.date_table;  
  n_tab dbms_sql.number_table;  
  v_tab dbms_sql.varchar2_table;  
  t_bulk_size pls_integer := 200;  
  t_r integer;  
begin  
  context.insmode := t_insert;  
  t_loc := addr2loc(p_address);  
  if t_loc.sht_nr is null then return; end if;  
  t_loc := align_loc(t_loc, t_insert );  
--  t_c := dbms_sql.open_cursor;  
--  dbms_sql.parse( t_c, p_sql, dbms_sql.native );  
  t_c := DBMS_SQL.TO_CURSOR_NUMBER(p_table); 
  dbms_sql.describe_columns2( t_c, t_col_cnt, t_desc_tab );  
  t_cloc := t_loc;  
  for c in 1 .. t_col_cnt  
  loop  
    if t_header    
    then  
      add_field(t_desc_tab( c ).col_name, t_cloc);  
      t_cloc := next_col( t_cloc, t_insert );  
    end if;  
    case  
      when t_desc_tab( c ).col_type in ( 2, 100, 101 )  
      then  
        dbms_sql.define_array( t_c, c, n_tab, t_bulk_size, 1 );  
      when t_desc_tab( c ).col_type in ( 12, 178, 179, 180, 181 , 231 )  
      then  
        dbms_sql.define_array( t_c, c, d_tab, t_bulk_size, 1 );  
      when t_desc_tab( c ).col_type in ( 1, 8, 9, 96, 112 )  
      then  
        dbms_sql.define_array( t_c, c, v_tab, t_bulk_size, 1 );  
      else  
        null;  
    end case;  
  end loop;  
--  
  if t_header then t_loc := next_row(t_loc, t_insert); end if;  
--  t_r := dbms_sql.execute( t_c );  
  loop  
    t_r := dbms_sql.fetch_rows( t_c );  
    t_cloc := t_loc;  
    for c in 1 .. t_col_cnt  
    loop  
      t_rloc := t_cloc;  
      case  
        when t_desc_tab( c ).col_type in ( 2, 100, 101 )  
        then  
          dbms_sql.column_value( t_c, c, n_tab );  
          for i in 0 .. t_r - 1  
          loop  
            add_field(n_tab( i + n_tab.first() ), t_rloc);  
            t_rloc := next_row(t_rloc, t_insert);  
          end loop;  
          n_tab.delete;  
   
        when t_desc_tab( c ).col_type in ( 12, 178, 179, 180, 181 , 231 )  
        then  
          dbms_sql.column_value( t_c, c, d_tab );  
          for i in 0 .. t_r - 1  
          loop  
            add_field(d_tab( i + d_tab.first() ), t_rloc);  
            t_rloc := next_row(t_rloc, t_insert);  
          end loop;  
          d_tab.delete;  
   
        when t_desc_tab( c ).col_type in ( 1, 8, 9, 96, 112 )  
        then  
          dbms_sql.column_value( t_c, c, v_tab );  
          for i in 0 .. t_r - 1  
          loop  
            add_field( v_tab( i + v_tab.first() ), t_rloc);  
            t_rloc := next_row(t_rloc, t_insert);  
          end loop;  
          v_tab.delete;  
   
        else  
          for i in 0 .. t_r-1  
          loop  
            add_field('[unsupported]', t_rloc);  
            t_rloc := next_row(t_rloc, t_insert);  
          end loop;  
      end case;  
      t_cloc := next_col(t_cloc, t_insert);  
    end loop;  
    t_loc.row_nr := t_loc.row_nr + case when t_insert then 0 else t_r end;  
    t_loc.row_off := t_loc.row_off + case when not t_insert then 0 else t_r end;  
    exit when t_r < t_bulk_size;  
  end loop;  
  dbms_sql.close_cursor( t_c );  
  
exception  
  when others then  
    if dbms_sql.is_open( t_c ) then  
      dbms_sql.close_cursor( t_c );  
    end if;  
    raise;  
end;  
-- 
Function get_sheet_path(p_sheetName varchar2, p_xml xmlType := null) return varchar2 
is 
  t_xml xmlType; 
  t_rid varchar2(20); 
  t_xmlName varchar2(200); 
begin 
  t_xml := nvl(p_xml, get1xml(context.template, 'xl/workbook.xml' )); 
  select extract(t_xml 
      ,'/workbook/sheets/sheet[@name="'||p_sheetName||'"]/@r:id' 
      ,c_sheetns||' '||c_rns).getStringVal() 
    into t_rid 
    from dual; 
  t_xml := get1xml(context.template, 'xl/_rels/workbook.xml.rels' ); 
  select extract(t_xml 
      ,'/Relationships/Relationship[@Id="'||t_rId||'"]/@Target' 
      ,c_relsns).getStringVal() 
    into t_xmlName 
    from dual; 
--debug(t_rid||' '||'xl/'||t_xmlName); 
  return case when t_xmlName is null then null else 'xl/'||t_xmlName end; 
end; 
-- 
Function get_sheet_path(p_sht_nr pls_integer) return varchar2 
is 
  t_range tp_one_range; 
begin 
  t_range := context.ran_ind(p_sht_nr); 
  return get_sheet_path(t_range.range_name); 
end; 
--  
Procedure load_workbook(p_xlsx BLOB)  
is  
  t_val varchar2(4000);  
  t_nd dbms_xmldom.domnode;  
--  t_nd1 dbms_xmldom.domnode;  
  t_nl dbms_xmldom.domnodelist;  
--  t_nl1 dbms_xmldom.domnodelist;  
  t_dateFmtId number:=null;  
  t_ind pls_integer;  
  t_loc tp_loc;  
begin 
  t_nd := blob2node( as_zip.get_file( p_xlsx, 'xl/workbook.xml' ) );  
  context.date1904 := lower( dbms_xslprocessor.valueof( t_nd, '/workbook/workbookPr/@date1904', c_sheetns ) ) in ( 'true', '1' );  
  context.current_sht := nvl(
      dbms_xslprocessor.valueof( t_nd, '/workbook/bookViews/workbookView/@activeTab', c_sheetns ) + 1 
-- Google Docs, Excel Online
    , dbms_xslprocessor.valueof( t_nd, '/workbook/sheets/sheet[not(@state) or @state="visible"][1]/@sheetId', c_sheetns ) 
  ); 
-- load sheets 
  t_nl := dbms_xslprocessor.selectnodes( t_nd, '/workbook/sheets/sheet', c_sheetns );  
  for i in 0 .. dbms_xmldom.getlength( t_nl ) - 1  
  loop  
    t_ind := i + 1; 
--debug(t_ind);  
    add_range(dbms_xslprocessor.valueof( dbms_xmldom.item( t_nl, i ), '@name' ), '');  
 end loop; 
-- load named ranges 
  t_nl := dbms_xslprocessor.selectnodes( t_nd, '/workbook/definedNames/definedName', c_sheetns );  
  for i in 0 .. dbms_xmldom.getlength( t_nl ) - 1  
  loop  
    add_range(dbms_xslprocessor.valueof( dbms_xmldom.item( t_nl, i ), '@name' )  
      , dbms_xslprocessor.valueof( dbms_xmldom.item( t_nl, i ), '.' ) );  
  end loop;  
  dbms_xmldom.freeDocument(dbms_xmldom.getOwnerDocument(t_nd));  
-- load sheets merged cells
  for i in 1..context.ran_ind.count
  loop
    exit when not context.ran_ind(i).isSheet;
    t_nd := blob2node(as_zip.get_file( p_xlsx, get_sheet_path( i ) ));  
    t_nl := dbms_xslprocessor.selectnodes(t_nd, '/worksheet/mergeCells/mergeCell',c_sheetns);  
    for j in 0..dbms_xmldom.getlength( t_nl )-1  
    loop  
      t_loc:=addr2loc(dbms_xslprocessor.valueof( dbms_xmldom.item( t_nl, j ),'@ref',c_sheetns));  
      context.merges(t_ind)(t_loc.row_nr)(t_loc.col_nr).row_off:=t_loc.row_off;  
      context.merges(t_ind)(t_loc.row_nr)(t_loc.col_nr).col_off:=t_loc.col_off;  
    end loop;
    dbms_xmldom.freeDocument(dbms_xmldom.getOwnerDocument(t_nd));  
  end loop;
 -- define default date style  
  t_nd := blob2node( as_zip.get_file( p_xlsx, 'xl/styles.xml' ) );  
  t_nl := dbms_xslprocessor.selectnodes( t_nd, '/styleSheet/numFmts/numFmt', c_sheetns );  
  for i in 0 .. dbms_xmldom.getlength( t_nl ) - 1  
  loop  
    t_val := upper(dbms_xslprocessor.valueof( dbms_xmldom.item( t_nl, i ), '@formatCode' ));  
    if (  instr( t_val, 'DD' ) > 0  
       or instr( t_val, 'MM' ) > 0  
       or instr( t_val, 'YY' ) > 0  
       )  
    then  
      t_dateFmtId:=dbms_xslprocessor.valueof( dbms_xmldom.item( t_nl, i ), '@numFmtId' ) ;  
--debug(t_dateFmtId);  
      exit;  
    end if;  
  end loop;  
  t_dateFmtId := nvl(t_dateFmtId, 14); 
  t_nl := dbms_xslprocessor.selectnodes( t_nd, '/styleSheet/cellXfs/xf/@numFmtId', c_sheetns );  
  for i in 0 .. dbms_xmldom.getlength( t_nl ) - 1  
  loop  
    if t_dateFmtId = dbms_xmldom.getnodevalue( dbms_xmldom.item( t_nl, i ) )  
    then  
      context.date_style := i;  
      exit;  
    end if;  
  end loop;  
  dbms_xmldom.freeDocument(dbms_xmldom.getOwnerDocument(t_nd)); 
-- load strings count 
  t_nd := blob2node( as_zip.get_file( p_xlsx, 'xl/sharedStrings.xml' ) );  
  if not dbms_xmldom.isnull( t_nd )  
  then  
    context.str_cnt_of := dbms_xmldom.getlength(dbms_xslprocessor.selectnodes( t_nd, '/sst/si', c_sheetns ));  
  end if;  
  dbms_xmldom.freeDocument(dbms_xmldom.getOwnerDocument(t_nd));
end; 
--  
Procedure loadSheetStyles(p_xlsx BLOB, p_sht_nr pls_integer)  
as  
  t_loc tp_loc;  
  t_ftype char(1);  
--  t_ns varchar2(200) := 'xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"';  
  t_nd dbms_xmldom.domnode;  
  t_nl2 dbms_xmldom.domnodelist;  
  t_nl3 dbms_xmldom.domnodelist;  
  t_s varchar2(50);  
  t_nr number;  
  t_val varchar2(4000);  
  t_r varchar2(50);  
  t_t varchar2(50);  
  t_ro pls_integer;  
begin  
  clear_styles;  
  t_nd := blob2node( as_zip.get_file(p_xlsx, get_sheet_path(p_sht_nr) ) );  
  t_nl3 := dbms_xslprocessor.selectnodes( t_nd, '/worksheet/sheetData/row' );  
  for r in 0 .. dbms_xmldom.getlength( t_nl3 ) - 1 -- rows  
  loop  
    t_nr := dbms_xslprocessor.valueof(dbms_xmldom.item( t_nl3, r ),'@r');  
    t_val := dbms_xslprocessor.valueof(dbms_xmldom.item( t_nl3, r ),'@ht');  
    context.styles.rows(t_nr).height :=  
       to_number( t_val, translate( t_val, '.012345678,-+', 'D999999999' ), 'NLS_NUMERIC_CHARACTERS=.,' );  
    context.styles.rows(t_nr).style := dbms_xslprocessor.valueof(dbms_xmldom.item( t_nl3, r ),'@s');  
-- debug(r);  
    t_nl2 := dbms_xslprocessor.selectnodes( dbms_xmldom.item( t_nl3, r ), 'c' );  
    t_loc.sht_nr := p_sht_nr;  
    t_loc.row_nr := t_nr;  
    if dbms_xmldom.getlength( t_nl2 ) = 0 and not field_exists(t_loc)  
    then  
      context.fields(p_sht_nr)(t_nr)(0)(0) := null; -- no cols, row exists  
    end if;  
    for j in 0 .. dbms_xmldom.getlength( t_nl2 ) - 1 -- cols  
    loop  
      t_r := dbms_xslprocessor.valueof( dbms_xmldom.item( t_nl2, j ), '@r', c_sheetns );  
      t_loc := addr2loc(t_r);  
      t_loc.sht_nr := p_sht_nr;  
-- debug(t_r||' '||t_loc.row_nr||' '||t_loc.col_nr);  
      t_val := dbms_xslprocessor.valueof( dbms_xmldom.item( t_nl2, j ), 'v' );  
      t_t := dbms_xslprocessor.valueof( dbms_xmldom.item( t_nl2, j ), '@t' );  
      t_s := dbms_xslprocessor.valueof( dbms_xmldom.item( t_nl2, j ), '@s' );  
      if t_t is null or t_t='n'-- ( 'str', 'inlineStr', 'e' )  
      then  
        t_ftype := 'N';  
      elsif t_t = 's'  
      then  
        t_ftype := 'S';  
      else  
        t_ftype := null;  
      end if;  
      if t_ftype is not null then  
        t_nr := to_number( t_val, translate( t_val, '.012345678,-+', 'D999999999' ), 'NLS_NUMERIC_CHARACTERS=.,' );  
      else  
        t_nr := null;  
      end if;  
--debug(t_r||' '||t_t||' '||t_s||' '||t_val);  
      if not field_exists(t_loc)  
      then  
-- or replacing with initial cell value  
        context.fields(t_loc.sht_nr)(t_loc.row_nr)(0)(t_loc.col_nr).value := t_nr;  
        context.fields(t_loc.sht_nr)(t_loc.row_nr)(0)(t_loc.col_nr).type := t_ftype;  
      end if;  
      add_style(t_s, t_ftype, t_loc.row_nr, t_loc.col_nr);  
    end loop; -- col  
  end loop; -- row  
  dbms_xmldom.freeDocument(dbms_xmldom.getOwnerDocument(t_nd));  
end;  
--  
Procedure init  
as 
begin  
  clear_context;  
  if dbms_lob.istemporary(context.template) = 1  
    then dbms_lob.trim(context.template,0);  
    else dbms_lob.createtemporary(context.template,false);  
  end if;  
end;  
--  
Procedure init(p_xtemplate BLOB, p_options varchar2:='')  
as 
begin  
  init; 
  context.options:=p_options;  
  dbms_lob.append(context.template, p_xtemplate); 
  if get1xml( context.template, 'xl/_rels/workbook.xml.rels' ) is null then 
    init; 
    raise_application_error(c_errcode,'#WORKBOOK! unknown structure'); 
  end if; 
  load_workbook(context.template);  
end;  
--  
Procedure replace_node(p_xml in out xmlType, p_xpath varchar2, p_newnode CLOB)  
as  
begin  
  select updateXML(p_xml, p_xpath, p_newnode, 'xmlns="'||p_xml.getnamespace()||'"') 
    into p_xml from dual;  
end;  
-- p_target like 'worksheets/sheet5.xml' or 'sharedStrings.xml' rId return  
Function add_rels_type(p_target varchar2, p_type varchar2, p_ctype varchar2) return number  
is  
  t_xml xmltype;  
  t_rid number;  
  t_xname varchar2(200); 
  t_str xmltype; 
begin  
  t_xname := 'xl/_rels/workbook.xml.rels';  
  t_xml := get1xml( context.template, t_xname ); 
  select substr(extract(t_xml 
      ,'/Relationships/Relationship[@Target="'||p_target||'"]/@Id' 
      ,c_relsns).getStringVal(), 4) 
    into t_rid 
    from dual; 
  if t_rid is not null then return t_rid; end if; 
--   
  select max(to_number(substr(id,4))) + 1 into t_rid  
    from XMLTable(  
        xmlnamespaces(default 'http://schemas.openxmlformats.org/package/2006/relationships'),      
         '/Relationships/Relationship' passing t_xml columns  
         id varchar2(8) path './@Id'); 
--          
  select appendChildXML(t_xml  
       , '/Relationships'  
       , xmltype('<Relationship Id="rId'||t_rid||'" '||  
         'Type="'||p_type||'" '||  
         'Target="'||p_target||'"/>')  
       , c_relsns)  
     into t_xml  
     from dual ;  
--debug(t_xml.getClobVal(1,1));  
  replace1xml(context.template, t_xname, t_xml);  
--   
  t_xname:='[Content_Types].xml';  
  t_xml:=get1xml( context.template, t_xname );  
  select appendChildXML(t_xml  
       , '/Types'  
       , xmlType('<Override ContentType="'||p_ctype||'" PartName="/xl/'||p_target||'"/>')  
       , 'xmlns="'||t_xml.getnamespace()||'"')  
     into t_xml  
     from dual ;  
--debug(t_xml.getClobVal(1,1));  
  replace1xml(context.template, t_xname, t_xml);  
  return t_rid;  
end;  
--  sheet names unquoted
Procedure finish_sheet(p_sheet_name varchar2, p_new_sheet_name varchar2:='')  
as  
  t_row_of pls_integer;  
  t_clob CLOB;  
  t_xml XMLType;  
  t_merges CLOB;  
  t_merges_cnt pls_integer := 0;  
  t_row_style tp_one_row_style;  
  t_row_style_null tp_one_row_style;  
  s pls_integer;  
  r pls_integer;  
  ro pls_integer;  
  c pls_integer;  
  t_fld tp_one_field;  
begin 
  s := context.ranges(upper(p_sheet_name)); 
  if not context.fields.exists(s) then return; end if; -- sheet not filled 
  dbms_lob.createtemporary(t_clob, false);  
  dbms_lob.createtemporary(t_merges, false);  
  loadSheetStyles(context.template, s);  
  t_row_of := 0;  
  t_clob := '<sheetData '||c_sheetns||'>';  
  t_merges_cnt := 0;  
  t_merges := '<mergeCells '||c_sheetns||' count="xxx">'; --init CLOB for append 
  r := context.fields(s).first();  
  while r is not null          
  loop  
    begin  
      t_row_style:=context.styles.rows(r);  
    exception when no_data_found then t_row_style := t_row_style_null;  
    end;  
    ro := context.fields(s)(r).first();  
    while ro is not null  
    loop  
      dbms_lob.append(t_clob  
        ,'<row r="'||to_number(t_row_of+r+ro)||'" spans="1:1024" '  
          ||case  
            when t_row_style.height>0  
            then ' customHeight="1" ht="'  
              ||to_char(t_row_style.height ,'TM9', 'NLS_NUMERIC_CHARACTERS=.,' )||'" '  
            else '' end  
--          ||case  
--            when t_row_style.style is not null  
--            then ' s="' || t_row_style.style ||'" '  
--            else '' end  
          ||'>'  
      );  
      c := context.fields(s)(r)(ro).first();  
      while c is not null  
      loop  
        if c > 0 then  
          t_fld := context.fields(s)(r)(ro)(c);  
--debug((r+ro+t_row_of)||' '||ro||' '||c||' '||t_fld.value);  
          dbms_lob.append(t_clob  
             , '<c r="'  
             ||cell2alfan(t_row_of+r+ro,c)||'" '||get_cell_style(r, c, t_fld)  
             ||case  
                 when t_fld.value is null then '/>'  
                 else '><v>'  
                   ||to_char(t_fld.value, 'TM9', 'NLS_NUMERIC_CHARACTERS=.,' )  
                   ||'</v></c>'  
               end  
          );  
          if merge_exists(s,r,c)  
          then  
            if ro = 0 then  
              null;  
--              if field_exists(s, r, 1, c) then  
-- if insert mode, reset vertical merge  
--                context.merges( s )( r )( c ).row_off:=0;  
--              end if;  
            else  
              for i in 1..context.merges(s)(r)(c).col_off loop  
                dbms_lob.append( t_clob  
                  , '<c r="'||cell2alfan(t_row_of+r+ro,c+i)  
                  ||'" s="'||context.styles.cells(r)(c+i).style||'"/>'  
                );  
              end loop;  
            end if;  
            dbms_lob.append(t_merges  
               ,'<mergeCell ref="'||cell2alfan(t_row_of+r+ro,c)||':'  
                ||cell2alfan( t_row_of + r + ro + context.merges(s)(r)(c).row_off  
                   , c+context.merges(s)(r)(c).col_off)  
                ||'"/>');  
            t_merges_cnt := t_merges_cnt+1;  
          end if;  
        end if;  
        c := context.fields(s)(r)(ro).next(c);  
      end loop;  
      dbms_lob.append(t_clob, '</row>');  
      ro := context.fields(s)(r).next(ro);  
    end loop; -- rows  
    t_row_of := t_row_of+context.fields(s)(r).count()-1;  
    r := context.fields(s).next(r);  
  end loop;  
  dbms_lob.append(t_clob,'</sheetData>');  
  dbms_lob.append(t_merges,'</mergeCells>');  
  t_merges:=replace(t_merges,' count="xxx">',' count="'||t_merges_cnt||'">');  
  t_xml:=get1xml(context.template, get_sheet_path(p_sheet_name));  
  replace_node(t_xml, '/worksheet/sheetData', t_clob);  
  replace_node(t_xml, '/worksheet/mergeCells', t_merges); 
  if p_new_sheet_name is not null then 
    select deleteXML(t_xml, '/worksheet/drawing', c_sheetns) into t_xml from dual; 
  end if; 
  replace1xml(context.template  
    , get_sheet_path(nvl(p_new_sheet_name, p_sheet_name)), t_xml);  
  dbms_lob.freetemporary(t_clob);  
  dbms_lob.freetemporary(t_merges);  
  clear_styles;  
end;  
--  
Procedure save_shared_strings  
as  
  t_xname varchar2(200):='xl/sharedStrings.xml';  
  t_xml xmltype:=null;  
  t_clob CLOB;  
  s pls_integer;  
begin  
  if context.str_ind.count() = 0 then return; end if;  
  dbms_lob.createtemporary(t_clob, false);  
  if context.str_cnt_of > 0 then  
    t_xml := get1xml(context.template, t_xname);  
    dbms_lob.append(t_clob,replace(t_xml.getclobval(4,0),'</sst>',''));  
  else  
    s := add_rels_type('sharedStrings.xml'  
      ,'http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings'  
      ,'application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml'); 
    s := context.str_ind.count(); 
    dbms_lob.append(t_clob,  
'<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="'||s||'" uniqueCount="'||s||'">'); 
  end if;  
  s := context.str_ind.first();  
  while s is not null  
  loop  
    dbms_lob.append(t_clob,'<si><t>'||dbms_xmlgen.convert(context.str_ind(s))||'</t></si>');  
    s := context.str_ind.next(s);  
  end loop;    
  dbms_lob.append(t_clob,'</sst>');  
  replace1xml(context.template, t_xname, xmlType(t_clob));  
--   
  dbms_lob.freetemporary(t_clob);  
end;  
--  
Procedure finish_xlsx  
as  
  s pls_integer;  
begin  
  s := context.fields.first();  
  while s is not null loop 
    finish_sheet(context.ran_ind(s).range_name);  
    s:=context.fields.next(s);  
  end loop; -- sheets  
  save_shared_strings;  
end;  
--  
Procedure in_sheet  
( p_sheet_name varchar2  
, p_newsheet_name varchar2  
, p_options varchar2 := ''  
)  
as
  t_xname varchar2(200);
  t_xml xmltype;  
  t_nid number; -- new sheetId 
  t_rid number; -- new rId 
  t_pos number; -- new sheet position 
  t_sht_nr number; -- source sheet index 
  t_str varchar2(500 char); 
  t_sheet_name varchar2(100) := replace(p_sheet_name,'''');
  t_newsheet_name varchar2(100) := replace(p_newsheet_name,'''');
begin  
--debug(getSheetXMLTarget(p_sheet_name)); 
  begin 
    t_sht_nr := context.ranges(upper(t_sheet_name)); -- source sheet or named range exists? 
    t_pos := context.fields(t_sht_nr).first;  -- source sheet is filled? 
  exception 
    when others then raise_application_error(c_errcode, '#SHEET! not filled sheet: '||p_sheet_name); 
  end; 
-- destination sheet or named range exists? 
  if context.ranges.exists(upper(t_newsheet_name))  
     or get_sheet_path(t_newsheet_name) is not null then 
    raise_application_error(c_errcode, '#SHEET! name exists: '||p_newsheet_name); 
  end if;
  if length(t_newsheet_name) > 31
    or regexp_instr(t_newsheet_name,'\/|\\|\*|\[|\]|\:|\?|''') > 0 
  then
    raise_application_error(c_errcode, '#SHEET! bad name: '||p_newsheet_name); 
  end if;
  t_xname := 'xl/workbook.xml'; 
  t_xml := get1xml( context.template, t_xname );  
-- get new sheetId 
  select max(sheetid) + 1 into t_nid from 
    ( select rownum rn, extractValue(value(s),'/sheet/@sheetId',c_sheetns) sheetid 
        from table(xmlsequence(extract(t_xml,'/workbook/sheets/sheet',c_sheetns))) s 
    ); 
-- get new rId  
  t_rid := add_rels_type('worksheets/sheet'||t_nid||'.xml'  
      ,'http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet'  
      ,'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml');  
-- insert new sheet node 
  t_str := '<sheet name="'||replace(t_newsheet_name,'"','&quot;')  
           ||'" sheetId="'||t_nid||'" state="visible" r:id="rId'||t_rid||'" '||c_rns||'/>';  
  if instr(p_options,'b') > 0  
  then  
    select insertXMLbefore(t_xml  
        , '/workbook/sheets/sheet[@name="'||t_sheet_name||'"]'  
        , xmltype(t_str)  
        , c_sheetns||' '||c_rns)  
      into t_xml  
      from dual ;  
  else 
    select insertXMLafter(t_xml  
        , '/workbook/sheets/sheet[@name="'||t_sheet_name||'"]'  
        , xmltype(t_str)  
        , c_sheetns||' '||c_rns)  
      into t_xml  
      from dual ;  
  end if;        
-- hide source sheet  
  if instr(p_options,'h') > 0 then  
    begin 
      select insertChildXML(t_xml  
        , '/workbook/sheets/sheet[@name="'||t_sheet_name||'"]' 
        , '@state'  
        , 'hidden'  
        , c_sheetns )  
      into t_xml from dual;  
    exception -- assume attribute @state exists 
      when others then 
        select updateXML(t_xml  
          , '/workbook/sheets/sheet[@name="'||t_sheet_name||'"]/@state' 
          , 'hidden'  
          , c_sheetns )  
      into t_xml from dual;  
    end; 
  end if; 
-- make new sheet active   
  select rn into t_pos from 
    ( select rownum rn, extractValue(value(s),'/sheet/@name',c_sheetns) name 
        from table(xmlsequence(extract(t_xml,'/workbook/sheets/sheet',c_sheetns))) s 
    ) where name = t_newsheet_name; 
--debug(t_pos); 
  select updateXML(t_xml  
    , '/workbook/bookViews/workbookView/@activeTab'  
    , t_pos - 1 
    , c_sheetns )  
    into t_xml from dual; 
--debug(t_xml.getClobVal(1,1));  
  replace1xml(context.template, t_xname, t_xml);  
-- save new sheet data and clear 
  finish_sheet(t_sheet_name, t_newsheet_name);  
  clear_sheet_fields(t_sht_nr);  
end;      
--  
Procedure finish (p_xfile in out nocopy BLOB)  
as  
begin  
  finish_xlsx;  
  p_xfile := context.template;    
  init;       
end;  
-- 
Function address   
( p_row pls_integer          -- > 0     
, p_col pls_integer          -- > 0
, p_address varchar2 
, p_options varchar2 := ''   --  o l m 1
) return varchar2
is
  t_mloc tp_loc;
  t_loc tp_loc;
begin
  if p_address = '!' then
    return cell2alfan(null,null,context.ran_ind(context.current_sht).range_name);
  end if;
  t_loc := addr2loc(p_address,false);
  if t_loc.sht_nr is null or not(p_col > 0 and p_row > 0) then return null; end if;
  t_loc.row_nr := t_loc.row_nr + p_row - 1;
  t_loc.col_nr := t_loc.col_nr + p_col - 1;
  t_mloc := get_merge(t_loc);
  if t_mloc.col_nr != t_loc.col_nr and instr(p_options,'o') > 0 then -- align outside merge
    t_loc.col_nr := t_mloc.col_nr + t_mloc.col_off + 1;
  end if;
  if instr(p_options,'l') > 0 then -- align to left cell
    t_loc.col_nr := t_mloc.col_nr;
  end if;
  if instr(p_options,'t') > 0 then -- align to top cell
    t_loc.row_nr := t_mloc.row_nr;
  end if;
  return loc2addr(t_loc, p_options);
end;
--
Function new_workbook return BLOB  
/* 
  xlsx created by Google Docs. Convert to base64: 
  win10: certutil -f -encode <in_file.xlsx> <out_file.b64> 
  linux: openssl enc -base64 -in <in_file.xlsx> -out <out_file.b64> 
  Varchar2 literal MAX size = 4000 characters
*/ 
is 
  Function base642blob(p_base64 varchar2) return BLOB 
  is 
  begin 
--debug(length(p_base64)); 
    return utl_encode.base64_decode(utl_raw.cast_to_raw(
        replace(replace(p_base64, chr(10)),' '))); 
  end; 
begin  
  return base642blob( 
'UEsDBBQACAgIADQmfEsAAAAAAAAAAAAAAAAYAAAAeGwvZHJhd2luZ3MvZHJhd2lu
ZzEueG1sndBBTsMwEAXQE3CHaPatAwuEoqbdRJwADjDYk9jCY1szLm1vj0XJvsry
62uevuZwunLsfkg05DTC876HjpLNLqRlhM+P990bdFoxOYw50Qg3Ujgdnw5XJ8NF
J+nafdKhxRF8rWUwRq0nRt3nQqm1cxbG2qIsxglemszRvPT9q9EihE49UZ3uDfx7
uEFjDGm9f2hNnudgacr2zJTqHRGKWNsv1Ieiq2Y3rLEepa4APyQwyve57Gzm0jZ8
hRjq7Q9bGbfwhiUu4CLIYI6/UEsHCB+aYCjMAAAA7gEAAFBLAwQUAAgICAA0JnxL
AAAAAAAAAAAAAAAAGAAAAHhsL2RyYXdpbmdzL2RyYXdpbmcyLnhtbJ3QQU7DMBAF
0BNwh2j2rQMLhKKm3UScAA4w2JPYwmNbMy5tb49Fyb7K8utrnr7mcLpy7H5INOQ0
wvO+h46SzS6kZYTPj/fdG3RaMTmMOdEIN1I4HZ8OVyfDRSfp2n3SocURfK1lMEat
J0bd50KptXMWxtqiLMYJXprM0bz0/avRIoROPVGd7g38e7hBYwxpvX9oTZ7nYGnK
9syU6h0RiljbL9SHoqtmN6yxHqWuAD8kMMr3uexs5tI2fIUY6u0PWxm38IYlLuAi
yGCOv1BLBwgfmmAozAAAAO4BAABQSwMEFAAICAgANCZ8SwAAAAAAAAAAAAAAABgA
AAB4bC93b3Jrc2hlZXRzL3NoZWV0MS54bWyNks1ywiAQgJ+g78BwN8RW25pJ4qGO
U2+dTn/OCBvDyE8GiIlvXxI1UyeX3JZl+fhYNl23SqITWCeMzvA8ijECzQwX+pDh
76/t7BUj56nmVBoNGT6Dw+v8IW2MPboSwKMA0C7DpfdVQohjJSjqIlOBDjuFsYr6
sLQH4ioLlPeHlCSPcfxMFBUaXwiJncIwRSEYbAyrFWh/gViQ1Ad9V4rK3WiqHeGU
YNY4U/iIGXUlBQNGoGXQC73eCSk2xUhRe6yrWUBWwWIvpPDn3mvAnDJcW51cGbNB
ozuThPuTk5K34na+mOY9auaKrO7sA4mOHzCdRdlAUtMwQxuv/5qnPfJHQOP+xagb
nb0xx26x4xmOMclTMqrd9k3+sIjVzhv1DuJQ+jCiGHEoaC39m5G/gvsy5BbR4mnI
f5pmKF5GL8sO3xM31NMQc0ubMN7IJiLcbnd83gsME53/AVBLBwiqiI/PWAEAABUD
AABQSwMEFAAICAgANCZ8SwAAAAAAAAAAAAAAACMAAAB4bC93b3Jrc2hlZXRzL19y
ZWxzL3NoZWV0MS54bWwucmVsc43PSwrCMBAG4BN4hzB7k9aFiDTtRoRupR5gSKYP
bB4k8dHbm42i4MLlzM98w181DzOzG4U4OSuh5AUwssrpyQ4Szt1xvQMWE1qNs7Mk
YaEITb2qTjRjyjdxnHxkGbFRwpiS3wsR1UgGI3eebE56FwymPIZBeFQXHEhsimIr
wqcB9ZfJWi0htLoE1i2e/rFd30+KDk5dDdn044XQAe+5WCYxDJQkcP7avcOSZxZE
XYmvivUTUEsHCK2o602zAAAAKgEAAFBLAwQUAAgICAA0JnxLAAAAAAAAAAAAAAAA
GAAAAHhsL3dvcmtzaGVldHMvc2hlZXQyLnhtbI2TS26DMBBAT9A7WN4Hkyb9IaCq
GlXtrqr6WbtmCFawB9kG0tt3IAlqlQ0LSx4z8/wY2+n93tSsA+c12owvo5gzsAoL
bbcZ/3h/Wtxy5oO0hazRQsZ/wPP7/CLt0e18BRAYAazPeBVCkwjhVQVG+ggbsPSl
RGdkoNBthW8cyGIsMrW4jONrYaS2/EBI3BwGlqVWsEHVGrDhAHFQy0D6vtKNP9HM
/gxntHLosQyRQnMkkYESsFcwCt3+EzJqjpGRbtc2C0I2ZPGtax1+Rq8J02W8dTY5
MhaTxlCT0P5JZ+pT8n65nud91sw7cffPnkjy/Afms6SaSGYeZmrj8VzzdER+auj9
nzkbrs434m4IXoqMx1zkqTjLfRqb/OqYan1A8wx6WwW6opwVUMq2Do9Yf+kiVLS2
jtaraf0N+yn5Krq5GvAjcSODzFOHPXMDJ0/VMHkgoj/EXb6+vFldR3EqOlJSNCj7
JHcoL5zs6XEwl2hydy/FctSf3kP+C1BLBwgiKqtMegEAAFMDAABQSwMEFAAICAgA
NCZ8SwAAAAAAAAAAAAAAACMAAAB4bC93b3Jrc2hlZXRzL19yZWxzL3NoZWV0Mi54
bWwucmVsc43PSwrCMBAG4BN4hzB7k7YLEWnajQjdSj3AkEwf2CYhiY/e3mwUCy5c
zvzMN/xl/ZwndicfRmsk5DwDRkZZPZpewqU9bffAQkSjcbKGJCwUoK425ZkmjOkm
DKMLLCEmSBhidAchghpoxsCtI5OSzvoZYxp9LxyqK/YkiizbCf9tQLUyWaMl+Ebn
wNrF0T+27bpR0dGq20wm/nghtMdHKpZI9D1FCZy/d5+w4IkFUZViVbF6AVBLBwiF
AfUVtAAAACoBAABQSwMEFAAICAgANCZ8SwAAAAAAAAAAAAAAABQAAAB4bC9zaGFy
ZWRTdHJpbmdzLnhtbA3LQQ7CIBBA0RN4BzJ7C7owxpR21xPoASZlLCQwEGZi9Pay
/Hn58/ot2XyoS6rs4TI5MMR7DYkPD6/ndr6DEUUOmCuThx8JrMtpFlEzVhYPUbU9
rJU9UkGZaiMe8q69oI7sh5XWCYNEIi3ZXp272YKJwS5/UEsHCK+9gnR0AAAAgAAA
AFBLAwQUAAgICAA0JnxLAAAAAAAAAAAAAAAADQAAAHhsL3N0eWxlcy54bWydVMtu
nDAU/YL+g+VFdzNmRlHUJEAUVaLqJl1kKnVrjBms+EFtk0K/vteYCdBJlVHZ2Pfh
c47vvSa975VEL9w6YXSGd9sEI66ZqYQ+Zvj7odh8wsh5qisqjeYZHrjD9/mH1PlB
8qeGc48AQbsMN963t4Q41nBF3da0XEOkNlZRD6Y9EtdaTisXDilJ9klyTRQVGkeE
2353RdkZjhLMGmdqv2VGEVPXgvFzpBtyQyg7IalzmDfkKGqfu3YDsC31ohRS+GFU
hfNUd6pQ3iFmOu2hLq8uFJevFTivrzCKgJ9NFWoD38efnfF3m7gotTKrCpM8JRN2
'||'ntZGzxR7HB156n6jFyoBP4FuwAFmpLHIHssMF0UyfsGtqeIx8cEKKkfoCPA2zL/z
xyXoEVKu9YAjT6E6nltdgIGm/WFo4b4aJiLCjHnvZEtxbPwXS4fFkXEB5tLYCmZw
We7oCqlTEArBpXwKc/ejXqX2NYo5oSswwAH0tIWbTdu5cWDQtpXDA0jSikeY6CpM
tALvki6SL3j3/8fb1xcKyFN6CqIw6/AevwWq8bBrrNDPB1MIP9rwfr1gobWl8d4o
jH5Z2h54P4bDXfr6Irk7/Pd8Xyw4Wo+dKrktxkfx3jVOushU2kWDV+199c5iwixn
+DHQSIzKTkgvdIytOgeYVT83LUbnX1f+B1BLBwhI5uS4/QEAAP8EAABQSwMEFAAI
CAgANCZ8SwAAAAAAAAAAAAAAAA8AAAB4bC93b3JrYm9vay54bWyNkt1ugjAYhq9g
90B6rlVnFkcET5YlnixLtl1AbT+ksT+kX2V49/tAIDpPOGpLeR+elne7a6xJagio
vcvYcr5gCTjplXbHjP18v882LMEonBLGO8jYBZDt8qftrw+ng/enhPIOM1bGWKWc
oyzBCpz7ChztFD5YEWkZjhyrAEJhCRCt4avF4oVboR27EtIwheGLQkt48/JswcUr
JIARkeyx1BUONNs84KyWwaMv4lx625PIQHJoJHRCmzshK6cYWRFO52pGyIosDtro
eOm8RkydsXNwac+YjRptJqXvp7U1w8vNcj3N++EyX/nrnT2RxOMBprOEHEl2Gma8
xv6/5mNHPgPPtx0f+7GtVKQ21Rr1wQBLnLC0/Gr3llS4dtwr6iNLQqppEvbqmfF/
6VIrBe4uvLgJr27C6zbMBwcFhXagPiiH9FwKIztHPhjnf1BLBwgotAkzTAEAABgD
AABQSwMEFAAICAgANCZ8SwAAAAAAAAAAAAAAABoAAAB4bC9fcmVscy93b3JrYm9v
ay54bWwucmVsc72SwWrDMAyGn2DvYHRfnKRjjFGnlzHotcsewNhKHJrYxtLa5e3n
MbalUMoOZSchGX3/h/B68z6N4oCJhuAVVEUJAr0JdvC9gtf2+fYBBLH2Vo/Bo4IZ
CTbNzXqHo+a8Q26IJDLEkwLHHB+lJONw0lSEiD6/dCFNmnObehm12eseZV2W9zIt
GdCcMMXWKkhbW4Fo54h/YYeuGww+BfM2oeczEZJ4HrO/aHXqkRV89UXmgDwfX181
3umE9oVTPu7SYjm+JLO6pswxpD05RP4V+Rl9quZSXZK5+2eZ+ltGnny95gNQSwcI
nyKs4OIAAADCAgAAUEsDBBQACAgIADQmfEsAAAAAAAAAAAAAAAALAAAAX3JlbHMv
LnJlbHONz0EOgjAQBdATeIdm9lJwYYyhsDEmbA0eoLZDIUCnaavC7e1SjQuXk/nz
fqasl3liD/RhICugyHJgaBXpwRoB1/a8PQALUVotJ7IoYMUAdbUpLzjJmG5CP7jA
EmKDgD5Gd+Q8qB5nGTJyaNOmIz/LmEZvuJNqlAb5Ls/33L8bUH2YrNECfKMLYO3q
8B+bum5QeCJ1n9HGHxVfiSRLbzAKWCb+JD/eiMYsocCrkn88WL0AUEsHCKRvoSCy
AAAAKAEAAFBLAwQUAAgICAA0JnxLAAAAAAAAAAAAAAAAEwAAAFtDb250ZW50X1R5
cGVzXS54bWzFVNtqAjEQ/YL+w5LXYqI+lFJcfejlsS3UfsA0mXWDuZGJuvv3za5a
qFiooPiUmZyZc04mIZNZY02xxkjau5KN+JAV6KRX2i1K9jl/GdyzghI4BcY7LFmL
xGbTm8m8DUhFbnZUsjql8CAEyRotEPcBXUYqHy2knMaFCCCXsEAxHg7vhPQuoUuD
1HGw6eQJK1iZVDxu9zvqkkEIRktI2ZfIZKx4bjK4tdnl4h99a6cOzAx2RnhE09dQ
rQPdHgpklDqFtzyZqBWeJOGrSktUXq5sbuEUIoKiGjFZwzc+Lvt4q/kOMb2CzaSi
MeIHJNEvI7476ZV9jC/ng2qIqD5SzA+Ojnn5VXBOHyrCJnMe09xBtA/Oeg8n6F5y
7qk1eHzgPXLJG88rt6DdX0/vy/vlXl/0H830G1BLBwhzLo65NAEAAKgEAABQSwEC
FAAUAAgICAA0JnxLH5pgKMwAAADuAQAAGAAAAAAAAAAAAAAAAAAAAAAAeGwvZHJh
d2luZ3MvZHJhd2luZzEueG1sUEsBAhQAFAAICAgANCZ8Sx+aYCjMAAAA7gEAABgA
AAAAAAAAAAAAAAAAEgEAAHhsL2RyYXdpbmdzL2RyYXdpbmcyLnhtbFBLAQIUABQA
CAgIADQmfEuqiI/PWAEAABUDAAAYAAAAAAAAAAAAAAAAACQCAAB4bC93b3Jrc2hl
ZXRzL3NoZWV0MS54bWxQSwECFAAUAAgICAA0JnxLrajrTbMAAAAqAQAAIwAAAAAA
AAAAAAAAAADCAwAAeGwvd29ya3NoZWV0cy9fcmVscy9zaGVldDEueG1sLnJlbHNQ
SwECFAAUAAgICAA0JnxLIiqrTHoBAABTAwAAGAAAAAAAAAAAAAAAAADGBAAAeGwv
d29ya3NoZWV0cy9zaGVldDIueG1sUEsBAhQAFAAICAgANCZ8S4UB9RW0AAAAKgEA
ACMAAAAAAAAAAAAAAAAAhgYAAHhsL3dvcmtzaGVldHMvX3JlbHMvc2hlZXQyLnht
bC5yZWxzUEsBAhQAFAAICAgANCZ8S6+9gnR0AAAAgAAAABQAAAAAAAAAAAAAAAAA
iwcAAHhsL3NoYXJlZFN0cmluZ3MueG1sUEsBAhQAFAAICAgANCZ8S0jm5Lj9AQAA
/wQAAA0AAAAAAAAAAAAAAAAAQQgAAHhsL3N0eWxlcy54bWxQSwECFAAUAAgICAA0
JnxLKLQJM0wBAAAYAwAADwAAAAAAAAAAAAAAAAB5CgAAeGwvd29ya2Jvb2sueG1s
UEsBAhQAFAAICAgANCZ8S58irODiAAAAwgIAABoAAAAAAAAAAAAAAAAAAgwAAHhs
L19yZWxzL3dvcmtib29rLnhtbC5yZWxzUEsBAhQAFAAICAgANCZ8S6RvoSCyAAAA
KAEAAAsAAAAAAAAAAAAAAAAALA0AAF9yZWxzLy5yZWxzUEsBAhQAFAAICAgANCZ8
S3Mujrk0AQAAqAQAABMAAAAAAAAAAAAAAAAAFw4AAFtDb250ZW50X1R5cGVzXS54
bWxQSwUGAAAAAAwADAA2AwAAjA8AAAAA' 
  ); 
end; 
   
end "AM_FILL";  
/