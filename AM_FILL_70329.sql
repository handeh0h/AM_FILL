CREATE OR REPLACE PACKAGE  "AM_FILL" as
/*
*    Purpose: using the Microsoft XLSX files as a templates for data output
*
*    Uses AS_ZIP package designed by Anton Scheffer 
*    Download from: http://technology.amis.nl/wp-content/uploads/2010/06/as_zip7.txt
*
*    Based on the code of packages designed by Anton Scheffer
*        https://technology.amis.nl/wp-content/uploads/2011/02/as_xlsx11.txt
*        https://technology.amis.nl/wp-content/uploads/2013/01/as_read_xlsx9.txt 
*
*    Author: miktim@mail.ru, Petrozavodsk State University
*    Date: 2013-06-24
*    Updated:
*      2015-02-27 support Oracle non UTF-8 charSets.
*      2015-03-18 procedure in_sheet added. Thanks to odie_63 (ORACLE Community)
*                     (https://community.oracle.com/message/12933878)
*      2016-02-01 free xmlDocument objects
*      2016-06-07 fixed bugs: get sheet xml name, sheets without merges (align_loc)
*                 Thanks github.com/Zulus88
*      2016-06-08 fixed bug: calc new rId (IN_SHEET)
*      2017-03-28 fixed bug: 200 rows 'limitation' (IN_TABLE),
*                 options added (IN_SHEET),
*                 build-in workbook changed
*    
******************************************************************************
* Copyright (C) 2011 - 2013 by Anton Scheffer
*               2013 - 2017 by MikTim
* Lisence: MIT
******************************************************************************
*/
version constant varchar2(16):='2.70329';
/* 
   INIT: Initialize package by xlsx template
   p_options:
      e - enable exception on #REF!, otherwise ignore filling
*/      
Procedure init
( p_xtemplate BLOB
, p_options varchar2:=''  
);
/* INIT: Clear internal structures */
Procedure init;
/* 
   IN_FIELD: Fill in cell or upper left cell of named area
   p_cell_addr: A1 style cell address (sheet_name!cell_address) or range name
   p_options:
     i - row insert mode (sequentially on every call), default - overwrite
         WARNING: insert mode cut vertical merges
*/
Procedure in_field
( p_value date
, p_cell_addr varchar2
, p_options varchar2:='');
Procedure in_field
( p_value number
, p_cell_addr varchar2
, p_options varchar2:='');
Procedure in_field
( p_value varchar2
, p_cell_addr varchar2
, p_options varchar2:='');
/* 
   IN_TABLE: Fill in table
   p_sql: sql query text (without trailing semicolon) 
   p_cell_addr: default A1 of current sheet
   p_options:
     h - print headings (field names)
     i - row insert mode.
         WARNING: insert mode cut vertical merges (one record - one row)
*/
Procedure in_table
( p_sql CLOB
, p_cell_addr varchar2 := ''
, p_options varchar2 := '');
/*
   IN_SHEET: Save filled sheet with new name after source sheet,
             clears data from source sheet, a new sheet becomes active & visible.
         WARNING: the new sheet name does not check (length, allowed chars)
    
   p_options:
     h - hide source sheet
     b - insert BEFORE source sheet
*/
Procedure in_sheet
( p_sheet_name varchar2      -- source sheet name
, p_newsheet_name varchar2   -- new sheet name 
, p_options varchar2:=''     -- options
);
/*
   FINISH: Generate workbook, clear internal structures
         WARNING: all formulas from filled sheets will be removed
*/
Procedure finish
( p_xfile in out nocopy BLOB -- filled in xlsx returns
);
/*
   ADDRESS: Calculate relative (sheet or named range) address
            
*/
Function address
( p_row pls_integer
, p_col pls_integer
, p_range_name varchar2 := '' -- if omitted, the current sheet is used
) return varchar2;            -- A1 style address
/*
   NEW_WORKBOOK: empty workbook returned 
   Workbook has only sheet 'Sheet1'. A1 cell contains 'YYYY-MM-DD' formatted date. 
*/
Function new_workbook return BLOB;
/*
******************************************************************************
* Package usage:
*
* begin
* ...
* -- specify xlsx BLOB as template for filling
*    am_fill.init(<some_xlsx_blob>); 
* ...
* -- insert rows of query into Sheet1 from row 12 column B use styles of cells B12, C12, D12, ...
*    am_fill.in_table('select <some_fields> from <some_table>','Sheet1!B12','i');  
* ...
* -- fill named range cell with 2 cols & 10 rows offset using destination cell style
*    am_fill.in_field(<some_variable>, 'range1!C11');
* ...
* -- manually insert rows using 'Sheet1!B12' cell style 
* -- in this case, result of previous am_fill.x_table first column will be overprinted
*    for c in (select  <some_field> from <some_table>)
*    loop
*        am_fill.in_field(c.<some_field>, 'Sheet1!B12','i');
*    end loop;
* ...
* -- generate and return filled xlsx BLOB
*    am_fill.finish(<some_filled_blob>);   
* end;
*/
end;
/
CREATE OR REPLACE PACKAGE BODY  "AM_FILL" is
 
c_ssurl constant varchar2(200) := 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
c_ssns  constant varchar2(200) := 'xmlns="'||c_ssurl||'"';
c_rsurl constant varchar2(200) := 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
c_rsns  constant varchar2(200) := 'xmlns:r="'||c_rsurl||'"';
 
type tp_one_range is record
  ( range_name varchar2(200)  -- sheet or range name
  , isSheet boolean := false
  , sht_nr  pls_integer := 1  -- sheet number
  , row_nr  pls_integer := 1  -- upper left row
  , col_nr  pls_integer := 1  -- upper left col
  );
type tp_all_ranges is table of tp_one_range index by pls_integer;
--
type tp_loc is record
  ( sht_nr pls_integer
  , row_nr pls_integer := null
  , row_off pls_integer := 0  --  v_size-1 for merges, inserted row index
  , col_nr pls_integer := null
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
  , insmode boolean := false       -- insert, overwrite
  , date1904 boolean := true       -- workbook date format
  , date_style pls_integer := null -- default date style
  , current_sht pls_integer := 1
  , strings tp_strings
  , str_ind tp_str_ind
  , str_cnt_of pls_integer
  , ranges tp_strings
  , ran_ind tp_all_ranges
  , merges tp_merges
  , fields tp_fsheets
  , styles tp_sheet
  );
context tp_context;
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
    if (t_style.style is null and t_style.type = 'D') then
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
--***
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
  ( p_msfile in out nocopy blob
  , p_filename varchar2
  , p_xml xmlType
  )
is
  t_blob BLOB;
  t_xml xmltype;
  c_xsl constant xmltype := xmltype(  
'<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform">
   <xsl:output  indent="no" omit-xml-declaration="no" standalone="yes"/>
   <xsl:template match="@*|node()">  
      <xsl:copy>  
        <xsl:apply-templates select="@*|node()"/>  
     </xsl:copy>  
   </xsl:template>  
</xsl:stylesheet>');  
begin
  select xmltransform(p_xml, c_xsl) into t_xml from dual;
  t_blob := p_xml.getBlobVal(nls_charset_id('AL32UTF8'),4,0);
  replace1file( p_msfile, p_filename, t_blob );
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
function alfan2col( p_col pls_integer )
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
function alfan2cell(p_row pls_integer, p_col pls_integer, p_range_name varchar2:='') return varchar2
is
begin
   return case when p_range_name is not null then p_range_name||'!' else '' end
     ||alfan2col(p_col)||p_row;
end;
--  
function col2alfan( p_col varchar2 )
return pls_integer
is
begin
  return ascii( substr( p_col, -1 ) ) - 64
       + nvl( ( ascii( substr( p_col, -2, 1 ) ) - 64 ) * 26, 0 )
       + nvl( ( ascii( substr( p_col, -3, 1 ) ) - 64 ) * 676, 0 );
end;
-- 
function addr2loc(p_range_name varchar2) return tp_loc
is
  t_loc tp_loc;
  t_loc1 tp_loc;
  t_loc2 tp_loc;
  t_ind pls_integer;
  t_tmp1 varchar2(100);
  t_tmp2 varchar2(100);
  t_tmp3 varchar2(100);
-- !!!???
  t_rowcolp varchar2(100) := '\$?([[:alpha:]]{1,3})\$?([[:digit:]]*)';
  t_rcellp varchar2(100) := '(\$?[[:alpha:]]{1,3}\$?[[:digit:]]*)(:(\$?[[:alpha:]]{1,3}\$?[[:digit:]]*))?';
  t_rangep varchar2(200) := '^((\$?[[:alpha:]]{1,3}\$?[[:digit:]]*)(:\$?[[:alpha:]]{1,3}\$?[[:digit:]]*)?)$'||
--      '|^(('')?([[:alnum:]_ ]+)\5?){1}(!((\$?[[:alpha:]]{1,3}\$?[[:digit:]]*)(:\$?[[:alpha:]]{1,3}\$?[[:digit:]]*)?))?$';
      '|^(('')?([\w ]+|[[:alnum:]_ ]+)\5?){1}(!((\$?[[:alpha:]]{1,3}\$?[[:digit:]]*)(:\$?[[:alpha:]]{1,3}\$?[[:digit:]]*)?))?$';
begin
--dbms_output.put_line(p_range_name);
  if regexp_instr(nvl(p_range_name,'A1'),t_rangep)>0 then
    t_tmp1 := regexp_replace(p_range_name, t_rangep, '\6');
    t_tmp2 := regexp_replace(p_range_name, t_rangep, '\1\8');
    t_tmp3 := regexp_replace(t_tmp2, t_rcellp, '\3');
    t_tmp2 := regexp_replace(t_tmp2, t_rcellp, '\1');
    t_tmp3 := nvl(t_tmp3,t_tmp2);
-- sheet or range
    if t_tmp1 is not null then
      if context.ranges.exists( upper(t_tmp1) ) then
        t_ind := context.ranges( upper(t_tmp1 ));
        t_loc.sht_nr := context.ran_ind(t_ind).sht_nr;
        t_loc.row_nr := context.ran_ind(t_ind).row_nr;
        t_loc.col_nr := context.ran_ind(t_ind).col_nr;
      end if;
    else
      t_loc.sht_nr := context.current_sht;
      t_loc.row_nr := 1;
      t_loc.col_nr := 1;
    end if;
-- upper left cell
    if t_tmp2 is not null then
      t_loc1.col_nr := col2alfan(regexp_replace(t_tmp2,t_rowcolp,'\1'));
      t_loc1.row_nr := to_number(regexp_replace(t_tmp2,t_rowcolp,'\2'));
      t_loc2.col_nr := col2alfan(regexp_replace(t_tmp3,t_rowcolp,'\1'));
      t_loc2.row_nr := to_number(regexp_replace(t_tmp3,t_rowcolp,'\2'));
      t_loc.col_nr := t_loc.col_nr + least(t_loc1.col_nr, t_loc2.col_nr)-1;
      t_loc.row_nr := t_loc.row_nr + least(t_loc1.row_nr, t_loc2.row_nr)-1;
      t_loc.col_off := greatest(t_loc1.col_nr,t_loc2.col_nr) - least(t_loc1.col_nr,t_loc2.col_nr);
      t_loc.row_off := greatest(t_loc1.row_nr,t_loc2.row_nr) - least(t_loc1.row_nr,t_loc2.row_nr);
    end if;
  end if;
--dbms_output.put_line(p_range_name||'|'||t_loc.sht_nr||'|'||t_loc.row_nr||'|'||t_loc.col_nr||'|'||t_loc.col_off||'|'||t_loc.row_off);
  if t_loc.sht_nr is null or t_loc.row_nr is null or t_loc.col_nr is null 
  then
    if instr(context.options,'e') > 0
    then
      raise_application_error(-20001,'AM_FILL #REF!: '||p_range_name);
    else
      t_loc.sht_nr := null;
    end if;
  end if;
  return t_loc;
end;
--
Procedure add_range
  ( p_range_name varchar2
  , p_range_def varchar2
  , p_sht_nr pls_integer :=0
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
    if p_sht_nr = 0
    then
      t_loc := addr2loc(p_range_def);
--dbms_output.put_line(p_range_name||' '||p_range_def||' '||t_loc.sht_nr||' '||t_loc.row_nr||' '||t_loc.col_nr);
      context.ran_ind( t_cnt ).sht_nr := t_loc.sht_nr;
      context.ran_ind( t_cnt ).row_nr := t_loc.row_nr;
      context.ran_ind( t_cnt ).col_nr := t_loc.col_nr;
    else
      context.ran_ind( t_cnt ).sht_nr := p_sht_nr;
      context.ran_ind( t_cnt ).isSheet := true;
    end if;        
  end if;
end;
--
Function sheet_name
( p_shtnr pls_integer := context.current_sht
) return varchar2
is
  i pls_integer;
begin
  i := context.ran_ind.first();
  while i is not null and context.ran_ind(i).isSheet loop
    if context.ran_ind(i).sht_nr = p_shtnr then
      return context.ran_ind(i).range_name;
    end if;
    i := context.ran_ind.next(i);
  end loop;
  return '';
end;
--
Function merge_exists(s pls_integer, r pls_integer, c pls_integer) return boolean
is
begin
  return context.merges(s)(r).exists(c);
exception when no_data_found then return false;
end;
 -- Return merge as tp_loc
Function merge2loc
  ( p_sht pls_integer
  , p_row pls_integer
  , p_col pls_integer
  ) return tp_loc
is
  t_loc tp_loc;
begin
  t_loc.sht_nr := p_sht;
  t_loc.row_nr := p_row;
  t_loc.col_nr := p_col;
  t_loc.row_off := context.merges(p_sht)(p_row)(p_col).row_off;
  t_loc.col_off := context.merges(p_sht)(p_row)(p_col).col_off;
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
--dbms_output.put_line(alfan2cell(t_loc.row_nr,t_loc.col_nr)||':'||alfan2cell(t_loc.row_nr+t_loc.row_off,t_loc.col_nr+t_loc.col_off));
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
/*  t_dc := p_cloc.row_nr - p_mloc.row_nr ;
  t_merge.row_off := greatest(t_dc - 1, 0);
  context.merges(p_mloc.sht_nr)(p_mloc.row_nr)(p_mloc.col_nr) := t_merge;
-- copy_style
  if t_dc > 0 then
    t_merge.row_off := 0;
    context.merges(p_mloc.sht_nr)(p_cloc.row_nr)(p_mloc.col_nr) := t_merge;
-- copy style
  end if;
  if t_dc < p_mloc.row_off then
    t_merge.row_off := p_mloc.row_off - t_dc;
    context.merges(p_mloc.sht_nr)(p_cloc.row_nr + 1)(p_mloc.col_nr) := t_merge;
-- copy style
  end if;
*/
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
--dbms_output.put_line(t_mloc.row_off);
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
  t_mloc tp_loc;
begin
  t_mloc := get_merge( p_loc );
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
  context.fields( t_loc.sht_nr )( t_loc.row_nr )( t_loc.row_off )( t_loc.col_nr ) := t_field;
--dbms_output.put_line(t_loc.sht_nr||'|'||t_loc.row_nr||'|'||t_loc.row_off||'|'||t_loc.col_nr||'|'||t_loc.col_off);
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
Procedure in_field(p_value number, p_cell_addr varchar2, p_options varchar2:='')
is
  t_loc tp_loc;
begin
  context.insmode := instr(p_options,'i') > 0;
  t_loc := addr2loc(p_cell_addr);
  if t_loc.sht_nr is null then return; end if;
  t_loc := align_loc(t_loc);
  if context.insmode then set_cell_off(t_loc); end if;
  add_field( p_value, t_loc);
end;
--
Procedure in_field(p_value date, p_cell_addr varchar2, p_options varchar2:='')
is
  t_loc tp_loc;
begin
  context.insmode := instr(p_options,'i') > 0;
  t_loc := addr2loc(p_cell_addr);
  if t_loc.sht_nr is null then return; end if;
  t_loc := align_loc(t_loc);
  if context.insmode then set_cell_off(t_loc); end if;
  add_field( p_value, t_loc );
end;
--
Procedure in_field(p_value varchar2, p_cell_addr varchar2, p_options varchar2:='')
is
  t_loc tp_loc;
begin
  context.insmode := instr(p_options,'i') > 0;
  t_loc := addr2loc(p_cell_addr);
  if t_loc.sht_nr is null then return; end if;
  t_loc := align_loc(t_loc);
  if context.insmode then set_cell_off(t_loc); end if;
  add_field( p_value, t_loc );
end;
--
Procedure in_table(p_sql CLOB, p_cell_addr varchar2:='', p_options varchar2:='')
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
  t_loc := addr2loc(p_cell_addr);
  if t_loc.sht_nr is null then return; end if;
  t_loc := align_loc(t_loc, t_insert );
  t_c := dbms_sql.open_cursor;
  dbms_sql.parse( t_c, p_sql, dbms_sql.native );
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
  t_r := dbms_sql.execute( t_c );
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
/*
exception
  when others then
    if dbms_sql.is_open( t_c ) then
      dbms_sql.close_cursor( t_c );
    end if;
    raise;
*/
end;
--
function address
( p_row pls_integer
, p_col pls_integer
, p_range_name varchar2:=''
) return varchar2
is
  t_ran tp_one_range;
begin
  t_ran.range_name := p_range_name;
  t_ran.row_nr:=p_row;
  t_ran.col_nr:=p_col;
  if p_range_name = '!' then
    t_ran.range_name := sheet_name(context.current_sht);
  elsif context.ranges.exists(upper(p_range_name))then
    t_ran := context.ran_ind(context.ranges(upper(p_range_name)));
    if not t_ran.isSheet then
       t_ran.range_name := sheet_name(t_ran.sht_nr);
       for i in 1..nvl(p_col,0)-1 loop
         if merge_exists(t_ran.sht_nr,t_ran.row_nr,t_ran.col_nr) then
           t_ran.col_nr:=t_ran.col_nr
             + context.merges(t_ran.sht_nr)(t_ran.row_nr)(t_ran.col_nr).col_off;
         end if;
         t_ran.col_nr:=t_ran.col_nr+1;
       end loop;
       for i in 1..nvl(p_row,0)-1 loop
         if merge_exists(t_ran.sht_nr,t_ran.row_nr,t_ran.col_nr) then
           t_ran.row_nr:=t_ran.row_nr
             + context.merges(t_ran.sht_nr)(t_ran.row_nr)(t_ran.col_nr).row_off;
         end if;
         t_ran.row_nr:=t_ran.row_nr+1;
       end loop;
    end if;
  end if;
  return case when t_ran.range_name is not null then t_ran.range_name||'!' else '' end
     ||alfan2col(t_ran.col_nr)||t_ran.row_nr;
end;
 
--
Procedure read_names_xlsx(p_xlsx BLOB)
is
  t_val varchar2(4000);
  t_pat varchar2(200);
  t_ns varchar2(200) := 'xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"';
  t_nd dbms_xmldom.domnode;
  t_nd1 dbms_xmldom.domnode;
  t_nl dbms_xmldom.domnodelist;
  t_nl1 dbms_xmldom.domnodelist;
  t_dateFmtId number:=null;
  t_ind pls_integer;
  t_loc tp_loc;
  type tp_sheetxmlno is table of number index by pls_integer;
  t_sheetno tp_sheetxmlno;
  t_xml xmltype;
begin
  t_xml := get1xml( context.template, 'xl/_rels/workbook.xml.rels' );
  for t in
  ( select substr(target,17,length(target)-20) n, substr(id,4) id
    from XMLTable(
        xmlnamespaces(default 'http://schemas.openxmlformats.org/package/2006/relationships'),    
         '/Relationships/Relationship' passing t_xml columns
        id varchar2(5) path './@Id'
      , target varchar2(128) path './@Target')
    where target like 'worksheets/%')
  loop
     t_sheetno(t.id) := t.n;
  end loop;
  t_nd := blob2node( as_zip.get_file( p_xlsx, 'xl/workbook.xml' ) );
  context.date1904 := lower( dbms_xslprocessor.valueof( t_nd, '/workbook/workbookPr/@date1904', t_ns ) ) in ( 'true', '1' );
  t_nl := dbms_xslprocessor.selectnodes( t_nd, '/workbook/sheets/sheet', t_ns );
  for i in 0 .. dbms_xmldom.getlength( t_nl ) - 1
  loop
-- sheets in ranges
--    t_ind := dbms_xslprocessor.valueof( dbms_xmldom.item( t_nl, i ), '@sheetId', 'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"' );
    t_ind := substr(dbms_xslprocessor.valueof( dbms_xmldom.item( t_nl, i ), '@r:id', 'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"' ),4);
--dbms_output.put_line(t_ind);
    t_ind := t_sheetno(t_ind);
    add_range(dbms_xslprocessor.valueof( dbms_xmldom.item( t_nl, i ), '@name' ), '',t_ind);
    t_nd1 := blob2node(as_zip.get_file( p_xlsx, 'xl/worksheets/sheet'||t_ind||'.xml' ) );
    t_nl1 := dbms_xslprocessor.selectnodes(t_nd1, '/worksheet/mergeCells/mergeCell',t_ns);
    for j in 0..dbms_xmldom.getlength( t_nl1 )-1
    loop
        t_loc:=addr2loc(dbms_xslprocessor.valueof( dbms_xmldom.item( t_nl1, j ),'@ref',t_ns));
        context.merges(t_ind)(t_loc.row_nr)(t_loc.col_nr).row_off:=t_loc.row_off;
        context.merges(t_ind)(t_loc.row_nr)(t_loc.col_nr).col_off:=t_loc.col_off;
--dbms_output.put_line(alfan2cell(t_loc.row_nr,t_loc.col_nr)||':'||alfan2cell(t_loc.row_nr+t_loc.row_off,t_loc.col_nr+t_loc.col_off));
    end loop;
    dbms_xmldom.freeDocument(dbms_xmldom.getOwnerDocument(t_nd1));
  end loop;
  t_nl := dbms_xslprocessor.selectnodes( t_nd, '/workbook/definedNames/definedName', t_ns );
  for i in 0 .. dbms_xmldom.getlength( t_nl ) - 1
  loop
    add_range(dbms_xslprocessor.valueof( dbms_xmldom.item( t_nl, i ), '@name' )
      , dbms_xslprocessor.valueof( dbms_xmldom.item( t_nl, i ), '.' ) );
--dbms_output.put_line(dbms_xslprocessor.valueof( dbms_xmldom.item( t_nl, i ), '@name' )||'|'||dbms_xslprocessor.valueof( dbms_xmldom.item( t_nl, i ), '.' ));
  end loop;
  dbms_xmldom.freeDocument(dbms_xmldom.getOwnerDocument(t_nd));
-- date styles
  t_nd := blob2node( as_zip.get_file( p_xlsx, 'xl/styles.xml' ) );
  t_nl := dbms_xslprocessor.selectnodes( t_nd, '/styleSheet/numFmts/numFmt', t_ns );
  for i in 0 .. dbms_xmldom.getlength( t_nl ) - 1
  loop
    t_val := dbms_xslprocessor.valueof( dbms_xmldom.item( t_nl, i ), '@formatCode' );
    if (  instr( t_val, 'DD' ) > 0
       or instr( t_val, 'MM' ) > 0
       or instr( t_val, 'YY' ) > 0
       )
    then
      t_dateFmtId:=dbms_xslprocessor.valueof( dbms_xmldom.item( t_nl, i ), '@numFmtId' ) ;
--dbms_output.put_line(t_dateFmtId);
      exit;
    end if;
  end loop;
  t_nl := dbms_xslprocessor.selectnodes( t_nd, '/styleSheet/cellXfs/xf/@numFmtId', t_ns );
  for i in 0 .. dbms_xmldom.getlength( t_nl ) - 1
  loop
    if t_dateFmtId = dbms_xmldom.getnodevalue( dbms_xmldom.item( t_nl, i ) )
    then
      context.date_style := i;
      exit;
    end if;
  end loop;
  dbms_xmldom.freeDocument(dbms_xmldom.getOwnerDocument(t_nd));
  t_nd := blob2node( as_zip.get_file( p_xlsx, 'xl/sharedStrings.xml' ) );
  if not dbms_xmldom.isnull( t_nd )
  then
    context.str_cnt_of := dbms_xmldom.getlength(dbms_xslprocessor.selectnodes( t_nd, '/sst/si', t_ns ));
  end if;
  dbms_xmldom.freeDocument(dbms_xmldom.getOwnerDocument(t_nd));
end;
--
Procedure read_styles_xlsx(p_xlsx BLOB, p_sheet_nr pls_integer)
as
  t_loc tp_loc;
  t_ftype char(1);
  t_ns varchar2(200) := 'xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"';
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
  t_nd := blob2node( as_zip.get_file( p_xlsx, 'xl/worksheets/sheet' || p_sheet_nr || '.xml' ) );
  t_nl3 := dbms_xslprocessor.selectnodes( t_nd, '/worksheet/sheetData/row' );
  for r in 0 .. dbms_xmldom.getlength( t_nl3 ) - 1 -- rows
  loop
    t_nr := dbms_xslprocessor.valueof(dbms_xmldom.item( t_nl3, r ),'@r');
    t_val := dbms_xslprocessor.valueof(dbms_xmldom.item( t_nl3, r ),'@ht');
    context.styles.rows(t_nr).height :=
       to_number( t_val, translate( t_val, '.012345678,-+', 'D999999999' ), 'NLS_NUMERIC_CHARACTERS=.,' );
    context.styles.rows(t_nr).style := dbms_xslprocessor.valueof(dbms_xmldom.item( t_nl3, r ),'@s');
-- dbms_output.put_line(r);
    t_nl2 := dbms_xslprocessor.selectnodes( dbms_xmldom.item( t_nl3, r ), 'c' );
    t_loc.sht_nr := p_sheet_nr;
    t_loc.row_nr := t_nr;
    if dbms_xmldom.getlength( t_nl2 ) = 0 and not field_exists(t_loc)
    then
      context.fields(p_sheet_nr)(t_nr)(0)(0) := null; -- no cols, row exists
    end if;
    for j in 0 .. dbms_xmldom.getlength( t_nl2 ) - 1 -- cols
    loop
      t_r := dbms_xslprocessor.valueof( dbms_xmldom.item( t_nl2, j ), '@r', t_ns );
      t_loc := addr2loc(t_r);
      t_loc.sht_nr := p_sheet_nr;
-- dbms_output.put_line(t_r||' '||t_loc.row_nr||' '||t_loc.col_nr);
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
-- dbms_output.put_line(t_r||' '||t_t||' '||t_s||' '||t_val);
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
  read_names_xlsx(context.template);
end;
--
Function replace_XML_node(p_xml in out xmltype, p_xpath varchar2, p_newnode CLOB) return xmlType
as
  t_xtmp xmlType;
begin
  select updateXML(p_xml ,p_xpath, p_newnode, 'xmlns="'||p_xml.getnamespace()||'"') into t_xtmp from dual;
  return t_xtmp;
end;
-- p_target like 'worksheets/sheet1.xml' or 'sharedStrings.xml' rId return
Function add_RelsType(p_target varchar2, p_type varchar2, p_ctype varchar2) return number
is
  t_xml xmltype;
  t_rid number;
  t_xname varchar2(200);
begin
  t_xname := 'xl/_rels/workbook.xml.rels';
  t_xml := get1xml( context.template, t_xname );
--  select count(id)+1 into t_rid
  select max(to_number(substr(id,4))) + 1 into t_rid
    from XMLTable(
        xmlnamespaces(default 'http://schemas.openxmlformats.org/package/2006/relationships'),    
         '/Relationships/Relationship' passing t_xml columns
        id varchar2(8) path './@Id');
  select
     appendChildXML(t_xml
       , '/Relationships'
       , xmltype('<Relationship Id="rId'||t_rid||'" '||
         'Type="'||p_type||'" '||
         'xmlns="'||t_xml.getnamespace()||'" '||
         'Target="'||p_target||'"/>')
       , 'xmlns="'||t_xml.getnamespace()||'"')
     into t_xml
     from dual ;
--dbms_output.put_line(t_xml.getClobVal(1,1));
  replace1xml(context.template, t_xname, t_xml);
 
  t_xname:='[Content_Types].xml';
  t_xml:=get1xml( context.template, t_xname );
  select
     appendChildXML(t_xml
       , '/Types'
       , xmltype('<Override PartName="/xl/'||p_target||'" '||
          'xmlns="'||t_xml.getnamespace()||'" '||
          'ContentType="'||p_ctype||'"/>')
       , 'xmlns="'||t_xml.getnamespace()||'"')
     into t_xml
     from dual ;
--dbms_output.put_line(t_xml.getClobVal(1,1));
  replace1xml(context.template, t_xname, t_xml);
  return t_rid;
end;
--
Procedure finish_sheet(p_sn pls_integer, p_new_sn pls_integer:=null)
as
  t_row_of pls_integer;
  t_clob CLOB;
  t_xml XMLType;
  t_merges CLOB;
  t_merges_cnt pls_integer := 0;
  t_row_style tp_one_row_style;
  t_row_style_null tp_one_row_style;
  s pls_integer := p_sn;
  r pls_integer;
  ro pls_integer;
  c pls_integer;
  t_fld tp_one_field;
begin
  dbms_lob.createtemporary(t_clob, false);
  dbms_lob.createtemporary(t_merges, false);
  read_styles_xlsx(context.template, s);
  t_row_of := 0;
  t_clob := '<sheetData '||c_ssns||'>';
  t_merges_cnt := 0;
  t_merges := '<mergeCells '||c_ssns||' count="xxx">';
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
-- dbms_output.put_line((r+ro+t_row_of)||' '||ro||' '||c||' '||t_fld.value);
          dbms_lob.append(t_clob
             , '<c r="'
             ||alfan2cell(t_row_of+r+ro,c)||'" '||get_cell_style(r, c, t_fld)
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
                  , '<c r="'||alfan2cell(t_row_of+r+ro,c+i)
                  ||'" s="'||context.styles.cells(r)(c+i).style||'"/>'
                );
              end loop;
            end if;
            dbms_lob.append(t_merges
               ,'<mergeCell ref="'||alfan2cell(t_row_of+r+ro,c)||':'
                ||alfan2cell( t_row_of + r + ro + context.merges(s)(r)(c).row_off
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
  t_xml:=get1xml(context.template,'xl/worksheets/sheet' || s || '.xml');
  t_xml:=replace_XML_node(t_xml, '/worksheet/sheetData', t_clob);
  t_xml:=replace_XML_node(t_xml, '/worksheet/mergeCells', t_merges);
  replace1xml(context.template
    , 'xl/worksheets/sheet' || nvl(p_new_sn, p_sn) || '.xml', t_xml);
  dbms_lob.freetemporary(t_clob);
  dbms_lob.freetemporary(t_merges);
  clear_styles;
end;
--
Procedure save_sharedStrings
as
  t_xname varchar2(200):='xl/sharedStrings.xml';
  t_xml xmltype:=null;
  t_clob CLOB;
  s pls_integer;
begin
  if context.str_ind.count() = 0 then return; end if;
  t_xml := get1xml(context.template, t_xname);
  dbms_lob.createtemporary(t_clob, false);
  if t_xml is not null then
    dbms_lob.append(t_clob,replace(t_xml.getclobval(4,0),'</sst>',''));
  else
    s := add_RelsType('sharedStrings.xml'
      ,'http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings'
      ,'application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml');
    dbms_lob.append(t_clob,
'<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="1" uniqueCount="1">');
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
    finish_sheet(s);
    s:=context.fields.next(s);
  end loop; -- sheets
  save_sharedStrings;
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
  t_nid number;
  t_rid number;
  t_sid number;
  t_ind number;
  t_pos number;
  t_str varchar2(500 char);
begin
  t_sid := addr2loc(p_sheet_name).sht_nr;
  if t_sid is null then return; end if;
  begin
    if addr2loc(p_newsheet_name).sht_nr is not null then return; end if;
  exception
    when others then null;
  end;
  t_xname:='xl/workbook.xml';
  t_xml:=get1xml( context.template, t_xname );
  t_str := sheet_name(t_sid); --???
  select sheetid, rownum
     into t_sid, t_pos
     from XMLTable(
         xmlnamespaces(default 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'),    
         '/workbook/sheets/sheet' passing t_xml columns
         sheetid number path './@sheetId',
         sheetname varchar2(200) path './@name')
     where sheetname=t_str;
 
  select max(sheetid)+1 into t_nid
     from XMLTable(
         xmlnamespaces(default 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'),    
         '/workbook/sheets/sheet' passing t_xml columns
         sheetid number path './@sheetId');
 
  t_rid := add_RelsType('worksheets/sheet'||t_nid||'.xml'
      ,'http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet'
      ,'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml');
--
  t_str := '<sheet '||c_ssns||' '||c_rsns||' name="'||p_newsheet_name
           ||'" sheetId="'||t_nid||'" state="visible" r:id="rId'||t_rid||'"/>';
  if instr(p_options,'b') > 0
  then
    select  
      insertXMLbefore(t_xml
        ,'/workbook/sheets/sheet[@sheetId='||t_sid||']'
        ,xmltype(t_str)
        ,c_ssns||' '||c_rsns)
      into t_xml
      from dual ;
    t_pos := t_pos-1;
  else
    select  
      insertXMLafter(t_xml
        ,'/workbook/sheets/sheet[@sheetId='||t_sid||']'
        ,xmltype(t_str)
        ,c_ssns||' '||c_rsns)
      into t_xml
      from dual ;
  end if;      
-- hide source sheet
  if instr(p_options,'h') > 0 then
    select
      updateXML(t_xml
          , '/workbook/sheets/sheet[@sheetId='||t_sid||']/@state'
          , 'hidden'
          , c_ssns )
      into t_xml from dual;
    select
       updateXML(t_xml
          , '/workbook/bookViews/workbookView/@activeTab'
          , t_pos
          , c_ssns )
       into t_xml from dual;
  end if;
--dbms_output.put_line(t_xml.getClobVal(1,1));
  replace1xml(context.template, t_xname, t_xml);
  finish_sheet(t_sid, t_nid);
  clear_sheet_fields(t_sid);
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
Function new_workbook return BLOB
is
Function base642blob(p_clob CLOB) return BLOB
as
l_blob BLOB;
l_str varchar2(32767);
l_buflen pls_integer:=130;--132;  
l_amount pls_integer:=l_buflen;
l_offset pls_integer:=1;
begin
dbms_lob.createtemporary(l_blob,true,dbms_lob.session);
dbms_lob.read(p_clob,l_amount,1,l_str);
l_offset:=l_amount+1;
l_blob:= utl_encode.base64_decode(utl_raw.cast_to_raw(l_str));
loop
  dbms_lob.read(p_clob,l_amount,l_offset,l_str);
  dbms_lob.append(l_blob,utl_encode.base64_decode(utl_raw.cast_to_raw(l_str)));
  l_offset:=l_offset+l_amount;
end loop;
return l_blob;
exception when no_data_found then
  return l_blob;
end;
begin
return base642blob(
'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');
end;
 
end "AM_FILL";
/

