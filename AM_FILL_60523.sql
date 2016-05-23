CREATE OR REPLACE PACKAGE  "AM_FILL" as
/*
*    Purpose: using the Microsoft XLSX files as a templates for data output
*
*    Uses AS_ZIP package designed by Anton Scheffer 
*    Download from: http://technology.amis.nl/wp-content/uploads/2010/06/as_zip7.txt
*
*    Based on the code of packages as_xlsx & as_read_xlsx designed by Anton Scheffer
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
*      2016-05-23 fixed bug: sheet xml name, thanks github.com/Zulus88
*    
******************************************************************************
* Copyright (C) 2011 - 2013 by Anton Scheffer
*               2013 - 2016 by MikTim
* Permission is hereby granted, free of charge, to any person obtaining a copy
* of this software and associated documentation files (the "Software"), to deal
* in the Software without restriction, including without limitation the rights
* to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
* copies of the Software, and to permit persons to whom the Software is
* furnished to do so, subject to the following conditions:
* The above copyright notice and this permission notice shall be included in
* all copies or substantial portions of the Software.
* THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
* IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
* FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
* AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
* LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
* OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
* THE SOFTWARE.
******************************************************************************
*/
version constant varchar2(16):='2.0.60523';
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
   p_cell_addr: A1 style cell address (sheet_name!cell_address) or area name
   p_options:
     i - insert rows mode (sequentially on every call), otherwise - overwrite
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
   p_sql: sql query text
   p_cell_addr: default A1 of current sheet
   p_options:
     h - print header (field names)
     i - insert rows mode. restriction: one record - one row
*/
Procedure in_table
( p_sql CLOB
, p_cell_addr varchar2:=null
, p_options varchar2:='');
/*
   IN_SHEET: Save (duplicate) filled sheet with new name,
             clear data from source sheet
*/
Procedure in_sheet
( p_sheet_name varchar2    -- source sheet name
, p_newsheet_name varchar2 -- new sheet name
);
/*
   FINISH: Generate workbook, clear internal structures
   notice: all formulas from filled sheets will be removed
*/
Procedure finish
( p_xfile in out nocopy BLOB -- filled in xlsx return
);
/*
   ADDRESS: Calculate relative (sheet or named area) address
   if p_range_name omit, current sheet used
*/
Function address
( p_row pls_integer
, p_col pls_integer
, p_range_name varchar2 := ''
) return varchar2;
/* NEW_WORKBOOK: Empty workbook returned */
Function new_workbook return BLOB;
/*
******************************************************************************
* Package usage:
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
  ( range_name varchar2(200)
  , isSheet boolean := false
  , sht_nr  pls_integer := 1
  , row_nr  pls_integer := 1
  , col_nr  pls_integer := 1
  );
type tp_all_ranges is table of tp_one_range index by pls_integer;
type tp_cell_loc is record
  ( sht_nr pls_integer 
  , row_nr pls_integer
  , col_nr pls_integer
  , row_off pls_integer
  , col_off pls_integer
  );
type tp_one_field is record
  ( type char(1)  -- N,D,S ?? I - initial cell value
  , value number
  );
type tp_fcols is table of tp_one_field index by pls_integer;
type tp_frows_of is table of tp_fcols index by pls_integer;
type tp_frows is table of tp_frows_of index by pls_integer;
type tp_fsheets is table of tp_frows index by pls_integer; 
--
type tp_strings is table of pls_integer index by varchar2(32767);
type tp_str_ind is table of varchar2(32767) index by pls_integer;
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
  ( row_off pls_integer
  , col_off pls_integer
  );
type tp_col_merge is table of tp_one_merge  index by pls_integer;
type tp_row_merge is table of tp_col_merge index by pls_integer;
type tp_merges is table of tp_row_merge index by pls_integer;
type tp_context is record
  ( template BLOB
  , options varchar2(10)
  , tpl_type char(1)               -- reserved
  , date1904 boolean := true       -- date format
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
function alfan_col( p_col pls_integer )
return varchar2
is
begin
  return case
           when p_col > 702 then chr( 64 + trunc( ( p_col - 27 ) / 676 ) ) || chr( 65 + mod( trunc( ( p_col - 1 ) / 26 ) - 1, 26 ) ) || chr( 65 + mod( p_col - 1, 26 ) )
           when p_col > 26  then chr( 64 + trunc( ( p_col - 1 ) / 26 ) ) || chr( 65 + mod( p_col - 1, 26 ) )
           else chr( 64 + p_col )
         end;
end;
--
function alfan_cell(p_row pls_integer, p_col pls_integer, p_range_name varchar2:='') return varchar2
is
begin
   return case when p_range_name is not null then p_range_name||'!' else '' end
     ||alfan_col(p_col)||p_row;
end;
--  
function col_alfan( p_col varchar2 )
return pls_integer
is
begin
  return ascii( substr( p_col, -1 ) ) - 64
       + nvl( ( ascii( substr( p_col, -2, 1 ) ) - 64 ) * 26, 0 )
       + nvl( ( ascii( substr( p_col, -3, 1 ) ) - 64 ) * 676, 0 );
end;
-- 
function name2loc(p_range_name varchar2) return tp_cell_loc
is
  t_loc tp_cell_loc;
  t_loc1 tp_cell_loc;
  t_loc2 tp_cell_loc;
  t_ind pls_integer;
  t_tmp1 varchar2(100); 
  t_tmp2 varchar2(100);
  t_tmp3 varchar2(100);
  t_rowcolp varchar2(100) := '\$?([[:alpha:]]{1,3})\$?([[:digit:]]*)';
  t_rcellp varchar2(100) := '(\$?[[:alpha:]]{1,3}\$?[[:digit:]]*)(:(\$?[[:alpha:]]{1,3}\$?[[:digit:]]*))?';
  t_rangep varchar2(200) := '^((\$?[[:alpha:]]{1,3}\$?[[:digit:]]*)(:\$?[[:alpha:]]{1,3}\$?[[:digit:]]*)?)$'||
--      '|^(('')?([[:alnum:]_ ]+)\5?){1}(!((\$?[[:alpha:]]{1,3}\$?[[:digit:]]*)(:\$?[[:alpha:]]{1,3}\$?[[:digit:]]*)?))?$';
      '|^(('')?([\w ]+|[[:alnum:]_ ]+)\5?){1}(!((\$?[[:alpha:]]{1,3}\$?[[:digit:]]*)(:\$?[[:alpha:]]{1,3}\$?[[:digit:]]*)?))?$';
begin
--dbms_output.put_line(p_range_name);
  if regexp_instr(p_range_name,t_rangep)>0 then
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
      t_loc1.col_nr := col_alfan(regexp_replace(t_tmp2,t_rowcolp,'\1'));
      t_loc1.row_nr := to_number(regexp_replace(t_tmp2,t_rowcolp,'\2'));
      t_loc2.col_nr := col_alfan(regexp_replace(t_tmp3,t_rowcolp,'\1'));
      t_loc2.row_nr := to_number(regexp_replace(t_tmp3,t_rowcolp,'\2'));
      t_loc.col_nr := t_loc.col_nr + least(t_loc1.col_nr, t_loc2.col_nr)-1;
      t_loc.row_nr := t_loc.row_nr + least(t_loc1.row_nr, t_loc2.row_nr)-1;
      t_loc.col_off := greatest(t_loc1.col_nr,t_loc2.col_nr) - least(t_loc1.col_nr,t_loc2.col_nr);
      t_loc.row_off := greatest(t_loc1.row_nr,t_loc2.row_nr) - least(t_loc1.row_nr,t_loc2.row_nr);
    end if;
  end if;
--dbms_output.put_line(p_range_name||'|'||t_loc.sht_nr||'|'||t_loc.row_nr||'|'||t_loc.col_nr||'|'||t_loc.col_off||'|'||t_loc.row_off);
  if (t_loc.sht_nr is null or t_loc.row_nr is null or t_loc.col_nr is null)
     and instr(context.options,'e') > 0
  then
      raise_application_error(-20001,'AM_FILL #REF!: '||p_range_name);
  else 
     t_loc.sht_nr := nvl(t_loc.sht_nr,1);
     t_loc.row_nr := nvl(t_loc.row_nr,1);
     t_loc.col_nr := nvl(t_loc.col_nr,1);
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
  t_loc tp_cell_loc;
begin
  if not context.ranges.exists( upper(p_range_name) )
  then
    t_cnt := context.ranges.count()+1;  
    context.ran_ind( t_cnt ).range_name := p_range_name;
    context.ranges( upper(nvl( p_range_name, '' ) )) := t_cnt;
    if p_sht_nr = 0
    then
      t_loc := name2loc(p_range_def);
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
Function cell_merged(s pls_integer, r pls_integer, c pls_integer) return boolean
is
begin
  return context.merges(s)(r).exists(c);
exception when no_data_found then return false;
end;
--
Function align_loc( p_loc tp_cell_loc ) return tp_cell_loc
is
  t_loc tp_cell_loc:=p_loc;
  r pls_integer;
  c pls_integer;
  t_merge tp_one_merge;
begin
  t_loc.col_off := 0;
  t_loc.row_off := 0;
  r := context.merges(p_loc.sht_nr).first();
  while r is not null and r <= p_loc.row_nr loop
    c := context.merges(p_loc.sht_nr)(r).first();
    while c is not null and c <= p_loc.col_nr loop
      t_merge := context.merges(p_loc.sht_nr)(r)(c);
      if (p_loc.row_nr between r and r+t_merge.row_off)
        and (p_loc.col_nr between c and c+t_merge.col_off)
      then
        t_loc.row_nr := r;
        t_loc.col_nr := c;
        t_loc.row_off := t_merge.row_off;
        t_loc.col_off := t_merge.col_off;
--dbms_output.put_line(alfan_cell(t_loc.row_nr,t_loc.col_nr)||':'||alfan_cell(t_loc.row_nr+t_loc.row_off,t_loc.col_nr+t_loc.col_off));
        return t_loc;
      end if;
      c := context.merges(p_loc.sht_nr)(r).next(c);
    end loop;
    r := context.merges(p_loc.sht_nr).next(r);
  end loop;
  return t_loc;
end;
--
Function field_exists(s pls_integer,r pls_integer, ro pls_integer:=0, c pls_integer:=0)
return boolean
is
begin
  return context.fields(s)(r)(ro).exists(c);
exception when others then return false;
end;
 
Function next_col(p_loc tp_cell_loc) return tp_cell_loc
is
  t_loc tp_cell_loc := p_loc;
begin
  if cell_merged(t_loc.sht_nr,t_loc.row_nr,t_loc.col_nr) then
    t_loc.col_nr:=t_loc.col_nr+context.merges(t_loc.sht_nr)(t_loc.row_nr)(t_loc.col_nr).col_off;
  end if;
  t_loc.col_nr:=t_loc.col_nr+1;
  return align_loc(t_loc);
end;
 
Function next_row(p_loc tp_cell_loc) return tp_cell_loc
is
  t_loc tp_cell_loc := p_loc;
begin
  if cell_merged(t_loc.sht_nr,t_loc.row_nr,t_loc.col_nr) then
    t_loc.row_nr:=t_loc.row_nr+context.merges(t_loc.sht_nr)(t_loc.row_nr)(t_loc.col_nr).row_off;
  end if;
  t_loc.row_nr:=t_loc.row_nr+1;
  return align_loc(t_loc);
end;
--
Procedure add_value
  ( p_value number
  , p_type char
  , p_loc tp_cell_loc
  , p_row_off pls_integer := 0
  , p_col_off pls_integer := 0
  )
is
  t_field tp_one_field;
begin
  context.current_sht:=p_loc.sht_nr;
  t_field.type := p_type;
  t_field.value := p_value;
--dbms_output.put_line(p_loc.sht_nr||' '||p_loc.row_nr||' '||p_loc.col_nr);
  context.fields( p_loc.sht_nr )( p_loc.row_nr )( p_row_off )( p_loc.col_nr ) := t_field;
end;
--
Procedure add_field
  ( p_value date
  , p_loc tp_cell_loc
  , p_row_off pls_integer:=0
  , p_col_off pls_integer:=0
  )
is
begin
  add_value( p_value - case when context.date1904 then to_date('01-01-1904','DD-MM-YYYY')
          else to_date('01-01-1900','DD-MM-YYYY') end + 2
     ,'D', p_loc, p_row_off, p_col_off);
end;
--
Procedure add_field
  ( p_value number
  , p_loc tp_cell_loc
  , p_row_off pls_integer:=0
  , p_col_off pls_integer:=0
  )
is
begin
  add_value(p_value, 'N', p_loc, p_row_off, p_col_off);
end;
--
Procedure add_field
  ( p_value varchar2
  , p_loc tp_cell_loc
  , p_row_off pls_integer:=0
  , p_col_off pls_integer:=0
  )
is
begin
  add_value(add_string(p_value), 'S', p_loc, p_row_off, p_col_off);
end;
--
Function next_row_of(p_loc tp_cell_loc, p_col_off pls_integer:=0) return pls_integer
is
  i pls_integer:=0;
begin
  while field_exists(p_loc.sht_nr, p_loc.row_nr, i, p_loc.col_nr + p_col_off)
  loop
    i := i + 1;
  end loop;
--dbms_output.put_line(alfan_cell(p_loc.row_nr, p_loc.col_nr)||':'||i);
  return i;
end;
--
Procedure in_field(p_value number, p_cell_addr varchar2, p_options varchar2:='')
is
  t_overprint boolean := not instr(p_options,'i') > 0;
  t_loc tp_cell_loc := align_loc(name2loc(p_cell_addr));
begin
  add_field( p_value, t_loc
      , case when t_overprint then 0 else next_row_of(t_loc) end, 0);
end;
--
Procedure in_field(p_value date, p_cell_addr varchar2, p_options varchar2:='')
is
  t_overprint boolean := not instr(p_options,'i') > 0;
  t_loc tp_cell_loc := align_loc(name2loc(p_cell_addr));
begin
  add_field( p_value, t_loc 
      , case when t_overprint then 0 else next_row_of(t_loc) end, 0);
end;
--
Procedure in_field(p_value varchar2, p_cell_addr varchar2, p_options varchar2:='')
is
  t_overprint boolean := not instr(p_options,'i') > 0;
  t_loc tp_cell_loc := align_loc(name2loc(p_cell_addr));
begin
  add_field( p_value, t_loc
      , case when t_overprint then 0 else next_row_of(t_loc) end, 0);
end;
--
Procedure in_table(p_sql CLOB, p_cell_addr varchar2:=null, p_options varchar2:='')
as
  t_header pls_integer := case when instr(p_options, 'h') > 0 then 1 else 0 end ;
  t_overprint boolean := not instr(p_options, 'i') > 0;
  t_loc tp_cell_loc;
  t_floc tp_cell_loc;
  t_c integer;
  t_col_cnt integer;
  t_desc_tab dbms_sql.desc_tab2;
  d_tab dbms_sql.date_table;
  n_tab dbms_sql.number_table;
  v_tab dbms_sql.varchar2_table;
  t_bulk_size pls_integer := 200;
  t_r integer;
begin
  t_loc := align_loc(name2loc(p_cell_addr));
  t_c := dbms_sql.open_cursor;
  dbms_sql.parse( t_c, p_sql, dbms_sql.native );
  dbms_sql.describe_columns2( t_c, t_col_cnt, t_desc_tab );
  t_floc := t_loc;
  for c in 1 .. t_col_cnt
  loop
    if t_header > 0 
    then
      add_field(t_desc_tab( c ).col_name, t_floc, 0);
      t_floc := next_col( t_floc );
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
  if t_header > 0 and t_overprint then t_loc := next_row(t_loc); end if;
  t_r := dbms_sql.execute( t_c );
  loop
    t_r := dbms_sql.fetch_rows( t_c );
    if t_r > 0
    then
      for c in 1 .. t_col_cnt
      loop
        t_floc := t_loc;
        case
          when t_desc_tab( c ).col_type in ( 2, 100, 101 )
          then
            dbms_sql.column_value( t_c, c, n_tab );
            for i in 0 .. t_r - 1
            loop
                add_field(n_tab( i + n_tab.first() ), t_floc
                    , case when t_overprint then 0 else i + t_header end );
                if t_overprint then t_floc := next_row(t_floc); end if;
            end loop;
            n_tab.delete;
            t_loc := next_col(t_loc);
          when t_desc_tab( c ).col_type in ( 12, 178, 179, 180, 181 , 231 )
          then
            dbms_sql.column_value( t_c, c, d_tab );
            for i in 0 .. t_r - 1
            loop
                add_field(d_tab( i + d_tab.first() ), t_floc
                    , case when t_overprint then 0 else i + t_header end );
                if t_overprint then t_floc := next_row(t_floc); end if;
            end loop;
            d_tab.delete;
            t_loc := next_col(t_loc);
          when t_desc_tab( c ).col_type in ( 1, 8, 9, 96, 112 )
          then
            dbms_sql.column_value( t_c, c, v_tab );
            for i in 0 .. t_r - 1
            loop
                add_field( v_tab( i + v_tab.first() ), t_floc
                    , case when t_overprint then 0 else i + t_header end );
                if t_overprint then t_floc := next_row(t_floc); end if;
            end loop;
            v_tab.delete;
            t_loc := next_col(t_loc);
          else
            null;
        end case;
      end loop;
    end if;
    exit when t_r != t_bulk_size;
  end loop;
  dbms_sql.close_cursor( t_c );
/*
exception
  when others
  then
    if dbms_sql.is_open( t_c )
    then
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
         if cell_merged(t_ran.sht_nr,t_ran.row_nr,t_ran.col_nr) then
           t_ran.col_nr:=t_ran.col_nr
             + context.merges(t_ran.sht_nr)(t_ran.row_nr)(t_ran.col_nr).col_off;
         end if;
         t_ran.col_nr:=t_ran.col_nr+1;
       end loop;
       for i in 1..nvl(p_row,0)-1 loop
         if cell_merged(t_ran.sht_nr,t_ran.row_nr,t_ran.col_nr) then
           t_ran.row_nr:=t_ran.row_nr
             + context.merges(t_ran.sht_nr)(t_ran.row_nr)(t_ran.col_nr).row_off;
         end if;
         t_ran.row_nr:=t_ran.row_nr+1;
       end loop;
    end if;
  end if;
  return case when t_ran.range_name is not null then t_ran.range_name||'!' else '' end
     ||alfan_col(t_ran.col_nr)||t_ran.row_nr;
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
  t_loc tp_cell_loc;
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
    t_ind := dbms_xslprocessor.valueof( dbms_xmldom.item( t_nl, i ), '@sheetId', 'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"' );
    t_ind := t_sheetno(t_ind);
    add_range(dbms_xslprocessor.valueof( dbms_xmldom.item( t_nl, i ), '@name' ), '',t_ind);
    t_nd1 := blob2node(as_zip.get_file( p_xlsx, 'xl/worksheets/sheet'||t_ind||'.xml' ) );
    t_nl1 := dbms_xslprocessor.selectnodes(t_nd1, '/worksheet/mergeCells/mergeCell',t_ns);
    for j in 0..dbms_xmldom.getlength( t_nl1 )-1
    loop
        t_loc:=name2loc(dbms_xslprocessor.valueof( dbms_xmldom.item( t_nl1, j ),'@ref',t_ns));
        context.merges(t_ind)(t_loc.row_nr)(t_loc.col_nr).row_off:=t_loc.row_off;
        context.merges(t_ind)(t_loc.row_nr)(t_loc.col_nr).col_off:=t_loc.col_off;
--dbms_output.put_line(alfan_cell(t_loc.row_nr,t_loc.col_nr)||':'||alfan_cell(t_loc.row_nr+t_loc.row_off,t_loc.col_nr+t_loc.col_off));
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
  t_loc tp_cell_loc;
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
    if  dbms_xmldom.getlength( t_nl2 ) = 0 
        and not field_exists(p_sheet_nr, t_nr, 0, 0) 
    then
      context.fields(p_sheet_nr)(t_nr)(0)(0) := null; -- no cols, row exists
    end if;
    for j in 0 .. dbms_xmldom.getlength( t_nl2 ) - 1 -- cols
    loop
      t_r := dbms_xslprocessor.valueof( dbms_xmldom.item( t_nl2, j ), '@r', t_ns );
      t_loc := name2loc(t_r);
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
      if not field_exists(t_loc.sht_nr, t_loc.row_nr, 0, t_loc.col_nr)
      then
-- or replacing with initial cell value
        add_value(t_nr, t_ftype, t_loc);
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
        id varchar2(3) path './@Id');
  select
     appendChildXML(t_xml
       , '/Relationships'
       , xmltype('<Relationship Id="rId'||t_rid||'" '||
         'Type="'||p_type||'" '||
         'xmlns="'||t_xml.getnamespace()||'" '||
         'Target="'||p_target||'"/>')
       , 'xmlns="'||t_xml.getnamespace()||'"')
/*     xmlquery('xquery version "1.0"; (: :)
       declare default element namespace
         "http://schemas.openxmlformats.org/package/2006/relationships";
       copy $tmp := $x
         modify insert node
           <Relationship Id="rId{$rid}" Type="{$typ}" Target="{$tgt}"/>
         as last into
           $tmp/Relationships
       return $tmp'
     passing  t_xml as "x", t_rid as "rid", p_type as "typ", p_target as "tgt"
     returning content)
*/
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
/*      xmlquery('xquery version "1.0"; (: :)
       declare default element namespace
         "http://schemas.openxmlformats.org/package/2006/content-types";
       copy $tmp := $x
         modify insert node
           <Override PartName="/xl/{$tgt}" ContentType="$typ"/>
         as last into
          $tmp/Types
       return $tmp'
     passing  t_xml as "x", p_target as "tgt", p_ctype as "typ"
     returning content)
*/
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
             ||alfan_cell(t_row_of+r+ro,c)||'" '||get_cell_style(r, c, t_fld)
             ||case
                 when t_fld.value is null then '/>'
                 else '><v>'
                   ||to_char(t_fld.value, 'TM9', 'NLS_NUMERIC_CHARACTERS=.,' )
                   ||'</v></c>'
               end
          );
          if cell_merged(s,r,c)
          then
            if ro = 0 then
              if field_exists(s, r, 1, c) then
-- if insert mode, reset vertical merge
                context.merges( s )( r )( c ).row_off:=0;
              end if;
            else
              for i in 1..context.merges(s)(r)(c).col_off loop
                dbms_lob.append( t_clob
                  , '<c r="'||alfan_cell(t_row_of+r+ro,c+i)
                  ||'" s="'||context.styles.cells(r)(c+i).style||'"/>'
                );
              end loop;
            end if;
            dbms_lob.append(t_merges
               ,'<mergeCell ref="'||alfan_cell(t_row_of+r+ro,c)||':'
                ||alfan_cell( t_row_of + r + ro + context.merges(s)(r)(c).row_off 
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
Procedure in_sheet(p_sheet_name varchar2, p_newsheet_name varchar2)
as
  t_xname varchar2(200);
  t_xml xmltype;
  t_nid number;
  t_rid number;
  t_sid number;
  t_ind number;
begin
  if p_sheet_name is null or p_newsheet_name is null then return; end if;
  t_xname:='xl/workbook.xml';
  t_xml:=get1xml( context.template, t_xname );
--dbms_output.put_line(t_xml.getClobVal());
  select sheetid
     into t_sid
     from XMLTable(
         xmlnamespaces(default 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'),    
         '/workbook/sheets/sheet' passing t_xml columns
         sheetid number path './@sheetId',
         name varchar2(200) path './@name')
     where name=p_sheet_name;
 
  select max(sheetid)+1 into t_nid
     from XMLTable(
         xmlnamespaces(default 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'),    
         '/workbook/sheets/sheet' passing t_xml columns
         sheetid number path './@sheetId');
 
  t_rid := add_RelsType('worksheets/sheet'||t_nid||'.xml'
      ,'http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet'
      ,'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml');
--
  select  
     insertXMLafter(t_xml
         ,'/workbook/sheets/sheet[@sheetId='||t_sid||']'
         ,xmltype('<sheet name="'||p_newsheet_name||'" sheetId="'||t_nid||'" '||
             c_ssns||' '||c_rsns||' r:id="rId'||t_rid||'"/>')
         ,c_ssns)
/*     xmlquery('xquery version "1.0"; (: :)
       declare default element namespace 
         "http://schemas.openxmlformats.org/spreadsheetml/2006/main"; 
       declare namespace
          r = "http://schemas.openxmlformats.org/officeDocument/2006/relationships" ;
       copy $tmp := $x
         modify insert node 
           <sheet name="{$snm}" sheetId="{$nid}" r:id="rId{$rid}"/>
         after 
           $tmp/workbook/sheets/sheet[@sheetId=$sid] 
       return $tmp' 
     passing  t_xml as "x", t_sid as "sid"
      , p_newsheet_name as "snm", t_nid as "nid", t_rid as "rid"
     returning content)
*/
     into t_xml
     from dual ;
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

