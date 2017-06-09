set define off
create or replace package body excel_xml as

  f utl_file.file_type;

  function escape_(text varchar2) return varchar2 is
  begin

    return replace(replace(replace(
          '&', '&amp;'),
          '<', '&lt;' ),
          '>', '&gt;' );

  end escape_;

  procedure start_(dir_name varchar2, filename varchar2) is -- {
  begin

    f := utl_file.fopen(dir_name, filename, 'W', 32767);

    write_(q'{<?xml version="1.0"?>
<?mso-application progid="Excel.Sheet"?>);
<Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet"
 xmlns:o="urn:schemas-microsoft-com:office:office"
 xmlns:x="urn:schemas-microsoft-com:office:excel"
 xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet"
 xmlns:html="http://www.w3.org/TR/REC-html40">}');

  end start_; -- }

  procedure end_ is -- {
  begin

    utl_file.fclose(f);
    f := null;

  end end_; -- }

  procedure write_(text varchar2) is -- {
  begin
    utl_file.put_line(f, text);
  end write_; -- }

  procedure cd(cell_data varchar2) is -- {
  begin
    if cell_data is null then cn; return; end if;

    write_('<Cell><Data ss:Type="String">' || cell_data || '</Data></Cell>');
  end cd; -- }

  procedure cd(cell_data number, style varchar2 := null) is -- {
    style_ varchar2(100);
  begin
    if cell_data is null then cn; return; end if;

    if style is not null then
       style_ := ' ss:StyleID="' || style || '"';
    end if;

    write_('<Cell' || style_ || '><Data ss:Type="Number">' || round(cell_data, 13) || '</Data></Cell>');
  end cd; -- }

  procedure cd(cell_data date) is -- {
  begin
    if cell_data is null then cn; return; end if;

    write_('<Cell ss:StyleID="date_dd_mm_yyyy"><Data ss:Type="DateTime" >' || to_char(cell_data, 'yyyy-mm-dd') || '</Data></Cell>');

  end cd; -- }

  procedure cn is -- {
  begin

    write_('<Cell></Cell>');

  end cn; -- }
  

end excel_xml;
/
sho err
