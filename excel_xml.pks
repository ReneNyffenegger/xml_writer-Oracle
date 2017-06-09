create or replace package excel_xml as

  procedure start_(dir_name varchar2, filename varchar2);
  procedure end_;

  procedure write_(text varchar2);

  procedure cd(cell_data varchar2);
  procedure cd(cell_data number  , style varchar2 := null);
  procedure cd(cell_data date    );

  procedure cn;
  

end excel_xml;
/
