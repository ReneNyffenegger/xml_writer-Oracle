
create directory excel_xml_out as 'c:\temp';

begin


  excel_xml.start_('EXCEL_XML_OUT', 'test_01.xml');

  excel_xml.write_('<Styles>');
  excel_xml.write_('<Style ss:ID="date_dd_mm_yyyy"><NumberFormat ss:Format="dd.mm.yyyy;@"/></Style>');
  excel_xml.write_('</Styles>');

  excel_xml.write_('<Worksheet ss:Name="' || 'Report' || '">');

  excel_xml.write_(q'{<Table ss:DefaultColumnWidth="60">
   <Column  ss:Width="91"  />
   <Column  ss:Width="182"  />
   <Column  ss:AutoFitWidth="1"  />
   <Column  ss:AutoFitWidth="1"  />}');
  
  excel_xml.write_('<Row>');
  excel_xml.cd(1);
  excel_xml.cd('one');
  excel_xml.cd(date '2017-03-18');
  excel_xml.write_('</Row>');

  excel_xml.write_('<Row>');
  excel_xml.cd(2);
  excel_xml.cd('two');
  excel_xml.cd(cast(null as date));
  excel_xml.write_('</Row>');

  excel_xml.write_('<Row>');
  excel_xml.cd(3.3);
  excel_xml.cd(cast (null as varchar2));
  excel_xml.cd(date '2017-08-28');
  excel_xml.write_('</Row>');

  excel_xml.write_('<Row>');
  excel_xml.cd(42.12345678901234567890123456790123456789);
  excel_xml.cd('42.....');
  excel_xml.cd(date '2017-08-28');
  excel_xml.write_('</Row>');

  excel_xml.write_('<Row>');
  excel_xml.cd(142.123456789012345678901234567901234567890);
  excel_xml.cd('42.....');
  excel_xml.cd(date '2017-08-28');
  excel_xml.write_('</Row>');

  excel_xml.write_('<Row>');
  excel_xml.cd(7142.123456789012345678901234567901234567890);
  excel_xml.cd('42.....');
  excel_xml.cd(date '2017-08-28');
  excel_xml.write_('</Row>');

  excel_xml.write_('<Row>');
  excel_xml.cd(1234567890.123456789012345678901234567901234567890);
  excel_xml.cd('42.....');
  excel_xml.cd(date '2017-08-28');
  excel_xml.write_('</Row>');

  excel_xml.write_('</Table>');
  excel_xml.write_('</Worksheet>');
  excel_xml.write_('</Workbook>');

  excel_xml.end_;

end;
/

drop directory excel_xml_out;
