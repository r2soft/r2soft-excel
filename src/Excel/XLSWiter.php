<?php


namespace R2Soft\Excel;
use XLSXWriter;

class XLSWiter extends XLSXWriter
{

    protected function buildWorkbookXML()
    {
        $i=0;
        $workbook_xml="";
        $workbook_xml.='<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'."\n";
        $workbook_xml.='<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">';
        $workbook_xml.='<fileVersion appName="Calc"/><workbookPr backupFile="false" showObjects="all" date1904="false"/><workbookProtection/>';
        $workbook_xml.='<bookViews><workbookView activeTab="0" firstSheet="0" showHorizontalScroll="true" showSheetTabs="true" showVerticalScroll="true" tabRatio="212" windowHeight="8192" windowWidth="16384" xWindow="0" yWindow="0"/></bookViews>';
        $workbook_xml.='<sheets>';
        foreach($this->sheets as $sheet_name=>$sheet) {
            $sheetname = self::sanitize_sheetname($sheet->sheetname);
            $workbook_xml.='<sheet name="'.self::xmlspecialchars($sheetname).'" sheetId="'.($i+1).'" state="visible" r:id="rId'.($i+2).'"/>';
            $i++;
        }
        $workbook_xml.='</sheets>';
        $workbook_xml.='<definedNames>';

        $i=0;
        foreach($this->sheets as $sheet_name=>$sheet) {
            if ($sheet->auto_filter) {
                $sheetname = self::sanitize_sheetname($sheet->sheetname);
//				$workbook_xml.='<definedName name="_xlnm._FilterDatabase" localSheetId="0" hidden="1">\''.self::xmlspecialchars($sheetname).'\'!$A$1:' . self::xlsCell($sheet->row_count - 1, count($sheet->columns) - 1, true) . '</definedName>';
                $workbook_xml.='<definedName name="_xlnm._FilterDatabase" localSheetId="'.$i.'" hidden="1">\''.self::xmlspecialchars($sheetname).'\'!$A$1:' . self::xlsCell($sheet->row_count - 1, count($sheet->columns) - 1, true) . '</definedName>';

                $i++;
            }
        }
        $workbook_xml.='</definedNames>';
        $workbook_xml.='<calcPr iterateCount="100" refMode="A1" iterate="false" iterateDelta="0.001"/></workbook>';
        return $workbook_xml;
    }

}