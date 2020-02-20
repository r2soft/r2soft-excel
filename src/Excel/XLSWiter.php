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

    public function writeSheetHeader($sheet_name, array $header_types, $col_options = null)
    {
        if (empty($sheet_name) || empty($header_types) || !empty($this->sheets[$sheet_name]))
            return;

        $suppress_row = isset($col_options['suppress_row']) ? intval($col_options['suppress_row']) : false;
        if (is_bool($col_options))
        {
            self::log( "Warning! passing $suppress_row=false|true to writeSheetHeader() is deprecated, this will be removed in a future version." );
            $suppress_row = intval($col_options);
        }
        $style = &$col_options;

        $col_widths = isset($col_options['widths']) ? (array)$col_options['widths'] : array();
        $auto_filter = isset($col_options['auto_filter']) ? intval($col_options['auto_filter']) : false;
        $freeze_rows = isset($col_options['freeze_rows']) ? intval($col_options['freeze_rows']) : false;
        $freeze_columns = isset($col_options['freeze_columns']) ? intval($col_options['freeze_columns']) : false;
        self::initializeSheet($sheet_name, $col_widths, $auto_filter, $freeze_rows, $freeze_columns);
        $sheet = &$this->sheets[$sheet_name];
        $sheet->columns = $this->initializeColumnTypes($header_types);
        if (!$suppress_row)
        {
            $header_row = array_keys($header_types);

            $sheet->file_writer->write('<row collapsed="false" customFormat="false" customHeight="false" hidden="false" ht="12.1" outlineLevel="0" r="' . (1) . '">');
            foreach ($header_row as $c => $v) {
                $cell_style_idx = empty($style) ? $sheet->columns[$c]['default_cell_style'] : $this->addCellStyle( 'GENERAL', json_encode(isset($style[0]) ? $style[$c] : $style) );
                $this->writeCell($sheet->file_writer, 0, $c, $v, $number_format_type='n_string', $cell_style_idx);
            }
            $sheet->file_writer->write('</row>');
            $sheet->row_count++;
        }
        $this->current_sheet = $sheet_name;
    }

    public function writeSheetRow($sheet_name, array $row, $row_options=null)
    {
        if (empty($sheet_name))
            return;

        self::initializeSheet($sheet_name);
        $sheet = &$this->sheets[$sheet_name];
        if (count($sheet->columns) < count($row)) {
            $default_column_types = $this->initializeColumnTypes( array_fill($from=0, $until=count($row), 'GENERAL') );//will map to n_auto
            $sheet->columns = array_merge((array)$sheet->columns, $default_column_types);
        }

        if (!empty($row_options))
        {
            $ht = isset($row_options['height']) ? floatval($row_options['height']) : 12.1;
            $customHt = isset($row_options['height']) ? true : false;
            $hidden = isset($row_options['hidden']) ? (bool)($row_options['hidden']) : false;
            $collapsed = isset($row_options['collapsed']) ? (bool)($row_options['collapsed']) : false;
            $sheet->file_writer->write('<row collapsed="'.($collapsed).'" customFormat="false" customHeight="'.($customHt).'" hidden="'.($hidden).'" ht="'.($ht).'" outlineLevel="0" r="' . ($sheet->row_count + 1) . '">');
        }
        else
        {
            $sheet->file_writer->write('<row collapsed="false" customFormat="false" customHeight="false" hidden="false" ht="12.1" outlineLevel="0" r="' . ($sheet->row_count + 1) . '">');
        }

        $style = &$row_options;

        $c=0;
        foreach ($row as $v) {

            if(isset($style['format-number'])){
                $number_format = $row_options['format-number'];
            }else{
                $number_format = $sheet->columns[$c]['number_format'];
            }

            $number_format_type = $sheet->columns[$c]['number_format_type'];
            $cell_style_idx = empty($style) ? $sheet->columns[$c]['default_cell_style'] : $this->addCellStyle( $number_format, json_encode(isset($style[0]) ? $style[$c] : $style) );
            $this->writeCell($sheet->file_writer, $sheet->row_count, $c, $v, $number_format_type, $cell_style_idx);
            $c++;
        }
        $sheet->file_writer->write('</row>');
        $sheet->row_count++;
        $this->current_sheet = $sheet_name;
    }

}