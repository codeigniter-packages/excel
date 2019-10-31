<?php

defined('BASEPATH') OR exit('No direct script access allowed');
/*
 * To change this template, choose Tools | Templates
 * and open the template in the editor.
 */


class Excel extends PHPExcel {

    /**
     * @header array();头部字段名称数组
     * @query;查询结果
     * @filename：导出的excel名字
     */
    function query_to_excel($header, $query_result, $filename) {

        $objPHPExcel = new PHPEXCEL();
        $sheet=$objPHPExcel->getActiveSheet();
        
        $sheet->fromArray($header, NULL, 'A1');
        $sheet->fromArray($query_result, NULL, 'A2');

        $objWriter = new PHPExcel_Writer_Excel2007($objPHPExcel);
        header("Pragma: public");
        header("Expires: 0");
        header("Cache-Control:must-revalidate, post-check=0, pre-check=0");
        header("Content-Type:application/force-download");
        header("Content-Type:application/vnd.ms-excel;");
        header("Content-Type:application/octet-stream");
        header("Content-Type:application/download");
        header("Content-Disposition:attachment;filename=" . urlencode($filename));
        header("Content-Transfer-Encoding:binary");
        $objWriter->save("php://output");
        exit;
    }

     function query_to_csv($header, $query_result, $filename) {
        header('Content-Type:text/csv; charset=gbk');
        header("Content-Disposition:attachment;filename=" . urlencode($filename));
        $out = fopen('php://output', 'w');
        foreach ($header as $key => $value) {
            $header[$key] = iconv("utf-8","gbk",$value);
        }
        fputcsv($out, $header);
        foreach ($query_result as $item) {
            $item = (array) $item;
            foreach ($item as $key => $value) {
                //只转换日期格式
                if(strtotime($value)!==FALSE && strstr($value,'-')!==FALSE){
                    $item[$key] ="\r".$value;
                }else{
                    $item[$key]=iconv("utf-8","gbk",$value);
                }

            }
            fputcsv($out, $item);
        }
        fclose($out);
        exit;
    }

    /**
      @ys_array=array('name'=>'姓名','age'=>'年龄');映射字段名称
     */
    function execl_to_array($filePath, $ys_array) {
        
        set_time_limit(0);
        $PHPReader = new PHPExcel_Reader_Excel2007();
        if (!$PHPReader->canRead($filePath)) {
            $PHPReader = new PHPExcel_Reader_Excel5();
            if (!$PHPReader->canRead($filePath)) {
                echo 'no Excel';
                return;
            }
        }
        //设置了这个以后，对于richtext就只读文本了！
        //$PHPReader->setReadDataOnly('true'); //如果只读数据将导致无法判断日期
        $PHPExcel = $PHPReader->load($filePath);

        //取得excel第一个文档
        $currentSheet = $PHPExcel->getSheet(0);

        //取得一共有多少列
        $columns = PHPExcel_Cell::columnIndexFromString($currentSheet->getHighestColumn());

        //取得一共有多少行
        $rows = $currentSheet->getHighestRow();

        $sqlarray = array();
        $header_fields = array();

        //取得列表数组，$header_fields的key为列名的排序，value为映射的英文字段名
        for ($field_index = 0; $field_index < $columns; $field_index++) {
            $item_column = $currentSheet->getCellByColumnAndRow($field_index, 1);
            $item_column_name = $item_column->getValue() instanceof PHPExcel_RichText ? $item_column->getValue()->getPlainText() : $item_column->getValue();
            if (isset($ys_array[$item_column_name])) {
                $header_fields[$field_index] = $ys_array[$item_column_name];
            }
        }


        //将字段内容放在返回数组的第一个
        $sqlarray[0] = $header_fields;

        //将数据填充到数组中
        for ($currentRow = 2; $currentRow <= $rows; $currentRow++) {
            foreach ($header_fields as $key => $value) {
                $cell = $currentSheet->getCellByColumnAndRow($key, $currentRow);
                if (PHPExcel_Shared_Date::isDateTime($cell)) {
                    $sqlarray[$currentRow - 1][$value] = (string) PHPExcel_Style_NumberFormat::toFormattedString($cell->getCalculatedValue(), 'YYYY-MM-DD');
                } else {
                    $sqlarray[$currentRow - 1][$value] = $cell->getValue() instanceof PHPExcel_RichText ? $cell->getValue()->getPlainText() : $cell->getValue();
                }
            }
        }


        return $sqlarray;
    }

}
