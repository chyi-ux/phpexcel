<?php

require_once dirname(__FILE__) . '/Classes/PHPExcel.php';

$objPHPExcel = new PHPExcel();

$objPHPExcel->setActiveSheetIndex(0);

$sheet=$objPHPExcel->getActiveSheet();

$sheet->setCellValue('A1','t');
$sheet->setCellValue('B2','你好');
$sheet->setCellValue('C3','s');
$sheet->setCellValue('D4','哈哈');

$sheet->setTitle('test');

$objPHPExcel->setActiveSheetIndex(0);

header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');

header('Content-Disposition: attachment;filename="test.xls"');

$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel5');

$objWriter->save("php://output");

exit;

?>