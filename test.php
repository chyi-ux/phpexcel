<?php

if (isset($_GET['type'])) {
	$type = $_GET['type'];

	require_once dirname(__FILE__) . '/Classes/PHPExcel.php';

	if ($type == 'xls') {
		excel2003();
	}elseif ($type == 'xlsx') {
		excel2007();
	}
}

function excel2003(){

	$objPHPExcel = new PHPExcel();

	$objPHPExcel->setActiveSheetIndex(0);

	$sheet=$objPHPExcel->getActiveSheet();

	$sheet->setCellValue('A1','t');
	$sheet->setCellValue('B2','你好');
	$sheet->setCellValue('C3','s');
	$sheet->setCellValue('D4','哈哈');

	$sheet->setTitle('test');

	$objPHPExcel->setActiveSheetIndex(0);

	// 生成2003excel格式的xls檔案
	header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');

	header('Content-Disposition: attachment;filename="excel2003.xls"');

	$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel5');

	$objWriter->save("php://output");

	exit;
}

function excel2007(){

	$objPHPExcel = new PHPExcel();

	$objPHPExcel->setActiveSheetIndex(0);

	$sheet=$objPHPExcel->getActiveSheet();

	$sheet->setCellValue('A1','t');
	$sheet->setCellValue('B2','你好');
	$sheet->setCellValue('C3','s');
	$sheet->setCellValue('D4','哈哈');

	$sheet->setTitle('test');

	$objPHPExcel->setActiveSheetIndex(0);

	// 生成2007excel格式的xlsx檔案
	header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');

	header('Content-Disposition: attachment;filename="excel2007.xlsx"');

	header('Cache-Control: max-age=0');

	$objWriter = PHPExcel_IOFactory:: createWriter($objPHPExcel, 'Excel2007');

	$objWriter->save( 'php://output');

	exit;
}

function pdf(){

	$objPHPExcel = new PHPExcel();

	$objPHPExcel->setActiveSheetIndex(0);

	$sheet=$objPHPExcel->getActiveSheet();

	$sheet->setCellValue('A1','t');
	$sheet->setCellValue('B2','你好');
	$sheet->setCellValue('C3','s');
	$sheet->setCellValue('D4','哈哈');

	$sheet->setTitle('test');

	$objPHPExcel->setActiveSheetIndex(0);

	header('Content-Type: application/pdf');
	header('Content-Disposition: attachment;filename="test.pdf"');
	header('Cache-Control: max-age=0');
	$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'PDF');
	$objWriter->save('php://output');
	exit;
	// 生成一個pdf檔案
	// $objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'PDF');
	// $objWriter->save('test.pdf');
}

function csv(){

	$objPHPExcel = new PHPExcel();

	$objPHPExcel->setActiveSheetIndex(0);

	$sheet=$objPHPExcel->getActiveSheet();

	$sheet->setCellValue('A1','t');
	$sheet->setCellValue('B2','你好');
	$sheet->setCellValue('C3','s');
	$sheet->setCellValue('D4','哈哈');

	$sheet->setTitle('test');

	$objPHPExcel->setActiveSheetIndex(0);

	$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'CSV')->setDelimiter(',' )  //設定分隔符
	->setEnclosure('"' )    //設定包圍符
	->setLineEnding("\r\n" )//設定行分隔符
	->setSheetIndex(0)      //設定活動表
	->save(str_replace('.php' , '.csv' , __FILE__));
}

?>