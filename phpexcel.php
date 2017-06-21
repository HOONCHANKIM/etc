<?php 
header('Content-Type: application/vnd.ms-excel');   // Redirect output to a client’s web browser (Excel5)
header('Content-Disposition: attachment;filename=ex.xlsx');
header('Cache-Control: max-age=0');

require_once("./phpexcel/Classes/PHPExcel.php");  /** PHPExcel IMPORT     **/
date_default_timezone_set("Asia/Seoul");                                    /** timezone_setting    **/
/*  Create new PHPExcel object 생성 및 엑셀 메타 정보 지정 시작   */
$objPHPExcel = new PHPExcel();
$objPHPExcelProperties = $objPHPExcel->getProperties();
$objPHPExcelProperties->setCreator("류");
$objPHPExcelProperties->setLastModifiedBy("류님");
$objPHPExcelProperties->setTitle("엑셀다운로드예제");
$objPHPExcelProperties->setSubject("엑셀다운로드예제");
$objPHPExcelProperties->setDescription("엑셀다운로드예제로 만든 엑셀");
$objPHPExcelProperties->setKeywords("PHPExcel사용");
$objPHPExcelProperties->setCategory("PHPExcel");
/*  Create new PHPExcel object 생성 및 엑셀 메타 정보 지정 끝    */

/*  첫번째 SHEET 테이블 제목 입력 시작   */
$objPHPExcelSheetFirst = $objPHPExcel->setActiveSheetIndex(0);
$objPHPExcelSheetFirst->setCellValue("A1", "이름");
$objPHPExcelSheetFirst->setCellValue("B1", "체중");
$objPHPExcelSheetFirst->setCellValue("C1", "키");
$objPHPExcelSheetFirst->setCellValue("D1", "나이");
$objPHPExcelSheetFirst->setCellValue("E1", "취미");
$objPHPExcelSheetFirst->setCellValue("F1", "기호식품");
$objPHPExcelSheetFirst->setCellValue("G1", "집주소");
/*  첫번째 SHEET 테이블 제목 입력 끝     */


/* 실제 엑셀 내용을 만드는 부분 시작   */
/*  첫번째 시트 내용을 만드는 부분 시작 */

$objPHPExcelSheetFirst->setCellValue("A2", "김철수");
$objPHPExcelSheetFirst->setCellValue("B2", "80kg");
$objPHPExcelSheetFirst->setCellValue("C2", "180cm");
$objPHPExcelSheetFirst->setCellValue("D2", "29살");
$objPHPExcelSheetFirst->setCellValue("E2", "기타연주");
$objPHPExcelSheetFirst->setCellValue("F2", "담배");
$objPHPExcelSheetFirst->setCellValue("G2", "서울");

$objPHPExcelSheetFirst->setCellValue("A3", "최영희");
$objPHPExcelSheetFirst->setCellValue("B3", "50kg");
$objPHPExcelSheetFirst->setCellValue("C3", "160cm");
$objPHPExcelSheetFirst->setCellValue("D3", "28살");
$objPHPExcelSheetFirst->setCellValue("E3", "여행");
$objPHPExcelSheetFirst->setCellValue("F3", "홍차");
$objPHPExcelSheetFirst->setCellValue("G3", "부산");

/*   첫번째 시트 내용을 만드는 부분 끝   */


/*  실제 엑셀 내용을 만드는 부분 시작    */

/*  엑셀 스타일 지정 시작   */
$columnList = array('A','B','C','D','E','F','G');

/* 엑셀 열에 넓이를 지정 시작 */
$objPHPExcelSheetFirst->getColumnDimension('A')->setWidth(14.71);
$objPHPExcelSheetFirst->getColumnDimension('B')->setWidth(18.5);
$objPHPExcelSheetFirst->getColumnDimension('C')->setWidth(30);
$objPHPExcelSheetFirst->getColumnDimension('D')->setWidth(45);
/*  엑셀 열에 넓이를 지정 끝   */

// 헤드에 들어갈 스타일
$style_header = array(
		'fill' => array(
				'type' => PHPExcel_Style_Fill::FILL_SOLID,
				'color' => array('rgb'=>'b4b4b4'),
		),
		'font' => array(
				'bold' => true,
		),
		'borders' => array(
				'outline' => array( 'style' => PHPExcel_Style_Border::BORDER_THIN )
		)
);
// 바디에 들어갈 스타일
$style_content = array(
		'borders' => array(
				'outline' => array( 'style' => PHPExcel_Style_Border::BORDER_THIN )
		)
);

foreach ($columnList as $column){
	for ($idx = 1; $idx < 4; $idx++) {
		if($idx == 1){
			$objPHPExcelSheetFirst->getStyle($column.(string)$idx)->applyFromArray( $style_header );
		}else{
			$objPHPExcelSheetFirst->getStyle($column.(string)$idx)->applyFromArray( $style_content );
		}
	}
}
/*  엑셀 스타일 지정 끝   */

$objPHPExcelSheetFirst ->setTitle("sheet1");
$objPHPExcel->setActiveSheetIndex(0);
$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007');
$objWriter->save('php://output');
exit;
?>