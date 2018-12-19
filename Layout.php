<?php

require './plugins/PHPExcel-1.8/Classes/PHPExcel.php';
$excel = new PHPExcel();

$con = mysqli_connect("localhost", "root", "", "mathematica");
if (!$con) {
    echo mysqli_error($con);
    exit;
}
$row = 2;
$excel->setActiveSheetIndex(0);

$excel->getActiveSheet()
        ->setCellValue('A1', '#')
        ->setCellValue('B1', 'Race Name')
        ->setCellValue('C1', 'Horse name')
        ->setCellValue('I1', 'Race ID')
        ->setCellValue('J1', 'Horse ID')
        ->setCellValue('M1', 'N/V');

$query = mysqli_query($con, "SELECT 
    *
FROM
    horse th
        INNER JOIN
    country tc ON th.country_ID = tc.country_ID
        INNER JOIN
    pora tp ON th.pora_ID = tp.pora_ID");
while ($rows = mysqli_fetch_object($query)) {
    $style = array('Times new roman' => array('size' => 10,'bold' => true,'color' => array('rgb' => 'ff0000')));
    $excel->getActiveSheet()
            ->setCellValue('B' . $row, $rows->country_Name)
            ->setCellValue('C' . $row, $rows->horse_name)
            ->setCellValue('I' . $row, $rows->pora)
            ->setCellValue('J' . $row, $rows->horse_counter);

    $row++;
}
$excel->getActiveSheet()->getStyle('A1:M1')->applyFromArray(
        array(
            'font' => array(
                'bold' => 24,
                'name'  => 'Times new roman',
            )
        )
);



$date = date('Y-m-d H:i:s');

header('Content-Type:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
header('Content-Disposition:attachment;filename=' . $date . 'test.xlsx');
header('Cache-Control:max-age=0');

$file = PHPExcel_IOFactory::createWriter($excel, 'Excel2007');
$file->save('php://output');
?>
