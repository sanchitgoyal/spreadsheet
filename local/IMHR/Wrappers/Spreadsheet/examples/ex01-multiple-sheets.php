<?php

include_once( dirname($_SERVER['DOCUMENT_ROOT']) . '/data/lib/vendor/autoload.php');

use IMHR\Wrappers\Spreadsheet\Column;
use IMHR\Wrappers\Spreadsheet\IMHRWorkbook;
use IMHR\Wrappers\Spreadsheet\Worksheet;

$workbook = new IMHRWorkbook();

$columns = array(
    new Column("year", "STRING"),
    new Column("month", "STRING"),
    new Column("amount", "NUMERIC"),
    new Column("first_event", "DATETIME_ORACLE"),
    new Column("second_event", "DATE_ORACLE")
);

$sheet1 = new Worksheet("Sheet1");
foreach ($columns as $index => $column) {
    $sheet1->addColumn($column);
}

$workbook->addWorksheet($sheet1);

$data1 = array(
    array('2003','1','-50.5','2010-01-01 23:00:00','2012-12-31'),
    array('2003','=B2', '23.5','2010-01-01 00:00:00','2012-12-31'),
    array('2003',"'=B2", '23.5','2010-01-01 00:00:00','2012-12-31'),
);

foreach ($data1 as $index => $row) {
    $workbook->addRow("Sheet1", $row);
}

$sheet2 = new Worksheet('Sheet2');
$workbook->addWorksheet($sheet2);

$data2 = array(
    array('2003','01','343.12','4000000000'),
    array('2003','02','345.12','2000000000'),
);

foreach ($data2 as $index => $row) {
    $workbook->addRow("Sheet2", $row);
}

$workbook->exportToStdout();