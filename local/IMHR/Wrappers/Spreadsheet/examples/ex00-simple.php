<?php

include_once( dirname($_SERVER['DOCUMENT_ROOT']) . '/data/lib/vendor/autoload.php');

use IMHR\Wrappers\Spreadsheet\Column;
use IMHR\Wrappers\Spreadsheet\IMHRWorkbook;
use IMHR\Wrappers\Spreadsheet\Worksheet;

$columns = array(
    new Column("c1-text", "STRING"),
    new Column("c2-numeric", "NUMERIC"),
    new Column("c2-numeric", "NUMERIC-1P"),
    new Column("c2-numeric", "NUMERIC-2P"),
    new Column("c3-date", "DATE_ORACLE"),
    new Column("c3-datetime", "DATETIME_ORACLE"),
    new Column("c3-usd", "CURRENCY_US")
);

$sheet = new Worksheet("Sheet1");
foreach ($columns as $index => $column) {
    $sheet->addColumn($column);
}

$workbook = new IMHRWorkbook();
$workbook->addWorksheet($sheet);

$rows = array(
    array('x101',3.1456, 3.1456, 3.1456, '2018-01-07','2018-01-07 14:00:00', '$12.00'),
    array('x201',3.1456,3.1456, 3.1456, '2018-02-07','2018-02-07 13:00:00', '$11.50'),
    array('x301',3.1456,3.1456, 3.1456, '2018-03-07','2018-03-07 12:00:00', '$10.99'),
    array('x401',3.1456,3.1456, 3.1456, '2018-04-07','2018-04-07 11:00:00', '$9'),
    array('x501',3.1456,3.1456, 3.1456, '2018-05-07','2018-05-07 10:00:00', '$8.99'),
    array('x601',3.1456,3.1456, 3.1456, '2018-06-07','2018-06-07 09:00:00', '$7.86'),
    array('x701',3.1456,3.1456, 3.1456, '2018-07-07','2018-07-07 08:00:00', '$6'),
);

foreach ($rows as $index => $row) {
    $workbook->addRow("Sheet1", $row);
}

$workbook->exportToStdout();