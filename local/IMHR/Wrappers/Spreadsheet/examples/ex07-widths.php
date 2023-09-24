<?php

include_once( dirname($_SERVER['DOCUMENT_ROOT']) . '/data/lib/vendor/autoload.php');

use IMHR\Wrappers\Spreadsheet\Column;
use IMHR\Wrappers\Spreadsheet\IMHRWorkbook;
use IMHR\Wrappers\Spreadsheet\Worksheet;

$columns = array(
    new Column("c1", "NUMERIC", 10),
    new Column("c2", "NUMERIC", 20),
    new Column("c3", "NUMERIC", 30),
    new Column("c4", "NUMERIC", 40)
);

$sheet = new Worksheet("Sheet1");
foreach ($columns as $index => $column) {
    $sheet->addColumn($column);
}

$workbook = new IMHRWorkbook();
$workbook->addWorksheet($sheet);
$workbook->addRow("Sheet1", array(300,234,456,789));

$workbook->exportToStdout();