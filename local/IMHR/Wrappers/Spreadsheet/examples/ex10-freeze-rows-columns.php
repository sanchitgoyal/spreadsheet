<?php

include_once( dirname($_SERVER['DOCUMENT_ROOT']) . '/data/lib/vendor/autoload.php');

use IMHR\Wrappers\Spreadsheet\Column;
use IMHR\Wrappers\Spreadsheet\IMHRWorkbook;
use IMHR\Wrappers\Spreadsheet\Worksheet;

$columns = array(
    new Column("c1", "STRING"),
    new Column("c2", "NUMERIC"),
    new Column("c3", "NUMERIC"),
    new Column("c4", "NUMERIC"),
    new Column("c5", "NUMERIC")
);

$sheet = new Worksheet("Sheet1");
foreach ($columns as $index => $column) {
    $sheet->addColumn($column);
}

$sheet->freezeColumn(1);
$sheet->freezeRow(1);

$workbook = new IMHRWorkbook();

$workbook->addWorksheet($sheet);

$chars = 'abcdefgh';

for ($i = 0; $i < 250; $i++) {
    $workbook->addRow('Sheet1', array(
        str_shuffle($chars),
        rand() % 10000,
        rand() % 10000,
        rand() % 10000,
        rand() % 10000
    ));
}

$workbook->exportToStdout();