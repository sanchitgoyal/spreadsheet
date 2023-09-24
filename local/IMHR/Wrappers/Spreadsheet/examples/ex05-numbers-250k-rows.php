<?php

include_once( dirname($_SERVER['DOCUMENT_ROOT']) . '/data/lib/vendor/autoload.php');

use IMHR\Wrappers\Spreadsheet\Column;
use IMHR\Wrappers\Spreadsheet\IMHRWorkbook;
use IMHR\Wrappers\Spreadsheet\Worksheet;

set_time_limit(0);

$columns = array(
    new Column("c1", "NUMERIC"),
    new Column("c2", "NUMERIC"),
    new Column("c3", "NUMERIC"),
    new Column("c4", "NUMERIC")
);

$sheet = new Worksheet("Sheet1");
foreach ($columns as $index => $column) {
    $sheet->addColumn($column);
}

$workbook = new IMHRWorkbook();
$workbook->addWorksheet($sheet);

for ($i = 0; $i < 250000; $i++) {
    $workbook->addRow("Sheet1", array(rand() % 10000, rand() % 10000, rand() % 10000, rand() % 10000));
}

$workbook->exportToStdout();