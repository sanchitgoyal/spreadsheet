<?php
include_once( dirname($_SERVER['DOCUMENT_ROOT']) . '/data/lib/vendor/autoload.php');

use IMHR\Wrappers\Spreadsheet\Column;
use IMHR\Wrappers\Spreadsheet\IMHRWorkbook;
use IMHR\Wrappers\Spreadsheet\Worksheet;

$chars = 'abcdefgh';

$columns = array(
    new Column("col-string", "STRING", 15),
    new Column("col-numbers", "NUMERIC", 15),
    new Column("col-timestamps", "DATETIME_ORACLE", 30)
);

$sheet = new Worksheet("Sheet1");
foreach ($columns as $index => $column) {
    $sheet->addColumn($column);
}

$sheet->setAutoFilter(true);

$workbook = new IMHRWorkbook();
$workbook->addWorksheet($sheet);

for($i=0; $i<1000; $i++)
{
    $workbook->addRow('Sheet1', array(
        str_shuffle($chars),
        rand()%10000,
        date('Y-m-d H:i:s',time()-(rand()%31536000))
    ));
}

$workbook->exportToStdout();