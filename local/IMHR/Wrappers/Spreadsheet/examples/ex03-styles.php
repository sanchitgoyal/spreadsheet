<?php

include_once( dirname($_SERVER['DOCUMENT_ROOT']) . '/data/lib/vendor/autoload.php');

use IMHR\Wrappers\Spreadsheet\IMHRWorkbook;
use IMHR\Wrappers\Spreadsheet\Worksheet;
use IMHR\Wrappers\Spreadsheet\Styles\Border;
use IMHR\Wrappers\Spreadsheet\Styles\BorderStyle;
use IMHR\Wrappers\Spreadsheet\Styles\Color;
use IMHR\Wrappers\Spreadsheet\Styles\Font;
use IMHR\Wrappers\Spreadsheet\Styles\FontStyle;
use IMHR\Wrappers\Spreadsheet\Styles\Style;

$sheet = new Worksheet("Sheet1");
$workbook = new IMHRWorkbook();

$style = new Style();
$style->setFont(Font::ARIAL);
$style->setFontSize(10);
$style->setFontStyle(FontStyle::BOLD);
$style->setBorder(array(Border::BOTTOM));
$style->setBorderStyle(BorderStyle::THICK);
$style->setColor(Color::GREEN);


$sheet->setStyle(Worksheet::ALL_ROWS, Worksheet::ALL_COLUMNS, $style);
$sheet->setRowHeight(2, 40);
$sheet->hideRow(3);

$workbook->addWorksheet($sheet);

$workbook->addRow("Sheet1", array("Test Value 1", 101));
$workbook->addRow("Sheet1", array("Test Value 2", 201));
$workbook->addRow("Sheet1", array("Test Value 3", 301));
$workbook->addRow("Sheet1", array("Test Value 4", 401));
$workbook->addRow("Sheet1", array("Test Value 5", 501));
$workbook->addRow("Sheet1", array("Test Value 6", 601));

$workbook->exportToStdout();