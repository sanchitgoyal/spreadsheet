<?php

namespace IMHR\Wrappers\Spreadsheet;

use IMHR\Wrappers\Spreadsheet\Worksheet;

use DomainException;
use InvalidArgumentException;

/**
 * Class IMHRWorkbook
 * @package IMHR\Wrappers\Spreadsheet
 */
class IMHRWorkbook {
    /**
     * Property to store the spreadsheet title
     * @var string $title
     */
    protected $title;
    /**
     * Property to store the spreadsheet subject
     * @var string $subject
     */
    protected $subject;
    /**
     * Property to store the author of the spreadsheet (default=IMHR)
     * @var string $author
     */
    protected $author = "IMHR";
    /**
     * Property to store the company of the author (default: IMHR)
     * @var string $company
     */
    protected $company = "IMHR";
    /**
     * Property to store the spreadsheet description
     * @var string $description
     */
    protected $description;
    /**
     * We do not store the actual rows in a data structure
     * We store the count of rows written so that we can
     * figure out the next index and keep track of other related metadata
     *
     * @var int $row_count
     */
    protected $row_count = 0;

    /**
     * Property that keeps track of the
     * worksheets contained in this spreadsheet.
     * @var array $sheets
     */
    protected $sheets = array();

    /**
     * Property that does all the work behind the scenes
     * i.e. The base XLSXWriter
     * @var \XLSXWriter $writer
     */
    protected $writer;

    /**
     * IMHRWorkbook constructor.
     */
    public function __construct()
    {
        // Initialize the PHPXSLXWriter
        $this->writer = new \XLSXWriter();

        // Set temp folder for processing
        $this->writer->setTempDir(__DIR__ . '/temp');
    }

    // <editor-fold defaultstate="collapsed" desc="Getters">

    /**
     * Returns the title of the spreadsheet for the public
     * @return string
     */
    public function getTitle()
    {
        return $this->title;
    }

    /**
     * Returns the subject of spreadsheet
     *
     * @return string
     */
    public function getSubject()
    {
        return $this->subject;
    }

    /**
     * Returns the author of the spreadsheet
     * @return string
     */
    public function getAuthor(){
        return $this->author;
    }

    /**
     * Returns the company of the author
     * @return string
     */
    public function getCompany()
    {
        return $this->company;
    }

    /**
     * Returns the description of the spreadsheet
     * @return string
     */
    public function getDescription()
    {
        return $this->description;
    }

    /**
     * Outputs the spreadsheet to stdout
     * By default it adds the http headers to the output
     * But that functionality can be turned off at will
     * @param bool $add_headers
     */
    public function exportToStdout($add_headers = true, $filename = "export.xlsx") {
        // Add headers to the mix if so desired
        if($add_headers) {
            header('Content-type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
            header('Content-disposition: attachment; filename="' . $filename . '"');
            header('Content-Transfer-Encoding: binary');
            header('Cache-Control: must-revalidate');
        }

        // Blurt out the spreadsheet to stdout
        $this->writer->writeToStdOut();
    }

    /**
     * Outputs spreadsheet to a file
     * File is placed in system temp directory
     * @param $filename
     * @return void
     */
    public function exportToFile($filename = 'export.xlsx') {
        $this->writer->writeToFile(sys_get_temp_dir() . '/' . $filename);
    }

    /**
     * Figures out the index of the next row
     * to be written to the spreadsheet
     * @return int
     */
    private function getNextRowIndex() {
        return $this->row_count + 1;
    }

    /**
     * Takes the name of the worksheet and lets you
     * know if that sheet exists in the spreadsheet workbook
     * @param string $sheet_name
     * @return bool
     * @throws InvalidArgumentException
     */
    public function hasWorkSheet($sheet_name) {
        // Check if a valid sheet name has been passed as the argument
        if(!Worksheet::isValidWorksheetName($sheet_name)) {
            throw new InvalidArgumentException('Invalid sheet name provided as argument');
        }

        // Look within the sheet array
        return array_key_exists($sheet_name, $this->sheets);
    }

    /**
     * Takes the name of the worksheet returns
     * it from the storage, i.e. if one exists
     * @param string $sheet_name
     * @return Worksheet
     * @throws InvalidArgumentException
     */
    private function getWorksheetByName($sheet_name) {
        // Validate worksheet name
        if(!Worksheet::isValidWorksheetName($sheet_name)) {
            throw new InvalidArgumentException('Invalid worksheet name specified as argument');
        }
        // Return the sheet if exists
        else return $this->sheets[$sheet_name];
    }

    // </editor-fold>

    // <editor-fold defaultstate="collapsed" desc="Setters">

    /**
     * Sets the title for the spreadsheet
     * @param string $title
     */
    public function setTitle($title)
    {
        if(!$this->isValidTitle($title)) {
            throw new InvalidArgumentException('Invalid spreadsheet title');
        }

        $this->title = $title;
    }

    /**
     * Sets the subject for the spreadhsheet
     * @param string $subject
     * @throws InvalidArgumentException
     */
    public function setSubject($subject)
    {
        // Validate subject passed as arg
        if(!$this->isValidSubject($subject)) {
            throw new InvalidArgumentException('Invalid spreadsheet subject');
        }

        // Set the value
        $this->subject = $subject;
    }

    /**
     * Sets the author for the spreadsheet (default is IMHR)
     * @param string $author
     * @throws InvalidArgumentException
     */
    public function setAuthor($author)
    {
        // Validate author
        if(!$this->isValidAuthor($author)) {
            throw new InvalidArgumentException('Invalid spreadsheet author');
        }

        // Set the value
        $this->author = $author;
    }

    /**
     * Sets the company name of the author (default is IMHR)
     * @param string  $company
     * @throws InvalidArgumentException
     */
    public function setCompany($company)
    {
        // Validate the company name
        if(!$this->isValidCompany($company)) {
            throw new InvalidArgumentException('Invalid spreadsheet company');
        }

        // Set the value
        $this->company = $company;
    }

    /**
     * Sets the description for the spreadsheet
     * @param string $description
     * @throws InvalidArgumentException
     */
    public function setDescription($description)
    {
        // Validate description
        if(!$this->isValidDescription($description)) {
            throw new InvalidArgumentException('Invalid spreadsheet description');
        }

        // Set the value
        $this->description = $description;
    }

    /**
     * Method fills data into the worksheet
     * For efficiency it does not store it in the memory long term
     * Processes it and throws it to base writer for file storage
     * @param string $sheet_name
     * @param array $row
     * @throws InvalidArgumentException
     */
    public function addRow($sheet_name, $row) {
        // Validate worksheet name
        if(!Worksheet::isValidWorksheetName($sheet_name)) {
            throw new InvalidArgumentException('Invalid worksheet name specified as argument');
        }
        // Validate worksheet existence with the spreadsheet collection
        elseif(!$this->hasWorkSheet($sheet_name)) {
            throw new InvalidArgumentException('Worksheet does not exist in the spreadsheet collection');
        }
        // Validate the row data against column formats
        elseif (!$this->isValidRowData($sheet_name, $row)) {
            throw new InvalidArgumentException('Spreadsheet row data did not match column formats');
        }

        // Fetch sheet in context
        $sheet = $this->getWorksheetByName($sheet_name);

        // Initialize temp storage for extracted values
        $values = array();
        // Initialize temp storage for extracted styles
        $styles = array();
        // Start by fetching next writable row index
        $row_index = $this->getNextRowIndex();
        // Start by index 1 for column (mandated by excel)
        $col_index = 1;
        // Process each column one by one
        foreach ($row as $i => $col_value) {
            // Set values
            $values[$col_index-1] = $col_value;

            // Get style corresponding the row, column
            $style = $sheet->getStyle($row_index, $col_index);
            // Make sure that style is atleast a blank array for
            // processing sake, if not set by the user
            $style = empty($style) ? array() : $style->toArray();
            // Store styles after processing
            $styles[$col_index-1] = $style;

            // Go to next column
            $col_index++;
        }

        // Fetch options set on the entire row (if any)
        $options = array(
            'height' => $sheet->getRowHeight($row_index),
            'hidden' => $sheet->isHiddenRow($row_index)
        );

        // This is just to filter out any empty/unset options
        $options = array_filter($options);

        // Merge options with style
        // This is so that we can transform the data in the format the base
        // writer is expecting
        $styles = empty($options) ? $styles : array_merge($styles, $options);

        // Use PHPXLSXWriter to write the data row
        $this->writer->writeSheetRow($sheet_name, $values, $styles);

        // Bump the row index
        $this->row_count++;
    }

    /**
     * This method takes a worksheet object as
     * argument and adds it to the spreadsheet; if not already there
     * Act of adding it finalizes and writes the headers of the worksheet to the file
     *
     * @param \IMHR\Wrappers\Spreadsheet\Worksheet $worksheet
     * @throws InvalidArgumentException
     */
    public function addWorksheet(Worksheet $worksheet) {
        // Validate if argument is a valid worksheet
        if(empty($worksheet) || !($worksheet instanceof Worksheet)) {
            throw new InvalidArgumentException('Worksheet needs to be a valid Worksheet object');
        }

        // Split the worksheet into header parts
        // as expected by the base spreadsheet writer
        $header_parts = $this->convertWorkSheetToHeaderParts($worksheet);

        // Make sure that the header parts are set before touching the base writer
        // It is possible that the header will be empty if the user has decided not
        // to include one in the spreadsheet
        if(is_array($header_parts)) {
            // Writer header via base spreadsheet writer
            $this->writer->writeSheetHeader(
                $header_parts['name'],
                $header_parts['header_types'],
                $header_parts['options']
            );
        }

        // We clone the worksheet here in order to make sure
        // that the worksheet headers arent updated by the user after
        // finalizing them in the first place
        $this->sheets[$worksheet->getName()] = clone $worksheet;
    }

    /**
     * Internal method that converts a worksheet object into header parts
     * @param \IMHR\Wrappers\Spreadsheet\Worksheet $sheet
     * @return array
     */
    private function convertWorkSheetToHeaderParts(Worksheet $sheet) {
        // Initialize return value with emptyness
        $header_parts = array();

        // Initialize processing vars
        $header_types = array();
        $widths = array();

        // Set name part
        $header_parts['name'] = $sheet->getName();

        // Get sheet columns for processing
        $columns = $sheet->getColumns();

        // Loop over columns for processing
        foreach ($columns as $index => $column) {
            // Separate out header types
            $header_types[$column->getName()] = ColumnFormats::getExcelFormat($column->getFormat());

            // Separate out column widths
            $width = $column->getWidth();
            // If a column width has been set
            // then extract that value
            if(!empty($width)) {
                $widths[] = $width;
            }
            // If a column width has not been
            // specified then default to 0
            else {
                $widths[] = 0;
            }
        }

        // Set extracted header types in return var
        $header_parts['header_types'] = $header_types;

        // Set extracted col options in return var
        $header_parts['options'] = array(
            'widths'        => $widths,
            'auto_filter'   => $sheet->getAutoFilter(),
            'freeze_rows'   => $sheet->getFrozenRow(),
            'freeze_columns'=> $sheet->getFrozenColumn()
        );

        // return extracted sheet header parts
        return $header_parts;
    }

    // </editor-fold>

    // <editor-fold defaultstate="collapsed" desc="Validators">
    /**
     * Validates the spreadsheet title
     * @param $title
     * @return bool
     * @throws InvalidArgumentException
     */
    private function isValidTitle($title) {
        // Title should be a string
        if(!is_string($title)) {
            throw new InvalidArgumentException('Spreadsheet title needs to be a string');
        }

        // trim the title before storage in order to
        // avoid extra spaces that may have been introduced by the user
        $title = trim($title);

        // Check for emptyness because title
        // cannot be blank if intended to be set
        return !empty($title);
    }

    /**
     * Method to validate spreadsheet subject
     * @param string $subject
     * @return bool
     * @throws InvalidArgumentException
     */
    private function isValidSubject($subject) {
        // Subject should be a string
        if(!is_string($subject)) {
            throw new InvalidArgumentException('Spreadsheet subject needs to be a string');
        }

        // trim for unnecessary extra spaces
        $subject = trim($subject);

        // Subject cannot be empty if intended to be set
        return !empty($subject);
    }

    /**
     * Validates spreadsheet author
     * @param string $author
     * @return bool
     * @throws InvalidArgumentException
     */
    private function isValidAuthor($author) {
        // Author name should be a string
        if(!is_string($author)) {
            throw new InvalidArgumentException('Spreadsheet author needs to be a string');
        }

        // Trim author name for unneccessary extra spaces
        $author = trim($author);

        // Author name cannot be empty if intended to be set
        return !empty($author);
    }

    /**
     * Validates spreadsheet author's company name
     * @param string $company
     * @return bool
     * @throws InvalidArgumentException
     */
    private function isValidCompany($company) {
        // Company name should be a string
        if(!is_string($company)) {
            throw new InvalidArgumentException('Spreadsheet company needs to be a string');
        }

        // trim for unnecessary extra spaces
        $company = trim($company);

        // Cannot be empty if intended to be set
        return !empty($company);
    }

    /**
     * Validates spreadsheet description
     * @param string $description
     * @return bool
     * @throws InvalidArgumentException
     */
    private function isValidDescription($description) {
        // Should be a string
        if(!is_string($description)) {
            throw new InvalidArgumentException('Spreadsheet description needs to be a string');
        }

        // Trim for unnecessary spaces
        $description = trim($description);

        // Cannot be empty if intended to be set
        return !empty($description);
    }

    /**
     * Validates row data by matching against intended column formats
     * @param string $sheet_name
     * @param array $row
     * @return bool
     * @throws InvalidArgumentException
     */
    private function isValidRowData($sheet_name, array $row) {
        // Validate if the row data is in the form of an array
        if(!is_array($row)) {
            throw new InvalidArgumentException('Spreadsheet row data is expected to be an array');
        }
        else {
            // Fetch the worksheet in context
            $sheet = $this->getWorksheetByName($sheet_name);

            // Validate data against the set sheet column formats
            return $sheet->isValidRowData($row);
        }
    }
    // </editor-fold>
}
