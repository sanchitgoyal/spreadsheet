<?php

namespace IMHR\Wrappers\Spreadsheet;

use IMHR\Wrappers\Spreadsheet\Styles\Style;
use InvalidArgumentException;

/**
 * Class Worksheet
 * @package IMHR\Wrappers\Spreadsheet
 */
class Worksheet {
    /**
     * This variable stores worksheet name
     * @var string
     */
    protected $name;

    /**
     * This variable stores the columns configured
     * by the user within the worksheet
     * @var array
     */
    protected $columns = array();

    /**
     * This variable stores the option for turning
     * auto filter on/off for the worksheet
     *
     * @var bool
     */
    protected $auto_filter = false;

    /**
     * This variable stores freeze row option for rows.
     * It takes a single value and freezes that row during rendering
     * @var int
     */
    protected $freeze_rows = null;

    /**
     * This variable stores freeze option for columns.
     * It takes a single value and freezed that column during rendering
     *
     * @var int
     */
    protected $freeze_columns = null;

    /**
     * This array stores user specified
     * display styling for rows, columns and cells
     *
     * @var array
     */
    protected $styles = array();

    /**
     * This array stores display options for rows
     *
     * @var array[]
     */
    protected $row_options = array(
        'hidden' => array(),
        'height' => array()
    );

    /**
     * Excel indexes rows from 1 onwards
     * Number below represents All Rows selection
     */
    const ALL_ROWS = -999;
    /**
     * Excel indexes columns from 1 onwards
     * Number below represents All columns' selection
     */
    const ALL_COLUMNS = -999;

    /**
     * Worksheet constructor.
     * @param string $name
     */
    public function __construct($name)
    {
        // Validate and set the name of the worksheet
        // based on the param specified
        $this->setName($name);
    }

    // <editor-fold defaultstate="collapsed" desc="Getters">

    /**
     * This method returns the name of the worksheet
     * @return string
     */
    public function getName()
    {
        return $this->name;
    }

    /**
     * This method tells which row has been frozen for the worksheet (if any)
     * @return int
     */
    public function getFrozenRow()
    {
        return $this->freeze_rows;
    }

    /**
     * This method tells which column has been frozen for the worksheet (if any)
     * @return int
     */
    public function getFrozenColumn()
    {
        return $this->freeze_columns;
    }

    /**
     * This method tells if the auto filter
     * has been set to appear for the worksheet
     * @return bool
     */
    public function getAutoFilter()
    {
        return $this->auto_filter;
    }

    /**
     * This method returns list of column
     * objects configured in the worksheet
     * @return array
     * @throws \UnexpectedValueException
     */
    public function getColumns()
    {
        // Validate if the internal columns var is an array
        if(is_array($this->columns)) {
            // If yes then return the columns
            return $this->columns;
        }
        else {
            // Otherwise throw an exception
            throw new \UnexpectedValueException('Unexpected value encountered in the list of worksheet columns');
        }

    }

    /**
     * This method returns the count of
     * columns configured in the worksheet
     * @return int
     * @throws \UnexpectedValueException
     */
    public function getColumnCount(){
        // Get columns
        $columns = $this->getColumns();

        // Return count
        return count($columns);
    }

    /**
     * This method verifies whether or not a
     * column is configured in the worksheet
     * @param Column
     * @return bool
     * @throws InvalidArgumentException
     */
    private function hasColumn(Column $search_column) {
        // Verify if the search column is a valid one
        if(empty($search_column) || !($search_column instanceof Column)) {
            throw new InvalidArgumentException('Invalid column specified for the existence check');
        }
        // If the configured columns are empty no need to proceed with the search
        elseif(empty($this->columns) || !is_array($this->columns)) {
            return false;
        }
        else {
            // Search within each column one by one
            foreach ($this->columns as $index => $column) {
                // COmpare column names to compare existence
                // Sheets cannot have multiple columns with same name
                if(strcmp($search_column->getName(), $column->getName()) === 0) {
                    return true;
                }
                else {
                    // Obv. if nothing found then return false
                    return false;
                }
            }
        }
    }

    /**
     * This method returns the style configured on a cell,
     * i.e. if one is configured at all. The user can pass row
     * and columns indexes to fetch the configured style.
     *
     * User can pass ALL_ROWS and ALL_COLUMNS index
     * also to set and retrieve styles that correspond
     * to the cells in the entrie row
     *
     * @param int $row_index
     * @param int $column_index
     * @return Style|null
     * @throws InvalidArgumentException
     */
    public function getStyle($row_index, $column_index) {
        // Validate row index
        if($row_index !== self::ALL_ROWS && !self::isValidRowIndex($row_index)) {
            throw new InvalidArgumentException('Invalid row index passed as argument');
        }
        // Validate column index
        elseif($column_index !== self::ALL_COLUMNS && !self::isValidColumnIndex($column_index)) {
            throw new InvalidArgumentException('Invalid column index passed as argument');
        }

        // 1st preference explicit cell style setting
        if(isset($this->styles[$row_index][$column_index])) {
            return $this->styles[$row_index][$column_index];
        }
        // 2nd preference row style setting
        elseif (isset($this->styles[$row_index][self::ALL_COLUMNS])) {
            return $this->styles[$row_index][self::ALL_COLUMNS];
        }
        // 3rd preference column style setting
        elseif(isset($this->styles[self::ALL_ROWS][$column_index])) {
            return $this->styles[self::ALL_ROWS][$column_index];
        }
        // 4th preference row style setting
        elseif(isset($this->styles[self::ALL_ROWS][self::ALL_COLUMNS])) {
            return $this->styles[self::ALL_ROWS][self::ALL_COLUMNS];
        }
        // default is null
        else return null;
    }

    /**
     * Returns height setting on the row by index
     * @param int $row_index
     * @return bool|mixed
     * @throws InvalidArgumentException
     */
    public function getRowHeight($row_index) {
        // Validate row index
        if($row_index !== self::ALL_ROWS && !self::isValidRowIndex($row_index)) {
            throw new InvalidArgumentException('Invalid row index passed as argument');
        }

        // If height is set explicitly for the row then return it
        if(isset($this->row_options['height'][$row_index])) {
            return $this->row_options['height'][$row_index];
        }
        // Otherwise if a height has been set for all rows then return it instead
        elseif(isset($this->row_options['height'][self::ALL_ROWS])) {
            return $this->row_options['height'][self::ALL_ROWS];
        }
        // Return false, if nothing has been configured at all
        else return false;
    }

    /**
     * This method checks whether the row has been set to hide on render
     * @param int $row_index
     * @return bool
     * @throws InvalidArgumentException
     */
    public function isHiddenRow($row_index) {
        // Validate row index
        if($row_index !== self::ALL_ROWS && !self::isValidRowIndex($row_index)) {
            throw new InvalidArgumentException('Invalid row index passed as argument');
        }

        // Check for the option setting and return boolean
        return array_key_exists($row_index, $this->row_options['hidden']);
    }

    // </editor-fold>

    // <editor-fold defaultstate="collapsed" desc="Setters">

    /**
     * Method sets the name of the worksheet
     * @param string $name
     * @throws InvalidArgumentException
     */
    public function setName($name)
    {
        // Validate name for excel spec
        if(!self::isValidWorksheetName($name)) {
            throw new \InvalidArgumentException('Not a valid sheet name');
        }

        $this->name = $name;
    }

    /**
     * Mark row to freeze on render
     * @param int $index
     * @throws InvalidArgumentException
     */
    public function freezeRow($index) {
        // Check row index validity
        if(!self::isValidRowIndex($index)) {
            throw new \InvalidArgumentException('Invalid row index specified');
        }

        // Set frozen row index
        $this->freeze_rows = $index;
    }

    /**
     * Mark column to freeze on render
     * @param int $index
     * @throws InvalidArgumentException
     */
    public function freezeColumn($index) {
        // Check column index validity
        if(!self::isValidColumnIndex($index)) {
            throw new \InvalidArgumentException('Invalid column index specified');
        }

        // Set column freeze index
        $this->freeze_columns = $index;
    }

    /**
     * Method marks sheet to auto filter content
     * @param bool $auto_filter
     * @throws InvalidArgumentException
     */
    public function setAutoFilter($auto_filter) {
        // Validate if value passed as argument is actually a boolean
        if(!is_bool($auto_filter)) {
            throw new InvalidArgumentException('A boolean value is expected for autofilter setting');
        }

        // Set autofiter value from arg
        $this->auto_filter = $auto_filter;
    }

    /**
     * Method takes column object as argument
     * and adds it to the worksheet
     * @param Column $column
     * @return void
     * @throws InvalidArgumentException
     */
    public function addColumn(Column $column) {
        // Validate column object argument
        if(empty($column) || !($column instanceof Column)) {
            throw new \InvalidArgumentException('Invalid column passed as argument');
        }
        // Check if column already exists
        elseif ($this->hasColumn($column)) {
            throw new \InvalidArgumentException('Column already exists in the sheet');
        }
        else {
            // We need to make sure that the column
            // indexes are integers 1, 2, 3 ...
            // Hence we use incremented column count as the index
            $count = $this->getColumnCount();
            $this->columns[$count] = $column;
        }
    }

    /**
     * Method adds style to cells, rows and columns
     * @param int $row_index
     * @param int $column_index
     * @param Style $style
     * @throws InvalidArgumentException
     */
    public function setStyle($row_index = self::ALL_ROWS, $column_index = self::ALL_COLUMNS, Style $style) {
        // Validate row index
        if($row_index !== self::ALL_ROWS && !self::isValidRowIndex($row_index)) {
            throw new InvalidArgumentException('Invalid row index');
        }
        // Validate column index
        elseif($column_index !== self::ALL_COLUMNS && !self::isValidColumnIndex($column_index)) {
            throw new InvalidArgumentException('Invalid column index');
        }
        // Validate style
        elseif(empty($style) || !($style instanceof Style)) {
            throw new InvalidArgumentException('Invalid style argument specified');
        }
        // Set style in memory
        else {
            $this->styles[$row_index][$column_index] = $style;
        }
    }

    /**
     * Method sets the height for specified row
     * @param $row_index
     * @param int $height
     * @throws InvalidArgumentException
     */
    public function setRowHeight($row_index, $height) {
        // Validate row index
        if($row_index !== self::ALL_ROWS && !self::isValidRowIndex($row_index)) {
            throw new InvalidArgumentException('Invalid row index passed as argument');
        }
        // Validate height value for the row
        if(!self::isValidRowHeight($height)) {
            throw new InvalidArgumentException('Invalid row height passed as argument');
        }

        // If everything checks out then set height
        $this->row_options['height'][$row_index] = $height;
    }

    /**
     * Method sets the hidden option on the specified row
     * @param int $row_index
     * @throws InvalidArgumentException
     */
    public function hideRow($row_index) {
        // Validate row index
        if($row_index !== self::ALL_ROWS && !self::isValidRowIndex($row_index)) {
            throw new InvalidArgumentException('Invalid row index passed as argument');
        }

        // Set row option if everything checks out
        $this->row_options['hidden'][] = $row_index;
    }

    // </editor-fold>

    // <editor-fold defaultstate="collapsed" desc="Validators">

    /**
     * Method validates worksheet name
     * @param string $name
     * @return bool
     * @throws InvalidArgumentException
     */
    public static function isValidWorksheetName($name) {
        // Name should be a string
        if(!is_string($name)) {
            throw new InvalidArgumentException('Worksheet name needs to be a string');
        }

        // Should not be blank
        return !empty($name);
    }

    /**
     * Method validates row index based on excel specs
     * @param int $index
     * @return bool
     * @throws InvalidArgumentException
     */
    public static function isValidRowIndex($index) {
        // Row index should be an integer
        if(!is_integer($index)) {
            throw new InvalidArgumentException('Row index needs to be an integer');
        }

        // Index should be between 1 and 1048577
        return $index > 0 && $index < 1048577;
    }

    /**
     * Method validates column index based on excel specs
     * @param int $index
     * @return bool
     * @throws InvalidArgumentException
     */
    public static function isValidColumnIndex($index) {
        // Index should be an integer
        if(!is_integer($index)) {
            throw new InvalidArgumentException('Column index needs to be an integer');
        }

        // Index should be between 1 and 16385
        return $index > 0 && $index < 16385;
    }

    /**
     * Method validates row data based on column
     * formats configured in the worksheet (if any)
     * @param array $row
     * @return bool
     * @throws InvalidArgumentException
     */
    public function isValidRowData(array $row) {
        // Validate for data being an array
        if(empty($row) || !is_array($row)) {
            throw new InvalidArgumentException('Row data needs to be an array');
        }

        // Fetch the sheets configured columns
        $columns = $this->getColumns();

        // If there is no header set for the sheet
        // Then there is no format to adhere to
        if(!isset($columns)) {
            return true;
        }

        // Take any associative array keys out of picture
        $row = array_values($row);

        $count=0;
        foreach ($row as $index => $value) {
            // Null values need not be validated
            if(is_null($value)) {
                continue;
            }

            // Fetch corresponding column
            $column = $columns[$index];

            // Either a column header should exist for the column data
            // and column data should match the format specified
            if(!empty($columns) && !$column->valueCompliesToFormat($value)) {
                return false;
            }
            // Or at the very least it should be a text value
            elseif(!ColumnFormats::checkDataFormat($value, "STRING")) {
                return false;
            }
            // If its all fine and dandy with this column's data
            // then move on to the next column
            else {
                continue;
            }
        }

        // If everything is good return true
        return true;
    }

    /**
     * Validates the height of a row specified by the user
     * @param int $height
     * @return bool
     * @throws InvalidArgumentException
     */
    public function isValidRowHeight($height) {
        // CHeck if the specified is an integer
        if(!is_integer($height)) {
            throw new InvalidArgumentException('Row height needs to be a valid integer between 0 and 409');
        }
        // Check if height is between 0 and 409 as per the excel specs
        else return $height >= 0 && $height <= 409;
    }

    // </editor-fold>
}
