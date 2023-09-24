<?php


namespace IMHR\Wrappers\Spreadsheet;

use IMHR\Wrappers\Spreadsheet\ColumnFormats;

use InvalidArgumentException;

/**
 * Class Column
 * Represents a spreadsheet column
 *
 * @package IMHR\Wrappers\Spreadsheet
 * @access public
 */
class Column
{
    /**
     * @var string $name Name of the column
     */
    protected $name;

    /**
     * @var string $format format of the column
     */
    protected $format;

    /**
     * @var int $width Width of the column
     */
    protected $width;

    /**
     * Column class constructor.
     * @param string $name Name of the column (string + max 255 chars)
     * @param string $format Column data type (string + one of approved types from ColumnFormat class)
     * @param int $width Width of the column; excel usually expects whole numbers for this.
     */
    public function __construct($name, $format = "STRING", $width = 10)
    {
        // Validate name
        if (!static::isValidName($name)) {
            throw new InvalidArgumentException('Not a valid column name');
        }
        // Validate column format
        elseif (!ColumnFormats::isValidFormatName($format)) {
            throw new InvalidArgumentException('Not a valid column format');
        }
        // Validate width
        elseif (!static::isValidWidth($width)) {
            throw new InvalidArgumentException('Not a valid value for the column width');
        }

        // Set format
        $this->format = $format;

        // Set name
        $this->name = $name;

        // Set width
        $this->width = $width;
    }

    // <editor-fold defaultstate="collapsed" desc="Getters">

    /**
     * Returns the name of the column
     * @return string
     */
    public function getName() { return $this->name; }

    /**
     * Returns the format of the column
     * @return string
     */
    public function getFormat() { return $this->format; }

    /**
     * Returns the width of the column
     * @return int
     */
    public function getWidth() { return $this->width; }

    // </editor-fold>

    // <editor-fold defaultstate="collapsed" desc="Validators">

    /**
     * Checks if the name provided in the args is
     * worthy of becoming a column name
     * @param string $name
     * @return bool
     * @throws InvalidArgumentException if column name provided is not a string
     */
    public static function isValidName($name) {
        // Name should be a string
        if(!is_string($name)) {
            throw new InvalidArgumentException('Column name needs to be string');
        }
        // Use trimmed version of the argument
        // because name can be just blank spaces
        $name = trim($name);

        // Check for emptyness
        return !empty($name);
    }

    /**
     * Validates column width
     * @param int $width
     * @return bool
     */
    public static function isValidWidth($width) {
        // Width should be an integer
        if(!is_integer($width)) {
            throw new InvalidArgumentException('Column width needs to be an integer');
        }

        // Column width should be bewteen 1-255
        return $width >= 1 && $width <= 255;
    }

    /**
     * Check if the specified column value matches the expected format
     * @param $value
     * @return bool
     * @throws InvalidArgumentException if format name is not one of the allowed ones
     */
    public function valueCompliesToFormat($value) {
        if(!is_null($value)) {
            return ColumnFormats::checkDataFormat($value, $this->format);
        }
    }

    // </editor-fold>
}
