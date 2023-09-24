<?php


namespace IMHR\Wrappers\Spreadsheet;

use InvalidArgumentException;
use UnexpectedValueException;

/**
 * Class ColumnFormats
 * @package IMHR\Wrappers\Spreadsheet
 */
class ColumnFormats
{
    /**
     * Data structure to store available column formats
     * @var array
     */
    private static $available_formats = array();

    /**
     * Method to initialize static class stuff
     * @throws UnexpectedValueException
     */
    public static function init() {
        // Read format config file
        $config = file_get_contents(__DIR__. "/ColumnFormats.json");
        // Validate if the config file was read OK
        $config = is_string($config) && !empty($config) ? json_decode($config, true) : false;
        // Validate if JSON conversion was successful
        if(JSON_ERROR_NONE !== json_last_error() || $config === false) {
            throw new UnexpectedValueException('Column format config file could not be processed');
        }
        else {
            // Fetch formats from configuration file input
            self::$available_formats = (array) $config;
        }
    }

    /**
     * Method to supply excel format name
     * stored under the internal format key
     * @param string $format
     * @return mixed
     * @throws InvalidArgumentException
     */
    public static function getExcelFormat($format) {
        // Validate format name before processing
        if(!self::isValidFormatName($format)) {
            throw new InvalidArgumentException('Not a valid column format');
        }
        else {
            // Case insensitive processing
            $format = strtoupper($format);
            // Find and return excel format config
            return self::$available_formats[$format]['excel'];
        }
    }

    /**
     * Validates the name of the column format
     * @param string $format
     * @return bool
     * @throws InvalidArgumentException
     */
    public static function isValidFormatName($format) {
        // Check if format is actually a string value
        if(!is_string($format)) {
            throw new InvalidArgumentException('Column format name is expected to be a string');
        }
        else {
            // Case insensitive processing of format name
            $format = strtoupper($format);
            // Format name should exist as array key
            return array_key_exists($format, self::$available_formats);
        }
    }

    /**
     * Validates the format of the data
     * passed as an argument
     * @param $value
     * @param $format
     * @return bool
     * @throws InvalidArgumentException
     */
    public static function checkDataFormat($value, $format) {
        // Validate format name
        if(!self::isValidFormatName($format)) {
            throw new InvalidArgumentException('Not a valid column format');
        }
        else {
            // Case insensite processing of format name
            $format = strtoupper($format);
            // Fetch the format structure
            $format = self::$available_formats[$format];
            // Use regex component to match data format and return accordingly
            return preg_match($format['regex'], $value) === 1;
        }
    }
}

// Initialize column format class.
ColumnFormats::init();
