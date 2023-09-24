<?php


namespace IMHR\Wrappers\Spreadsheet\Styles;

use ReflectionClass;

abstract class StyleTemplate
{
    public static function isValid($value) {
        if(!is_string($value)) {
            throw new \InvalidArgumentException('Value is expected to be a string');
        }

        try {
            $reflection = new ReflectionClass(get_called_class());
        } catch (\ReflectionException $e) {
            throw new \UnexpectedValueException('Unexpected error occured while fetching pre-set style values');
        }
        $constants = $reflection->getConstants();

        foreach ($constants as $name => $value) {
            if(strcmp($value, $value) === 0) {
                return true;
            }
        }

        return false;
    }
}
