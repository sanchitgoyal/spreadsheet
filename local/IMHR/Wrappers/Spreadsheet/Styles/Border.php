<?php


namespace IMHR\Wrappers\Spreadsheet\Styles;

use InvalidArgumentException;

class Border extends StyleTemplate
{
    const LEFT = "left";
    const RIGHT = "right";
    const TOP = "top";
    const BOTTOM = "bottom";

    public static function isValid($values)
    {
        if(empty($values)) {
            throw new InvalidArgumentException("Border style value cannot be blank");
        }

        $values = is_array($values) ? $values : array($values);

        $reflection = new \ReflectionClass(__CLASS__);
        $constants = $reflection->getConstants();

        foreach ($values as $index => $row) {
            if(!in_array($row, $constants)) {
                return false;
            }
        }

        return true;
    }
}