<?php


namespace IMHR\Wrappers\Spreadsheet\Styles;


use InvalidArgumentException;

class FontStyle extends StyleTemplate
{
    const BOLD = "bold";
    const ITALIC = "italic";
    const UNDERLINE = "underline";
    const STRIKETHROUGH = "strikethrough";

    public static function isValid($value)
    {
        if(empty($value)) {
            throw new InvalidArgumentException("Font style cannot be blank");
        }

        $value = is_array($value) ? $value : array($value);

        $reflection = new \ReflectionClass(__CLASS__);
        $constants = $reflection->getConstants();

        foreach ($value as $index => $row) {
            if(!in_array($row, $constants)) {
                return false;
            }
        }

        return true;
    }
}
