<?php


namespace IMHR\Wrappers\Spreadsheet\Styles;


class FontSize extends StyleTemplate
{
    /**
     * @param $value
     * @return bool
     */
    public static function isValid($value) {
        if(!is_integer($value)) {
            throw new \InvalidArgumentException('Font size needs to be an integer');
        }
        else return $value >= 1 && $value <= 409;
    }
}
