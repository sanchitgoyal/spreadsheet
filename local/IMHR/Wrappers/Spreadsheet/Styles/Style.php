<?php


namespace IMHR\Wrappers\Spreadsheet\Styles;

use InvalidArgumentException;

class Style
{
    protected $value = array(
        'font' => Font::ARIAL,
        'font-size' => 8,
        'font-style' => '',
        'border' => '',
        'border-style' => '',
        'border-color' => '',
        'color' => '',
        'fill' => '',
        'halign' => '',
        'valign' => 'top',
        'wrap_text' => false
    );

    public function setFont($value) {
        if(!Font::isValid($value)) {
            throw new \InvalidArgumentException('Not a valid font value');
        }

        $this->value['font'] = $value;
    }

    public function setTextWrap($value) {
        if(!is_bool($value)) {
            throw new \InvalidArgumentException('Not a valid wrap value');
        }
 
        $this->value['wrap_text'] = $value;

       
    }

    public function setFontSize($value) {
        if(!FontSize::isValid($value)) {
            throw new \InvalidArgumentException('Not a valid font size value');
        }

        $this->value['font-size'] = $value;
    }

    public function setFontStyle($values) {
        if(!FontStyle::isValid($values)) {
            throw new InvalidArgumentException('Invalid font style specified');
        }

        $values = is_array($values) ? $values : array($values);

        $this->value['font-style'] = implode(",", $values);
    }

    public function setBorder($values) {
        if(!Border::isValid($values)) {
            throw new InvalidArgumentException('Invalid border value specified');
        }

        $values = is_array($values) ? $values : array($values);

        $this->value['border'] = implode(",", $values);
    }

    public function setBorderStyle($value) {
        if(!BorderStyle::isValid($value)) {
            throw new \InvalidArgumentException('Not a valid border style value');
        }

        $this->value['border-style'] = $value;
    }

    public function setHorizontalAlign($value) {
        if(!HAlign::isValid($value)) {
            throw new \InvalidArgumentException('Not a valid horizontal align value');
        }

        $this->value['halign'] = $value;
    }

    public function setVerticalAlign($value) {
        if(!VAlign::isValid($value)) {
            throw new \InvalidArgumentException('Not a valid Vertical align value');
        }

        $this->value['valign'] = $value;
    }

    public function setColor($value) {
        if(!Color::isValid($value)) {
            throw new \InvalidArgumentException('Not a valid color value');
        }

        $this->value['color'] = $value;
    }

    public function setFill($value) {
        if(!Color::isValid($value)) {
            throw new \InvalidArgumentException('Not a valid color value');
        }

        $this->value['fill'] = $value;
    }

    public function setBorderColor($value) {
        if(!Color::isValid($value)) {
            throw new \InvalidArgumentException('Not a valid color value');
        }

        $this->value['border-color'] = $value;
    }

    public function toArray() {
        return array_filter($this->value);
    }
}
