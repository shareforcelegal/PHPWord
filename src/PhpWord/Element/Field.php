<?php
/**
 * This file is part of PHPWord - A pure PHP library for reading and writing
 * word processing documents.
 *
 * PHPWord is free software distributed under the terms of the GNU Lesser
 * General Public License version 3 as published by the Free Software Foundation.
 *
 * For the full copyright and license information, please read the LICENSE
 * file that was distributed with this source code. For the full list of
 * contributors, visit https://github.com/PHPOffice/PHPWord/contributors.
 *
 * @see         https://github.com/PHPOffice/PHPWord
 * @copyright   2010-2018 PHPWord contributors
 * @license     http://www.gnu.org/licenses/lgpl.txt LGPL version 3
 */

namespace Shareforce\PhpWord\Element;

use Shareforce\PhpWord\Style\Font;

/**
 * Field element
 *
 * @since 0.11.0
 * @see  http://www.schemacentral.com/sc/ooxml/t-w_CT_SimpleField.html
 */
class Field extends AbstractElement
{
    /**
     * Field properties and options. Depending on type, a field can have different properties
     * and options
     *
     * @var array
     */
    protected $fieldsArray = array(
        'PAGE' => array(
           'properties' => array(
               'format' => array('Arabic', 'ArabicDash', 'alphabetic', 'ALPHABETIC', 'roman', 'ROMAN'),
           ),
           'options' => array('PreserveFormat'),
        ),
        'NUMPAGES' => array(
           'properties' => array(
               'format' => array('Arabic', 'ArabicDash', 'CardText', 'DollarText', 'Ordinal', 'OrdText',
                   'alphabetic', 'ALPHABETIC', 'roman', 'ROMAN', 'Caps', 'FirstCap', 'Lower', 'Upper', ),
               'numformat' => array('0', '0,00', '#.##0', '#.##0,00', '€ #.##0,00(€ #.##0,00)', '0%', '0,00%'),
           ),
           'options' => array('PreserveFormat'),
        ),
        'DATE' => array(
            'properties' => array(
                'dateformat' => array(
                    /* Generic formats */
                    'yyyy-MM-dd', 'yyyy-MM', 'MMM-yy', 'MMM-yyyy', 'h:mm am/pm', 'h:mm:ss am/pm', 'HH:mm', 'HH:mm:ss',
                    /* Day-Month-Year formats */
                    'dddd d MMMM yyyy', 'd MMMM yyyy', 'd-MMM-yy', 'd MMM. yy',
                    'd-M-yy', 'd-M-yy h:mm', 'd-M-yy h:mm:ss', 'd-M-yy h:mm am/pm', 'd-M-yy h:mm:ss am/pm', 'd-M-yy HH:mm', 'd-M-yy HH:mm:ss',
                    'd/M/yy', 'd/M/yy h:mm', 'd/M/yy h:mm:ss', 'd/M/yy h:mm am/pm', 'd/M/yy h:mm:ss am/pm', 'd/M/yy HH:mm', 'd/M/yy HH:mm:ss',
                    'd-M-yyyy', 'd-M-yyyy h:mm', 'd-M-yyyy h:mm:ss', 'd-M-yyyy h:mm am/pm', 'd-M-yyyy h:mm:ss am/pm', 'd-M-yyyy HH:mm', 'd-M-yyyy HH:mm:ss',
                    'd/M/yyyy', 'd/M/yyyy h:mm', 'd/M/yyyy h:mm:ss', 'd/M/yyyy h:mm am/pm', 'd/M/yyyy h:mm:ss am/pm', 'd/M/yyyy HH:mm', 'd/M/yyyy HH:mm:ss',
                    /* Month-Day-Year formats */
                    'dddd, MMMM d yyyy', 'MMMM d yyyy', 'MMM-d-yy', 'MMM. d yy',
                    'M-d-yy', 'M-d-yy h:mm', 'M-d-yy h:mm:ss', 'M-d-yy h:mm am/pm', 'M-d-yy h:mm:ss am/pm', 'M-d-yy HH:mm', 'M-d-yy HH:mm:ss',
                    'M/d/yy', 'M/d/yy h:mm', 'M/d/yy h:mm:ss', 'M/d/yy h:mm am/pm', 'M/d/yy h:mm:ss am/pm', 'M/d/yy HH:mm', 'M/d/yy HH:mm:ss',
                    'M-d-yyyy', 'M-d-yyyy h:mm', 'M-d-yyyy h:mm:ss', 'M-d-yyyy h:mm am/pm', 'M-d-yyyy h:mm:ss am/pm', 'M-d-yyyy HH:mm', 'M-d-yyyy HH:mm:ss',
                    'M/d/yyyy', 'M/d/yyyy h:mm', 'M/d/yyyy h:mm:ss', 'M/d/yyyy h:mm am/pm', 'M/d/yyyy h:mm:ss am/pm', 'M/d/yyyy HH:mm', 'M/d/yyyy HH:mm:ss',
                ),
            ),
            'options' => array('PreserveFormat', 'LunarCalendar', 'SakaEraCalendar', 'LastUsedFormat'),
        ),
        'MACROBUTTON' => array(
            'properties' => array('macroname' => ''),
        ),
        'XE' => array(
            'properties' => array(),
            'options'    => array('Bold', 'Italic'),
        ),
        'INDEX' => array(
            'properties' => array(),
            'options'    => array('PreserveFormat'),
        ),
        'STYLEREF' => array(
            'properties' => array('StyleIdentifier' => ''),
            'options'    => array('PreserveFormat'),
        ),
        'REF' => array(
            'properties' => array('name' => ''),
            'options'    => array('f', 'h', 'n', 'p', 'r', 't', 'w'),
        ),
    );

    /**
     * Field type
     *
     * @var string
     */
    protected $type;

    /**
     * Field text
     *
     * @var TextRun|string
     */
    protected $text;

    /**
     * Field properties
     *
     * @var array
     */
    protected $properties = array();

    /**
     * Field options
     *
     * @var array
     */
    protected $options = array();

    /**
     * Font style
     *
     * @var string|\Shareforce\PhpWord\Style\Font
     */
    protected $fontStyle;

    /**
     * Set Font style
     *
     * @param string|array|\Shareforce\PhpWord\Style\Font $style
     * @return string|\Shareforce\PhpWord\Style\Font
     */
    public function setFontStyle($style = null)
    {
        if ($style instanceof Font) {
            $this->fontStyle = $style;
        } elseif (is_array($style)) {
            $this->fontStyle = new Font('text');
            $this->fontStyle->setStyleByArray($style);
        } elseif (null === $style) {
            $this->fontStyle = null;
        } else {
            $this->fontStyle = $style;
        }

        return $this->fontStyle;
    }

    /**
     * Get Font style
     *
     * @return string|\Shareforce\PhpWord\Style\Font
     */
    public function getFontStyle()
    {
        return $this->fontStyle;
    }

    /**
     * Create a new Field Element
     *
     * @param string $type
     * @param array $properties
     * @param array $options
     * @param TextRun|string|null $text
     * @param string|array|\Shareforce\PhpWord\Style\Font $fontStyle
     */
    public function __construct($type = null, $properties = array(), $options = array(), $text = null, $fontStyle = null)
    {
        $this->setType($type);
        $this->setProperties($properties);
        $this->setOptions($options);
        $this->setText($text);
        $this->setFontStyle($fontStyle);
    }

    /**
     * Set Field type
     *
     * @param string $type
     *
     * @throws \InvalidArgumentException
     * @return string
     */
    public function setType($type = null)
    {
        if (isset($type)) {
            if (isset($this->fieldsArray[$type])) {
                $this->type = $type;
            } else {
                throw new \InvalidArgumentException("Invalid type '$type'");
            }
        }

        return $this->type;
    }

    /**
     * Get Field type
     *
     * @return string
     */
    public function getType()
    {
        return $this->type;
    }

    /**
     * Set Field properties
     *
     * @param array $properties
     *
     * @throws \InvalidArgumentException
     * @return self
     */
    public function setProperties($properties = array())
    {
        if (is_array($properties)) {
            foreach (array_keys($properties) as $propkey) {
                if (!(isset($this->fieldsArray[$this->type]['properties'][$propkey]))) {
                    throw new \InvalidArgumentException("Invalid property '$propkey'");
                }
            }
            $this->properties = array_merge($this->properties, $properties);
        }

        return $this->properties;
    }

    /**
     * Get Field properties
     *
     * @return array
     */
    public function getProperties()
    {
        return $this->properties;
    }

    /**
     * Set Field options
     *
     * @param array $options
     *
     * @throws \InvalidArgumentException
     * @return self
     */
    public function setOptions($options = array())
    {
        if (is_array($options)) {
            foreach (array_keys($options) as $optionkey) {
                if (!(isset($this->fieldsArray[$this->type]['options'][$optionkey])) && substr($optionkey, 0, 1) !== '\\') {
                    throw new \InvalidArgumentException("Invalid option '$optionkey', possible values are " . implode(', ', $this->fieldsArray[$this->type]['options']));
                }
            }
            $this->options = array_merge($this->options, $options);
        }

        return $this->options;
    }

    /**
     * Get Field properties
     *
     * @return array
     */
    public function getOptions()
    {
        return $this->options;
    }

    /**
     * Set Field text
     *
     * @param string|TextRun $text
     *
     * @throws \InvalidArgumentException
     * @return null|string|TextRun
     */
    public function setText($text = null)
    {
        if (isset($text)) {
            if (is_string($text) || $text instanceof TextRun) {
                $this->text = $text;
            } else {
                throw new \InvalidArgumentException('Invalid text');
            }
        }

        return $this->text;
    }

    /**
     * Get Field text
     *
     * @return string|TextRun
     */
    public function getText()
    {
        return $this->text;
    }
}
