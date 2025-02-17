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

namespace Shareforce\PhpWord\Writer\HTML\Style;

use Shareforce\PhpWord\Style\AbstractStyle as Style;

/**
 * Style writer
 *
 * @since 0.10.0
 */
abstract class AbstractStyle
{
    /**
     * Parent writer
     *
     * @var \Shareforce\PhpWord\Writer\AbstractWriter
     */
    private $parentWriter;

    /**
     * Style
     *
     * @var array|\Shareforce\PhpWord\Style\AbstractStyle
     */
    private $style;

    /**
     * Write style
     */
    abstract public function write();

    /**
     * Create new instance
     *
     * @param array|\Shareforce\PhpWord\Style\AbstractStyle $style
     */
    public function __construct($style = null)
    {
        $this->style = $style;
    }

    /**
     * Set parent writer.
     *
     * @param \Shareforce\PhpWord\Writer\AbstractWriter $writer
     */
    public function setParentWriter($writer)
    {
        $this->parentWriter = $writer;
    }

    /**
     * Get parent writer
     *
     * @return \Shareforce\PhpWord\Writer\AbstractWriter
     */
    public function getParentWriter()
    {
        return $this->parentWriter;
    }

    /**
     * Get style
     *
     * @return array|\Shareforce\PhpWord\Style\AbstractStyle $style
     */
    public function getStyle()
    {
        if (!$this->style instanceof Style && !is_array($this->style)) {
            return '';
        }

        return $this->style;
    }

    /**
     * Takes array where of CSS properties / values and converts to CSS string
     *
     * @param array $css
     * @return string
     */
    protected function assembleCss($css)
    {
        $pairs = array();
        $string = '';
        foreach ($css as $key => $value) {
            if ($value != '') {
                $pairs[] = $key . ': ' . $value;
            }
        }
        if (!empty($pairs)) {
            $string = implode('; ', $pairs) . ';';
        }

        return $string;
    }

    /**
     * Get value if ...
     *
     * @param bool|null $condition
     * @param string $value
     * @return string
     */
    protected function getValueIf($condition, $value)
    {
        return $condition == true ? $value : '';
    }
}
