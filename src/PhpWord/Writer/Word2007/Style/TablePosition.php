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

namespace Shareforce\PhpWord\Writer\Word2007\Style;

/**
 * TablePosition style writer
 */
class TablePosition extends AbstractStyle
{
    /**
     * Write style.
     */
    public function write()
    {
        $style = $this->getStyle();
        if (!$style instanceof \Shareforce\PhpWord\Style\TablePosition) {
            return;
        }

        $values = array();
        $properties = array(
          'leftFromText',
          'rightFromText',
          'topFromText',
          'bottomFromText',
          'vertAnchor',
          'horzAnchor',
          'tblpXSpec',
          'tblpX',
          'tblpYSpec',
          'tblpY',
        );
        foreach ($properties as $property) {
            $method = 'get' . $property;
            if (method_exists($style, $method)) {
                $values[$property] = $style->$method();
            }
        }
        $values = array_filter($values);

        if ($values) {
            $xmlWriter = $this->getXmlWriter();
            $xmlWriter->startElement('w:tblpPr');
            foreach ($values as $property => $value) {
                $xmlWriter->writeAttribute('w:' . $property, $value);
            }
            $xmlWriter->endElement();
        }
    }
}
