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

namespace Shareforce\PhpWord\Writer\RTF\Style;

/**
 * Line numbering style writer
 *
 * @since 0.10.0
 */
class Tab extends AbstractStyle
{
    /**
     * Write style.
     */
    public function write()
    {
        $style = $this->getStyle();
        if (!$style instanceof \Shareforce\PhpWord\Style\Tab) {
            return;
        }
        $tabs = array(
            \Shareforce\PhpWord\Style\Tab::TAB_STOP_RIGHT   => '\tqr',
            \Shareforce\PhpWord\Style\Tab::TAB_STOP_CENTER  => '\tqc',
            \Shareforce\PhpWord\Style\Tab::TAB_STOP_DECIMAL => '\tqdec',
        );
        $content = '';
        if (isset($tabs[$style->getType()])) {
            $content .= $tabs[$style->getType()];
        }
        $content .= '\tx' . round($style->getPosition());

        return $content;
    }
}
