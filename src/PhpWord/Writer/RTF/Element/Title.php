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

namespace Shareforce\PhpWord\Writer\RTF\Element;

/**
 * Title element RTF writer; extends from text
 *
 * @since 0.11.0
 */
class Title extends Text
{
    protected function getStyles()
    {
        /** @var \Shareforce\PhpWord\Element\Title $element Type hint */
        $element = $this->element;
        $style = $element->getStyle();
        $style = str_replace('Heading', 'Heading_', $style);
        $style = \Shareforce\PhpWord\Style::getStyle($style);
        if ($style instanceof \Shareforce\PhpWord\Style\Font) {
            $this->fontStyle = $style;
            $pstyle = $style->getParagraph();
            if ($pstyle instanceof \Shareforce\PhpWord\Style\Paragraph && $pstyle->hasPageBreakBefore()) {
                $sect = $element->getParent();
                if ($sect instanceof \Shareforce\PhpWord\Element\Section) {
                    $elems = $sect->getElements();
                    if ($elems[0] === $element) {
                        $pstyle = clone $pstyle;
                        $pstyle->setPageBreakBefore(false);
                    }
                }
            }
            $this->paragraphStyle = $pstyle;
        }
    }

    /**
     * Write element
     *
     * @return string
     */
    public function write()
    {
        /** @var \Shareforce\PhpWord\Element\Title $element Type hint */
        $element = $this->element;
        $elementClass = str_replace('\\Writer\\RTF', '', get_class($this));
        if (!$element instanceof $elementClass || !is_string($element->getText())) {
            return '';
        }

        $this->getStyles();

        $content = '';

        $content .= $this->writeOpening();
        $endout = '';
        $style = $element->getStyle();
        if (is_string($style)) {
            $style = str_replace('Heading', '', $style);
            if (is_numeric($style)) {
                $style = (int) $style - 1;
                if ($style >= 0 && $style <= 8) {
                    $content .= '{\\outlinelevel' . $style;
                    $endout = '}';
                }
            }
        }

        $content .= '{';
        $content .= $this->writeFontStyle();
        $content .= $this->writeText($element->getText());
        $content .= '}';
        $content .= $this->writeClosing();
        $content .= $endout;

        return $content;
    }
}
