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

namespace Shareforce\PhpWord;

use Shareforce\PhpWord\SimpleType\Jc;

/**
 * Test class for Shareforce\PhpWord\Style
 *
 * @coversDefaultClass \Shareforce\PhpWord\Style
 * @runTestsInSeparateProcesses
 */
class StyleTest extends \PHPUnit\Framework\TestCase
{
    /**
     * Add and get paragraph, font, link, title, and table styles
     *
     * @covers ::addParagraphStyle
     * @covers ::addFontStyle
     * @covers ::addLinkStyle
     * @covers ::addNumberingStyle
     * @covers ::addTitleStyle
     * @covers ::addTableStyle
     * @covers ::setDefaultParagraphStyle
     * @covers ::countStyles
     * @covers ::getStyle
     * @covers ::resetStyles
     * @covers ::getStyles
     * @test
     */
    public function testStyles()
    {
        $paragraph = array('alignment' => Jc::CENTER);
        $font = array('italic' => true, '_bold' => true);
        $table = array('bgColor' => 'CCCCCC');
        $numbering = array(
            'type'   => 'multilevel',
            'levels' => array(
                array(
                    'start'     => 1,
                    'format'    => 'decimal',
                    'restart'   => 1,
                    'suffix'    => 'space',
                    'text'      => '%1.',
                    'alignment' => Jc::START,
                ),
            ),
        );

        $styles = array(
            'Paragraph' => 'Paragraph',
            'Font'      => 'Font',
            'Link'      => 'Font',
            'Table'     => 'Table',
            'Heading_1' => 'Font',
            'Normal'    => 'Paragraph',
            'Numbering' => 'Numbering',
        );

        Style::addParagraphStyle('Paragraph', $paragraph);
        Style::addFontStyle('Font', $font);
        Style::addLinkStyle('Link', $font);
        Style::addNumberingStyle('Numbering', $numbering);
        Style::addTitleStyle(1, $font);
        Style::addTableStyle('Table', $table);
        Style::setDefaultParagraphStyle($paragraph);

        $this->assertCount(count($styles), Style::getStyles());
        foreach ($styles as $name => $style) {
            $this->assertInstanceOf("Shareforce\\PhpWord\\Style\\{$style}", Style::getStyle($name));
        }
        $this->assertNull(Style::getStyle('Unknown'));

        Style::resetStyles();
        $this->assertCount(0, Style::getStyles());
    }

    /**
     * Test default paragraph style
     *
     * @covers ::setDefaultParagraphStyle
     * @test
     */
    public function testDefaultParagraphStyle()
    {
        $paragraph = array('alignment' => Jc::CENTER);

        Style::setDefaultParagraphStyle($paragraph);

        $this->assertInstanceOf('Shareforce\\PhpWord\\Style\\Paragraph', Style::getStyle('Normal'));
    }
}
