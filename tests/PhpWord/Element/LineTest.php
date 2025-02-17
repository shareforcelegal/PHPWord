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

/**
 * Test class for Shareforce\PhpWord\Element\Line
 *
 * @coversDefaultClass \Shareforce\PhpWord\Element\Line
 * @runTestsInSeparateProcesses
 */
class LineTest extends \PHPUnit\Framework\TestCase
{
    /**
     * Create new instance
     */
    public function testConstruct()
    {
        $oLine = new Line();

        $this->assertInstanceOf('Shareforce\\PhpWord\\Element\\Line', $oLine);
        $this->assertNull($oLine->getStyle());
    }

    /**
     * Get style name
     */
    public function testStyleText()
    {
        $oLine = new Line('lineStyle');

        $this->assertEquals('lineStyle', $oLine->getStyle());
    }

    /**
     * Get style array
     */
    public function testStyleArray()
    {
        $oLine = new Line(
            array(
                'width'            => \Shareforce\PhpWord\Shared\Converter::cmToPixel(14),
                'height'           => \Shareforce\PhpWord\Shared\Converter::cmToPixel(4),
                'positioning'      => 'absolute',
                'posHorizontalRel' => 'page',
                'posVerticalRel'   => 'page',
                'flip'             => true,
                'marginLeft'       => \Shareforce\PhpWord\Shared\Converter::cmToPixel(5),
                'marginTop'        => \Shareforce\PhpWord\Shared\Converter::cmToPixel(3),
                'wrappingStyle'    => \Shareforce\PhpWord\Style\Image::WRAPPING_STYLE_SQUARE,
                'beginArrow'       => \Shareforce\PhpWord\Style\Line::ARROW_STYLE_BLOCK,
                'endArrow'         => \Shareforce\PhpWord\Style\Line::ARROW_STYLE_OVAL,
                'dash'             => \Shareforce\PhpWord\Style\Line::DASH_STYLE_LONG_DASH_DOT_DOT,
                'weight'           => 10,
            )
        );

        $this->assertInstanceOf('Shareforce\\PhpWord\\Style\\Line', $oLine->getStyle());
    }
}
