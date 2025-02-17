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

namespace Shareforce\PhpWord\Shared;

/**
 * Test class for Shareforce\PhpWord\Shared\Converter
 *
 * @coversDefaultClass \Shareforce\PhpWord\Shared\Converter
 */
class ConverterTest extends \PHPUnit\Framework\TestCase
{
    /**
     * Test unit conversion functions with various numbers
     */
    public function testUnitConversions()
    {
        $values = array();
        $values[] = 0; // zero value
        $values[] = rand(1, 100) / 100; // fraction number
        $values[] = rand(1, 100); // integer

        foreach ($values as $value) {
            $result = Converter::cmToTwip($value);
            $this->assertEquals($value / 2.54 * 1440, $result);

            $result = Converter::cmToInch($value);
            $this->assertEquals($value / 2.54, $result);

            $result = Converter::cmToPixel($value);
            $this->assertEquals($value / 2.54 * 96, $result);

            $result = Converter::cmToPoint($value);
            $this->assertEquals($value / 2.54 * 72, $result);

            $result = Converter::cmToEmu($value);
            $this->assertEquals(round($value / 2.54 * 96 * 9525), $result);

            $result = Converter::inchToTwip($value);
            $this->assertEquals($value * 1440, $result);

            $result = Converter::inchToCm($value);
            $this->assertEquals($value * 2.54, $result);

            $result = Converter::inchToPixel($value);
            $this->assertEquals($value * 96, $result);

            $result = Converter::inchToPoint($value);
            $this->assertEquals($value * 72, $result);

            $result = Converter::inchToEmu($value);
            $this->assertEquals(round($value * 96 * 9525), $result);

            $result = Converter::pixelToTwip($value);
            $this->assertEquals($value / 96 * 1440, $result);

            $result = Converter::pixelToCm($value);
            $this->assertEquals($value / 96 * 2.54, $result);

            $result = Converter::pixelToPoint($value);
            $this->assertEquals($value / 96 * 72, $result);

            $result = Converter::pixelToEmu($value);
            $this->assertEquals(round($value * 9525), $result);

            $result = Converter::pointToTwip($value);
            $this->assertEquals($value * 20, $result);

            $result = Converter::pointToCm($value);
            $this->assertEquals($value * 0.035277778, $result, '', 0.00001);

            $result = Converter::pointToPixel($value);
            $this->assertEquals($value / 72 * 96, $result);

            $result = Converter::pointToEmu($value);
            $this->assertEquals(round($value / 72 * 96 * 9525), $result);

            $result = Converter::emuToPixel($value);
            $this->assertEquals(round($value / 9525), $result);

            $result = Converter::picaToPoint($value);
            $this->assertEquals($value / 6 * 72, $result, '', 0.00001);

            $result = Converter::degreeToAngle($value);
            $this->assertEquals((int) round($value * 60000), $result);

            $result = Converter::angleToDegree($value);
            $this->assertEquals(round($value / 60000), $result);
        }
    }

    /**
     * Test htmlToRGB()
     */
    public function testHtmlToRGB()
    {
        $flse = false;
        $this->assertEquals(array(255, 153, 221), Converter::htmlToRgb('#FF99DD')); // With #
        $this->assertEquals(array(224, 170, 29), Converter::htmlToRgb('E0AA1D')); // 6 characters
        $this->assertEquals(array(102, 119, 136), Converter::htmlToRgb('678')); // 3 characters
        $this->assertEquals($flse, Converter::htmlToRgb('0F9D')); // 4 characters
        $this->assertEquals(array(0, 0, 0), Converter::htmlToRgb('unknow')); // 6 characters, invalid
        $this->assertEquals(array(139, 0, 139), Converter::htmlToRgb(\Shareforce\PhpWord\Style\Font::FGCOLOR_DARKMAGENTA)); // Constant
    }

    /**
     * Test css size to point
     */
    public function testCssSizeParser()
    {
        $this->assertNull(Converter::cssToPoint('10em'));
        $this->assertEquals(0, Converter::cssToPoint('0'));
        $this->assertEquals(10, Converter::cssToPoint('10pt'));
        $this->assertEquals(7.5, Converter::cssToPoint('10px'));
        $this->assertEquals(720, Converter::cssToPoint('10in'));
        $this->assertEquals(7.2, Converter::cssToPoint('0.1in'));
        $this->assertEquals(120, Converter::cssToPoint('10pc'));
        $this->assertEquals(28.346457, Converter::cssToPoint('10mm'), '', 0.000001);
        $this->assertEquals(283.464567, Converter::cssToPoint('10cm'), '', 0.000001);
        $this->assertEquals(40, Converter::cssToPixel('30pt'));
        $this->assertEquals(1.27, Converter::cssToCm('36pt'));
        $this->assertEquals(127000, Converter::cssToEmu('10pt'));
    }
}
