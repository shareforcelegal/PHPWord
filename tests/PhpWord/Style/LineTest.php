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

namespace Shareforce\PhpWord\Style;

/**
 * Test class for Shareforce\PhpWord\Style\Image
 *
 * @coversDefaultClass \Shareforce\PhpWord\Style\Image
 * @runTestsInSeparateProcesses
 */
class LineTest extends \PHPUnit\Framework\TestCase
{
    /**
     * Test setting style with normal value
     */
    public function testSetGetNormal()
    {
        $object = new Line();

        $properties = array(
            'connectorType' => \Shareforce\PhpWord\Style\Line::CONNECTOR_TYPE_STRAIGHT,
            'beginArrow'    => \Shareforce\PhpWord\Style\Line::ARROW_STYLE_BLOCK,
            'endArrow'      => \Shareforce\PhpWord\Style\Line::ARROW_STYLE_OVAL,
            'dash'          => \Shareforce\PhpWord\Style\Line::DASH_STYLE_LONG_DASH_DOT_DOT,
            'weight'        => 10,
            'color'         => 'red',
        );
        foreach ($properties as $key => $value) {
            $set = "set{$key}";
            $get = "get{$key}";
            $object->$set($value);
            $this->assertEquals($value, $object->$get());
        }
    }

    /**
     * Test setStyleValue method
     */
    public function testSetStyleValue()
    {
        $object = new Line();

        $properties = array(
            'connectorType' => \Shareforce\PhpWord\Style\Line::CONNECTOR_TYPE_STRAIGHT,
            'beginArrow'    => \Shareforce\PhpWord\Style\Line::ARROW_STYLE_BLOCK,
            'endArrow'      => \Shareforce\PhpWord\Style\Line::ARROW_STYLE_OVAL,
            'dash'          => \Shareforce\PhpWord\Style\Line::DASH_STYLE_LONG_DASH_DOT_DOT,
            'weight'        => 10,
            'color'         => 'red',
        );
        foreach ($properties as $key => $value) {
            $get = "get{$key}";
            $object->setStyleValue("{$key}", $value);
            $this->assertEquals($value, $object->$get());
        }
    }

    /**
     * Test set/get flip
     */
    public function testSetGetFlip()
    {
        $expected = true;
        $object = new Line();
        $object->setFlip($expected);
        $this->assertEquals($expected, $object->isFlip());
    }

    /**
     * Test set/get connectorType
     */
    public function testSetGetConnectorType()
    {
        $expected = \Shareforce\PhpWord\Style\Line::CONNECTOR_TYPE_STRAIGHT;
        $object = new Line();
        $object->setConnectorType($expected);
        $this->assertEquals($expected, $object->getConnectorType());
    }

    /**
     * Test set/get weight
     */
    public function testSetGetWeight()
    {
        $expected = 10;
        $object = new Line();
        $object->setWeight($expected);
        $this->assertEquals($expected, $object->getWeight());
    }

    /**
     * Test set/get color
     */
    public function testSetGetColor()
    {
        $expected = 'red';
        $object = new Line();
        $object->setColor($expected);
        $this->assertEquals($expected, $object->getColor());
    }

    /**
     * Test set/get dash
     */
    public function testSetGetDash()
    {
        $expected = \Shareforce\PhpWord\Style\Line::DASH_STYLE_LONG_DASH_DOT_DOT;
        $object = new Line();
        $object->setDash($expected);
        $this->assertEquals($expected, $object->getDash());
    }

    /**
     * Test set/get beginArrow
     */
    public function testSetGetBeginArrow()
    {
        $expected = \Shareforce\PhpWord\Style\Line::ARROW_STYLE_BLOCK;
        $object = new Line();
        $object->setBeginArrow($expected);
        $this->assertEquals($expected, $object->getBeginArrow());
    }

    /**
     * Test set/get endArrow
     */
    public function testSetGetEndArrow()
    {
        $expected = \Shareforce\PhpWord\Style\Line::ARROW_STYLE_CLASSIC;
        $object = new Line();
        $object->setEndArrow($expected);
        $this->assertEquals($expected, $object->getEndArrow());
    }
}
