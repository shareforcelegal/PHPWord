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

use Shareforce\PhpWord\SimpleType\Jc;

/**
 * Test class for Shareforce\PhpWord\Style\Image
 *
 * @coversDefaultClass \Shareforce\PhpWord\Style\Image
 * @runTestsInSeparateProcesses
 */
class ImageTest extends \PHPUnit\Framework\TestCase
{
    /**
     * Test setting style with normal value
     */
    public function testSetGetNormal()
    {
        $object = new Image();

        $properties = array(
            'width'              => 200,
            'height'             => 200,
            'alignment'          => Jc::START,
            'marginTop'          => 240,
            'marginLeft'         => 240,
            'wrappingStyle'      => 'inline',
            'wrapDistanceLeft'   => 10,
            'wrapDistanceRight'  => 20,
            'wrapDistanceTop'    => 30,
            'wrapDistanceBottom' => 40,
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
        $object = new Image();

        $properties = array(
            'width'              => 200,
            'height'             => 200,
            'alignment'          => Jc::START,
            'marginTop'          => 240,
            'marginLeft'         => 240,
            'position'           => 10,
            'positioning'        => \Shareforce\PhpWord\Style\Image::POSITION_ABSOLUTE,
            'posHorizontal'      => \Shareforce\PhpWord\Style\Image::POSITION_HORIZONTAL_CENTER,
            'posVertical'        => \Shareforce\PhpWord\Style\Image::POSITION_VERTICAL_TOP,
            'posHorizontalRel'   => \Shareforce\PhpWord\Style\Image::POSITION_RELATIVE_TO_COLUMN,
            'posVerticalRel'     => \Shareforce\PhpWord\Style\Image::POSITION_RELATIVE_TO_IMARGIN,
            'wrapDistanceLeft'   => 10,
            'wrapDistanceRight'  => 20,
            'wrapDistanceTop'    => 30,
            'wrapDistanceBottom' => 40,
        );
        foreach ($properties as $key => $value) {
            $get = "get{$key}";
            $object->setStyleValue("{$key}", $value);
            $this->assertEquals($value, $object->$get());
        }
    }

    /**
     * Test setWrappingStyle exception
     *
     * @expectedException \InvalidArgumentException
     */
    public function testSetWrappingStyleException()
    {
        $object = new Image();
        $object->setWrappingStyle('foo');
    }
}
