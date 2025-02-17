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

use PHPUnit\Framework\Assert;

/**
 * Test class for Shareforce\PhpWord\Style\Language
 *
 * @coversDefaultClass \Shareforce\PhpWord\Style\Language
 */
class LanguageTest extends \PHPUnit\Framework\TestCase
{
    /**
     * Test get/set
     */
    public function testGetSetProperties()
    {
        $object = new Language();
        $properties = array(
            'latin'          => array(null, 'fr-BE'),
            'eastAsia'       => array(null, 'ja-JP'),
            'bidirectional'  => array(null, 'ar-SA'),
            'langId'         => array(null, 1036),
        );
        foreach ($properties as $property => $value) {
            list($default, $expected) = $value;
            $get = "get{$property}";
            $set = "set{$property}";

            $this->assertEquals($default, $object->$get()); // Default value

            $object->$set($expected);

            $this->assertEquals($expected, $object->$get()); // New value
        }
    }

    /**
     * Test throws exception if wrong locale is given
     *
     * @expectedException \InvalidArgumentException
     */
    public function testWrongLanguage()
    {
        $language = new Language();
        $language->setLatin('fra');
    }

    /**
     * Tests that a language can be set with just a 2 char code
     */
    public function testShortLanguage()
    {
        //when
        $language = new Language();
        $language->setLatin('fr');

        //then
        Assert::assertEquals('fr-FR', $language->getLatin());
    }
}
