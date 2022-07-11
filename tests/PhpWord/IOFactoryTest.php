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

/**
 * Test class for Shareforce\PhpWord\IOFactory
 *
 * @runTestsInSeparateProcesses
 */
class IOFactoryTest extends \PHPUnit\Framework\TestCase
{
    /**
     * Create existing writer
     */
    public function testExistingWriterCanBeCreated()
    {
        $this->assertInstanceOf(
            'Shareforce\\PhpWord\\Writer\\Word2007',
            IOFactory::createWriter(new PhpWord(), 'Word2007')
        );
    }

    /**
     * Create non-existing writer
     *
     * @expectedException \Shareforce\PhpWord\Exception\Exception
     */
    public function testNonexistentWriterCanNotBeCreated()
    {
        IOFactory::createWriter(new PhpWord(), 'Word2006');
    }

    /**
     * Create existing reader
     */
    public function testExistingReaderCanBeCreated()
    {
        $this->assertInstanceOf(
            'Shareforce\\PhpWord\\Reader\\Word2007',
            IOFactory::createReader('Word2007')
        );
    }

    /**
     * Create non-existing reader
     *
     * @expectedException \Shareforce\PhpWord\Exception\Exception
     */
    public function testNonexistentReaderCanNotBeCreated()
    {
        IOFactory::createReader('Word2006');
    }

    /**
     * Load document
     */
    public function testLoad()
    {
        $file = __DIR__ . '/_files/templates/blank.docx';
        $this->assertInstanceOf(
            'Shareforce\\PhpWord\\PhpWord',
            IOFactory::load($file)
        );
    }
}
