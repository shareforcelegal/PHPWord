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

namespace Shareforce\PhpWord\Writer\ODText\Part;

use Shareforce\PhpWord\Writer\ODText;

/**
 * Test class for Shareforce\PhpWord\Writer\ODText\Part\AbstractPart
 *
 * @coversDefaultClass \Shareforce\PhpWord\Writer\ODText\Part\AbstractPart
 */
class AbstractPartTest extends \PHPUnit\Framework\TestCase
{
    /**
     * covers   ::setParentWriter
     * covers   ::getParentWriter
     */
    public function testSetGetParentWriter()
    {
        $object = $this->getMockForAbstractClass('Shareforce\\PhpWord\\Writer\\ODText\\Part\\AbstractPart');
        $object->setParentWriter(new ODText());
        $this->assertEquals(new ODText(), $object->getParentWriter());
    }

    /**
     * covers   ::getParentWriter
     *
     * @expectedException \Exception
     * @expectedExceptionMessage No parent WriterInterface assigned.
     */
    public function testSetGetParentWriterNull()
    {
        $object = $this->getMockForAbstractClass('Shareforce\\PhpWord\\Writer\\ODText\\Part\\AbstractPart');
        $object->getParentWriter();
    }
}
