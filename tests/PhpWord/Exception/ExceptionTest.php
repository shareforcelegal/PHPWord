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

namespace Shareforce\PhpWord\Exception;

/**
 * Test class for Shareforce\PhpWord\Exception\Exception
 *
 * @coversDefaultClass \Shareforce\PhpWord\Exception\Exception
 * @runTestsInSeparateProcesses
 */
class ExceptionTest extends \PHPUnit\Framework\TestCase
{
    /**
     * Throw new exception
     *
     * @expectedException \Shareforce\PhpWord\Exception\Exception
     * @covers            \Shareforce\PhpWord\Exception\Exception
     */
    public function testThrowException()
    {
        throw new Exception();
    }
}
