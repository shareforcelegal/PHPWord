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

namespace Shareforce\PhpWord\Shared\Microsoft;

/**
 * Test class for \Shareforce\PhpWord\Shared\Microsoft
 * @coversDefaultClass \Shareforce\PhpWord\Shared\Microsoft
 */
class PasswordEncoderTest extends \PHPUnit\Framework\TestCase
{
    /**
     * Test that a password can be hashed without specifying any additional parameters
     */
    public function testEncodePassword()
    {
        //given
        $password = 'test';

        //when
        $hashPassword = PasswordEncoder::hashPassword($password);

        //then
        $this->assertEquals('M795/MAlmGU8RIsY9Q9uDLHC7bk=', $hashPassword);
    }

    /**
     * Test that a password can be hashed with a custom salt
     */
    public function testEncodePasswordWithSalt()
    {
        //given
        $password = 'test';
        $salt = base64_decode('uq81pJRRGFIY5U+E9gt8tA==');

        //when
        $hashPassword = PasswordEncoder::hashPassword($password, PasswordEncoder::ALGORITHM_SHA_1, $salt);

        //then
        $this->assertEquals('QiDOcpia1YzSVJPiKPwWebl9p/0=', $hashPassword);
    }

    /**
     * Test that the encoder falls back on SHA-1 if a non supported algorithm is given
     */
    public function testDefaultsToSha1IfUnsupportedAlgorithm()
    {
        //given
        $password = 'test';
        $salt = base64_decode('uq81pJRRGFIY5U+E9gt8tA==');

        //when
        $hashPassword = PasswordEncoder::hashPassword($password, PasswordEncoder::ALGORITHM_MAC, $salt);

        //then
        $this->assertEquals('QiDOcpia1YzSVJPiKPwWebl9p/0=', $hashPassword);
    }

    /**
     * Test that the encoder falls back on SHA-1 if a non supported algorithm is given
     */
    public function testEncodePasswordWithNullAsciiCodeInPassword()
    {
        //given
        $password = 'test' . chr(0);
        $salt = base64_decode('uq81pJRRGFIY5U+E9gt8tA==');

        //when
        $hashPassword = PasswordEncoder::hashPassword($password, PasswordEncoder::ALGORITHM_MAC, $salt, 1);

        //then
        $this->assertEquals('rDV9sgdDsztoCQlvRCb1lF2wxNg=', $hashPassword);
    }
}
