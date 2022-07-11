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

use Shareforce\PhpWord\Exception\Exception;
use Shareforce\PhpWord\Reader\ReaderInterface;
use Shareforce\PhpWord\Writer\WriterInterface;

abstract class IOFactory
{
    /**
     * Create new writer
     *
     * @param PhpWord $phpWord
     * @param string $name
     *
     * @throws \Shareforce\PhpWord\Exception\Exception
     *
     * @return WriterInterface
     */
    public static function createWriter(PhpWord $phpWord, $name = 'Word2007')
    {
        if ($name !== 'WriterInterface' && !in_array($name, array('ODText', 'RTF', 'Word2007', 'HTML', 'PDF'), true)) {
            throw new Exception("\"{$name}\" is not a valid writer.");
        }

        $fqName = "Shareforce\\PhpWord\\Writer\\{$name}";

        return new $fqName($phpWord);
    }

    /**
     * Create new reader
     *
     * @param string $name
     *
     * @throws Exception
     *
     * @return ReaderInterface
     */
    public static function createReader($name = 'Word2007')
    {
        return self::createObject('Reader', $name);
    }

    /**
     * Create new object
     *
     * @param string $type
     * @param string $name
     * @param \Shareforce\PhpWord\PhpWord $phpWord
     *
     * @throws \Shareforce\PhpWord\Exception\Exception
     *
     * @return \Shareforce\PhpWord\Writer\WriterInterface|\Shareforce\PhpWord\Reader\ReaderInterface
     */
    private static function createObject($type, $name, $phpWord = null)
    {
        $class = "Shareforce\\PhpWord\\{$type}\\{$name}";
        if (class_exists($class) && self::isConcreteClass($class)) {
            return new $class($phpWord);
        }
        throw new Exception("\"{$name}\" is not a valid {$type}.");
    }

    /**
     * Loads PhpWord from file
     *
     * @param string $filename The name of the file
     * @param string $readerName
     * @return \Shareforce\PhpWord\PhpWord $phpWord
     */
    public static function load($filename, $readerName = 'Word2007')
    {
        /** @var \Shareforce\PhpWord\Reader\ReaderInterface $reader */
        $reader = self::createReader($readerName);

        return $reader->load($filename);
    }

    /**
     * Check if it's a concrete class (not abstract nor interface)
     *
     * @param string $class
     * @return bool
     */
    private static function isConcreteClass($class)
    {
        $reflection = new \ReflectionClass($class);

        return !$reflection->isAbstract() && !$reflection->isInterface();
    }
}
