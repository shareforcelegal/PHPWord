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

namespace Shareforce\PhpWord\Writer\PDF;

use Shareforce\PhpWord\PhpWord;
use Shareforce\PhpWord\Settings;
use Shareforce\PhpWord\Writer\PDF;

/**
 * Test class for Shareforce\PhpWord\Writer\PDF\TCPDF
 *
 * @runTestsInSeparateProcesses
 */
class TCPDFTest extends \PHPUnit\Framework\TestCase
{
    /**
     * Test construct
     */
    public function testConstruct()
    {
        // TCPDF version 6.3.5 doesn't support PHP 5.3, fixed via https://github.com/tecnickcom/TCPDF/pull/197,
        // pending new release.
        if (version_compare(PHP_VERSION, '5.4.0', '<')) {
            return;
        }

        // TCPDF version 6.3.5 doesn't support PHP 8.0, fixed via https://github.com/tecnickcom/TCPDF/pull/293,
        // pending new release.
        if (version_compare(PHP_VERSION, '8.0.0', '>=')) {
            return;
        }

        $file = __DIR__ . '/../../_files/tcpdf.pdf';

        $phpWord = new PhpWord();
        $section = $phpWord->addSection();
        $section->addText('Test 1');

        $rendererName = Settings::PDF_RENDERER_TCPDF;
        $rendererLibraryPath = realpath(PHPWORD_TESTS_BASE_DIR . '/../vendor/tecnickcom/tcpdf');
        Settings::setPdfRenderer($rendererName, $rendererLibraryPath);
        $writer = new PDF($phpWord);
        $writer->save($file);

        $this->assertFileExists($file);

        unlink($file);
    }
}
