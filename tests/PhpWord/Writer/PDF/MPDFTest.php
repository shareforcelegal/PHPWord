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
 * Test class for Shareforce\PhpWord\Writer\PDF\MPDF
 *
 * @runTestsInSeparateProcesses
 */
class MPDFTest extends \PHPUnit\Framework\TestCase
{
    /**
     * Test construct
     */
    public function testConstruct()
    {
        $file = __DIR__ . '/../../_files/mpdf.pdf';

        $phpWord = new PhpWord();
        $section = $phpWord->addSection();
        $section->addText('Test 1');
        $section->addPageBreak();

        $rendererName = Settings::PDF_RENDERER_MPDF;
        $rendererLibraryPath = realpath(PHPWORD_TESTS_BASE_DIR . '/../vendor/mpdf/mpdf');
        Settings::setPdfRenderer($rendererName, $rendererLibraryPath);
        $writer = new PDF($phpWord);
        $writer->save($file);

        $this->assertFileExists($file);

        unlink($file);
    }
}
