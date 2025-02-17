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

namespace Shareforce\PhpWord\Writer;

use Shareforce\PhpWord\AbstractWebServerEmbeddedTest;
use Shareforce\PhpWord\PhpWord;
use Shareforce\PhpWord\SimpleType\Jc;
use Shareforce\PhpWord\TestHelperDOCX;

/**
 * Test class for Shareforce\PhpWord\Writer\Word2007
 *
 * @runTestsInSeparateProcesses
 */
class Word2007Test extends AbstractWebServerEmbeddedTest
{
    /**
     * Tear down after each test
     */
    public function tearDown()
    {
        TestHelperDOCX::clear();
    }

    /**
     * Construct
     */
    public function testConstruct()
    {
        $object = new Word2007(new PhpWord());

        $writerParts = array(
            'ContentTypes' => 'ContentTypes',
            'Rels'         => 'Rels',
            'DocPropsApp'  => 'DocPropsApp',
            'Document'     => 'Document',
            'Styles'       => 'Styles',
            'Numbering'    => 'Numbering',
            'Settings'     => 'Settings',
            'WebSettings'  => 'WebSettings',
            'Header'       => 'Header',
            'Footer'       => 'Footer',
            'Footnotes'    => 'Footnotes',
            'Endnotes'     => 'Footnotes',
        );
        foreach ($writerParts as $part => $type) {
            $this->assertInstanceOf(
                "Shareforce\\PhpWord\\Writer\\Word2007\\Part\\{$type}",
                $object->getWriterPart($part)
            );
            $this->assertInstanceOf(
                'Shareforce\\PhpWord\\Writer\\Word2007',
                $object->getWriterPart($part)->getParentWriter()
            );
        }
    }

    /**
     * Save
     */
    public function testSave()
    {
        $localImage = __DIR__ . '/../_files/images/earth.jpg';
        $remoteImage = self::getRemoteGifImageUrl();
        $phpWord = new PhpWord();
        $phpWord->addFontStyle('Font', array('size' => 11));
        $phpWord->addParagraphStyle('Paragraph', array('alignment' => Jc::CENTER));
        $section = $phpWord->addSection();
        $section->addText('Test 1', 'Font', 'Paragraph');
        $section->addTextBreak();
        $section->addText('Test 2');
        $section = $phpWord->addSection();
        $textrun = $section->addTextRun();
        $textrun->addText('Test 3');
        $footnote = $textrun->addFootnote();
        $footnote->addLink('https://github.com/PHPOffice/PHPWord');
        $header = $section->addHeader();
        $header->addImage($localImage);
        $footer = $section->addFooter();
        $footer->addImage($remoteImage);

        $writer = new Word2007($phpWord);
        $file = __DIR__ . '/../_files/temp.docx';
        $writer->save($file);

        $this->assertFileExists($file);

        unlink($file);
    }

    /**
     * Save using disk caching
     */
    public function testSaveUseDiskCaching()
    {
        $phpWord = new PhpWord();
        $section = $phpWord->addSection();
        $section->addText('Test');
        $footnote = $section->addFootnote();
        $footnote->addText('Test');

        $writer = new Word2007($phpWord);
        $writer->setUseDiskCaching(true);
        $file = __DIR__ . '/../_files/temp.docx';
        $writer->save($file);

        $this->assertFileExists($file);

        unlink($file);
    }

    /**
     * Check content types
     */
    public function testCheckContentTypes()
    {
        $images = array(
            'mars_noext_jpg'    => '1.jpg',
            'mars.jpg'          => '2.jpg',
            'mario.gif'         => '3.gif',
            'firefox.png'       => '4.png',
            'duke_nukem.bmp'    => '5.bmp',
            'angela_merkel.tif' => '6.tif',
        );
        $phpWord = new PhpWord();
        $section = $phpWord->addSection();
        foreach ($images as $source => $target) {
            $section->addImage(__DIR__ . "/../_files/images/{$source}");
        }

        $doc = TestHelperDOCX::getDocument($phpWord);
        $mediaPath = $doc->getPath() . '/word/media';

        foreach ($images as $source => $target) {
            $this->assertFileEquals(
                __DIR__ . "/../_files/images/{$source}",
                $mediaPath . "/section_image{$target}"
            );
        }
    }

    /**
     * Get writer part return null value
     */
    public function testGetWriterPartNull()
    {
        $object = new Word2007();
        $this->assertNull($object->getWriterPart());
    }

    /**
     * Set/get use disk caching
     */
    public function testSetGetUseDiskCaching()
    {
        $this->setOutputCallback(function () {
        });
        $phpWord = new PhpWord();
        $phpWord->addSection();
        $object = new Word2007($phpWord);
        $object->setUseDiskCaching(true, PHPWORD_TESTS_BASE_DIR);
        $writer = new Word2007($phpWord);
        $writer->save('php://output');

        $this->assertTrue($object->isUseDiskCaching());
    }

    /**
     * Use disk caching exception
     *
     * @expectedException \Shareforce\PhpWord\Exception\Exception
     */
    public function testSetUseDiskCachingException()
    {
        $dir = implode(DIRECTORY_SEPARATOR, array(PHPWORD_TESTS_BASE_DIR, 'foo'));

        $object = new Word2007();
        $object->setUseDiskCaching(true, $dir);
    }
}
