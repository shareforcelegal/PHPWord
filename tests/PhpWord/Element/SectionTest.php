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

namespace Shareforce\PhpWord\Element;

use Shareforce\PhpWord\PhpWord;
use Shareforce\PhpWord\Style;
use Shareforce\PhpWord\Style\Section as SectionStyle;

/**
 * @covers \Shareforce\PhpWord\Element\Section
 * @coversDefaultClass \Shareforce\PhpWord\Element\Section
 * @runTestsInSeparateProcesses
 */
class SectionTest extends \PHPUnit\Framework\TestCase
{
    public function testConstructorWithDefaultStyle()
    {
        $section = new Section(0);
        $this->assertInstanceOf('Shareforce\\PhpWord\\Style\\Section', $section->getStyle());
    }

    public function testConstructorWithArrayStyle()
    {
        $section = new Section(0, array('orientation' => 'landscape'));
        $style = $section->getStyle();
        $this->assertInstanceOf('Shareforce\\PhpWord\\Style\\Section', $style);
        $this->assertEquals('landscape', $style->getOrientation());
    }

    public function testConstructorWithObjectStyle()
    {
        $style = new SectionStyle();
        $section = new Section(0, $style);
        $this->assertSame($style, $section->getStyle());
    }

    /**
     * @covers ::setStyle
     */
    public function testSetStyle()
    {
        $expected = 'landscape';
        $object = new Section(0);
        $object->setStyle(array('orientation' => $expected, 'foo' => null));
        $this->assertEquals($expected, $object->getStyle()->getOrientation());
    }

    /**
     * @coversNothing
     */
    public function testAddElements()
    {
        $objectSource = __DIR__ . '/../_files/documents/reader.docx';
        $imageSource = __DIR__ . '/../_files/images/PhpWord.png';

        $section = new Section(0);
        $section->setPhpWord(new PhpWord());
        $section->addText(utf8_decode('ä'));
        $section->addLink(utf8_decode('http://äää.com'), utf8_decode('ä'));
        $section->addTextBreak();
        $section->addPageBreak();
        $section->addTable();
        $section->addListItem(utf8_decode('ä'));
        $section->addObject($objectSource);
        $section->addImage($imageSource);
        $section->addTitle(utf8_decode('ä'), 1);
        $section->addTextRun();
        $section->addFootnote();
        $section->addCheckBox(utf8_decode('chkä'), utf8_decode('Contentä'));
        $section->addTOC();

        $elementCollection = $section->getElements();
        $elementTypes = array(
            'Text',
            'Link',
            'TextBreak',
            'PageBreak',
            'Table',
            'ListItem',
            'OLEObject',
            'Image',
            'Title',
            'TextRun',
            'Footnote',
            'CheckBox',
            'TOC',
        );
        $elmCount = 0;
        foreach ($elementTypes as $elementType) {
            $this->assertInstanceOf("Shareforce\\PhpWord\\Element\\{$elementType}", $elementCollection[$elmCount]);
            $elmCount++;
        }
    }

    /**
     * @coversNothing
     * @expectedException \Shareforce\PhpWord\Exception\InvalidObjectException
     */
    public function testAddObjectException()
    {
        $source = __DIR__ . '/_files/xsl/passthrough.xsl';
        $section = new Section(0);
        $section->addObject($source);
    }

    /**
     * Add title with predefined style
     *
     * @coversNothing
     */
    public function testAddTitleWithStyle()
    {
        Style::addTitleStyle(1, array('size' => 14));
        $section = new Section(0);
        $section->setPhpWord(new PhpWord());
        $section->addTitle('Test', 1);
        $elementCollection = $section->getElements();

        $this->assertInstanceOf('Shareforce\\PhpWord\\Element\\Title', $elementCollection[0]);
    }

    /**
     * @covers ::addHeader
     * @covers ::addFooter
     * @covers ::hasDifferentFirstPage
     */
    public function testAddHeaderFooter()
    {
        $object = new Section(0);
        $elements = array('Header', 'Footer');

        foreach ($elements as $element) {
            $method = "add{$element}";
            $this->assertInstanceOf("Shareforce\\PhpWord\\Element\\{$element}", $object->$method());
        }
        $this->assertFalse($object->hasDifferentFirstPage());
    }

    /**
     * @covers ::addHeader
     * @covers ::hasDifferentFirstPage
     */
    public function testHasDifferentFirstPageFooter()
    {
        $object = new Section(1);
        $object->addFooter(Header::FIRST);
        $this->assertTrue($object->hasDifferentFirstPage());
    }

    /**
     * @covers ::addHeader
     * @covers ::hasDifferentFirstPage
     */
    public function testHasDifferentFirstPage()
    {
        $object = new Section(1);
        $header = $object->addHeader();
        $header->setType(Header::FIRST);
        $this->assertTrue($object->hasDifferentFirstPage());
    }

    /**
     * @covers ::addHeader
     * @expectedException \Exception
     * @expectedExceptionMessage Invalid header/footer type.
     */
    public function testAddHeaderException()
    {
        $object = new Section(1);
        $object->addHeader('ODD');
    }

    /**
     * @covers \Shareforce\PhpWord\Element\AbstractContainer::removeElement
     */
    public function testRemoveElementByIndex()
    {
        $section = new Section(1);
        $section->addText('firstText');
        $section->addText('secondText');

        $this->assertEquals(2, $section->countElements());
        $section->removeElement(1);

        $this->assertEquals(1, $section->countElements());
    }

    /**
     * @covers \Shareforce\PhpWord\Element\AbstractContainer::removeElement
     */
    public function testRemoveElementByElement()
    {
        $section = new Section(1);
        $firstText = $section->addText('firstText');
        $secondText = $section->addText('secondText');

        $this->assertEquals(2, $section->countElements());
        $section->removeElement($firstText);

        $this->assertEquals(1, $section->countElements());
        $this->assertEquals($secondText->getElementId(), $section->getElement(1)->getElementId());
    }
}
