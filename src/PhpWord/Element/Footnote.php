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

use Shareforce\PhpWord\Style\Paragraph;

class Footnote extends AbstractContainer
{
    /**
     * @var string Container type
     */
    protected $container = 'Footnote';

    /**
     * Paragraph style
     *
     * @var string|\Shareforce\PhpWord\Style\Paragraph
     */
    protected $paragraphStyle;

    /**
     * Is part of collection
     *
     * @var bool
     */
    protected $collectionRelation = true;

    /**
     * Create new instance
     *
     * @param string|array|\Shareforce\PhpWord\Style\Paragraph $paragraphStyle
     */
    public function __construct($paragraphStyle = null)
    {
        $this->paragraphStyle = $this->setNewStyle(new Paragraph(), $paragraphStyle);
        $this->setDocPart($this->container);
    }

    /**
     * Get paragraph style
     *
     * @return string|\Shareforce\PhpWord\Style\Paragraph
     */
    public function getParagraphStyle()
    {
        return $this->paragraphStyle;
    }

    /**
     * Get Footnote Reference ID
     *
     * @deprecated 0.10.0
     * @codeCoverageIgnore
     *
     * @return int
     */
    public function getReferenceId()
    {
        return $this->getRelationId();
    }

    /**
     * Set Footnote Reference ID
     *
     * @deprecated 0.10.0
     * @codeCoverageIgnore
     *
     * @param int $rId
     */
    public function setReferenceId($rId)
    {
        $this->setRelationId($rId);
    }
}
