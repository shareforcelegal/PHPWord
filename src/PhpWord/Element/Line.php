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

use Shareforce\PhpWord\Style\Line as LineStyle;

/**
 * Line element
 */
class Line extends AbstractElement
{
    /**
     * Line style
     *
     * @var \Shareforce\PhpWord\Style\Line
     */
    private $style;

    /**
     * Create new line element
     *
     * @param mixed $style
     */
    public function __construct($style = null)
    {
        $this->style = $this->setNewStyle(new LineStyle(), $style);
    }

    /**
     * Get line style
     *
     * @return \Shareforce\PhpWord\Style\Line
     */
    public function getStyle()
    {
        return $this->style;
    }
}
