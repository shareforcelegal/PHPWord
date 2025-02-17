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

namespace Shareforce\PhpWord\Collection;

/**
 * Collection abstract class
 *
 * @since 0.10.0
 */
abstract class AbstractCollection
{
    /**
     * Items
     *
     * @var \Shareforce\PhpWord\Element\AbstractContainer[]
     */
    private $items = array();

    /**
     * Get items
     *
     * @return \Shareforce\PhpWord\Element\AbstractContainer[]
     */
    public function getItems()
    {
        return $this->items;
    }

    /**
     * Get item by index
     *
     * @param int $index
     * @return \Shareforce\PhpWord\Element\AbstractContainer
     */
    public function getItem($index)
    {
        if (array_key_exists($index, $this->items)) {
            return $this->items[$index];
        }

        return null;
    }

    /**
     * Set item.
     *
     * @param int $index
     * @param \Shareforce\PhpWord\Element\AbstractContainer $item
     */
    public function setItem($index, $item)
    {
        if (array_key_exists($index, $this->items)) {
            $this->items[$index] = $item;
        }
    }

    /**
     * Add new item
     *
     * @param \Shareforce\PhpWord\Element\AbstractContainer $item
     * @return int
     */
    public function addItem($item)
    {
        $index = $this->countItems() + 1;
        $this->items[$index] = $item;

        return $index;
    }

    /**
     * Get item count
     *
     * @return int
     */
    public function countItems()
    {
        return count($this->items);
    }
}
