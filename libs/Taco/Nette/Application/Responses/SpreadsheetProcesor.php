<?php
/**
 * This file is part of the Taco Projects.
 *
 * Copyright (c) 2004, 2013 Martin Takáč (http://martin.takac.name)
 *
 * For the full copyright and license information, please view
 * the file LICENCE that was distributed with this source code.
 *
 * PHP version 5.3
 *
 * @author     Martin Takáč (martin@takac.name)
 */

namespace Taco\Nette\Application\Responses;


use Traversable;


/**
 * Naformátuje pole na Excel sešit.
 */
interface SpreadsheetProcesor
{


	function setProperties(array $props = array());



	function setHeaders(array $headers = array());



	/**
	 * @return string
	 */
	function getContentType();



	/**
	 * @return string
	 */
	function getExtension();



	/**
	 * Render cell
	 * @param mixed $value
	 * @return string
	 */
	function echo_(Traversable $xs);



}
