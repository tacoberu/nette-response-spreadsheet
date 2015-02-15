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


use Nette,
	Nette\Application,
	Nette\Http\IRequest,
	Nette\Http\IResponse;
use Traversable,
	ArrayIterator,
	InvalidArgumentException,
	DateTime;


/**
 * Spreadsheet response
 */
class SpreadsheetResponse extends Nette\Object implements Application\IResponse
{


	/**
	 * Implementace konvertoru do XLS.
	 * @var XlsProcesor
	 */
	private $procesor;


	/**
	 * Formát souboru. Viz SpreadsheetProcesor.
	 * @var string
	 */
	private $format = Null;


	/**
	 * Jméno souboru posílané současně s exportem,
	 * @var string
	 */
	private $filename;


	/**
	 * @var array
	 */
	private $properties = array();


	/**
	 * @var Traversable
	 */
	private $headers;


	/**
	 * @var Traversable
	 */
	private $data;


	/**
	 * @param Traversable $headers
	 * @param Traversable $data
	 * @param XlsProcesor $procesor
	 */
	function __construct($data, array $headers = array(), XlsProcesor $procesor = Null)
	{
		if (! $data instanceof Traversable && ! is_array($data)) {
			throw new InvalidArgumentException('Input data must be array or Traversable.');
		}

		if (is_array($data)) {
			$data = new ArrayIterator($data);
		}

		$this->headers = $headers;
		$this->data = $data;
		if ($procesor) {
			$this->procesor = $procesor;
		}
	}



	/**
	 * Volitelné přiřazení jména souboru.
	 */
	function setFilename($s)
	{
		$this->filename = (string)$s;
		return $this;
	}



	/**
	 * Formát souboru. Viz SpreadsheetProcesor.
	 */
	function setFormat($s)
	{
		$this->format = (string)$s;
		return $this;
	}



	/**
	 * @param string
	 */
	function setCreator($s)
	{
		$this->properties['creator'] = (string)$s;
		return $this;
	}



	/**
	 * @param string
	 */
	function setCreated(DateTime $m)
	{
		$this->properties['created'] = $m->format('Y-m-d H:i:s');
		return $this;
	}



	/**
	 * @param string
	 */
	function getModified(DateTime $m)
	{
		$this->properties['modified'] = $m->format('Y-m-d H:i:s');
		return $this;
	}



	/**
	 * @param string
	 */
	function setLastModifiedBy($s)
	{
		$this->properties['lastModifiedBy'] = (string)$s;
		return $this;
	}



	/**
	 * @param string
	 */
	function setTitle($s)
	{
		$this->properties['title'] = (string)$s;
		return $this;
	}



	/**
	 * @param string
	 */
	function setSubject($s)
	{
		$this->properties['subject'] = (string)$s;
		return $this;
	}



	/**
	 * @param string
	 */
	function setDescription($s)
	{
		$this->properties['description'] = (string)$s;
		return $this;
	}



	/**
	 * @param string
	 */
	function setKeywords($s)
	{
		$this->properties['keywords'] = (string)$s;
		return $this;
	}



	/**
	 * @param string
	 */
	function setCategory($s)
	{
		$this->properties['category'] = (string)$s;
		return $this;
	}



	/**
	 * @param Nette\Http\IRequest
	 * @param Nette\Http\IResponse
	 */
	function send(IRequest $httpRequest, IResponse $httpResponse)
	{
		$httpResponse->setContentType($this->getProcesor()->getContentType());
		$httpResponse->setHeader('Cache-Control', 'max-age=0');
		$httpResponse->setHeader('Content-Transfer-Encoding', 'binary');

		if ($this->filename) {
			$httpResponse->setHeader('Content-Disposition', 'attachment;filename="' . $this->filename . '.' . $this->getProcesor()->getExtension() . '"');
		}

		$this->getProcesor()
				->setProperties($this->properties)
				->setHeaders($this->headers)
				->echo_($this->data);
	}



	function test()
	{
		$this->getProcesor()->echo_($this->data);
	}



	// -- PRIVATE ------------------------------------------------------



	private function getProcesor()
	{
		if (empty($this->procesor)) {
			$this->procesor = new SpreadsheetProcesor($this->format);
		}
		return $this->procesor;
	}


}
