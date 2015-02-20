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
use Box\Spout\Common\Type;


/**
 * Spreadsheet response
 */
class SpreadsheetResponse extends Nette\Object implements Application\IResponse
{


	/**
	 * Implementace konvertoru do XLS.
	 * @var SpreadsheetProcesor
	 */
	private $procesor;


	/**
	 * Formát souboru. Viz SpreadsheetProcesor.
	 * @var string
	 */
	private $format = 'Excel2007';


	/**
	 * Jméno souboru posílané současně s exportem,
	 * @var string
	 */
	private $filename = 'spreadsheet';


	/**
	 * @var array
	 */
	private $properties = array();


	/**
	 * @var array
	 */
	private $data = array();


	/**
	 * @param array $headers
	 * @param Traversable|array $data
	 * @param SpreadsheetProcesor $procesor
	 */
	function __construct($rows, array $headers = array(), SpreadsheetProcesor $procesor = Null)
	{
		$this->addSheet($rows, $headers);
		if ($procesor) {
			$this->procesor = $procesor;
		}
	}



	/**
	 * Volitelné přiřazení jména souboru.
	 */
	function addSheet($rows, array $headers = array(), $name = Null)
	{
		if (! $rows instanceof Traversable && ! is_array($rows)) {
			throw new InvalidArgumentException('Input rows must be array or Traversable.');
		}

		if (is_array($rows)) {
			$rows = new ArrayIterator($rows);
		}

		$this->data[] = (object) array(
				'headers' => $headers,
				'rows' => $rows,
				'name' => $name,
				);
		return $this;
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
	function setModified(DateTime $m)
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
	 * Název dokumentu a prvního listu.
	 * @param string
	 */
	function setTitle($s)
	{
		$this->properties['title'] = (string)$s;
		$this->data[0]->name = (string)$s;
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

		// When forcing the download of a file over SSL,IE8 and lower browsers fail
		// if the Cache-Control and Pragma headers are not set.
		//
		// @see http://support.microsoft.com/KB/323308
		// @see https://github.com/liuggio/ExcelBundle/issues/45
		$httpResponse->setHeader('Pragma', 'public');

		$this->getProcesor()
				->setProperties($this->properties)
				->echo_($this->data);
	}



	/**
	 * Without headers.
	 */
	function test()
	{
		$this->getProcesor()
				->setProperties($this->properties)
				->setHeaders($this->headers)
				->echo_($this->data);
	}



	// -- PRIVATE ------------------------------------------------------



	private function getProcesor()
	{
		if (empty($this->procesor)) {
			switch (strtolower($this->format)) {
				case 'csv':
					$this->procesor = new SpoutSpreadsheetProcesor(Type::CSV);
					break;
				case 'excel2007':
					$this->procesor = new SpoutSpreadsheetProcesor(Type::XLSX);
					break;
				default:
					$this->procesor = new PHPExcelSpreadsheetProcesor($this->format);
			}
		}
		return $this->procesor;
	}


}
