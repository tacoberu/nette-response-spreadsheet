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


use Traversable,
	LogicException;
use Box\Spout\Writer\WriterFactory,
	Box\Spout\Common\Type;


/**
 * Naformátuje pole na Excel sešit použitím Spout knihovny. Je o něco méně
 * nenažraná, než PHPExcel, ale zase podporuje jen csv a XLSX.
 */
class SpoutSpreadsheetProcesor implements SpreadsheetProcesor
{

	/** @var array */
	private $headers = array();


	/** @var string */
	private $version = Type::XLSX;


	/** @var string */
	private $filename = 'spreadsheet';


	/**
	 * Constructor injection.
	 */
	function __construct($version = Null)
	{
		if ($version) {
			if (! in_array($version, array(Type::XLSX, Type::CSV))) {
				throw new LogicException("Unsupported type of spreadsheet: `$version'.");
			}
			$this->version = $version;
		}
	}



	function setProperties(array $props = array())
	{
		foreach ($props as $name => $value) {
			$this->{'set' . ucfirst($name)}($value);
		}
		return $this;
	}



	function setHeaders(array $headers = array())
	{
		if (! count($headers)) {
			return $this;
		}

		$this->headers = $headers;
		return $this;
	}



	/**
	 * @return string
	 */
	function getContentType()
	{
		switch ($this->version) {
			case 'Excel2007':
			case Type::XLSX:
				return 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
			case 'CSV':
			case Type::CSV:
				return 'text/csv';
			default:
				throw new LogicException("Unsupported type of sheet: `{$this->version}'.");
		}
	}



	/**
	 * @return string
	 */
	function getExtension()
	{
		switch ($this->version) {
			case 'Excel2007':
			case Type::XLSX:
				return 'xlsx';
			case 'CSV':
			case Type::CSV:
				return 'csv';
			default:
				throw new LogicException("Unsupported type of sheet: `{$this->version}'.");
		}
	}



	/**
	 * Render cell
	 * @param mixed $value
	 * @return string
	 */
	function echo_(Traversable $xs)
	{
		$procesor = $this->getProcesor();
		$procesor->openToBrowser($this->filename);
		if (count($this->headers)) {
			$procesor->addRow($this->headers);
		}
		$this->fill($procesor, $xs);
		$procesor->close();
	}



	// -- PRIVATE ------------------------------------------------------



	private function getProcesor()
	{
		if (empty($this->procesor)) {
			$this->procesor = WriterFactory::create($this->version);
		}
		return $this->procesor;
	}



	/**
	 * Naplnit sešit daty.
	 *
	 * @param Traversable $data - data k zpracování
	 * @return PHPExcel
	 */
	private function fill($sheet, Traversable $data)
	{
		foreach ($data as $row) {
			if (! is_array($row) && ! $row instanceof Traversable) {
				continue;
			}

			$line = [];
			foreach ($row as $name => $cell) {
				$line[] = $this->formatCellValue($name, $cell);
			}
			$sheet->addRow($line);
		}
	}



	private function formatCellValue($name, $cell)
	{
		if (! is_object($cell)) {
			return $cell;
		}

		if (method_exists($cell, '__toString')) {
			return (string)$cell;
		}

		if (method_exists($cell, 'render')) {
			ob_start();
				$cell->render();
			$s = ob_get_contents();
			ob_clean();
			return $s;
		}

		return (string)$cell;
	}



	private function setFilename($filename)
	{
		$this->filename = $filename;
		return $this;
	}



	private function setTitle($_)
	{
		// pass
		return $this;
	}

}
