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
	LogicException,
	stdClass;
use Box\Spout\Writer\WriterFactory,
	Box\Spout\Common\Type,
	Box\Spout\Common\Exception\UnsupportedTypeException,
	Box\Spout\Common\Helper\GlobalFunctionsHelper;


/**
 * Naformátuje pole na Excel sešit použitím Spout knihovny. Je o něco méně
 * nenažraná, než PHPExcel, ale zase podporuje jen csv a XLSX.
 */
class SpoutSpreadsheetProcesor implements SpreadsheetProcesor
{

	/** @var WriterInterface */
	private $procesor;


	/** @var string */
	private $version = Type::XLSX;


	/** @var array */
	private $postProcessing = array();


	/**
	 * Constructor injection.
	 */
	function __construct($version = Null, $procesor = Null)
	{
		if ($version) {
			if (! in_array($version, array(Type::XLSX, Type::CSV))) {
				throw new LogicException("Unsupported type of spreadsheet: `$version'.");
			}
			$this->version = $version;
		}
		$this->procesor = $procesor;
	}



	function setProperties(array $props = array())
	{
		foreach ($props as $name => $value) {
			$this->{'set' . ucfirst($name)}($value);
		}
		return $this;
	}



	function setPostProcessing(array $xs)
	{
		$this->postProcessing = $xs;
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
	function echo_(array $sheets)
	{
		$procesor = $this->getProcesor();
		$procesor->openToBrowser('');
		foreach ($sheets as $index => $pack) {
			if ($index > 0) {
				$procesor->addNewSheetAndMakeItCurrent();
			}

			if ($pack->name) {
				//~ $sheet->setTitle($pack->name);
			}

			if (count($pack->headers)) {
				$this->fillHeaders($procesor, $pack->headers);
			}

			$this->fillRows($procesor, $pack->rows);
		}

		$procesor->close();
	}



	// -- PRIVATE ------------------------------------------------------



	private function getProcesor()
	{
		if (empty($this->procesor)) {
			$this->procesor = $this->createProcesor($this->version);
		}
		return $this->procesor;
	}



	/**
	 * This creates an instance of the appropriate writer, given the type of the file to be read
	 *
	 * @param  string $writerType Type of the writer to instantiate
	 * @return \Box\Spout\Writer\CSV|\Box\Spout\Writer\XLSX
	 * @throws \Box\Spout\Common\Exception\UnsupportedTypeException
	 */
	private function createProcesor($writerType)
	{
		$writer = null;

		switch ($writerType) {
			case Type::CSV:
				$writer = new SpoutCSVWriter();
				break;
			case Type::XLSX:
				$writer = new SpoutXLSXWriter();
				break;
			default:
				throw new UnsupportedTypeException('No writers supporting the given type: ' . $writerType);
		}

		$writer->setGlobalFunctionsHelper(new GlobalFunctionsHelper());

		return $writer;
	}



	/**
	 * Naplnit první řádek hlavičkou.
	 * @return int Jaký je následující řádek.
	 */
	private function fillHeaders($sheet, array $headers = array())
	{
		if (! count($headers)) {
			return;
		}
		$sheet->addRow($headers);
	}



	/**
	 * Naplnit sešit daty.
	 *
	 * @param Traversable $data - data k zpracování
	 * @return PHPExcel
	 */
	private function fillRows($sheet, Traversable $data)
	{
		foreach ($data as $row) {
			if (! is_array($row) && ! $row instanceof Traversable && ! $row instanceof stdClass) {
				continue;
			}

			$line = [];
			foreach ($row as $name => $cell) {
				$line[] = $this->applyPostProcessing($this->formatCellValue($name, $cell));
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



	private function applyPostProcessing($val)
	{
		foreach ($this->postProcessing as $method) {
			if (is_string($method)) {
				$val = $method($val);
			}
		}
		return $val;
	}



	private function setTitle($_)
	{
		// pass
		return $this;
	}

}
