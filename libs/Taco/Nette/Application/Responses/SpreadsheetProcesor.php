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
	PHPExcel,
	PHPExcel_IOFactory,
	LogicException;


/**
 * Naformátuje pole na Excel sešit.
 */
class SpreadsheetProcesor
{

	/** @var int */
	private $firstRow = 1;


	/** @var string */
	private $version = 'Excel5';


	/**
	 * Constructor injection.
	 */
	function __construct($version = Null)
	{
		if ($version) {
			$this->version = $version;
		}
	}



	function setProperties(array $props = array())
	{
		$obj = $this->getProcesor()->getProperties();
		foreach ($props as $name => $value) {
			$obj->{'set' . ucfirst($name)}($value);
		}
		return $this;
	}



	function setHeaders(array $headers = array())
	{
		if (! count($headers)) {
			return $this;
		}

		$procesor = $this->getProcesor();
		$procesor->setActiveSheetIndex(0);
		$sheet = $procesor->getActiveSheet();

		$rowSymbol = $this->firstRow;
		$columnSymbol = 'A';
		foreach ($headers as $name) {
			$cell = $sheet->setCellValue($columnSymbol . $rowSymbol, $name, True);
			//~ $cell->getStyle()->getFont()->setBold(True);
			$columnSymbol++;
		}

		$this->firstRow++;
		return $this;
	}



	/**
	 * @return string
	 */
	function getContentType()
	{
		switch ($this->version) {
			case 'Excel2007':
				return 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
			case 'Excel5':
				return 'application/vnd.ms-excel';
			case 'Excel2003XML':
				return 'application/xml';
			case 'OOCalc':
				return 'application/vnd.oasis.opendocument.spreadsheet';
			case 'SYLK':
				return 'application/x-sylk';
			case 'Gnumeric':
				return 'application/x-gnumeric';
			case 'HTML':
				return 'text/html';
			case 'CSV':
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
				return 'xlsx';
			case 'Excel5':
				return 'xls';
			case 'Excel2003XML':
				return 'xml';
			case 'OOCalc':
				return 'ods';
			case 'SYLK':
				return 'slk';
			case 'Gnumeric':
				return 'gnumeric';
			case 'HTML':
				return 'html';
			case 'CSV':
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
		$procesor->setActiveSheetIndex(0);
		$this->fill($procesor->getActiveSheet(), $xs);

		$writer = PHPExcel_IOFactory::createWriter($procesor, $this->version);
		$writer->save('php://output');
	}



	// -- PRIVATE ------------------------------------------------------



	private function getProcesor()
	{
		if (empty($this->procesor)) {
			$this->procesor = new PHPExcel();
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
		$rowSymbol = $this->firstRow;
		foreach ($data as $row) {
			if (! is_array($row) && ! $row instanceof Traversable) {
				continue;
			}

			$columnSymbol = 'A';
			foreach ($row as $name => $cell) {
				$sheet->setCellValue($columnSymbol . $rowSymbol, $this->formatCellValue($name, $cell));
				$columnSymbol++;
			}

			$rowSymbol++;
		}
	}



	private function formatCellValue($name, $cell)
	{
		if (! is_object($cell)) {
			return $cell;
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

}
