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
 * Naformátuje pole na Excel sešit použitím PHPExcel knihovny.
 */
class PHPExcelSpreadsheetProcesor implements SpreadsheetProcesor
{

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
	 * @param array list of sheet.
	 * @return string
	 */
	function echo_(array $sheets)
	{
		$procesor = $this->getProcesor();

		foreach ($sheets as $index => $pack) {
			if (0 == $index) {
				$sheet = $procesor->getSheet();
			}
			else {
				$sheet = $procesor->createSheet();
			}

			if ($pack->name) {
				$sheet->setTitle($pack->name);
			}

			$next = $this->fillHeaders($sheet, $pack->headers);
			$this->fillRows($sheet, $pack->rows, $next);
		}

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
	 * Naplnit první řádek hlavičkou.
	 * @return int Jaký je následující řádek.
	 */
	private function fillHeaders($sheet, array $headers = array())
	{
		$rowSymbol = 1;

		if (! count($headers)) {
			return $rowSymbol;
		}

		$columnSymbol = 'A';
		foreach ($headers as $name) {
			$cell = $sheet->setCellValue($columnSymbol . $rowSymbol, $name, True);
			//~ $cell->getStyle()->getFont()->setBold(True);
			$columnSymbol++;
		}

		$rowSymbol++;
		return $rowSymbol;
	}



	/**
	 * Naplnit sešit daty.
	 *
	 * @param Traversable $data - data k zpracování
	 * @return PHPExcel
	 */
	private function fillRows($sheet, Traversable $data, $start = 1)
	{
		$rowSymbol = $start;
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

}
