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


require_once __dir__ . '/../../../../../vendor/autoload.php';
require_once __dir__ . '/../../../../../libs/Taco/Nette/Application/Responses/SpreadsheetResponse.php';


use PHPUnit_Framework_TestCase;
use Nette;
use ArrayIterator;


/**
 * @call phpunit --bootstrap ../../../../../bootstrap.php SpoutSpreadsheetResponseTest.php
 */
class SpoutSpreadsheetResponseTest extends PHPUnit_Framework_TestCase
{


	function testDefault()
	{
		$data1 = array(
			array('A0', 'B0', 'C0'),
			array('A1', 'B1', 'C1'),
			array('A2', 'B2', 'C2'),
		);

		$impl = $this->getMock('Taco\Nette\Application\Responses\SpoutCSVWriter'
				, array('openToBrowser', 'addRow', 'close')
				, array(), '', False);
		$impl->expects($this->at(1))
			->method('addRow')
			->with(array('A0', 'B0', 'C0'));
		$impl->expects($this->at(2))
			->method('addRow')
			->with(array('A1', 'B1', 'C1'));
		$impl->expects($this->at(3))
			->method('addRow')
			->with(array('A2', 'B2', 'C2'));

		$procesor = new SpoutSpreadsheetProcesor(Null, $impl);

		$data = array(
				(object) array(
						'name' => 'A',
						'headers' => array(),
						'rows' => new ArrayIterator($data1),
						),
				);

		$procesor->echo_($data);
	}



	function testRowAsIterator()
	{
		$data1 = array(
			new ArrayIterator(array('A0', 'B0', 'C0')),
			new ArrayIterator(array('A1', 'B1', 'C1')),
			new ArrayIterator(array('A2', 'B2', 'C2')),
		);

		$impl = $this->getMock('Taco\Nette\Application\Responses\SpoutCSVWriter'
				, array('openToBrowser', 'addRow', 'close')
				, array(), '', False);
		$impl->expects($this->at(1))
			->method('addRow')
			->with(array('A0', 'B0', 'C0'));
		$impl->expects($this->at(2))
			->method('addRow')
			->with(array('A1', 'B1', 'C1'));
		$impl->expects($this->at(3))
			->method('addRow')
			->with(array('A2', 'B2', 'C2'));

		$procesor = new SpoutSpreadsheetProcesor(Null, $impl);

		$data = array(
				(object) array(
						'name' => 'A',
						'headers' => array(),
						'rows' => new ArrayIterator($data1),
						),
				);

		$procesor->echo_($data);
	}



	function testRowAsStdClass()
	{
		$data1 = array(
			(object)array('A0', 'B0', 'C0'),
			(object)array('A1', 'B1', 'C1'),
			(object)array('A2', 'B2', 'C2'),
		);

		$impl = $this->getMock('Taco\Nette\Application\Responses\SpoutCSVWriter'
				, array('openToBrowser', 'addRow', 'close')
				, array(), '', False);
		$impl->expects($this->at(1))
			->method('addRow')
			->with(array('A0', 'B0', 'C0'));
		$impl->expects($this->at(2))
			->method('addRow')
			->with(array('A1', 'B1', 'C1'));
		$impl->expects($this->at(3))
			->method('addRow')
			->with(array('A2', 'B2', 'C2'));

		$procesor = new SpoutSpreadsheetProcesor(Null, $impl);

		$data = array(
				(object) array(
						'name' => 'A',
						'headers' => array(),
						'rows' => new ArrayIterator($data1),
						),
				);

		$procesor->echo_($data);
	}





}
