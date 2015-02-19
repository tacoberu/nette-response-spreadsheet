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
 * @call phpunit --bootstrap ../../../../../bootstrap.php SpreadsheetResponseTest.php
 */
class SpreadsheetResponseTest extends PHPUnit_Framework_TestCase
{


	function testDefault()
	{
		$request = $this->getMock('Nette\Http\Request', array(), array(), '', False);

		$response = $this->getMock('Nette\Http\Response', array('setContentType', 'setHeader'), array(), '', False);
		$response->expects($this->once())
			->method('setContentType')
			->with('application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
		$response->expects($this->at(1))
			->method('setHeader')
			->with('Cache-Control', 'max-age=0');
		$response->expects($this->at(2))
			->method('setHeader')
			->with('Content-Transfer-Encoding', 'binary');
		//~ $response->expects($this->at(3))
			//~ ->method('setHeader')
			//~ ->with('z', 'a');

		$procesor = $this->getMock('Taco\Nette\Application\Responses\SpreadsheetProcesor',
				array('setProperties', 'setHeaders', 'getContentType', 'getExtension', 'echo_'),
				array(), '', False);
		$procesor->expects($this->once())
			->method('getContentType')
			->will($this->returnValue('application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'));
		$procesor->expects($this->once())
			->method('setProperties')
			->with(array())
			->will($this->returnSelf());
		$procesor->expects($this->once())
			->method('setHeaders')
			->with(array())
			->will($this->returnSelf());
		$procesor->expects($this->once())
			->method('echo_')
			->with(new ArrayIterator(array()));

		$impl = new SpreadsheetResponse(array(), array(), $procesor);
		//~ print_r($response);
		$impl->send($request, $response);
	}


	function testWithFilename()
	{
		$request = $this->getMock('Nette\Http\Request', array(), array(), '', False);

		$response = $this->getMock('Nette\Http\Response', array('setContentType', 'setHeader'), array(), '', False);
		$response->expects($this->once())
			->method('setContentType')
			->with('application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
		$response->expects($this->at(1))
			->method('setHeader')
			->with('Cache-Control', 'max-age=0');
		$response->expects($this->at(2))
			->method('setHeader')
			->with('Content-Transfer-Encoding', 'binary');
		$response->expects($this->at(3))
			->method('setHeader')
			->with('Content-Disposition', 'attachment;filename="abc.ext.xslx"');

		$procesor = $this->getMock('Taco\Nette\Application\Responses\SpreadsheetProcesor',
				array('setProperties', 'setHeaders', 'getContentType', 'getExtension', 'echo_'),
				array(), '', False);
		$procesor->expects($this->once())
			->method('getContentType')
			->will($this->returnValue('application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'));
		$procesor->expects($this->once())
			->method('getExtension')
			->will($this->returnValue('xslx'));
		$procesor->expects($this->once())
			->method('setProperties')
			->with(array())
			->will($this->returnSelf());
		$procesor->expects($this->once())
			->method('setHeaders')
			->with(array())
			->will($this->returnSelf());
		$procesor->expects($this->once())
			->method('echo_')
			->with(new ArrayIterator(array()));

		$impl = new SpreadsheetResponse(array(), array(), $procesor);
		$impl->setFilename('abc.ext');
		//~ print_r($response);
		$impl->send($request, $response);
	}


}
