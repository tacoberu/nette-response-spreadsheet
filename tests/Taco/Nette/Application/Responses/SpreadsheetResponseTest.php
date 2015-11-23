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
		$response->expects($this->at(4))
			->method('setHeader')
			->with('Pragma', 'public');

		$procesor = $this->getMock('Taco\Nette\Application\Responses\SpreadsheetProcesor',
				array('setProperties', 'setHeaders', 'setPostProcessing', 'getContentType', 'getExtension', 'echo_'),
				array(), '', False);
		$procesor->expects($this->once())
			->method('getContentType')
			->will($this->returnValue('application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'));
		$procesor->expects($this->once())
			->method('setProperties')
			->with(array())
			->will($this->returnSelf());
		$procesor->expects($this->once())
			->method('setPostProcessing')
			->with(array())
			->will($this->returnSelf());
		$procesor->expects($this->once())
			->method('echo_')
			->with(array(
				(object) array(
					'headers' => array(),
					'rows' => new ArrayIterator(array()),
					'name' => Null,
				)
			));

		$impl = new SpreadsheetResponse(array(), array(), $procesor);
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
		$response->expects($this->at(4))
			->method('setHeader')
			->with('Pragma', 'public');

		$procesor = $this->getMock('Taco\Nette\Application\Responses\SpreadsheetProcesor',
				array('setProperties', 'setHeaders', 'setPostProcessing', 'getContentType', 'getExtension', 'echo_'),
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
			->method('setPostProcessing')
			->with(array())
			->will($this->returnSelf());
		$procesor->expects($this->once())
			->method('echo_')
			->with(array(
				(object) array(
					'headers' => array(),
					'rows' => new ArrayIterator(array()),
					'name' => Null,
				)
			));

		$impl = new SpreadsheetResponse(array(), array(), $procesor);
		$impl->setFilename('abc.ext');
		$impl->send($request, $response);
	}


	function testWithData()
	{
		$data = array(
			array('A0', 'B0', 'C0'),
			array('A1', 'B1', 'C1'),
			array('A2', 'B2', 'C2'),
		);

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
		$response->expects($this->at(4))
			->method('setHeader')
			->with('Pragma', 'public');

		$procesor = $this->getMock('Taco\Nette\Application\Responses\SpreadsheetProcesor',
				array('setProperties', 'setHeaders', 'setPostProcessing', 'getContentType', 'getExtension', 'echo_'),
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
			->method('setPostProcessing')
			->with(array())
			->will($this->returnSelf());
		$procesor->expects($this->once())
			->method('echo_')
			->with(array(
				(object) array(
					'headers' => array(),
					'rows' => new ArrayIterator($data),
					'name' => Null,
				)
			));

		$impl = new SpreadsheetResponse($data, array(), $procesor);
		$impl->setFilename('abc.ext');
		//~ print_r($response);
		$impl->send($request, $response);
	}


	function testWithMultipleData()
	{
		$data1 = array(
			array('A0', 'B0', 'C0'),
			array('A1', 'B1', 'C1'),
			array('A2', 'B2', 'C2'),
		);
		$data2 = array(
			array('A0', 'B0', 'C0'),
			array('A1', 'B1', 'C1'),
			array('A2', 'B2', 'C2'),
		);

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
		$response->expects($this->at(4))
			->method('setHeader')
			->with('Pragma', 'public');

		$procesor = $this->getMock('Taco\Nette\Application\Responses\SpreadsheetProcesor',
				array('setProperties', 'setHeaders', 'setPostProcessing', 'getContentType', 'getExtension', 'echo_'),
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
			->method('setPostProcessing')
			->with(array())
			->will($this->returnSelf());
		$procesor->expects($this->once())
			->method('echo_')
			->with(array(
				(object) array(
					'headers' => array(),
					'rows' => new ArrayIterator($data1),
					'name' => Null,
				),
				(object) array(
					'headers' => array('Une', 'Deux', 'Trois'),
					'rows' => new ArrayIterator($data2),
					'name' => Null,
				),
			));

		$impl = new SpreadsheetResponse($data1, array(), $procesor);
		$impl->setFilename('abc.ext');
		$impl->addSheet($data2, array('Une', 'Deux', 'Trois'));
		$impl->send($request, $response);
	}


}
