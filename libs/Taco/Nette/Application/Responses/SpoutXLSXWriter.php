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


use Box\Spout\Writer\XLSX;


/**
 * Class XLSX
 * This class provides base support to write data to XLSX files
 */
class SpoutXLSXWriter extends XLSX
{

    /**
     * Inits the writer and opens it to accept data.
     * By using this method, the data will be outputted directly to the browser.
     *
     * @return \Box\Spout\Writer\AbstractWriter
     * @throws \Box\Spout\Common\Exception\IOException If the writer cannot be opened
     */
    function openToBrowser()
    {
        $this->filePointer = $this->globalFunctionsHelper->fopen('php://output', 'w');
        $this->throwIfFilePointerIsNotAvailable();

        $this->openWriter();
        $this->isWriterOpened = true;

        return $this;
    }


}
