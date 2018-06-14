<?php

namespace Word;

class PDFConverter
{
    private $com;

    /**
     * need to install openoffice and run in the background
     * soffice -headless-accept="socket,host=127.0.0.1,port=8100;urp;" -nofirststartwizard
     */
    public function __construct()
    {
        try {
            $this->com = new COM('com.sun.star.ServiceManager');
        } catch (Exception $e) {
            die('Please be sure that OpenOffice.org is installed.');
        }
    }

    /**
     * Execute PDF file(absolute path) conversion
     * @param $source [source file]
     * @param $export [export file]
     */
    public function execute($source, $export)
    {
        $source = 'file:///' . str_replace('\\', '/', $source);
        $export = 'file:///' . str_replace('\\', '/', $export);
        $this->convertProcess($source, $export);
    }

    /**
     * Get the PDF pages
     * @param $pdf_path [absolute path]
     * @return int
     */
    public function getPages($pdf_path)
    {
        if (!file_exists($pdf_path)) return 0;
        if (!is_readable($pdf_path)) return 0;
        if ($fp = fopen($pdf_path, 'r')) {
            $page = 0;
            while (!feof($fp)) {
                $line = fgets($fp, 255);
                if (preg_match('/\/Count [0-9]+/', $line, $matches)) {
                    preg_match('/[0-9]+/', $matches[0], $matches2);
                    $page = ($page < $matches2[0]) ? $matches2[0] : $page;
                }
            }
            fclose($fp);
            return $page;
        }
        return 0;
    }

    private function setProperty($name, $value)
    {
        $struct = $this->com->Bridge_GetStruct('com.sun.star.beans.PropertyValue');
        $struct->Name = $name;
        $struct->Value = $value;
        return $struct;
    }

    private function convertProcess($source, $export)
    {
        $desktop_args = array($this->setProperty('Hidden', true));
        $desktop = $this->com->createInstance('com.sun.star.frame.Desktop');
        $export_args = array($this->setProperty('FilterName', 'writer_pdf_Export'));
        $program = $desktop->loadComponentFromURL($source, '_blank', 0, $desktop_args);
        $program->storeToURL($export, $export_args);
        $program->close(true);
    }


}
