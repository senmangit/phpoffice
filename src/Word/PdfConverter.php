<?php


namespace Word;

class PdfConverter
{


    public function getSystem()
    {
        return strtoupper(substr(PHP_OS, 0, 3)) === 'WIN' ? "Windows" : "Linux";
    }

    /**
     * Execute PDF file(absolute path) conversion
     * @param $source [source file]
     * @param $export [export file]
     */
    public function execute($source, $export)
    {
        $my_system = $this->getSystem();
        $system = new $my_system();
        $source = 'file:///' . str_replace('\\', '/', $source);
        $export = 'file:///' . str_replace('\\', '/', $export);
        return $system->execute($source, $export);
    }
}