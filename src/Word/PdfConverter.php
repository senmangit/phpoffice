<?php


namespace Word;

class PdfConverter
{

    /**
     * @return string
     */
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
        $className = "Word\\" . $my_system;
        $system = new   $className();
        return $system->execute($source, $export,$shell="");
    }
}