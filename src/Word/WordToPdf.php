<?php


namespace Word;

class WordToPdf
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
     * @param $shell  shell的绝对路径
     */
    public function execute($source, $export, $shell = "")
    {
        $my_system = $this->getSystem();
        $className = "Word\\" . $my_system;
        $system = new   $className();
        return $system->execute($source, $export, $shell);
    }

}