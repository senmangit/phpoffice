<?php

namespace Word;

class Linux
{

    /**
     * Execute PDF file(absolute path) conversion
     * @param $source [source file]
     * @param $export [export file]
     */
    public function execute($source, $export, $shell = " java -jar /home/jodconverter/jodconverter-2.2.2/lib/jodconverter-cli-2.2.2.jar ")
    {
        return exec($shell . " " . $source . " " . $export);
    }

}