<?php

namespace Word;

class Linux
{

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

}