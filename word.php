<?php
require 'vendor/autoload.php';

$arr = array('doc', 'docx', 'xls', 'xlsx', 'ppt', 'pptx');

$converter = new \Word\PDFConverter();
$source = __DIR__ . '/1.doc';
$export = __DIR__ . '/2.pdf';
$converter->execute($source, $export);
