<?php
require 'vendor/autoload.php';

$arr = array('doc', 'docx', 'xls', 'xlsx', 'ppt', 'pptx');

set_time_limit(300);
$converter = new \Word\PdfConverter();
$source = __DIR__ . DIRECTORY_SEPARATOR . '1.doc';
$export = __DIR__ . DIRECTORY_SEPARATOR . '2.pdf';
echo $converter->execute($source, $export);
