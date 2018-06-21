<?php
require 'vendor/autoload.php';

$arr = array('doc', 'docx', 'xls', 'xlsx', 'ppt', 'pptx');

set_time_limit(300);
$converter = new \Word\WordToPdf();
$source = __DIR__ . DIRECTORY_SEPARATOR . '1.doc';
$export = __DIR__ . DIRECTORY_SEPARATOR . '2.pdf';
var_dump($converter->execute($source, $export));

die;


//pdf转为图片
$pdftoimage=new \Word\PdfToImage();
$source = __DIR__ . DIRECTORY_SEPARATOR . 'test.png';
var_dump($pdftoimage->pdf2png($export,$source));