<?php
/**
 * Created by PhpStorm.
 * User: senman
 * Date: 2018/6/11
 * Time: 9:20
 */
require 'vendor/autoload.php';

$phpexcel = new \Excel\Excel();
$tableheader = [
    [
        "title" => "我是测试",
        "font_size" => 10,
        "font_name" => "微软雅黑",
        "font_color" => "FFFF0000",
    ],
    [
    "title" => "我是测试2",
    "font_size" => 10,
    "font_name" => "微软雅黑",
    "font_color" => "FFFF0000",
]

];


$data_style = [
    [
        "title" => "我是测试",
        "font_size" => 10,
        "font_name" => "微软雅黑",
        "font_color" => "FFFF0000",
    ]


];
$sheetname = "充值和提现记录表";
$data = [
    ["q"=>1,"qq"=>"haha"],
];
//export($data, $file_name, $fileheader, $sheetname, $is_save = 0, $save_path = "", $properties = [])
$phpexcel->export($data, "test.xls", $tableheader, $sheetname, 0, "./", $properties = [],$data_style);