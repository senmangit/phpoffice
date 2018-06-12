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
        "title" => "我是测试",//定义标题，必须要配置
        "font_size" => 10,//定义标题字体大小
        "font_name" => "微软雅黑",//定义标题
        "font_color" => "FFFF0000",//定义标题字体颜色
        "fill_color" => "00B050",//填充颜色

    ],
    [
        "title" => "我是测试2",
//        "font_size" => 10,//定义标题字体大小
//        "font_name" => "微软雅黑",//定义标题
//        "font_color" => "FFFF0000",//定义标题字体颜色
//        "fill_color" => "00B050",//填充颜色
    ]

];


$data_style = [

    "font_size" => 10,//定义数据部分字体大小
    "font_name" => "微软雅黑",//定义数据部分
    "font_color" => "FFFF0000",//定义数据部分字体颜色
    "fill_color" => "00B050",//定义数据部分填充颜色


];
$sheetname = "测试表";
$data = [
    ["senman" => 1, "senman1" => "2"],
];
$phpexcel->export($data, "test.xls", $tableheader, $sheetname, 0, "./", $properties = [], $data_style);

