# phpexcel

基于phpspreadsheet封装的excel导表组件。

1、安装

composer require senman/phpexcel dev-master

2、使用示例

```

<?php

require 'vendor/autoload.php'; //引入自动加载文件

$phpexcel = new \Excel\Excel(); //实例化入口文件

//设置表格标题和样式
$tableheader = [
    [
        "title" => "我是标题1",//定义标题，必须要配置
        
        "font_size" => 10,//定义标题字体大小
        
        "font_name" => "微软雅黑",//定义标题
        
        "font_color" => "FFFF0000",//定义标题字体颜色
        
        "fill_color" => "00B050",//填充颜色

    ],
    
    [
        "title" => "我是标题2",
        
//        "font_size" => 10,//定义标题字体大小

//        "font_name" => "微软雅黑",//定义标题

//        "font_color" => "FFFF0000",//定义标题字体颜色

//        "fill_color" => "00B050",//填充颜色
    ]

];


//设置数据部分的表格样式
$data_style = [
    "font_size" => 10,//定义数据部分字体大小
    "font_name" => "微软雅黑",//定义数据部分
    "font_color" => "FFFF0000",//定义数据部分字体颜色
    "fill_color" => "00B050",//定义数据部分填充颜色


];

//定义sheet的名称
$sheetname = "测试表";


//定义数据
$data = [
    ["senman" => 1, "senman1" => "2"],
];

$file_name="test.xls";//文件名称

//执行导出
$is_save=0;   //0：直接下载，1：生成文件后保存在服务器（同时需要配置保存路径）

$save_path="./";//文件保存路径，当is_save为1时生效，否则不生效


//表格属性，选配
$properties=[
   "creator"=>"创建人",
   "last_modified"=>"最后修改人",
   "title"=>"标题",
   "subject"=>"主题",
   "description"=>"描述",
   "keywords"=>"关键词",
   "category"=>"种类",
];

$phpexcel->export($data, $file_name, $tableheader, $sheetname, $is_save, $save_path, $properties, $data_style);

```

6、如有任何疑问欢迎加入QQ群：338461207 进行交流
if you have any questions, welcome to join QQ group: 338461207
