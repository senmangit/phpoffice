# phpexcel

简单易用的office套件，可实现导出excel、将word转为PDF等等功能。

一、安装

composer require senman/phpoffice dev-master

二、使用示例


1、导出表格


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

2、导入表格

import()方法有以下参数，除第一个资源路径参数为必须外，其余为选传参数

参数有：$source, $start_line = 2, $end_line = null, $start_column = 1, $end_column = null

$source：文件路径，（必须传入）

$start_line：从第几行开始

$end_line：到几行结束
$start_column：从第几列开始

$end_column：到第几列的数据

该函数返回一个数组
```
require 'vendor/autoload.php';
$phpexcel = new \Excel\Excel();

//表格导入测试
$source = __DIR__ . DIRECTORY_SEPARATOR . 'test.xls';
var_dump($phpexcel->import($source));

```
3、将word转为pdf

该功能基于借助了openoffice服务，具体服务详情自行查阅相关文档，

WIN参考：http://blog.monqin.com/?p=46

LINUX参考：http://blog.monqin.com/?p=69

支持将下列格式文件转为pdf：
'doc', 'docx', 'xls', 'xlsx', 'ppt', 'pptx'

```
$converter = new PDFConverter();//实例化PDF组件
$source = __DIR__ . '/1.doc';//需要转的资源文件
$export = __DIR__ . '/2.pdf';//转成功后需要存贮的路径即文件名
$converter->execute($source, $export);//执行转化操作

```
4、将pdf转为图片

 $pdf  待处理的PDF文件
 $path 待保存的图片路径
 $page 待导出的页面 -1为全部 0为第一页 1为第二页

```
$pdftoimage=new \Word\PdfToImage();
$pdf = __DIR__ . DIRECTORY_SEPARATOR . '2.pdf';//需要转为图片的PDF
$path = __DIR__ . DIRECTORY_SEPARATOR . 'test.png';//需要保存的图片路径
var_dump($pdftoimage->pdf2png($pdf,$path,$page=-1));
```


5、若是需要在linux安装openoffice环境，可直接执行该包里的office.sh脚本文件，可以避免很多的踩坑事件


6、如有任何疑问欢迎加入QQ群：338461207 进行交流
if you have any questions, welcome to join QQ group: 338461207

