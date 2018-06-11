<?php
/**
 * Created by PhpStorm.
 * User: senman
 * Date: 2018/6/11
 * Time: 9:20
 */
require 'vendor/autoload.php';

$phpexcel= new \Excel\Excel();
$tableheader = array('会员账号');
$sheetname="充值和提现记录表";
$data=[
    ["title"=>1],
];
//export($data, $file_name, $fileheader, $sheetname, $is_save = 0, $save_path = "", $properties = [])
$phpexcel->export($data, "test.xls",$tableheader, $sheetname,1,"./", $properties = []);