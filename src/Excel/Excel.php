<?php
/**
 * Created by PhpStorm.
 * User: senman
 * Date: 2018/6/8
 * Time: 17:56
 */

namespace Excel;

use PhpOffice\PhpSpreadsheet\Cell\Coordinate;
use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Reader\Exception;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Style\Alignment;
use PhpOffice\PhpSpreadsheet\Style\Fill;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

class Excel
{

    // $tableheader = array('会员账号', '金额', '收款信息', '类型', '处理状态', '处理时间', '申请时间');
    // $sheetname ="测试";
    // $savefile ="保存的文件名";
    //$data_style 数据的格式
    function export($data, $file_name, $fileheader, $sheetname, $is_save = 0, $save_path = "./", $properties = [], $data_style = [])
    {
        $objPHPExcel = new Spreadsheet();
        //$objPHPExcel = new \PHPExcel();
        if (is_null($file_name)) {
            $file_name = time();
        } else {
            //防止中文命名，下载时ie9及其他情况下的文件名称乱码
            iconv('UTF-8', 'GB2312', $file_name);
        }

        // 设置excel的属性：

        if (!empty($properties)) {
            if (!empty($properties['creator'])) {
                //创建人
                $objPHPExcel->getProperties()->setCreator($properties['creator']);
            }

            if (!empty($properties['last_modified'])) {
                //最后修改人
                $objPHPExcel->getProperties()->setLastModifiedBy($properties['last_modified']);
            }

            if (!empty($properties['title'])) {
                //标题
                $objPHPExcel->getProperties()->setTitle($properties['title']);
            }

            if (!empty($properties['subject'])) {
                //题目
                $objPHPExcel->getProperties()->setSubject($properties['subject']);
            }

            if (!empty($properties['description'])) {
                //描述
                $objPHPExcel->getProperties()->setDescription($properties['description']);
            }

            if (!empty($properties['keywords'])) {
                //关键字
                $objPHPExcel->getProperties()->setKeywords($properties['keywords']);
            }

            if (!empty($properties['category'])) {
                //种类
                $objPHPExcel->getProperties()->setCategory($properties['category']);
            }
        }

        //设置excel属性
        $objActSheet = $objPHPExcel->getActiveSheet();
        //设置sheet的name
        $objActSheet->setTitle($sheetname);

//        //合并单元格
//        $objPHPExcel->getActiveSheet()->mergeCells('A18:E22');
//        //分离单元格
//        $objPHPExcel->getActiveSheet()->unmergeCells('A28:B28');


        //根据有生成的excel多少列，$letter长度要大于等于这个值
        // $letter = array('A', 'B', 'C', 'D', 'E', 'F', 'G', 'H');
        $header_count = count($fileheader);//获取头数量
        $letter = $this->getTableHeader($header_count);

        //设置当前的sheet
        $objPHPExcel->setActiveSheetIndex(0);


        /**
         *         const COLOR_BLACK = 'FF000000';
         * const COLOR_WHITE = 'FFFFFFFF';
         * const COLOR_RED = 'FFFF0000';
         * const COLOR_DARKRED = 'FF800000';
         * const COLOR_BLUE = 'FF0000FF';
         * const COLOR_DARKBLUE = 'FF000080';
         * const COLOR_GREEN = 'FF00FF00';
         * const COLOR_DARKGREEN = 'FF008000';
         * const COLOR_YELLOW = 'FFFFFF00';
         * const COLOR_DARKYELLOW = 'FF808000';
         */


        //设置表头数据格式
        for ($i = 0; $i < $header_count; $i++) {

            //设置表头值，这里的setCellValue第二个参数不能使用iconv，否则excel中显示false
            $objActSheet->setCellValue("$letter[$i]1", $fileheader[$i]['title']);

            //设置表头字体大小
            if (isset($fileheader[$i]['font_size'])) {
                $objActSheet->getStyle("$letter[$i]1")->getFont()->setSize($fileheader[$i]['font_size']);
            }
//            else {
//                $objActSheet->getStyle("$letter[$i]1")->getFont()->setSize(12);
//            }

            //设置表头字体样式
            if (isset($fileheader[$i]['font_name'])) {
                $objActSheet->getStyle("$letter[$i]1")->getFont()->setName($fileheader[$i]['font_name']);
            }
//            else {
//                $objActSheet->getStyle("$letter[$i]1")->getFont()->setName("微软雅黑");
//            }

            //设置表头文字颜色
            if (isset($fileheader[$i]['font_color'])) {
                $objActSheet->getStyle("$letter[$i]1")->getFont()->getColor()->setARGB($fileheader[$i]['font_color']);
            } else {
                $objActSheet->getStyle("$letter[$i]1")->getFont()->getColor()->setARGB("FFFFFFFF");
            }

            //单元宽度自适应,1.8.1版本phpexcel中文支持勉强可以，自适应后单独设置宽度无效
//            if (isset($fileheader[$i]['font_auto_size'])) {
//                $objActSheet->getColumnDimension("$letter[$i]1")->setAutoSize($fileheader[$i]['font_auto_size']);
//            } else {
//                $objActSheet->getColumnDimension("$letter[$i]1")->setAutoSize(false);
//            }

            //设置自动换行
            if (isset($fileheader[$i]['wrap_text'])) {
                $objActSheet->getStyle("$letter[$i]1")->getAlignment()->setWrapText($fileheader[$i]['wrap_text']);
            }
//            else {
//                $objActSheet->getStyle("$letter[$i]1")->getAlignment()->setWrapText(false);
//            }

            //设置表头字体是否加粗
            if (isset($fileheader[$i]['font_bold'])) {
                $objActSheet->getStyle("$letter[$i]1")->getFont()->setBold($fileheader[$i]['font_bold']);
            } else {
                $objActSheet->getStyle("$letter[$i]1")->getFont()->setBold(true);
            }

            //设置表头文字垂直居中
            if (isset($fileheader[$i]['alignment_horizontal'])) {
                $objActSheet->getStyle("$letter[$i]1")->getAlignment()->setHorizontal($fileheader[$i]['alignment_horizontal']);
            } else {
                $objActSheet->getStyle("$letter[$i]1")->getAlignment()->setHorizontal(Alignment::HORIZONTAL_CENTER);
            }

            //设置文字上下居中
            //$objActSheet->getStyle($letter[$i])->getAlignment()->setVertical();//与下一行冲突

//            //设置border的颜色
            if (isset($fileheader[$i]['border_color'])) {
                $objPHPExcel->getActiveSheet()->getStyle("$letter[$i]1")->getBorders()->getLeft()->getColor()->setARGB($fileheader[$i]['border_color']);
            }
//            else {
//                $objPHPExcel->getActiveSheet()->getStyle("$letter[$i]1")->getBorders()->getLeft()->getColor()->setARGB('FF993300');
//            }

            //填充类型
            if (isset($fileheader[$i]['fill_type'])) {
                $objPHPExcel->getActiveSheet()->getStyle("$letter[$i]1")->getFill()->setFillType($fileheader[$i]['fill_type']);
            } else {
                $objPHPExcel->getActiveSheet()->getStyle("$letter[$i]1")->getFill()->setFillType(Fill::FILL_SOLID);
            }

            //填充颜色
            if (isset($fileheader[$i]['fill_color'])) {
                $objPHPExcel->getActiveSheet()->getStyle("$letter[$i]1")->getFill()->getStartColor()->setARGB($fileheader[$i]['fill_color']);
            } else {
                $objPHPExcel->getActiveSheet()->getStyle("$letter[$i]1")->getFill()->getStartColor()->setARGB('FF808080');
            }


            //保护cell
            //    $objPHPExcel->getActiveSheet()->getProtection()->setSheet(true); // Needs to be set to true in order to enable any worksheet protection!
            //    $objPHPExcel->getActiveSheet()->protectCells('A3:E13', '123456');

        }


        //设置除表头外数据格式
        for ($i = 0; $i < count($letter); $i++) {

            //设置字体大小
            if (isset($data_style['font_size'])) {
                $objActSheet->getStyle("$letter[$i]")->getFont()->setSize($data_style['font_size']);
            }
            //设置字体样式
            if (isset($data_style['font_name'])) {
                $objActSheet->getStyle("$letter[$i]")->getFont()->setName($data_style['font_name']);
            }
            //设置文字颜色
            if (isset($data_style['font_color'])) {
                $objActSheet->getStyle("$letter[$i]")->getFont()->getColor()->setARGB($data_style['font_color']);
            }
            //单元宽度自适应,1.8.1版本phpexcel中文支持勉强可以，自适应后单独设置宽度无效
            if (isset($data_style['font_auto_size'])) {
                $objActSheet->getColumnDimension("$letter[$i]")->setAutoSize($data_style['font_auto_size']);
            } else {
                $objActSheet->getColumnDimension("$letter[$i]")->setAutoSize(false);
            }

            //设置自动换行
            if (isset($data_style['wrap_text'])) {
                $objActSheet->getStyle("$letter[$i]")->getAlignment()->setWrapText($data_style['wrap_text']);
            } else {
                $objActSheet->getStyle("$letter[$i]")->getAlignment()->setWrapText(false);
            }

            //设置字体是否加粗
            if (isset($data_style['font_bold'])) {
                $objActSheet->getStyle("$letter[$i]")->getFont()->setBold($data_style['font_bold']);
            } else {
                $objActSheet->getStyle("$letter[$i]")->getFont()->setBold(false);
            }
            //设置文字垂直居中
            if (isset($data_style['alignment_horizontal'])) {
                $objActSheet->getStyle("$letter[$i]")->getAlignment()->setHorizontal($data_style['alignment_horizontal']);
            } else {
                $objActSheet->getStyle("$letter[$i]")->getAlignment()->setHorizontal(Alignment::HORIZONTAL_CENTER);
            }

            //设置文字上下居中
            //$objActSheet->getStyle($letter[$i])->getAlignment()->setVertical();//与下一行冲突

//            //设置border的颜色
            if (isset($data_style['border_color'])) {
                $objPHPExcel->getActiveSheet()->getStyle("$letter[$i]")->getBorders()->getLeft()->getColor()->setARGB($data_style['border_color']);
            }
//            else {
//                $objPHPExcel->getActiveSheet()->getStyle("$letter[$i]")->getBorders()->getLeft()->getColor()->setARGB('FF993300');
//            }

            //填充类型
            if (isset($data_style['fill_type'])) {
                $objPHPExcel->getActiveSheet()->getStyle("$letter[$i]")->getFill()->setFillType($data_style['fill_type']);
            } else {
                $objPHPExcel->getActiveSheet()->getStyle("$letter[$i]")->getFill()->setFillType(Fill::FILL_SOLID);
            }
            //填充颜色
            if (isset($data_style['fill_color'])) {
                $objPHPExcel->getActiveSheet()->getStyle("$letter[$i]")->getFill()->getStartColor()->setARGB($data_style['fill_color']);
            }
//            else {
//                $objPHPExcel->getActiveSheet()->getStyle("$letter[$i]")->getFill()->getStartColor()->setARGB('FF808080');
//            }


            //保护cell
            //    $objPHPExcel->getActiveSheet()->getProtection()->setSheet(true); // Needs to be set to true in order to enable any worksheet protection!
            //    $objPHPExcel->getActiveSheet()->protectCells('A3:E13', '123456');

        }


        foreach ($letter as $lk => $lv) {
            //单独设置D列宽度为15
            //$objActSheet->getColumnDimension($lv)->setWidth(100);
        }
//
//        $objActSheet->getColumnDimension('G')->setWidth(20);
//        $objActSheet->getColumnDimension('E')->setWidth(20);
//        $objActSheet->getColumnDimension('F')->setWidth(20);
//        $objActSheet->getColumnDimension('D')->setWidth(20);
//        $objActSheet->getColumnDimension('C')->setWidth(100);
//        $objActSheet->getColumnDimension('B')->setWidth(15);
//        $objActSheet->getColumnDimension('A')->setWidth(30);
        //这里$i初始值设置为2，$j初始值设置为0，


        //加图片
//        $objDrawing = new \PHPExcel_Worksheet_Drawing();
//        $objDrawing->setName('Logo');
//        $objDrawing->setDescription('Logo');
//        $objDrawing->setPath('./images/officelogo.jpg');
//        $objDrawing->setHeight(36);
//        $objDrawing->setWorksheet($objPHPExcel->getActiveSheet());
//        $objDrawing = new \PHPExcel_Worksheet_Drawing();
//        $objDrawing->setName('Paid');
//        $objDrawing->setDescription('Paid');
//        $objDrawing->setPath('./images/paid.png');
//        $objDrawing->setCoordinates('B15');
//        $objDrawing->setOffsetX(110);
//        $objDrawing->setRotation(25);
//        $objDrawing->getShadow()->setVisible(true);
//        $objDrawing->getShadow()->setDirection(45);
//        $objDrawing->setWorksheet($objPHPExcel->getActiveSheet());
        for ($i = 2; $i <= count($data) + 1; $i++) {
            $j = 0;
            foreach ($data[$i - 2] as $key => $value) {
                //是图片是加入图片到excel
                if ($key == 'img') {
                    if ($value != '') {
                        $value = iconv("UTF-8", "GB2312", $value); //防止中文命名的文件
                        // 图片生成
                        $objDrawing[$key] = new \ImagickDraw();
                        // 图片地址
                        $objDrawing[$key]->setPath('.\Uploads' . $value);
                        // 设置图片宽度高度
                        $objDrawing[$key]->setHeight('80px'); //照片高度
                        $objDrawing[$key]->setWidth('80px'); //照片宽度
                        // 设置图片要插入的单元格
                        $objDrawing[$key]->setCoordinates('D' . $i);
                        // 图片偏移距离
                        $objDrawing[$key]->setOffsetX(12);
                        $objDrawing[$key]->setOffsetY(12);
                        //下边两行不知道对图片单元格的格式有什么作用
                        //$objDrawing[$key]->getShadow()->setVisible(true);
                        //$objDrawing[$key]->getShadow()->setDirection(50);
                        $objDrawing[$key]->setWorksheet($objActSheet);
                    }
                } else {
                    $objActSheet->setCellValue("$letter[$j]$i", $value);
                }
                $j++;
            }
            //设置单元格高度，暂时没有找到统一设置高度方法
            //$objActSheet->getRowDimension($i)->setRowHeight('80px');
        }


        if ($is_save) {
            if (!empty($save_path)) {
                //判断目录是否存在,不存在则创建
                if (!file_exists($save_path)) {
                    mkdir($save_path, '0777', true);
                }
                // 保存excel在服务器上
                //$objWriter = new \PHPExcel_Writer_Excel2007($objPHPExcel);
                $objWriter = new Xlsx($objPHPExcel);
                //或者$objWriter = new PHPExcel_Writer_Excel5($excel);
                $objWriter->save($save_path . DIRECTORY_SEPARATOR . $file_name);
            } else {
                throw new \Exception("路径不存在");
            }
        } else {

            ob_end_clean();
            //下载的excel文件名称，为Excel5，后缀为xls，不过影响似乎不大
            header('Content-Disposition: attachment;filename="' . $file_name . '"');
            header('Cache-Control: max-age=0');
            header("Pragma: public");
            header("Expires: 0");
            header("Cache-Control:must-revalidate, post-check=0, pre-check=0");
            header("Content-Type:application/force-download");
            header("Content-Type:application/vnd.ms-execl");
            header("Content-Type:application/octet-stream");
            header("Content-Type:application/download");;
            header("Content-Transfer-Encoding:binary");
            // 用户下载excel
            $objWriter = new Xlsx($objPHPExcel);
            $objWriter->save('php://output');
        }

    }

    /**
     * @param $source
     * @param int $start_line
     * @param null $end_line
     * @param int $start_column
     * @param null $end_column
     * @return array
     * @throws Exception
     * @throws \PhpOffice\PhpSpreadsheet\Exception
     * 将表格数据转换为数组
     */

    function import($source, $start_line = 2, $end_line = null, $start_column = 1, $end_column = null)
    {

        $ext = pathinfo($source, PATHINFO_EXTENSION);
        if (strtolower($ext) == "xls") {
            $reader = IOFactory::createReader('Xls');
        } elseif (strtolower($ext) == "xlsx") {
            $reader = IOFactory::createReader('Xlsx');
        } else {
            throw new Exception("文件后缀错误");
        }
        $spreadsheet = $reader->load($source);
        $reader->setReadDataOnly(TRUE);
        $worksheet = $spreadsheet->getActiveSheet();
        $highestRow = $worksheet->getHighestRow(); // 总行数
        $highestColumn = $worksheet->getHighestColumn(); // 总列数
        $highestColumnIndex = Coordinate::columnIndexFromString($highestColumn); // e.g. 5
        if (!is_null($end_column)) {
            $column_count = $end_column;
        } else {
            $column_count = $highestColumnIndex;
        }

        if (!is_null($end_line)) {
            $line_count = $end_line;
        } else {
            $line_count = $highestRow;
        }


        $lines = $highestRow - 1;
        if ($lines <= 0) {
            exit('Excel表格中没有数据');
        }

        $data = [];
        for ($start_line_i = $start_line; $start_line_i <= $line_count; ++$start_line_i) {
            $colum_data = [];
            for ($start_column_i = $start_column; $start_column_i <= $column_count; ++$start_column_i) {
                $colum_data[] = $worksheet->getCellByColumnAndRow($start_column_i, $start_line_i)->getValue();
            }
            $data[] = $colum_data;
        }
        return $data;
    }

    /**
     * 循环获取头
     */
    public function getTableHeader($header_count)
    {
        $letter = [];
        for ($j = 0; $j < $header_count; $j++) {
            $letter[] = strtoupper($this->greatezimu($j));
        }
        return $letter;
    }


    //生成一个字母
    public function greatezimu($count)
    {
        $zimu_first = $this->greateFirst($count);//获取了前几位的字母
        $wei_count = intval($count % 26);//余数
        return $zimu_first . chr($wei_count + 65);//存入大写字母,65开始为A,$i = 65; $i < 91

    }

    /**
     * @param $count ,总的字母数
     * @param string $s
     * @return string
     * 递归获取前几位字母
     */
    public function greateFirst($count, $s = "")
    {
        $shouwei = intval($count / 26) - 1;//商，即代表前面的字母取值
        $s .= chr($shouwei + 65);
        if ($shouwei <= 26) {
            if ($shouwei < 0) {
                return "";
            }
            return $s;//存入大写字母,65开始为A,$i = 65; $i < 91
        } else {
            return $this->greateFirst($shouwei, $s);
        }

    }
}

