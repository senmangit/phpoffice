<?php
/**
 * Created by PhpStorm.
 * User: senman
 * Date: 2018/6/20
 * Time: 11:19
 */

namespace Word;


class PdfToImage
{
    /**
     * PDF2PNG
     * @param $pdf  待处理的PDF文件
     * @param $path 待保存的图片路径
     * @param $page 待导出的页面 -1为全部 0为第一页 1为第二页
     * @return      保存好的图片路径和文件名
     * 注：此处为坑 对于Imagick中的$pdf路径 和$path路径来说，   php版本为5+ 可以使用相对路径。php7+版本必须使用绝对路径。所以，建议大伙使用绝对路径。
     */
    function pdf2png($pdf, $image_path, $page = -1)
    {

        $path = dirname($image_path);
        if (!extension_loaded('imagick')) {
            return false;
        }
        if (!file_exists($pdf)) {
            return false;
        }
        if (!is_readable($pdf)) {
            return false;
        }
        $im = new Imagick();
        $im->setResolution(150, 150);
        $im->setCompressionQuality(100);
        if ($page == -1)
            $im->readImage($pdf);
        else
            $im->readImage($pdf . "[" . $page . "]");
        foreach ($im as $Key => $Var) {
            $Var->setImageFormat('png');
            $filename = $path . md5($Key . time()) . '.png';
            if ($Var->writeImage($filename) == true) {
                $Return[] = $filename;
            }
        }
        //返回转化图片数组，由于pdf可能多页，此处返回二维数组。
        $this->Spliceimg($Return, $image_path);
    }


    /**
     * Spliceimg
     * @param array $imgs pdf转化png  路径
     * @param string $path 待保存的图片路径
     * @return string  将多个图片拼接为成图的路径
     * 注：多页的pdf转化为图片后拼接方法
     */
    public function Spliceimg($imgs = array(), $img_path = '')
    {
        //自定义宽度
        $width = 1230;
        //获取总高度
        $pic_tall = 0;
        foreach ($imgs as $key => $value) {
            $info = getimagesize($value);
            $pic_tall += $width / $info[0] * $info[1];
        }
        // 创建长图
        $temp = imagecreatetruecolor($width, $pic_tall);
        //分配一个白色底色
        $color = imagecolorAllocate($temp, 255, 255, 255);
        imagefill($temp, 0, 0, $color);
        $target_img = $temp;
        $source = array();
        foreach ($imgs as $k => $v) {
            $source[$k]['source'] = Imagecreatefrompng($v);
            $source[$k]['size'] = getimagesize($v);
        }
        $num = 1;
        $tmp = 1;
        $tmpy = 2; //图片之间的间距
        $count = count($imgs);
        for ($i = 0; $i < $count; $i++) {
            imagecopy($target_img, $source[$i]['source'], $tmp, $tmpy, 0, 0, $source[$i]['size'][0], $source[$i]['size'][1]);
            $tmpy = $tmpy + $source[$i]['size'][1];
            //释放资源内存
            imagedestroy($source[$i]['source']);
        }
        $returnfile = $img_path . date('Y-m-d');
        if (!file_exists($returnfile)) {
            if (!$this->make_dir($returnfile)) {
                /* 创建目录失败 */
                return false;
            }
        }
        $return_imgpath = $returnfile . '/' . md5(time() . $pic_tall . 'pdftopng') . '.png';
        imagepng($target_img, $return_imgpath);
        return $return_imgpath;
    }


    /**
     * make_dir
     * @param string $folder 生成目录地址
     * 注：生成目录方法
     */
    public function make_dir($folder)
    {
        $reval = false;
        if (!file_exists($folder)) {
            /* 如果目录不存在则尝试创建该目录 */
            @umask(0);
            /* 将目录路径拆分成数组 */
            preg_match_all('/([^\/]*)\/?/i', $folder, $atmp);
            /* 如果第一个字符为/则当作物理路径处理 */
            $base = ($atmp[0][0] == '/') ? '/' : '';
            /* 遍历包含路径信息的数组 */
            foreach ($atmp[1] AS $val) {
                if ('' != $val) {
                    $base .= $val;
                    if ('..' == $val || '.' == $val) {
                        /* 如果目录为.或者..则直接补/继续下一个循环 */
                        $base .= '/';
                        continue;
                    }
                } else {
                    continue;
                }
                $base .= '/';
                if (!file_exists($base)) {
                    /* 尝试创建目录，如果创建失败则继续循环 */
                    if (@mkdir(rtrim($base, '/'), 0777)) {
                        @chmod($base, 0777);
                        $reval = true;
                    }
                }
            }
        } else {
            /* 路径已经存在。返回该路径是不是一个目录 */
            $reval = is_dir($folder);
        }
        clearstatcache();
        return $reval;
    }
}