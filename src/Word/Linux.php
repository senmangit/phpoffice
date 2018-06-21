<?php

namespace Word;

class Linux
{

    /**
     * @param $source
     * @param $export
     * @param string $shell 一定要写绝对路径
     * @return array  执行结果和路径
     * Execute PDF file(absolute path) conversion
     */
    public function execute($source, $export,$java_shell=" /usr/lib/java/jdk1.8.0_171/bin/java ")
    {
        $jod_path = dirname(dirname(dirname(__FILE__))) . DIRECTORY_SEPARATOR . "jodconverter" . DIRECTORY_SEPARATOR . "lib" . DIRECTORY_SEPARATOR . "jodconverter-cli-2.2.2.jar";
        $commond = " {$java_shell}  -jar  {$jod_path}   {$source}   {$export} ";
        exec($commond, $result, $status);
        return array("result" => $result, "status" => $status);
    }

    public function exec_commond($cfe)
    {
        $res = '';

        if ($cfe) {

            if (function_exists('system')) {

                @ob_start();

                @system($cfe);

                $res = @ob_get_contents();

                @ob_end_clean();

            } elseif (function_exists('passthru')) {

                @ob_start();

                @passthru($cfe);

                $res = @ob_get_contents();

                @ob_end_clean();

            } elseif (function_exists('shell_exec')) {

                $res = @shell_exec($cfe);

            } elseif (function_exists('exec')) {

                @exec($cfe, $res);

                $res = join("\n", $res);

            } elseif (@is_resource($f = @popen($cfe, "r"))) {

                $res = '';

                while (!@feof($f)) {

                    $res .= @fread($f, 1024);

                }

                @pclose($f);

            }

        }

        return $res;

    }

}
