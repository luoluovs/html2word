<?php
/**
 * Created by PhpStorm.
 * User: XH18
 * Date: 2021/1/21
 * Time: 14:47
 */
namespace html2word;

class Zip{

    private static $intance;
    private static $ZipArchive;

    public static function getInstance(){
        if(!self::$intance){
            self::$intance = new Zip();
        }
        return self::$intance;
    }

    public function __construct()
    {
        if(!static::$ZipArchive){
            static::$ZipArchive = new \ZipArchive();
        }
    }

    /**
     * 解压
     * @param string $filePath 压缩包所在地址 【绝对文件地址】d:/test/123.zip
     * @param string $path 解压路径 【绝对文件目录路径】d:/test
     * @return bool
     * @throws \Exception
     */
    public function unzip($filePath, $path) {
        if (empty($path) || empty($filePath)) {
            return false;
        }

        $zip = static::$ZipArchive;

        if ($zip->open($filePath) === true) {
            $zip->extractTo($path);
            $zip->close();
            return true;
        } else {
            throw new \Exception("解压失败");
        }
    }


    /**
     * createZip : 压缩
     * @param array $sources 压缩的文件地址 【绝对文件地址】d:/test/123.zip
     * @param string $destination 解压路径 【绝对文件目录路径】d:/test
     * @return bool
     * @throws \Exception
     ** created by zhangjian at 2021/1/21 14:57
     */
    public function createZip($sources, $destination) {
        if(!is_array($sources) || !isset($sources[0])) return false;
        $rootUrl = str_replace("\\","/",dirname($sources[0]))."/";

        $zip = static::$ZipArchive;
        if (!$zip->open($destination, \ZIPARCHIVE::CREATE)) {
            throw new \Exception("开始压缩失败");
        }
        foreach ($sources as $source){
            if (!extension_loaded('zip') || !file_exists($source)) {
                throw new \Exception("压缩的资源路径不存在");
            }

            $source = str_replace('\\', '/', realpath($source));

            if (is_dir($source) === true)
            {
                $files = new \RecursiveIteratorIterator(new \RecursiveDirectoryIterator($source), \RecursiveIteratorIterator::SELF_FIRST);

                foreach ($files as $file)
                {
                    $file = str_replace('\\', '/', realpath($file));

                    if( in_array(substr($file, strrpos($file, '/')+1), array('.', '..')) ){
                        continue;
                    }

                    //去掉最外层根文件夹
                    if(($file."/") == $rootUrl){
                        continue;
                    }

                    $tempZipName = str_replace($rootUrl, '', $file);
                    if (is_dir($file) === true)
                    {
                        $zip->addEmptyDir($tempZipName);
                    }
                    else if (is_file($file) === true)
                    {
                        $zip->addFromString($tempZipName, file_get_contents($file));
                    }
                }
            }
            else if (is_file($source) === true)
            {
                $zip->addFromString(basename($source), file_get_contents($source));
            }
        }

        return $zip->close();
    }

    /**
     * getChildDirs : 获取某个目录下的所有文件目录
     * @param $dirPath
     * @return array
     ** created by zhangjian at 2021/1/21 15:46
     */
    public function getChildDirs($dirPath){
        $files = scandir($dirPath);
        $files = array_filter($files,function ($item){
            return  ($item == "." || $item == "..") ? false : true;
        });
        $files = array_values(array_map(function ($item)use($dirPath){
            $item = $dirPath."/".$item;
            return $item;
        },$files));
        return $files;
    }

}