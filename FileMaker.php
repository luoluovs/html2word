<?php
/**
 * Created by PhpStorm.
 * User: XH18
 * Date: 2021/1/21
 * Time: 14:47
 */
namespace html2word;

class FileMaker{

    private static $intance;

    public static function getInstance(){
        if(!self::$intance){
            self::$intance = new FileMaker();
        }
        return self::$intance;
    }

    /**
     * delDir : 删除文件夹
     * @param string $Dir
     * @return  boolean
     * @throws \Exception
     */
    public function delDir($Dir)
    {
        $Dir = str_replace('', '/', $Dir);
        $Dir = substr($Dir, -1) == '/' ? $Dir : $Dir . '/';
        if (!is_dir($Dir)) {
            return false;
        }
        $dirHandle = opendir($Dir);
        while (false !== ($file = readdir($dirHandle))) {
            if ($file == '.' || $file == '..') {
                continue;
            }
            if (!is_dir($Dir . $file)) {
                unlink($Dir . $file);
            } else {
                $this->delDir($Dir . $file);
            }
        }
        closedir($dirHandle);
        return rmdir($Dir);
    }
}