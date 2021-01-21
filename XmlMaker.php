<?php
/**
 * Created by PhpStorm.
 * User: XH18
 * Date: 2021/1/21
 * Time: 15:42
 */
namespace html2word;

class XmlMaker{
    private static $intance;

    public static function getInstance(){
        if(!self::$intance){
            self::$intance = new XmlMaker();
        }
        return self::$intance;
    }

    public function __construct()
    {

    }

    /**
     * parseWordToXml : 解析docx为xml
     * @param  string $filePath
     * @return string 返回解压的xml路径
     * @throws \Exception
     ** created by zhangjian at 2021/1/21 16:05
     */
    public function parseWordToXml($filePath){
        require_once "Zip.php";
        //先替换为zip
        $unZipPath = substr($filePath,0,strripos($filePath,"."));
        $zipPath = $unZipPath.".zip";
        $zipObject = Zip::getInstance();
        rename($filePath,$zipPath);
        if(file_exists($filePath)){
            unlink($filePath);
        }
        $zipObject->unzip($zipPath,$unZipPath);
        unlink($zipPath);
        return $unZipPath;
    }

    /**
     * compressXmlToWord : 将xml压缩成word
     * @param  string $filePath 需要压缩的资源文件的根目录
     * @return string 返回解压的xml路径
     * @throws \Exception
     ** created by zhangjian at 2021/1/21 16:05
     */
    public function compressXmlToWord($filePath){
        require_once "Zip.php";
        require_once "FileMaker.php";
        //先压缩成zip
        $zipObject = Zip::getInstance();
        $fileObj = FileMaker::getInstance();
        $files = $zipObject->getChildDirs($filePath);
        $zipPath = $filePath.".zip";
        $zipObject->createZip($files,$zipPath);
        if(file_exists($filePath) && is_dir($filePath)){
            $fileObj->delDir($filePath);
        }
        $wordPath = $filePath.".DOCX";
        rename($zipPath,$wordPath);
        if(file_exists($zipPath)){
            unlink($zipPath);
        }
        return $wordPath;
    }

    /**
     * replaceVerticalCenter : 替换xml文件垂直居中
     * @param string $xmlPath
     ** created by zhangjian at 2021/1/21 15:44
     *
     */
    public function replaceVerticalCenter($xmlPath){
        $xmlContent = file_get_contents($xmlPath);

        $xmlContent = str_replace('<w:pPr><w:pStyle w:val="a3"/>',
            '<w:pPr><w:pStyle w:val="a3"/><w:textAlignment w:val="center"/>',
            $xmlContent);

        file_put_contents($xmlPath,$xmlContent);
    }


}