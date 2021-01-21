<?php
/**
 * Created by PhpStorm.
 * User: XH18
 * Date: 2021/1/21
 * Time: 14:59
 */
namespace html2word;

class OfficeComMaker{
    private static $intance;
    private static $Word;

    public static function getInstance(){
        if(!self::$intance){
            self::$intance = new OfficeComMaker();
        }
        return self::$intance;
    }

    public function getWordVersion(){
        return self::$Word->Version;
    }

    public function __construct()
    {
        //初始化com组件
        self::$Word = new \COM("word.application") or die("Please install Office");
    }

    /**
     * html2Word : html转word
     * @param $htmlPath html路径
     * @param $fileName 默认当前目录下
     * @throws \Exception
     ** created by zhangjian at 2021/1/21 15:07
     */
    public function html2Word($htmlPath,$fileName){
        //word另存为第二个传参需要32位
        $wordParameter = new \VARIANT(16, VT_I4);
        if(!file_exists($htmlPath)){
            throw new \Exception("html文件不存在");
        }
        try{
            self::$Word->visible =0 ;
            $doc = self::$Word->Documents->Open($htmlPath, false, false, false, "1", "1", true);

            $doc->final = false;
            $doc->Saved = true;
            //保持和系统路径的斜杠一样
            $doc->SaveAs2($fileName,$wordParameter);
        }catch (\Exception $e){
            //释放资源
            $this->releaseWord();
            var_dump($e);
        }
        $this->releaseWord();
    }

    /**
     * releaseWord : 释放资源
     ** created by zhangjian at 2021/1/21 15:35
     */
    private function releaseWord(){
        self::$Word->Quit();
        self::$Word = null;
    }
}