<?php
/**
 * Created by PhpStorm.
 * User: XH18
 * Date: 2021/1/21
 * Time: 14:59
 */

namespace html2word;

class OfficeComMaker
{
    private static $intance;
    private static $Word;
    private static $docUrl;
    private static $docName;

    private $wordType = [
        8 => "html",
        16 => "docx"
    ];

    /**
     * getInstance : 单例 实例化
     * @return OfficeComMaker
     ** created by zhangjian at 2021/1/22 14:07
     */
    public static function getInstance()
    {
        if (!self::$intance) {
            self::$intance = new OfficeComMaker();
        }
        return self::$intance;
    }

    public function getWordVersion()
    {
        return self::$Word->Version;
    }

    public function __construct()
    {
        //初始化com组件
        self::$Word = new \COM("word.application") or die("Please install Office");
    }

    /**
     * openFile : 打开文件
     * @param $filePath : 文件路径
     * @return object
     * @throws \Exception
     ** created by zhangjian at 2021/1/21 15:07
     */
    public function openFile($filePath)
    {
        //word另存为第二个传参需要32位

        $filePath = str_replace("\\", "/", $filePath);
        if (!file_exists($filePath)) {
            throw new \Exception("文件路径不存在");
        }
        try {
            self::$Word->visible = 0;
            $doc = self::$Word->Documents->Open($filePath, false, false, false, "1", "1", true);

            $doc->final = false;
            $doc->Saved = true;
            //保持和系统路径的斜杠一样
            self::$docUrl = dirname($filePath);
            self::$docName = substr(basename($filePath), 0, strrpos(basename($filePath), "."));
        } catch (\Exception $exception) {
            $this->handleException($exception);
        }
        return self::$intance;
    }

    /**
     * getFileContent : 获取文本内容
     * @return mixed
     ** created by zhangjian at 2021/1/22 15:03
     *
     */
    public function getFileContent(){
        try {
            return self::$Word->ActiveDocument->content->Text;
        } catch (\Exception $exception) {
            $this->handleException($exception);
        }
    }

    /**
     * Save : 另存为
     * @param $type :另存为的文件类型
     * @param $fileName : 文件名称 不填则为源文件名称
     * @throws \Exception
     * @return boolean
     ** created by zhangjian at 2021/1/22 13:46
     *
     */
    public function Save($type = "docx", $fileName = "")
    {

        if($index = array_search($type,$this->wordType) !== false){
            $wordParameter = new \VARIANT($index, VT_I4);
        }else{
            //16为docx文档
            $wordParameter = $index= 16;
        }

        //设置文档名称
        if (empty($fileName)) {
            if (!self::$docUrl) {
                self::$docUrl = dirname(__FILE__);
            }
            if (!self::$docName) {
                self::$docName = $this->generateDocName();
            }
            $fileName = self::$docUrl . "/" . self::$docName . ".".$type;
        }
        $fileName = str_replace("\\", "/", $fileName);
        try {
            self::$Word->ActiveDocument->SaveAs2($fileName, $wordParameter);
        } catch (\Exception $exception) {
            $this->handleException($exception);
        }

        $this->releaseWord();
        return true;
    }

    /**
     * BaseLineAlignment :  指定行中字体的垂直位置 （垂直）
     ** created by zhangjian at 2021/1/22 13:17
     */
    public function BaseLineAlignment()
    {
        try {
            self::$Word->ActiveDocument->content->ParagraphFormat->BaseLineAlignment = 1;
        } catch (\Exception $exception) {
            $this->handleException($exception);
        }
        return self::$intance;
    }

    /**
     * Alignment : 指定行中字体的水平位置(水平)
     ** created by zhangjian at 2021/1/22 13:17
     */
    public function Alignment()
    {
        try {
            self::$Word->ActiveDocument->content->ParagraphFormat->Alignment = 1;
        } catch (\Exception $exception) {
            $this->handleException($exception);
        }

        return self::$intance;
    }

    /**
     * LineSpacing : 行间距
     * @param $lineSpace : 最小12
     * @return object
     * @throws \Exception
     ** created by zhangjian at 2021/1/22 13:38
     *
     */
    public function LineSpacing($lineSpace = 12)
    {
        try {
            self::$Word->ActiveDocument->content->ParagraphFormat->LineSpacing = $lineSpace;
        } catch (\Exception $exception) {
            $this->handleException($exception);
        }

        return self::$intance;
    }

    /**
     * releaseWord : 释放资源
     ** created by zhangjian at 2021/1/21 15:35
     */
    private function releaseWord()
    {
        self::$Word->Quit();
        self::$Word = null;
    }

    /**
     * handleException : 处理异常
     * @param $exception
     * @return   \Exception
     ** created by zhangjian at 2021/1/22 14:17
     */
    private function handleException(\Exception $exception)
    {
        //释放资源
        $this->releaseWord();
        return json_encode($exception);
    }

    /**
     * generateDocName : 生成文档名称
     * @param $prefix : 前缀
     * @return float
     ** created by zhangjian at 2021/1/22 14:17
     */
    private function generateDocName($prefix = "")
    {
        //获取时间戳
        list($msec, $sec) = explode(' ', microtime());
        $msectime = (float)sprintf('%.0f', (floatval($msec) + floatval($sec)) * 1000);

        return $prefix . $msectime . rand(1111, 9999);
    }
}