<?php
header("Content-Type: text/html;charset=utf-8");
set_time_limit(0);
//不超时
error_reporting(E_ALL);
require "node/html2word/OfficeComMaker.php";

//打印所有的错误

$htmlPath = dirname(__FILE__)."/resources/10.html";

$office = \html2word\OfficeComMaker::getInstance();
if($office->openFile($htmlPath)->BaseLineAlignment()->Save("docx")){
    return "html转word成功";
}
return "html转word失败";

