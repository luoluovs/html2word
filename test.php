<?php
header("Content-Type: text/html;charset=utf-8");
set_time_limit(0);
//不超时
error_reporting(E_ALL);
require "node/html2word/OfficeComMaker.php";
require "HttpClient.php";

//打印所有的错误

//$url = "http://resapi.xh.com/index.php";
//$url = "http://192.168.10.122/htmltoword.php";
$content = file_get_contents("http://filecache.xht.com/xhfile/202101/26/202101261304-600fa8283d30c.html");

$htmlPath = dirname(__FILE__) . "/resources/10.html";
$docPath = dirname(__FILE__) . "/resources/10.docx";
file_put_contents($htmlPath, $content);
$office = \html2word\OfficeComMaker::getInstance();
if ($office->openFile($htmlPath)->BaseLineAlignment()->SaveLinkToLocal()->SaveAs("docx")) {
    //$office->openFile($docPath)->SaveLinkToLocal();
    return "html转word成功";
}
return "html转word失败";

/*$a = ["htmlContent"=>$content];
var_dump(json_encode($a, JSON_UNESCAPED_UNICODE ));exit;
$http = new HttpClient();
$http->setBaseUrl($url);
$res = $http->postJson(null,["htmlContent"=>$content]);
//var_dump($http);

var_dump($res->toArray());*/
