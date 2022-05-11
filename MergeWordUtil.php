<?php
require_once "TbsZip.php";

/**
 * MergeWordUtil version 1.0
 * Date    : 2022-05-08
 * Author  : ZQ
 * 合并word文档工具类
 */
final class MergeWordUtil{

    private function __construct(){
        
    }

    /**
     * 获取word文档的内容
     * @param TbsZip $zip TbsZip对象（防止创建多个TbsZip对象浪费内存资源）
     * @param String $word word文档的路径
     * @return boolean|String $content word文档的内容|false|文件不存在
     */
    public static function getContent($zip,$word){
        if(!file_exists($word)){
            echo "<P>".$word."文件不存在"."</p>";
            return false;
        }
        // Open the first document
        $zip->Open($word);
        $content = $zip->FileRead('word/document.xml');
        $zip->Close();

        // Extract the content of the first document
        $p = strpos($content, '<w:body');
        if ($p===false) return false;
        $p = strpos($content, '>', $p);
        $content = substr($content, $p+1);
        $p = strpos($content, '</w:body>');
        if ($p===false) return false;
        $content = substr($content, 0, $p);

        return $content;
    }

    /**
     * 合并多个word文档
     * @param TbsZip $zip TbsZip对象（防止创建多个TbsZip对象浪费内存资源）
     * @param $workId 项目ID
     * @param Array $word word文档的路径（数组）
     * @param String $resword 要生成的word文档的路径  格式：$dir.'result.docx'
     * @return boolean True|False
     */
    public static function mergeWord($zip,$workId,$words,$resword="upload/word/result.docx"){
        $content = "";
        for($i=1;$i<count($words);$i++){
            $word = MergeWordUtil::getWordPath($words[$i],$workId);
            $curContent = MergeWordUtil::getContent($zip,$word);
            if($curContent == false){
                return false;
            }
            
            $content = $content.$curContent;
        }

        $word0 = MergeWordUtil::getWordPath($words[0],$workId);
        if(!file_exists($word0)){
            echo "<P>".$word0."文件不存在"."</p>";
            return false;
        }
        // Insert into the second document
        $zip->Open($word0);
        $contents = $zip->FileRead('word/document.xml');
        $p = strpos($contents, '</w:body>');
        if ($p===false) return false;
        $contents = substr_replace($contents, $content, $p, 0);
        $zip->FileReplace('word/document.xml', $contents, TBSZIP_STRING);

        $result = MergeWordUtil::getWordPath($resword,$workId);
        // Save the merge into a third file
        $zip->Flush(TBSZIP_FILE, $result);
        $zip->Close();

        return true;
    }

    /**
     * 格式化word文档的路径
     * @param String $word word文档的路径
     * @param $workId 项目ID
     * @return String word文档的路径
     */
    public static function getWordPath($word, $workId){
        $p = strpos($word, '.');

        return substr_replace($word, $workId, $p, 0);
    }
}