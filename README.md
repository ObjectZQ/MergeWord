# MergeWord(PHP合并Word文档)
A PHP library for merge word documents（用于合并word文档的PHP库）
---merge docx files on one file（将多个docx文件合并成一个）---


## Usage（用法示例）
```php
require "MergeWord/MergeWordUtil.php"

$dir = "upload/word/";

$zip = new TbsZip();
$word1 = $dir."word1.docx";
$word2 = $dir."word2.docx";
$word3 = $dir."word3.docx";
$words = array($word1, $word2, $word3);

if(MergeWordUtil::mergeWord($zip, $words, $dir."result.docx")) {
    echo "Success!!!";
} else {
    echo "Fail!!!";
}

```

## Contributing（贡献）

---welcome to discuss a ideas, features and bugs（欢迎讨论一个想法、功能和缺陷）---
