# MergeWord
A PHP library for merge word documents
---merge docx files on one file---


## Usage
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

## Contributing

---welcome to discuss a ideas, features and bugs---
