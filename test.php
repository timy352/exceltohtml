<!DOCTYPE html>
<html lang="en">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=UTF-8" />
</head>

<body>
<?php
require_once('excelphp.php');
$rt = new ExcelPHP(false);
$text = $rt->readDocument('sample.xlsx','');
echo $text;
// The following two lines are optional if required to save the html text to a file (xlfile.php)
$myfile = fopen("xlfile.php", "w") or die("Unable to open file!");
fwrite($myfile, $text);
?>
</body>
