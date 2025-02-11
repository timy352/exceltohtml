# This is a class to read a Microsoft Excel XLSX file and output it in HTML format for the web

## DESCRIPTION

This php script will convert an Excel .XLSX file to html and display the resultant code (including images) in a web page. It will recognise nearly all the formatting, images etc. in the original Excel XLSX document including the most Conditional Formatting styles except Icon Sets and those based on formulas. At present it doesn't support graphs and charts. 

NOTE:- Needs at least php 5 and will run on up to (at least) php 8.2.

NOTE:- It will not read the older 'XLS' Excel format. In this case it will give an error message saying that it is a 'non zip file'.

I have a website featuring all my php scripts where you can also test this script with one of your own Excel documents - https://phpscripts.meccanoindex.co.uk. You can also contact me through my website.

FEATURES

1. It replicates nearly all text formatting - font (so long as its a common font), font size, bold, italic, single and double underlining, superscript, subscript, strikethrough and colour. It will also support hyperlinks.

2. It recognises and replicates virtually all number formatting, including %, scientific, currency, accountacy, time and date etc..

3. It will recognise how numbers/text are positioned in a cell - left, centre, right - top, centre, bottom.

4. It replicates most cell formatting including merged cells - background colour, border lines in solid, dotted and dashed in three thicknesses and the correct colour. It can reproduce diagonal lines in all colours, but so far can only reproduce solid lines which are in only one thickness. At present it can't reproduce two diagonal lines in a cell to form a cross, so it reverts to a single diagonal line. It also can't replicate cell background patterns. (Currently no easy way to do these in CSS).

5. It will reproduce most of the more common Conditional Formatting styles (either from a given number/string or the contents of a referenced cell):-
Greater Than, Greater Than or Equal to, Less Than, Less Than or Equal to, Between, Not Between, Equal, Not Equal, Contains a Text String, Does Not Contain a Text String, Begins with a Text String, Ends with a Text String, Duplicate Values, Unique Values.

6. It will now reproduce virtually all other conditional formatting except Icon Sets type and also any based on formulas. I.e. it supports Data Bar, Colour Scale, Top N(N%), Bottom N(N%), Above/below average and 1,2,3 Standard Deviation away.

7. It will just display the area of populated spreadsheet cells. Any blank columns or rows, either before or after the populated cells are not shown. In the 'Print' mode it will display the Excel 'Print Area' - assuming that this has been set.

8. It will display any images in the correct cell locations as per the spreadsheet. The images are also saved to the default folder 'images'. However this folder/directory name can be changed if desired.

9. The results of all formulas/calculations are shown.

10. It can display headers and footers together with any text formatting - selectable, see below.

11. If the spreadsheet contains more that one populated sheet, then the default is to display them all in succession one after the other. However an option is provided to just display one, which can be selected by its sheet number.

12. By default column widths and row heights are as per the original Excel spreadsheet. However an 'Auto' mode is available as an option where the browser is allowed to choose them itself. This option can be useful for wide spreadsheets and/or narrower screens.

13. If the default 'Print' option is selected, it will look similar to a printed Excel sheet following the 'Print Area' setting (if set), with headers and footers displayed, but no row and column references.

14. If the 'Spreadsheet' option is selected then it will display the name of the sheet at the top together with the correct row and column references in the usual gray background along the top and left hand side as per a spreadsheet. Headers and footers will not be shown in this mode.

15. It will follow the Excel spreadsheet as to whether gridlines are shown or not. When gridlines would normally be shown there is an option to not display them.

16. The resultant html code is designed to be used either as is, or (after saving) included in another html file.

If anyone finds any problems or has sugestions for enhancements, please contact me on timothy.edwards1@btinternet.com 

# BASIC USAGE
```
require_once('excelphp.php');
$rt = new ExcelPHP(false);
$text = $rt->readDocument('FILENAME','');
echo $text;
```

# DETAILED USAGE

## Include the class in your php script
```
require_once('excelphp.php');
```

## debug mode
Will display the various zipped XML files in the XLSX file which are used by the class for the document being converted and will also display the resultant HTML.
```
$rt = new ExcelPHP(true);
```

## without debug (Normal Use)
```
$rt = new ExcelPHP(false); or $rt = new ExcelPHP();
```

## Set output encoding (Default is UTF-8)
You can alter the encoding of the resultant HTML - eg. 'ISO-8859-1', 'windows-1252', etc. Although note that many special chacters and symbols may not then display correctly so the default should be used whenever practical.
```
$rt = new ExcelPHP(false, 'ISO-8859-1');
```

## Change directory for images (Default is 'images')
Will change the directory used for any images in the document.
```
$rt = new ExcelPHP(false, null, 'dir_name');
```

## Read xlsx file and return the html code - Default mode
See below for various options
```
$text = $rt->readDocument('FILENAME','');
```

## Display the html code on screen
```
echo $text;
```

##  Save the html code to a file (if required)
```
$myfile = fopen("xlfile.php", "w") or die("Unable to open file!");

fwrite($myfile, $text)
```

# OPTIONS

## Select which sheet(s) to display - Option 1
Default is to show all sheets. However you can also select an individual sheet. This is selected by the first character in the last parameter of readDocument(), i.e.:-
```
$text = $rt->readDocument('FILENAME','A'); Display all sheets in the Excel spreadsheet. (Default)
$text = $rt->readDocument('FILENAME','2'); Sheet two is shown selected for display here.
```

## Column Widths and Row Heights - Option 2
Default is to make the Column Widths and Row Heights approx the same as the original spreadsheet. This is selected by the second character in the last parameter of readDocument(), i.e.:-
```
$text = $rt->readDocument('FILENAME','AO'); The Column Widths and Row Heights will be approx the same as the original spreadsheet. (Default, so option character not necessarily needed)
$text = $rt->readDocument('FILENAME','AA'); 'Auto' mode, it will let the browser choose the Column Widths and Row Heights.
```

## 'Print' or 'Spreadsheet' mode - Option 3
Default is to make display similar to a printed spreadsheet. This is selected by the third character in the last parameter of readDocument(), i.e.:-
```
$text = $rt->readDocument('FILENAME','AOP'); 'Print' mode - Displays the 'Print Area' (if set). Includes headers and footers, but no row or column references or sheet name. (Default, so option character not necessarily needed)
$text = $rt->readDocument('FILENAME','AOS'); 'Spreadsheet' mode - Includes row and column references and the sheet name but no headers/footers.
```

## Gridlines display mode - Option 4
You can choose whether to have the gridlines or not. Default is to have them.
```
$text = $rt->readDocument('FILENAME','AOPG');  'Display Gridlines' mode. Will display gridlines (assuming that gridlines are shown in the Excel document. (Default, so option character not necessarily needed)
$text = $rt->readDocument('FILENAME','AOPN');  'No Gridlines' mode. Will not display gridlines
```
## UPDATE NOTES

Version 1.0.3 - A bug is cleared that prevented some images from being shown. Also it now follows the original Excel spreadsheet as to whether gridlines are shown or not. There is also a new option to not show gridlines when they would otherwise be shown. When in the 'Print' display mode, it will now just show the Excel defined 'Print Area' - assuming it is set. Also in some environments a lot of warning messages were displayed. These have now been cleared.

Version 1.0.2 - Removed a depreciated aspect of strpos.

Version 1.0.1 - It provides the following enhancements -
It will now recognise more number formatting. It will now replicate the following Conditional Formatting (except formula based ones) - Data Bar, Colour Scale, Top N(N%), Bottom N(N%), Above/below average and 1,2,3 Standard Deviations Away. It now gives better support/replication for currency units - particularly trailing ones. In accounting formatting a '0' amount will be now be displayed as '-' as per Excel. Also a number of bugs have been cleared.
