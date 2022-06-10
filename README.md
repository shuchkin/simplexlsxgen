# SimpleXLSXGen
[<img src="https://img.shields.io/github/license/shuchkin/simplexlsxgen" />](https://github.com/shuchkin/simplexlsxgen/blob/master/license.md) [<img src="https://img.shields.io/github/stars/shuchkin/simplexlsxgen" />](https://github.com/shuchkin/simplexlsxgen/stargazers) [<img src="https://img.shields.io/github/forks/shuchkin/simplexlsxgen" />](https://github.com/shuchkin/simplexlsxgen/network) [<img src="https://img.shields.io/github/issues/shuchkin/simplexlsxgen" />](https://github.com/shuchkin/simplexlsxgen/issues)

Export data to Excel XLSX file. PHP XLSX generator. No external tools and libraries.  
- XLSX reader [here](https://github.com/shuchkin/simplexlsx)
- XLS reader [here](https://github.com/shuchkin/simplexls)
- CSV reader/writer [here](https://github.com/shuchkin/simplecsv)

**Sergey Shuchkin** <sergey.shuchkin@gmail.com> 2020-2022<br/>

*Hey, bro, please ★ the package for my motivation :) and [donate](https://opencollective.com/simplexlsx) for more motivation!* 

## Basic Usage
```php
$books = [
    ['ISBN', 'title', 'author', 'publisher', 'ctry' ],
    [618260307, 'The Hobbit', 'J. R. R. Tolkien', 'Houghton Mifflin', 'USA'],
    [908606664, 'Slinky Malinki', 'Lynley Dodd', 'Mallinson Rendel', 'NZ']
];
$xlsx = Shuchkin\SimpleXLSXGen::fromArray( $books );
$xlsx->saveAs('books.xlsx'); // or downloadAs('books.xlsx') or $xlsx_content = (string) $xlsx 
```
![XLSX screenshot](books.png)

## Installation
The recommended way to install this library is [through Composer](https://getcomposer.org).
[New to Composer?](https://getcomposer.org/doc/00-intro.md)

This will install the latest supported version:
```bash
$ composer require shuchkin/simplexlsxgen
```
or download class [here](https://github.com/shuchkin/simplexlsxgen/blob/master/src/SimpleXLSXGen.php)

## Examples
### Data types
```php
$data = [
    ['Integer', 123],
    ['Float', 12.35],
    ['Percent', '12%'],
    ['Datetime', '2020-05-20 02:38:00'],
    ['Date','2020-05-20'],
    ['Time','02:38:00'],
    ['Datetime PHP', new DateTime('2021-02-06 21:07:00')],
    ['String', 'Long UTF-8 String in autoresized column'],
    ['Hyperlink', 'https://github.com/shuchkin/simplexlsxgen'],
    ['Hyperlink + Anchor', '<a href="https://github.com/shuchkin/simplexlsxgen">SimpleXLSXGen</a>'],
    ['RAW string', "\0".'2020-10-04 16:02:00']
];
Shuchkin\SimpleXLSXGen::fromArray( $data )->saveAs('datatypes.xlsx');
```
![XLSX screenshot](datatypes.png)

### Formatting
```php
$data = [
    ['Normal', '12345.67'],
    ['Bold', '<b>12345.67</b>'],
    ['Italic', '<i>12345.67</i>'],
    ['Underline', '<u>12345.67</u>'],
    ['Strike', '<s>12345.67</s>'],
    ['Bold + Italic', '<b><i>12345.67</i></b>'],
    ['Hyperlink', 'https://github.com/shuchkin/simplexlsxgen'],
    ['Italic + Hyperlink + Anchor', '<i><a href="https://github.com/shuchkin/simplexlsxgen">SimpleXLSXGen</a></i>'],
    ['Green', '<style color="#00FF00">12345.67</style>'],
    ['Bold Red Text', '<b><style color="#FF0000">12345.67</style></b>'],
    ['Blue Text and Yellow Fill', '<style bgcolor="#FFFF00" color="#0000FF">12345.67</style>'],
    ['Left', '<left>12345.67</left>'],
    ['Center', '<center>12345.67</center>'],
    ['Right', '<right>Right Text</right>'],
    ['Center + Bold', '<center><b>Name</b></center>'],
    ['Row height', '<style height="50">Row Height = 50</style>'],
    ['Top', '<style height="50"><top>Top</top></style>'],
    ['Middle + Center', '<style height="50"><middle><center>Middle + Center</center></middle></style>'],
    ['Bottom + Right', '<style height="50"><bottom><right>Bottom + Right</right></bottom></style>'],
    ['<center>MERGE CELLS MERGE CELLS MERGE CELLS MERGE CELLS MERGE CELLS</center>', null],
];
SimpleXLSXGen::fromArray( $data )
    ->setDefaultFont( 'Courier New' )
    ->setDefaultFontSize( 14 )
    ->setColWidth(1, 35) // 1 - num column, 35 - size in chars
    ->mergeCells('A20:B20')
    ->saveAs('styles_and_tags.xlsx');
```
![XLSX screenshot](styles.png)

### More examples
```php
// Fluid interface, output to browser for download
Shuchkin\SimpleXLSXGen::fromArray( $books )->downloadAs('table.xlsx');

// Fluid interface, multiple sheets
Shuchkin\SimpleXLSXGen::fromArray( $books )->addSheet( $books2 )->download();

// Alternative interface, sheet name, get xlsx content
$xlsx_cache = (string) (new Shuchkin\SimpleXLSXGen)->addSheet( $books, 'Modern style');

// Classic interface
use Shuchkin\SimpleXLSXGen
$xlsx = new SimpleXLSXGen();
$xlsx->addSheet( $books, 'Catalog 2021' );
$xlsx->addSheet( $books2, 'Stephen King catalog');
$xlsx->downloadAs('books_2021.xlsx');
exit();
```
### JS array to Excel (AJAX)
```php
<?php // array2excel.php
if (isset($_POST['array2excel'])) {
    require __DIR__.'/simplexlsxgen/src/SimpleXLSXGen.php';
    $data = json_decode($_POST['array2excel'], false);
    \Shuchkin\SimpleXLSXGen::fromArray($data)->downloadAs('file.xlsx');
    return;
}
?>
<html lang="en">
<head>
    <title>JS array to Excel</title>
</head>
<script>

function array2excel() {
    var books = [
        ["ISBN", "title", "author", "publisher", "ctry"],
        [618260307, "The Hobbit", "J. R. R. Tolkien", "Houghton Mifflin", "USA"],
        [908606664, "Slinky Malinki", "Lynley Dodd", "Mallinson Rendel", "NZ"]
    ];
    var json = JSON.stringify(books);

    var request = new XMLHttpRequest();

    request.onreadystatechange = function () {
        if (this.readyState === 4) {
            if (this.status === 200) {
                var file = new Blob([this.response], {type:  "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"});
                var fileURL = URL.createObjectURL(file);
                window.open(fileURL);
            } else {
                alert("Error: " + this.status + "  " + this.statusText);
            }
        }
    }
    request.open('POST', "array2excel.php");
    request.responseType = "blob";
    request.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");
    request.send("array2excel=" + encodeURIComponent(json));
}
</script>
<body>
<input type="button" onclick="array2excel()" value="array2excel" />
</body>
</html>
```

## Debug
```php
ini_set('error_reporting', E_ALL );
ini_set('display_errors', 1 );

$data = [
    ['Debug', 123]
];

Shuchkin\SimpleXLSXGen::fromArray( $data )->saveAs('debug.xlsx');
```
