# Changelog

## 1.4.12 (2024-07-28)
* tag ```<raw>``` for styled raw values

## 1.4.11 (2024-02-07)
* hyperlinks to local files
* num2name static now

## 1.4.10 (2023-12-31)
* added SimpleXLSXGen::create($title = null) to create empty book with title
* added SimpleXLSXGen::save to save xlsx in current folder as {$title}.xslx or {$curdate}.xlsx
* SimpleXLSXGen::esc and SimpleXLSXGen::date2excel static now
* added examples SimpleXLSXGen::create, SimpleXSLXGen::raw, SimpleXLSXGen:save in README.md
* fixed [fpassthru disabled issue](https://github.com/shuchkin/simplexlsxgen/issues/116)
* fixed empty book, A1 empty now, text _No data_ removed
* thx [Javier](https://github.com/xaviermdq)

## 1.3.20 (2023-12-12)
* force little endian numbers in zip headers

## 1.3.18 (2023-12-02)
* simple linebreaks

## 1.3.17 (2023-10-02)
* fixed [issue 128](https://github.com/shuchkin/simplexlsxgen/issues/128) date2excel type cast

## 1.3.16 (2023-09-12)
* preserve leading or traling spaces

## 1.3.15 (2023-04-19)
* added meta: setTitle, setSubject, setAuthor, setCompany, setManager, setKeywords, setDescription, setCategory, setApplication, setLastModifiedBy. Thx [Oleg Kosarev](https://github.com/DevOlegKosarev)

## 1.3.14 (2023-04-18)
* fixed &quot;This action doesn't work on multiple selection&quot; error

## 1.3.13 (2023-04-11)
* ```$xlsx->rightToLeft()``` - RTL mode. Column A is on the far right, Column B is one column left of Column A, and so on. Also, information in cells is displayed in the Right to Left format.

## 1.3.12 (2023-03-31)
* ```<style font-size="32">Big Text</style>``` - font size in cells, thx [Andrew Robinson](https://github.com/mrjemson)

## 1.3.11 (2023-03-28)
* freezePanes( corner_cell ) - freezePanes to keep an area of a worksheet visible while you scroll, corner_cell is not included, thx [Javier](https://github.com/xaviermdq)

## 1.3.10 (2022-12-14)
* added borders ```<style border="medium">Black Border</style>``` see colored [examples](https://github.com/shuchkin/simplexlsxgen#formatting)
* added formulas ```<f v="100">SUM(B1:B10)</f>``` see [examples](https://github.com/shuchkin/simplexlsxgen#data-types)
* added internal links ```<a href="sheet2!A1">Go to page 2</a>```
* added custom number formats ```<style nf="&quot;£&quot;#,##0.00">500</style>```
* added 3 currencies ```$data = [ ['$100.23', '2000.00 €', '1200.30 ₽'] ];```

## 1.2.16 (2022-08-12)
* added `autoFilter( $range )`
```php
$xlsx->autoFilter('A2:B10');
```
* fixed `0%` bug 

## 1.2.15 (2022-07-05)
* added wrap words in long strings `<wraptext>long long line</wraptext>`

## 1.2.14 (2022-06-10)
* added example [JS array to Excel (AJAX)](https://github.com/shuchkin/simplexlsxgen#js-array-to-excel-ajax)

## 1.2.13 (2022-06-01)
* setColWidth(num_col_started_1, size_in_chars) - set column width

## 1.2.12 (2022-05-17)
* Vertical align (tags top,middle,bottom) `<bottom>12345</bottom>`

## 1.2.11 (2022-05-01)
* Row height `<style height="50">Custom row height 50</style>`


## 1.2.10 (2022-04-24)
* Added colors `<style color="#FFFF00" bgcolor="#00FF00">Yellow text on blue background</style>`, thx [mrjemson](https://github.com/mrjemson)

## 1.1.12 (2022-03-15)
* Added `$xlsx->mergeCells('A1:C1')`

## 1.1.11 (2022-02-05)
* sheet name maximum length is 31 chars, mb_substr used now
* license fixed

## 1.1.10 (2022-02-05)
* namespace added, use Shuchkin\SimpleXLSXGen

## 1.0.23 (2022-02-01)
* fixed dates if year < 1900 and time only cells, thx [fapth](https://github.com/shuchkin/simplexlsxgen/issues/51)   

## 1.0.22 (2021-10-29)
* Escape \x00 and \x0B (vertical tab)

## 1.0.21 (2021-09-03)
*  Fixed saveAs / downloadAs / etc methods more than once

## 1.0.20 (2021-07-29)
* Fixed sheet names duplicates (Page, Page (1), Page (2)...) 

## 1.0.19 (2021-07-28)
* Fixed sheet names duplicates 

## 1.0.18 (2021-07-28)
* Fixed email regex

## 1.0.17 (2021-07-28)
* Fixed &quot; and &amp; in sheets names

## 1.0.16 (2021-07-01)
* Fixed &quot;&amp;&quot; in hyperlinks

## 1.0.15 (2021-06-22)
* Fixed *mailto* hyperlinks detection

## 1.0.14 (2021-06-08)
* Added *mailto* hyperlinks support (thx Howard Martin)
```php
SimpleXLSXGen::fromArray([
	['Mailto hyperlink', '<a href="mailto:sergey.shuchkin@gmail.com">Please email me</a>']
])->saveAs('test.xlsx');
```
## 1.0.13 (2021-05-29)
* Fixed hyperlinks in several sheets
* Added [Opencollective donation link](https://opencollective.com/simplexlsx)

## 1.0.12 (2021-05-19)
* Fixed hyperlink regex

## 1.0.11 (2021-05-14)
* Fixed 0.00% format, thx [marcrobledo](https://github.com/shuchkin/simplexlsxgen/pull/34), more examples in README.md

## 1.0.10 (2021-05-03)
Stable release

* Added hyperlinks and minimal formatting

## 0.9.25 (2021-02-26)
* Added PHP Datetime object values in a cells

## 0.9.24 (2021-02-26)
* Percent support


## 0.9.23 (2021-01-25)
* Fix local floats in XML

## 0.9.22 (2020-11-04)
* Added multiple sheets support, thx [Savino59](https://github.com/Savino59), class ready for extend now
 
## 0.9.21 (2020-10-17)
* Updated images

## 0.9.20 (2020-10-04)
* Disable type detection if string started with chr(0)

## 0.9.19 (2020-08-23)
* Numbers like SKU right aligned now

## 0.9.18 (2020-08-22)
* Fixed fast shared strings index
 
## 0.9.17 (2020-08-21)
* Fixed real numbers in 123.45 format detection, fast shared strings index (thx fredriksundin)
 
## 0.9.16 (2020-07-29)
* Fixed time detection in HH:MM:SS format

## 0.9.15 (2020-07-14)
* Escape of shared strings for special chars in cells [#1](https://github.com/shuchkin/simplexlsxgen/issues/1) 

## 0.9.14 (2020-05-31)
* Fixed num2name A-Z,AA-AZ column names, thx Ertan Yusufoglu

## 0.9.13 (2020-05-21)
* If string more 160 chars, save as inlineStr

## 0.9.12 (2020-05-21)
* Readme fixed

## 0.9.11 (2020-05-21)
* Removed XML unimportant attributes

## 0.9.10 (2020-05-20)
* Initial release