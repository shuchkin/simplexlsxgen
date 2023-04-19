<?php

/** @noinspection ReturnTypeCanBeDeclaredInspection */
/** @noinspection PhpMissingReturnTypeInspection */
/** @noinspection NullCoalescingOperatorCanBeUsedInspection */

/** @noinspection PhpIssetCanBeReplacedWithCoalesceInspection */

namespace Shuchkin;

/**
 * Class SimpleXLSXGen
 * Export data to MS Excel. PHP XLSX generator
 * Author: sergey.shuchkin@gmail.com
 */
class SimpleXLSXGen
{
    public $curSheet;
    protected $defaultFont;
    protected $defaultFontSize;
    protected $rtl;
    protected $sheets;
    protected $template;
    protected $NF; // numFmts
    protected $NF_KEYS;
    protected $XF; // cellXfs
    protected $XF_KEYS;
    protected $BR_STYLE;
    protected $SI; // shared strings
    protected $SI_KEYS;
    protected $extLinkId;

    protected $title;
    protected $subject;
    protected $author;
    protected $company;
    protected $manager;
    protected $description;
    protected $application;
    protected $keywords;
    protected $category;
    protected $lastModifiedBy;
    const N_NORMAL = 0; // General
    const N_INT = 1; // 0
    const N_DEC = 2; // 0.00
    const N_PERCENT_INT = 9; // 0%
    const N_PRECENT_DEC = 10; // 0.00%
    const N_DATE = 14; // mm-dd-yy
    const N_TIME = 20; // h:mm
    const N_RUB = 164;
    const N_DOLLAR = 165;
    const N_EURO = 166;
    const N_DATETIME = 22; // m/d/yy h:mm
    const F_NORMAL = 0;
    const F_HYPERLINK = 1;
    const F_BOLD = 2;
    const F_ITALIC = 4;
    const F_UNDERLINE = 8;
    const F_STRIKE = 16;
    const F_COLOR = 32;
    const FL_NONE = 0; // none
    const FL_SOLID = 1; // solid
    const FL_MEDIUM_GRAY = 2; // mediumGray
    const FL_DARK_GRAY = 4; // darkGray
    const FL_LIGHT_GRAY = 8; // lightGray
    const FL_GRAY_125 = 16; // gray125
    const FL_COLOR = 32;
    const A_DEFAULT = 0;
    const A_LEFT = 1;
    const A_RIGHT = 2;
    const A_CENTER = 4;
    const A_TOP = 8;
    const A_MIDDLE = 16;
    const A_BOTTOM = 32;
    const A_WRAPTEXT = 64;
    const B_NONE = 0;
    const B_THIN = 1;
    const B_MEDIUM = 2;
    //const
    const B_DASHED = 3;
    const B_DOTTED = 4;
    const B_THICK = 5;
    const B_DOUBLE = 6;
    const B_HAIR = 7;
    const B_MEDIUM_DASHED = 8;
    const B_DASH_DOT = 9;
    const B_MEDIUM_DASH_DOT = 10;
    const B_DASH_DOT_DOT = 11;
    const B_MEDIUM_DASH_DOT_DOT = 12;
    const B_SLANT_DASH_DOT = 13;

    public function __construct()
    {
        $this->subject = '';
        $this->title = '';
        $this->author = 'Sergey Shuchkin <sergey.shuchkin@gmail.com>';
        $this->company = 'Sergey Shuchkin <sergey.shuchkin@gmail.com>';
        $this->manager = 'Sergey Shuchkin <sergey.shuchkin@gmail.com>';
        $this->description = '';
        $this->keywords = '';
        $this->category = '';
        $this->lastModifiedBy = 'Sergey Shuchkin <sergey.shuchkin@gmail.com>';
        $this->application = __CLASS__;

        $this->curSheet = -1;
        $this->defaultFont = 'Calibri';
        $this->defaultFontSize = 10;
        $this->rtl = false;
        $this->sheets = [['name' => 'Sheet1', 'rows' => [], 'hyperlinks' => [], 'mergecells' => [], 'colwidth' => [], 'autofilter' => '']];
        $this->extLinkId = 0;
        $this->SI = [];        // sharedStrings index
        $this->SI_KEYS = []; //  & keys

        // https://c-rex.net/projects/samples/ooxml/e1/Part4/OOXML_P4_DOCX_numFmts_topic_ID0E6KK6.html
        $this->NF = [
            self::N_RUB => '#,##0.00\ "₽"',
            self::N_DOLLAR => '[$$-1]#,##0.00',
            self::N_EURO => '#,##0.00\ [$€-1]'
        ];
        $this->NF_KEYS = array_flip($this->NF);

        $this->BR_STYLE = [
            self::B_NONE => 'none',
            self::B_THIN => 'thin',
            self::B_MEDIUM => 'medium',
            self::B_DASHED => 'dashed',
            self::B_DOTTED => 'dotted',
            self::B_THICK => 'thick',
            self::B_DOUBLE => 'double',
            self::B_HAIR => 'hair',
            self::B_MEDIUM_DASHED => 'mediumDashed',
            self::B_DASH_DOT => 'dashDot',
            self::B_MEDIUM_DASH_DOT => 'mediumDashDot',
            self::B_DASH_DOT_DOT => 'dashDotDot',
            self::B_MEDIUM_DASH_DOT_DOT => 'mediumDashDotDot',
            self::B_SLANT_DASH_DOT => 'slantDashDot'
        ];

        $this->XF = [  // styles 0 - num fmt, 1 - align, 2 - font, 3 - fill, 4 - font color, 5 - bgcolor, 6 - border, 7 - font size
            [self::N_NORMAL, self::A_DEFAULT, self::F_NORMAL, self::FL_NONE, 0, 0, '', 0],
            [self::N_NORMAL, self::A_DEFAULT, self::F_NORMAL, self::FL_GRAY_125, 0, 0, '', 0], // hack
        ];
        $this->XF_KEYS[implode('-', $this->XF[0])] = 0; // & keys
        $this->XF_KEYS[implode('-', $this->XF[1])] = 1;
        $this->template = [
            '_rels/.rels' => '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
</Relationships>',
            'docProps/app.xml' => '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties">
<TotalTime>0</TotalTime>
<Application>{APP}</Application>
<Company>{COMPANY}</Company>
<Manager>{MANAGER}</Manager>
</Properties>',
            'docProps/core.xml' => '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
<dcterms:created xsi:type="dcterms:W3CDTF">{DATE}</dcterms:created>
    <dc:title>{TITLE}</dc:title>
    <dc:subject>{SUBJECT}</dc:subject>
    <dc:creator>{AUTHOR}</dc:creator>
    <cp:lastModifiedBy>{LAST_MODIFY_BY}</cp:lastModifiedBy>
    <cp:keywords>{KEYWORD}</cp:keywords>
    <dc:description>{DESCRIPTION}</dc:description>
    <cp:category>{CATEGORY}</cp:category>
    <dc:language>en-US</dc:language>
<dcterms:modified xsi:type="dcterms:W3CDTF">{DATE}</dcterms:modified>
<cp:revision>1</cp:revision>
</cp:coreProperties>',
            'xl/_rels/workbook.xml.rels' => '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
{RELS}
</Relationships>',
            'xl/worksheets/sheet1.xml' => '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<dimension ref="{REF}"/>
{SHEETVIEWS}
{COLS}
<sheetData>{ROWS}</sheetData>
{AUTOFILTER}{MERGECELLS}{HYPERLINKS}
</worksheet>',
            'xl/worksheets/_rels/sheet1.xml.rels' => '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">{HYPERLINKS}</Relationships>',
            'xl/sharedStrings.xml' => '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="{CNT}" uniqueCount="{CNT}">{STRINGS}</sst>',
            'xl/styles.xml' => '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
{NUMFMTS}
{FONTS}
{FILLS}
{BORDERS}
<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" /></cellStyleXfs>
{XF}
<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
</styleSheet>',
            'xl/workbook.xml' => '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<fileVersion appName="{APP}"/>
<sheets>
{SHEETS}
</sheets>
</workbook>',
            '[Content_Types].xml' => '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Override PartName="/rels/.rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
<Override PartName="/xl/_rels/workbook.xml.rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>
<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
{TYPES}
</Types>',
        ];
        // <col min="1" max="1" width="22.1796875" bestFit="1" customWidth="1"/>
        // <row r="1" spans="1:2" x14ac:dyDescent="0.35"><c r="A1" t="s"><v>0</v></c><c r="B1"><v>100</v></c></row><row r="2" spans="1:2" x14ac:dyDescent="0.35"><c r="A2" t="s"><v>1</v></c><c r="B2"><v>200</v></c></row>
        // <si><t>Простой шаблон</t></si><si><t>Будем делать генератор</t></si>
    }

    public static function fromArray(array $rows, $sheetName = null)
    {
        return (new static())->addSheet($rows, $sheetName);
    }

    public function addSheet(array $rows, $name = null)
    {
        $this->curSheet++;
        if ($name === null) { // autogenerated sheet names
            $name = 'Sheet' . ($this->curSheet + 1);
        } else {
            $name = mb_substr($name, 0, 31);
            $names = [];
            foreach ($this->sheets as $sh) {
                $names[mb_strtoupper($sh['name'])] = 1;
            }
            for ($i = 0; $i < 100; $i++) {
                $postfix = ' (' . $i . ')';
                $new_name = ($i === 0) ? $name : $name . $postfix;
                if (mb_strlen($new_name) > 31) {
                    $new_name = mb_substr($name, 0, 31 - mb_strlen($postfix)) . $postfix;
                }
                $NEW_NAME = mb_strtoupper($new_name);
                if (!isset($names[$NEW_NAME])) {
                    $name = $new_name;
                    break;
                }
            }
        }
        $this->sheets[$this->curSheet] = ['name' => $name, 'hyperlinks' => [], 'mergecells' => [], 'colwidth' => [], 'autofilter' => '', 'frozen' => ''];
        if (isset($rows[0]) && is_array($rows[0])) {
            $this->sheets[$this->curSheet]['rows'] = $rows;
        } else {
            $this->sheets[$this->curSheet]['rows'] = [];
        }
        return $this;
    }

    public function __toString()
    {
        $fh = fopen('php://memory', 'wb');
        if (!$fh) {
            return '';
        }
        if (!$this->_write($fh)) {
            fclose($fh);
            return '';
        }
        $size = ftell($fh);
        fseek($fh, 0);
        return (string)fread($fh, $size);
    }

    public function saveAs($filename)
    {
        $fh = fopen($filename, 'wb');
        if (!$fh) {
            return false;
        }
        if (!$this->_write($fh)) {
            fclose($fh);
            return false;
        }
        fclose($fh);
        return true;
    }

    public function download()
    {
        return $this->downloadAs(gmdate('YmdHi') . '.xlsx');
    }

    public function downloadAs($filename)
    {
        $fh = fopen('php://memory', 'wb');
        if (!$fh) {
            return false;
        }
        if (!$this->_write($fh)) {
            fclose($fh);
            return false;
        }
        $size = ftell($fh);
        header('Content-type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        header('Content-Disposition: attachment; filename="' . $filename . '"');
        header('Last-Modified: ' . gmdate('D, d M Y H:i:s \G\M\T', time()));
        header('Content-Length: ' . $size);
        while (ob_get_level()) {
            ob_end_clean();
        }
        fseek($fh, 0);
        fpassthru($fh);
        fclose($fh);
        return true;
    }

    protected function _write($fh)
    {
        $dirSignatureE = "\x50\x4b\x05\x06"; // end of central dir signature
        $zipComments = 'Generated by ' . __CLASS__ . ' PHP class, thanks sergey.shuchkin@gmail.com';
        if (!$fh) {
            return false;
        }
        $cdrec = '';    // central directory content
        $entries = 0;    // number of zipped files
        $cnt_sheets = count($this->sheets);
        foreach ($this->template as $cfilename => $template) {
            if ($cfilename === 'xl/_rels/workbook.xml.rels') {
                $s = '';
                for ($i = 0; $i < $cnt_sheets; $i++) {
                    $s .= '<Relationship Id="rId' . ($i + 1) . '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"' .
                        ' Target="worksheets/sheet' . ($i + 1) . ".xml\"/>\r\n";
                }
                $s .= '<Relationship Id="rId' . ($i + 1) . '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>' . "\r\n";
                $s .= '<Relationship Id="rId' . ($i + 2) . '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>';

                $template = str_replace('{RELS}', $s, $template);
                $this->_writeEntry($fh, $cdrec, $cfilename, $template);
                $entries++;
            } elseif ($cfilename === 'xl/workbook.xml') {
                $s = '';
                foreach ($this->sheets as $k => $v) {
                    $s .= '<sheet name="' . $this->esc($v['name']) . '" sheetId="' . ($k + 1) . '" r:id="rId' . ($k + 1) . '"/>';
                }
                $search = ['{SHEETS}', '{APP}'];
                $replace = [$s, $this->esc($this->application)];
                $template = str_replace($search, $replace, $template);
                $this->_writeEntry($fh, $cdrec, $cfilename, $template);
                $entries++;
            } elseif ($cfilename === 'docProps/app.xml') {
                $search = ['{APP}', '{COMPANY}', '{MANAGER}'];
                $replace = [$this->esc($this->application), $this->esc($this->company), $this->esc($this->manager)];
                $template = str_replace($search, $replace, $template);
                $this->_writeEntry($fh, $cdrec, $cfilename, $template);
                $entries++;
            } elseif ($cfilename === 'docProps/core.xml') {
                $search = ['{DATE}', '{AUTHOR}', '{TITLE}', '{SUBJECT}', '{KEYWORD}', '{DESCRIPTION}', '{CATEGORY}', '{LAST_MODIFY_BY}'];
                $replace = [gmdate('Y-m-d\TH:i:s\Z'), $this->esc($this->author), $this->esc($this->title), $this->esc($this->subject), $this->esc($this->keywords), $this->esc($this->description), $this->esc($this->category), $this->esc($this->lastModifiedBy)];
                $template = str_replace($search, $replace, $template);
                $this->_writeEntry($fh, $cdrec, $cfilename, $template);
                $entries++;
            } elseif ($cfilename === 'xl/sharedStrings.xml') {
                if (!count($this->SI)) {
                    $this->SI[] = 'No Data';
                }
                $si_cnt = count($this->SI);
                $si = '<si><t>' . implode("</t></si>\r\n<si><t>", $this->SI) . '</t></si>';
                $template = str_replace(['{CNT}', '{STRINGS}'], [$si_cnt, $si], $template);
                $this->_writeEntry($fh, $cdrec, $cfilename, $template);
                $entries++;
            } elseif ($cfilename === 'xl/worksheets/sheet1.xml') {
                foreach ($this->sheets as $k => $v) {
                    $filename = 'xl/worksheets/sheet' . ($k + 1) . '.xml';
                    $xml = $this->_sheetToXML($k, $template);
                    $this->_writeEntry($fh, $cdrec, $filename, $xml);
                    $entries++;
                }
                $xml = null;
            } elseif ($cfilename === 'xl/worksheets/_rels/sheet1.xml.rels') {
                foreach ($this->sheets as $k => $v) {
                    if ($this->extLinkId) {
                        $RH = [];
                        $filename = 'xl/worksheets/_rels/sheet' . ($k + 1) . '.xml.rels';
                        foreach ($v['hyperlinks'] as $h) {
                            if ($h['ID']) {
                                $RH[] = '<Relationship Id="' . $h['ID'] . '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="' . $this->esc($h['H']) . '" TargetMode="External"/>';
                            }
                        }
                        $xml = str_replace('{HYPERLINKS}', implode("\r\n", $RH), $template);
                        $this->_writeEntry($fh, $cdrec, $filename, $xml);
                        $entries++;
                    }
                }
                $xml = null;
            } elseif ($cfilename === '[Content_Types].xml') {
                $TYPES = ['<Override PartName="/_rels/.rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'];
                foreach ($this->sheets as $k => $v) {
                    $TYPES[] = '<Override PartName="/xl/worksheets/sheet' . ($k + 1) . '.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>';
                    if ($this->extLinkId) {
                        $TYPES[] = '<Override PartName="/xl/worksheets/_rels/sheet' . ($k + 1) . '.xml.rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>';
                    }
                }
                $template = str_replace('{TYPES}', implode("\r\n", $TYPES), $template);
                $this->_writeEntry($fh, $cdrec, $cfilename, $template);
                $entries++;
            } elseif ($cfilename === 'xl/styles.xml') {
                $NF = $XF = $FONTS = $F_KEYS = $FILLS = $FL_KEYS = [];
                $BR = ['<border><left/><right/><top/><bottom/><diagonal/></border>'];
                $BR_KEYS = [0 => 0];
                foreach ($this->NF as $k => $v) {
                    $NF[] = '<numFmt numFmtId="' . $k . '" formatCode="' . htmlspecialchars($v, ENT_QUOTES) . '"/>';
                }
                foreach ($this->XF as $xf) {
                    // 0 - num fmt, 1 - align, 2 - font, 3 - fill, 4 - font color, 5 - bgcolor, 6 - border, 7 - font size
                    // fonts
                    $F_KEY = $xf[2] . '-' . $xf[4] . '-' . $xf[7];
                    if (isset($F_KEYS[$F_KEY])) {
                        $F_ID = $F_KEYS[$F_KEY];
                    } else {
                        $F_ID = $F_KEYS[$F_KEY] = count($FONTS);
                        $FONTS[] = '<font><name val="' . $this->defaultFont . '"/><family val="2"/>'
                            . ($xf[7] ? '<sz val="' . $xf[7] . '"/>' : '<sz val="' . $this->defaultFontSize . '"/>')
                            . ($xf[2] & self::F_BOLD ? '<b/>' : '')
                            . ($xf[2] & self::F_ITALIC ? '<i/>' : '')
                            . ($xf[2] & self::F_UNDERLINE ? '<u/>' : '')
                            . ($xf[2] & self::F_STRIKE ? '<strike/>' : '')
                            . ($xf[2] & self::F_HYPERLINK ? '<u/>' : '')
                            . ($xf[2] & self::F_COLOR ? '<color rgb="' . $xf[4] . '"/>' : '')
                            . '</font>';
                    }
                    // fills
                    $FL_KEY = $xf[3] . '-' . $xf[5];
                    if (isset($FL_KEYS[$FL_KEY])) {
                        $FL_ID = $FL_KEYS[$FL_KEY];
                    } else {
                        $FL_ID = $FL_KEYS[$FL_KEY] = count($FILLS);
                        $FILLS[] = '<fill><patternFill patternType="'
                            . ($xf[3] === 0 ? 'none' : '')
                            . ($xf[3] & self::FL_SOLID ? 'solid' : '')
                            . ($xf[3] & self::FL_MEDIUM_GRAY ? 'mediumGray' : '')
                            . ($xf[3] & self::FL_DARK_GRAY ? 'darkGray' : '')
                            . ($xf[3] & self::FL_LIGHT_GRAY ? 'lightGray' : '')
                            . ($xf[3] & self::FL_GRAY_125 ? 'gray125' : '')
                            . '"'
                            . ($xf[3] & self::FL_COLOR ? '><fgColor rgb="' . $xf[5] . '"/><bgColor indexed="64"/></patternFill>' : ' />')
                            . '</fill>';
                    }
                    $align = '';
                    if ($xf[1] & self::A_LEFT) {
                        $align .= ' horizontal="left"';
                    } elseif ($xf[1] & self::A_RIGHT) {
                        $align .= ' horizontal="right"';
                    } elseif ($xf[1] & self::A_CENTER) {
                        $align .= ' horizontal="center"';
                    }
                    if ($xf[1] & self::A_TOP) {
                        $align .= ' vertical="top"';
                    } elseif ($xf[1] & self::A_MIDDLE) {
                        $align .= ' vertical="center"';
                    } elseif ($xf[1] & self::A_BOTTOM) {
                        $align .= ' vertical="bottom"';
                    }
                    if ($xf[1] & self::A_WRAPTEXT) {
                        $align .= ' wrapText="1"';
                    }

                    // border
                    $BR_ID = 0;
                    if ($xf[6] !== '') {
                        $b = $xf[6];
                        if (isset($BR_KEYS[$b])) {
                            $BR_ID = $BR_KEYS[$b];
                        } else {
                            $BR_ID = count($BR_KEYS);
                            $BR_KEYS[$b] = $BR_ID;
                            $border = '<border>';
                            $ba = explode(' ', $b);
                            if (!isset($ba[1])) {
                                $ba[] = $ba[0];
                                $ba[] = $ba[0];
                                $ba[] = $ba[0];
                            }
                            if (!isset($ba[4])) { // diagonal
                                $ba[] = 'none';
                            }
                            $sides = ['left' => 3, 'right' => 1, 'top' => 0, 'bottom' => 2, 'diagonal' => 4];
                            foreach ($sides as $side => $idx) {
                                $s = 'thin';
                                $c = '';
                                $va = explode('#', $ba[$idx]);
                                if (isset($va[1])) {
                                    $s = $va[0] === '' ? 'thin' : $va[0];
                                    $c = $va[1];
                                } elseif (in_array($va[0], $this->BR_STYLE, true)) {
                                    $s = $va[0];
                                } else {
                                    $c = $va[0];
                                }
                                if (strlen($c) === 6) {
                                    $c = 'FF' . $c;
                                }
                                if ($s && $s !== 'none') {
                                    $border .= '<' . $side . ' style="' . $s . '">'
                                        . '<color ' . ($c === '' ? 'auto="1"' : 'rgb="' . $c . '"') . '/>'
                                        . '</' . $side . '>';
                                } else {
                                    $border .= '<' . $side . '/>';
                                }
                            }
                            $border .= '</border>';
                            $BR[] = $border;
                        }
                    }
                    $XF[] = '<xf numFmtId="' . $xf[0] . '" fontId="' . $F_ID . '" fillId="' . $FL_ID . '" borderId="' . $BR_ID . '" xfId="0"'
                        . ($xf[0] > 0 ? ' applyNumberFormat="1"' : '')
                        . ($F_ID > 0 ? ' applyFont="1"' : '')
                        . ($FL_ID > 0 ? ' applyFill="1"' : '')
                        . ($BR_ID > 0 ? ' applyBorder="1"' : '')
                        . ($align ? ' applyAlignment="1"><alignment' . $align . '/></xf>' : '/>');
                }
                // wrap collections
                array_unshift($NF, '<numFmts count="' . count($NF) . '">');
                $NF[] = '</numFmts>';
                array_unshift($XF, '<cellXfs count="' . count($XF) . '">');
                $XF[] = '</cellXfs>';
                array_unshift($FONTS, '<fonts count="' . count($FONTS) . '">');
                $FONTS[] = '</fonts>';
                array_unshift($FILLS, '<fills count="' . count($FILLS) . '">');
                $FILLS[] = '</fills>';
                array_unshift($BR, '<borders count="' . count($BR) . '">');
                $BR[] = '</borders>';

                $template = str_replace(
                    ['{NUMFMTS}', '{FONTS}', '{XF}', '{FILLS}', '{BORDERS}'],
                    [implode("\r\n", $NF), implode("\r\n", $FONTS), implode("\r\n", $XF), implode("\r\n", $FILLS), implode("\r\n", $BR)],
                    $template
                );
                $this->_writeEntry($fh, $cdrec, $cfilename, $template);
                $entries++;
            } else {
                $this->_writeEntry($fh, $cdrec, $cfilename, $template);
                $entries++;
            }
        }
        $before_cd = ftell($fh);
        fwrite($fh, $cdrec);
        // end of central dir
        fwrite($fh, $dirSignatureE);
        fwrite($fh, pack('v', 0)); // number of this disk
        fwrite($fh, pack('v', 0)); // number of the disk with the start of the central directory
        fwrite($fh, pack('v', $entries)); // total # of entries "on this disk"
        fwrite($fh, pack('v', $entries)); // total # of entries overall
        fwrite($fh, pack('V', mb_strlen($cdrec, '8bit')));     // size of central dir
        fwrite($fh, pack('V', $before_cd));         // offset to start of central dir
        fwrite($fh, pack('v', mb_strlen($zipComments, '8bit'))); // .zip file comment length
        fwrite($fh, $zipComments);

        return true;
    }

    protected function _writeEntry($fh, &$cdrec, $cfilename, $data)
    {
        $zipSignature = "\x50\x4b\x03\x04"; // local file header signature
        $dirSignature = "\x50\x4b\x01\x02"; // central dir header signature

        $e = [];
        $e['uncsize'] = mb_strlen($data, '8bit');
        // if data to compress is too small, just store it
        if ($e['uncsize'] < 256) {
            $e['comsize'] = $e['uncsize'];
            $e['vneeded'] = 10;
            $e['cmethod'] = 0;
            $zdata = $data;
        } else { // otherwise, compress it
            $zdata = gzcompress($data);
            $zdata = substr(substr($zdata, 0, -4), 2); // fix crc bug (thanks to Eric Mueller)
            $e['comsize'] = mb_strlen($zdata, '8bit');
            $e['vneeded'] = 10;
            $e['cmethod'] = 8;
        }
        $e['bitflag'] = 0;
        $e['crc_32'] = crc32($data);

        // Convert date and time to DOS Format, and set then
        $lastmod_timeS = str_pad(decbin(date('s') >= 32 ? date('s') - 32 : date('s')), 5, '0', STR_PAD_LEFT);
        $lastmod_timeM = str_pad(decbin(date('i')), 6, '0', STR_PAD_LEFT);
        $lastmod_timeH = str_pad(decbin(date('H')), 5, '0', STR_PAD_LEFT);
        $lastmod_dateD = str_pad(decbin(date('d')), 5, '0', STR_PAD_LEFT);
        $lastmod_dateM = str_pad(decbin(date('m')), 4, '0', STR_PAD_LEFT);
        $lastmod_dateY = str_pad(decbin(date('Y') - 1980), 7, '0', STR_PAD_LEFT);

        $e['modtime'] = bindec("$lastmod_timeH$lastmod_timeM$lastmod_timeS");
        $e['moddate'] = bindec("$lastmod_dateY$lastmod_dateM$lastmod_dateD");
        $e['offset'] = ftell($fh);

        fwrite($fh, $zipSignature);
        fwrite($fh, pack('s', $e['vneeded'])); // version_needed
        fwrite($fh, pack('s', $e['bitflag'])); // general_bit_flag
        fwrite($fh, pack('s', $e['cmethod'])); // compression_method
        fwrite($fh, pack('s', $e['modtime'])); // lastmod_time
        fwrite($fh, pack('s', $e['moddate'])); // lastmod_date
        fwrite($fh, pack('V', $e['crc_32']));  // crc-32
        fwrite($fh, pack('I', $e['comsize'])); // compressed_size
        fwrite($fh, pack('I', $e['uncsize'])); // uncompressed_size
        fwrite($fh, pack('s', mb_strlen($cfilename, '8bit')));   // file_name_length
        fwrite($fh, pack('s', 0));  // extra_field_length
        fwrite($fh, $cfilename);    // file_name
        // ignoring extra_field
        fwrite($fh, $zdata);

        // Append it to central dir
        $e['external_attributes'] = (substr($cfilename, -1) === '/' && !$zdata) ? 16 : 32; // Directory or file name
        $e['comments'] = '';

        $cdrec .= $dirSignature;
        $cdrec .= "\x0\x0";                                     // version made by
        $cdrec .= pack('v', $e['vneeded']);                     // version needed to extract
        $cdrec .= "\x0\x0";                                     // general bit flag
        $cdrec .= pack('v', $e['cmethod']);                     // compression method
        $cdrec .= pack('v', $e['modtime']);                     // lastmod time
        $cdrec .= pack('v', $e['moddate']);                     // lastmod date
        $cdrec .= pack('V', $e['crc_32']);                      // crc32
        $cdrec .= pack('V', $e['comsize']);                     // compressed filesize
        $cdrec .= pack('V', $e['uncsize']);                     // uncompressed filesize
        $cdrec .= pack('v', mb_strlen($cfilename, '8bit'));     // file name length
        $cdrec .= pack('v', 0);                                 // extra field length
        $cdrec .= pack('v', mb_strlen($e['comments'], '8bit')); // file comment length
        $cdrec .= pack('v', 0);                                 // disk number start
        $cdrec .= pack('v', 0);                                 // internal file attributes
        $cdrec .= pack('V', $e['external_attributes']);         // internal file attributes
        $cdrec .= pack('V', $e['offset']);                      // relative offset of local header
        $cdrec .= $cfilename;
        $cdrec .= $e['comments'];
    }

    protected function _sheetToXML($idx, $template)
    {
        // locale floats fr_FR 1.234,56 -> 1234.56
        $_loc = setlocale(LC_NUMERIC, 0);
        setlocale(LC_NUMERIC, 'C');
        $COLS = [];
        $ROWS = [];
        //        $SHEETVIEWS = '<sheetViews><sheetView tabSelected="1" workbookViewId="0"'.($this->rtl ? ' rightToLeft="1"' : '').'>';
        $SHEETVIEWS = '';
        $PANE = '';
        if (count($this->sheets[$idx]['rows'])) {
            if ($this->sheets[$idx]['frozen'] !== '' || isset($this->sheets[$idx]['frozen'][0]) || isset($this->sheets[$idx]['frozen'][1])) {
                //                $AC = 'A1'; // Active Cell
                $x = $y = 0;
                if (is_string($this->sheets[$idx]['frozen'])) {
                    $AC = $this->sheets[$idx]['frozen'];
                    self::cell2coord($AC, $x, $y);
                } else {
                    if (isset($this->sheets[$idx]['frozen'][0])) {
                        $x = $this->sheets[$idx]['frozen'][0];
                    }
                    if (isset($this->sheets[$idx]['frozen'][1])) {
                        $y = $this->sheets[$idx]['frozen'][1];
                    }
                    $AC = self::coord2cell($x, $y);
                }
                if ($x > 0 || $y > 0) {
                    $split = '';
                    if ($x > 0) {
                        $split .= ' xSplit="' . $x . '"';
                    }
                    if ($y > 0) {
                        $split .= ' ySplit="' . $y . '"';
                    }
                    $activepane = 'bottomRight';
                    if ($x > 0 && $y === 0) {
                        $activepane = 'topRight';
                    }
                    if ($x === 0 && $y > 0) {
                        $activepane = 'bottomLeft';
                    }
                    $PANE .= '<pane' . $split . ' topLeftCell="' . $AC . '" activePane="' . $activepane . '" state="frozen"/>';
                    $PANE .= '<selection activeCell="' . $AC . '" sqref="' . $AC . '"/>';
                }
            }
            if ($this->rtl || $PANE) {
                $SHEETVIEWS .= '<sheetViews>
<sheetView workbookViewId="0"' . ($this->rtl ? ' rightToLeft="1"' : '');
                $SHEETVIEWS .= $PANE ? ">\r\n" . $PANE . "\r\n</sheetView>" : ' />';
                $SHEETVIEWS .= "\r\n</sheetViews>";
            }
            $COLS[] = '<cols>';
            $CUR_ROW = 0;
            $COL = [];
            foreach ($this->sheets[$idx]['rows'] as $r) {
                $CUR_ROW++;
                $row = '';
                $CUR_COL = 0;
                $RH = 0; // row height
                foreach ($r as $v) {
                    $CUR_COL++;
                    if (!isset($COL[$CUR_COL])) {
                        $COL[$CUR_COL] = 0;
                    }
                    $cname = $this->num2name($CUR_COL) . $CUR_ROW;
                    if ($v === null || $v === '') {
                        $row .= '<c r="' . $cname . '"/>';
                        continue;
                    }
                    $ct = $cv = $cf = null;
                    $N = $A = $F = $FL = $C = $BG = $FS = 0;
                    $BR = '';
                    if (is_string($v)) {
                        if ($v[0] === "\0") { // RAW value as string
                            $v = substr($v, 1);
                            $vl = mb_strlen($v);
                        } else {
                            if (strpos($v, '<') !== false) { // tags?
                                if (strpos($v, '<b>') !== false) {
                                    $F += self::F_BOLD;
                                }
                                if (strpos($v, '<i>') !== false) {
                                    $F += self::F_ITALIC;
                                }
                                if (strpos($v, '<u>') !== false) {
                                    $F += self::F_UNDERLINE;
                                }
                                if (strpos($v, '<s>') !== false) {
                                    $F += self::F_STRIKE;
                                }
                                if (preg_match('/<style([^>]+)>/', $v, $m)) {
                                    if (preg_match('/ color="([^"]+)"/', $m[1], $m2)) {
                                        $F += self::F_COLOR;
                                        $c = ltrim($m2[1], '#');
                                        $C = strlen($c) === 8 ? $c : ('FF' . $c);
                                    }
                                    if (preg_match('/ bgcolor="([^"]+)"/', $m[1], $m2)) {
                                        $FL += self::FL_COLOR;
                                        $c = ltrim($m2[1], '#');
                                        $BG = strlen($c) === 8 ? $c : ('FF' . $c);
                                    }
                                    if (preg_match('/ height="([^"]+)"/', $m[1], $m2)) {
                                        $RH = $m2[1];
                                    }
                                    if (preg_match('/ nf="([^"]+)"/', $m[1], $m2)) {
                                        $c = htmlspecialchars_decode($m2[1], ENT_QUOTES);
                                        $N = $this->getNumFmtId($c);
                                    }
                                    if (preg_match('/ border="([^"]+)"/', $m[1], $m2)) {
                                        $b = htmlspecialchars_decode($m2[1], ENT_QUOTES);
                                        if ($b && $b !== 'none') {
                                            $BR = $b;
                                        }
                                    }
                                    if (preg_match('/ font-size="([^"]+)"/', $m[1], $m2)) {
                                        $FS = (int)$m2[1];
                                        if ($RH === 0) { // fix row height
                                            $RH = ($FS > $this->defaultFontSize) ? round($FS * 1.50, 1) : 0;
                                        }
                                    }
                                }
                                if (strpos($v, '<left>') !== false) {
                                    $A += self::A_LEFT;
                                }
                                if (strpos($v, '<center>') !== false) {
                                    $A += self::A_CENTER;
                                }
                                if (strpos($v, '<right>') !== false) {
                                    $A += self::A_RIGHT;
                                }
                                if (strpos($v, '<top>') !== false) {
                                    $A += self::A_TOP;
                                }
                                if (strpos($v, '<middle>') !== false) {
                                    $A += self::A_MIDDLE;
                                }
                                if (strpos($v, '<bottom>') !== false) {
                                    $A += self::A_BOTTOM;
                                }
                                if (strpos($v, '<wraptext>') !== false) {
                                    $A += self::A_WRAPTEXT;
                                }
                                if (preg_match('/<a href="(https?:\/\/[^"]+)">(.*?)<\/a>/i', $v, $m)) {
                                    $h = explode('#', $m[1]);
                                    $this->extLinkId++;
                                    $this->sheets[$idx]['hyperlinks'][] = ['ID' => 'rId' . $this->extLinkId, 'R' => $cname, 'H' => $h[0], 'L' => isset($h[1]) ? $h[1] : ''];
                                    $F += self::F_HYPERLINK; // Hyperlink
                                }
                                if (preg_match('/<a href="(mailto?:[^"]+)">(.*?)<\/a>/i', $v, $m)) {
                                    $this->extLinkId++;
                                    $this->sheets[$idx]['hyperlinks'][] = ['ID' => 'rId' . $this->extLinkId, 'R' => $cname, 'H' => $m[1], 'L' => ''];
                                    $F += self::F_HYPERLINK; // mailto hyperlink
                                }
                                if (preg_match('/<a href="([^"]+![^"]+)">(.*?)<\/a>/i', $v, $m)) {
                                    $this->sheets[$idx]['hyperlinks'][] = ['ID' => null, 'R' => $cname, 'H' => null, 'L' => $m[1]];
                                    $F += self::F_HYPERLINK; // internal hyperlink
                                }
                                if (preg_match('/<f([^>]*)>/', $v, $m)) {
                                    $cf = strip_tags($v);
                                    $v = 'formula';
                                    if (preg_match('/ v="([^"]+)"/', $m[1], $m2)) {
                                        $v = $m2[1];
                                    }
                                } else {
                                    $v = strip_tags($v);
                                }
                            } // tags
                            $vl = mb_strlen($v);
                            if ($N) {
                                $cv = ltrim($v, '+');
                            } elseif ($v === '0' || preg_match('/^[-+]?[1-9]\d{0,14}$/', $v)) { // Integer as General
                                $cv = ltrim($v, '+');
                                if ($vl > 10) {
                                    $N = self::N_INT; // [1] 0
                                }
                            } elseif (preg_match('/^[-+]?(0|[1-9]\d*)\.(\d+)$/', $v, $m)) {
                                $cv = ltrim($v, '+');
                                if (strlen($m[2]) < 3) {
                                    $N = self::N_DEC;
                                }
                            } elseif (preg_match('/^\$[-+]?[0-9\.]+$/', $v)) { // currency $?
                                $N = self::N_DOLLAR;
                                $cv = ltrim($v, '+$');
                            } elseif (preg_match('/^[-+]?[0-9\.]+( ₽| €)$/u', $v, $m)) { // currency ₽ €?
                                if ($m[1] === ' ₽') {
                                    $N = self::N_RUB;
                                } elseif ($m[1] === ' €') {
                                    $N = self::N_EURO;
                                }
                                $cv = trim($v, ' +₽€');
                            } elseif (preg_match('/^([-+]?\d+)%$/', $v, $m)) {
                                $cv = round($m[1] / 100, 2);
                                $N = self::N_PERCENT_INT; // [9] 0%
                            } elseif (preg_match('/^([-+]?\d+\.\d+)%$/', $v, $m)) {
                                $cv = round($m[1] / 100, 4);
                                $N = self::N_PRECENT_DEC; // [10] 0.00%
                            } elseif (preg_match('/^(\d\d\d\d)-(\d\d)-(\d\d)$/', $v, $m)) {
                                $cv = $this->date2excel($m[1], $m[2], $m[3]);
                                $N = self::N_DATE; // [14] mm-dd-yy
                            } elseif (preg_match('/^(\d\d)\/(\d\d)\/(\d\d\d\d)$/', $v, $m)) {
                                $cv = $this->date2excel($m[3], $m[2], $m[1]);
                                $N = self::N_DATE; // [14] mm-dd-yy
                            } elseif (preg_match('/^(\d\d):(\d\d):(\d\d)$/', $v, $m)) {
                                $cv = $this->date2excel(0, 0, 0, $m[1], $m[2], $m[3]);
                                $N = self::N_TIME; // time
                            } elseif (preg_match('/^(\d\d\d\d)-(\d\d)-(\d\d) (\d\d):(\d\d):(\d\d)$/', $v, $m)) {
                                $cv = $this->date2excel($m[1], $m[2], $m[3], $m[4], $m[5], $m[6]);
                                $N = ((int)$m[1] === 0) ? self::N_TIME : self::N_DATETIME; // [22] m/d/yy h:mm
                            } elseif (preg_match('/^(\d\d)\/(\d\d)\/(\d\d\d\d) (\d\d):(\d\d):(\d\d)$/', $v, $m)) {
                                $cv = $this->date2excel($m[3], $m[2], $m[1], $m[4], $m[5], $m[6]);
                                $N = self::N_DATETIME; // [22] m/d/yy h:mm
                            } elseif (preg_match('/^[0-9+-.]+$/', $v)) { // Long ?
                                $A += ($A & (self::A_LEFT | self::A_CENTER)) ? 0 : self::A_RIGHT;
                            } elseif (preg_match('/^https?:\/\/\S+$/i', $v)) {
                                $h = explode('#', $v);
                                $this->extLinkId++;
                                $this->sheets[$idx]['hyperlinks'][] = ['ID' => 'rId' . $this->extLinkId, 'R' => $cname, 'H' => $h[0], 'L' => isset($h[1]) ? $h[1] : ''];
                                $F += self::F_HYPERLINK; // Hyperlink
                            } elseif (preg_match("/^[a-zA-Z0-9_\.\-]+@([a-zA-Z0-9][a-zA-Z0-9\-]*\.)+[a-zA-Z]{2,}$/", $v)) {
                                $this->extLinkId++;
                                $this->sheets[$idx]['hyperlinks'][] = ['ID' => 'rId' . $this->extLinkId, 'R' => $cname, 'H' => 'mailto:' . $v, 'L' => ''];
                                $F += self::F_HYPERLINK; // Hyperlink
                            }
                            if (($N === self::N_DATE || $N === self::N_DATETIME) && $cv < 0) {
                                $cv = null;
                                $N = 0;
                            }
                        }
                        if ($cv === null) {
                            $v = $this->esc($v);
                            if ($cf) {
                                $ct = 'str';
                                $cv = $v;
                            } elseif (mb_strlen($v) > 160) {
                                $ct = 'inlineStr';
                                $cv = $v;
                            } else {
                                $ct = 's'; // shared string
                                $cv = false;
                                $skey = '~' . $v;
                                if (isset($this->SI_KEYS[$skey])) {
                                    $cv = $this->SI_KEYS[$skey];
                                }
                                if ($cv === false) {
                                    $this->SI[] = $v;
                                    $cv = count($this->SI) - 1;
                                    $this->SI_KEYS[$skey] = $cv;
                                }
                            }
                        }
                    } elseif (is_int($v)) {
                        $vl = mb_strlen((string)$v);
                        $cv = $v;
                    } elseif (is_float($v)) {
                        $vl = mb_strlen((string)$v);
                        $cv = $v;
                    } elseif ($v instanceof \DateTime) {
                        $vl = 16;
                        $cv = $this->date2excel($v->format('Y'), $v->format('m'), $v->format('d'), $v->format('H'), $v->format('i'), $v->format('s'));
                        $N = self::N_DATETIME; // [22] m/d/yy h:mm
                    } else {
                        continue;
                    }
                    $COL[$CUR_COL] = max($vl, $COL[$CUR_COL]);
                    $cs = 0;
                    if (($N + $A + $F + $FL + $FS > 0) || $BR !== '') {
                        if ($FL === self::FL_COLOR) {
                            $FL += self::FL_SOLID;
                        }
                        if (($F & self::F_HYPERLINK) && !($F & self::F_COLOR)) {
                            $F += self::F_COLOR;
                            $C = 'FF0563C1';
                        }
                        $XF_KEY = $N . '-' . $A . '-' . $F . '-' . $FL . '-' . $C . '-' . $BG . '-' . $BR . '-' . $FS;
                        if (isset($this->XF_KEYS[$XF_KEY])) {
                            $cs = $this->XF_KEYS[$XF_KEY];
                        }
                        if ($cs === 0) {
                            $cs = count($this->XF);
                            $this->XF_KEYS[$XF_KEY] = $cs;
                            $this->XF[] = [$N, $A, $F, $FL, $C, $BG, $BR, $FS];
                        }
                    }
                    $row .= '<c r="' . $cname . '"' . ($ct ? ' t="' . $ct . '"' : '') . ($cs ? ' s="' . $cs . '"' : '') . '>'
                        . ($cf ? '<f>' . $cf . '</f>' : '')
                        . ($ct === 'inlineStr' ? '<is><t>' . $cv . '</t></is>' : '<v>' . $cv . '</v>') . "</c>\r\n";
                }
                $ROWS[] = '<row r="' . $CUR_ROW . '"' . ($RH ? ' customHeight="1" ht="' . $RH . '"' : '') . '>' . $row . "</row>";
            }
            foreach ($COL as $k => $max) {
                $w = isset($this->sheets[$idx]['colwidth'][$k]) ? $this->sheets[$idx]['colwidth'][$k] : min($max + 1, 60);
                $COLS[] = '<col min="' . $k . '" max="' . $k . '" width="' . $w . '" customWidth="1" />';
            }
            $COLS[] = '</cols>';
            $REF = 'A1:' . $this->num2name(count($COL)) . $CUR_ROW;
        } else {
            $ROWS[] = '<row r="1"><c r="A1" t="s"><v>0</v></c></row>';
            $REF = 'A1:A1';
        }

        $AUTOFILTER = '';
        if ($this->sheets[$idx]['autofilter']) {
            $AUTOFILTER = '<autoFilter ref="' . $this->sheets[$idx]['autofilter'] . '" />';
        }

        $MERGECELLS = [];
        if (count($this->sheets[$idx]['mergecells'])) {
            $MERGECELLS[] = '';
            $MERGECELLS[] = '<mergeCells count="' . count($this->sheets[$idx]['mergecells']) . '">';
            foreach ($this->sheets[$idx]['mergecells'] as $m) {
                $MERGECELLS[] = '<mergeCell ref="' . $m . '"/>';
            }
            $MERGECELLS[] = '</mergeCells>';
        }

        $HYPERLINKS = [];
        if (count($this->sheets[$idx]['hyperlinks'])) {
            $HYPERLINKS[] = '<hyperlinks>';
            foreach ($this->sheets[$idx]['hyperlinks'] as $h) {
                $HYPERLINKS[] = '<hyperlink ref="' . $h['R'] . '"' . ($h['ID'] ? ' r:id="' . $h['ID'] . '"' : '') . ' location="' . $this->esc($h['L']) . '" display="' . $this->esc($h['H'] . ($h['L'] ? ' - ' . $h['L'] : '')) . '" />';
            }
            $HYPERLINKS[] = '</hyperlinks>';
        }

        //restore locale
        setlocale(LC_NUMERIC, $_loc);

        return str_replace(
            ['{REF}', '{COLS}', '{ROWS}', '{AUTOFILTER}', '{MERGECELLS}', '{HYPERLINKS}', '{SHEETVIEWS}'],
            [
                $REF,
                implode("\r\n", $COLS),
                implode("\r\n", $ROWS),
                $AUTOFILTER,
                implode("\r\n", $MERGECELLS),
                implode("\r\n", $HYPERLINKS),
                $SHEETVIEWS
            ],
            $template
        );
    }

    public function num2name($num)
    {
        $numeric = ($num - 1) % 26;
        $letter = chr(65 + $numeric);
        $num2 = (int)(($num - 1) / 26);
        if ($num2 > 0) {
            return $this->num2name($num2) . $letter;
        }
        return $letter;
    }

    public function date2excel($year, $month, $day, $hours = 0, $minutes = 0, $seconds = 0)
    {
        $excelTime = (($hours * 3600) + ($minutes * 60) + $seconds) / 86400;
        if ((int)$year === 0) {
            return $excelTime;
        }
        // self::CALENDAR_WINDOWS_1900
        $excel1900isLeapYear = True;
        if (($year === 1900) && ($month <= 2)) {
            $excel1900isLeapYear = False;
        }
        $myExcelBaseDate = 2415020;
        // Julian base date Adjustment
        if ($month > 2) {
            $month -= 3;
        } else {
            $month += 9;
            --$year;
        }
        $century = substr($year, 0, 2);
        $decade = substr($year, 2, 2);
        //    Calculate the Julian Date, then subtract the Excel base date (JD 2415020 = 31-Dec-1899 Giving Excel Date of 0)
        $excelDate = floor((146097 * $century) / 4) + floor((1461 * $decade) / 4) + floor((153 * $month + 2) / 5) + $day + 1721119 - $myExcelBaseDate + $excel1900isLeapYear;
        return (float)$excelDate + $excelTime;
    }

    public function setDefaultFont($name)
    {
        $this->defaultFont = $name;
        return $this;
    }

    public function setDefaultFontSize($size)
    {
        $this->defaultFontSize = $size;
        return $this;
    }

    public function setTitle($title)
    {
        $this->title = $title;
        return $this;
    }
    public function setSubject($subject)
    {
        $this->subject = $subject;
        return $this;
    }
    public function setAuthor($author)
    {
        $this->author = $author;
        return $this;
    }
    public function setCompany($company)
    {
        $this->company = $company;
        return $this;
    }
    public function setManager($manager)
    {
        $this->manager = $manager;
        return $this;
    }
    public function setKeywords($keywords)
    {
        $this->keywords = $keywords;
        return $this;
    }
    public function setDescription($description)
    {
        $this->description = $description;
        return $this;
    }
    public function setCategory($category)
    {
        $this->category = $category;
        return $this;
    }

    public function setApplication($application)
    {
        $this->application = $application;
        return $this;
    }
    public function setLastModifiedBy($lastModifiedBy)
    {
        $this->lastModifiedBy = $lastModifiedBy;
        return $this;
    }

    public function autoFilter($range)
    {
        $this->sheets[$this->curSheet]['autofilter'] = $range;
        return $this;
    }

    public function mergeCells($range)
    {
        $this->sheets[$this->curSheet]['mergecells'][] = $range;
        return $this;
    }

    public function setColWidth($col, $width)
    {
        $this->sheets[$this->curSheet]['colwidth'][$col] = $width;
        return $this;
    }
    public function rightToLeft($value = true)
    {
        $this->rtl = $value;
        return $this;
    }

    public function esc($str)
    {
        // XML UTF-8: #x9 | #xA | #xD | [#x20-#xD7FF] | [#xE000-#xFFFD] | [#x10000-#x10FFFF]
        // but we use fast version
        return str_replace(['&', '<', '>', "\x00", "\x03", "\x0B"], ['&amp;', '&lt;', '&gt;', '', '', ''], $str);
    }

    public function getNumFmtId($code)
    {
        if (isset($this->NF[$code])) { // id?
            return (int)$code;
        }
        if (isset($this->NF_KEYS[$code])) {
            return $this->NF_KEYS[$code];
        }
        $id = 197 + count($this->NF); // custom
        $this->NF[$id] = $code;
        $this->NF_KEYS[$code] = $id;
        return $id;
    }


    public static function raw($value)
    {
        return "\0" . $value;
    }

    public static function cell2coord($cell, &$x, &$y)
    {
        $x = $y = 0;
        $lettercount = 0;
        $cell = str_replace([' ', '\t', '\r', '\n', '\v', '\0'], '', $cell);
        if (empty($cell)) {
            return;
        }
        $cell = strtoupper($cell);
        for ($i = 0, $len = strlen($cell); $i < $len; $i++) {
            if ($cell[$i] >= 'A' && $cell[$i] <= 'Z') {
                $lettercount++;
            }
        }
        if ($lettercount > 0) {
            $x = ord($cell[$lettercount - 1]) - ord('A');
            $e = 1;
            for ($i = $lettercount - 2; $i >= 0; $i--) {
                $x += (ord($cell[$i]) - ord('A') + 1) * (26 ** $e);
                $e++;
            }
        }
        if ($lettercount < strlen($cell)) {
            $y = ((int)substr($cell, $lettercount)) - 1;
        }
    }

    public static function coord2cell($x, $y)
    {
        $c = '';
        for ($i = $x; $i >= 0; $i = ((int)($i / 26)) - 1) {
            $c = chr(ord('A') + $i % 26) . $c;
        }
        return $c . ($y + 1);
    }

    public function freezePanes($cell)
    {
        $this->sheets[$this->curSheet]['frozen'] = $cell;
        return $this;
    }
}
