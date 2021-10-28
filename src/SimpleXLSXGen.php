<?php
/**
 * Class SimpleXLSXGen
 * Export data to MS Excel. PHP XLSX generator
 * Author: sergey.shuchkin@gmail.com
 */

class SimpleXLSXGen {

	public $curSheet;
	protected $defaultFont;
	protected $defaultFontSize;
	protected $sheets;
	protected $template;
	protected $F, $F_KEYS; // fonts
	protected $XF, $XF_KEYS; // cellXfs
	protected $SI, $SI_KEYS; // shared strings
	const N_NORMAL = 0; // General
	const N_INT = 1; // 0
	const N_DEC = 2; // 0.00
	const N_PERCENT_INT = 9; // 0%
	const N_PRECENT_DEC = 10; // 0.00%
	const N_DATE = 14; // mm-dd-yy
	const N_TIME = 20; // h:mm
	const N_DATETIME = 22; // m/d/yy h:mm
	const F_NORMAL = 0;
	const F_HYPERLINK = 1;
	const F_BOLD = 2;
	const F_ITALIC = 4;
	const F_UNDERLINE = 8;
	const F_STRIKE = 16;
	const A_DEFAULT = 0;
	const A_LEFT = 1;
	const A_RIGHT = 2;
	const A_CENTER = 3;


	public function __construct() {
		$this->curSheet = -1;
		$this->defaultFont = 'Calibri';
		$this->sheets = [ ['name' => 'Sheet1', 'rows' => [], 'hyperlinks' => [] ] ];
		$this->SI = [];		// sharedStrings index
		$this->SI_KEYS = []; //  & keys
		$this->F = [ self::F_NORMAL ]; // fonts
		$this->F_KEYS = [0]; // & keys
		$this->XF  = [ [self::N_NORMAL, self::F_NORMAL, self::A_DEFAULT] ]; // styles
		$this->XF_KEYS = ['N0F0A0' => 0 ]; // & keys

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
<Application>'.__CLASS__.'</Application></Properties>',
			'docProps/core.xml' => '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
<dcterms:created xsi:type="dcterms:W3CDTF">{DATE}</dcterms:created>
<dc:language>en-US</dc:language>
<dcterms:modified xsi:type="dcterms:W3CDTF">{DATE}</dcterms:modified>
<cp:revision>1</cp:revision>
</cp:coreProperties>',
			'xl/_rels/workbook.xml.rels' => '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
{SHEETS}',
			'xl/worksheets/sheet1.xml' => '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
	xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
><dimension ref="{REF}"/>{COLS}<sheetData>{ROWS}</sheetData>{HYPERLINKS}</worksheet>',
			'xl/worksheets/_rels/sheet1.xml.rels' => '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">{HYPERLINKS}</Relationships>',
			'xl/sharedStrings.xml' => '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="{CNT}" uniqueCount="{CNT}">{STRINGS}</sst>',
			'xl/styles.xml' => '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
{FONTS}
<fills count="1"><fill><patternFill patternType="none"/></fill></fills>
<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" /></cellStyleXfs>
{XF}
<cellStyles count="1">
	<cellStyle name="Normal" xfId="0" builtinId="0"/>
</cellStyles>
</styleSheet>',
			'xl/workbook.xml' => '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<fileVersion appName="'.__CLASS__.'"/><sheets>
{SHEETS}
</sheets></workbook>',
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
	public static function fromArray( array $rows, $sheetName = null ) {
		$xlsx = new static();
		return $xlsx->addSheet( $rows, $sheetName );
	}

	public function addSheet( array $rows, $name = null ) {

		$this->curSheet++;
		if ( $name === null ) { // autogenerated sheet names
			$name = 'Sheet'.($this->curSheet+1);
		} else {
			$names = [];
			foreach( $this->sheets as $sh ) {
				$names[ mb_strtoupper( $sh['name']) ] = 1;
			}
			for( $i = 0; $i < 100; $i++ ) {
				$new_name = ($i === 0) ? $name : $name .' ('.$i.')';
				$NEW_NAME = mb_strtoupper( $new_name );
				if ( !isset( $names[ $NEW_NAME ]) ) {
					$name = $new_name;
					break;
				}
			}
		}

		$this->sheets[$this->curSheet] = ['name' => $name, 'hyperlinks' => []];

		if ( is_array( $rows ) && isset( $rows[0] ) && is_array($rows[0]) ) {
			$this->sheets[$this->curSheet]['rows'] = $rows;
		} else {
			$this->sheets[$this->curSheet]['rows'] = [];
		}
		return $this;
	}

	public function __toString() {
		$fh = fopen( 'php://memory', 'wb' );
		if ( ! $fh ) {
			return '';
		}

		if ( ! $this->_write( $fh ) ) {
			fclose( $fh );
			return '';
		}
		$size = ftell( $fh );
		fseek( $fh, 0);

		return (string) fread( $fh, $size );
	}

	public function saveAs( $filename ) {
		$fh = fopen( $filename, 'wb' );
		if (!$fh) {
			return false;
		}
		if ( !$this->_write($fh) ) {
			fclose($fh);
			return false;
		}
		fclose($fh);

		return true;
	}

	public function download() {
		return $this->downloadAs( gmdate('YmdHi') . '.xlsx' );
	}

	public function downloadAs( $filename ) {
		$fh = fopen('php://memory','wb');
		if (!$fh) {
			return false;
		}

		if ( !$this->_write( $fh )) {
			fclose( $fh );
			return false;
		}

		$size = ftell($fh);

		header('Content-type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
		header('Content-Disposition: attachment; filename="'.$filename.'"');
		header('Last-Modified: ' . gmdate('D, d M Y H:i:s \G\M\T' , time() ));
		header('Content-Length: '.$size);

		while( ob_get_level() ) {
			ob_end_clean();
		}
		fseek($fh,0);
		fpassthru( $fh );

		fclose($fh);
		return true;
	}

	protected function _write( $fh ) {


		$dirSignatureE= "\x50\x4b\x05\x06"; // end of central dir signature
		$zipComments = 'Generated by '.__CLASS__.' PHP class, thanks sergey.shuchkin@gmail.com';

		if (!$fh) {
			return false;
		}

		$cdrec = '';	// central directory content
		$entries= 0;	// number of zipped files
		$cnt_sheets = count( $this->sheets );

		foreach ($this->template as $cfilename => $template ) {
			if ( $cfilename === 'xl/_rels/workbook.xml.rels' ) {
				$s = '';
				for ( $i = 0; $i < $cnt_sheets; $i++) {
					$s .= '<Relationship Id="rId'.($i+2).'" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"'.
						 ' Target="worksheets/sheet'.($i+1).".xml\"/>\n";
				}
				$s .= '<Relationship Id="rId'.($i+2).'" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/></Relationships>';
				$template = str_replace('{SHEETS}', $s, $template);
				$this->_writeEntry($fh, $cdrec, $cfilename, $template);
				$entries++;
			} elseif ( $cfilename === 'xl/workbook.xml' ) {
				$s = '';
				foreach ( $this->sheets as $k => $v ) {
					$s .= '<sheet name="' . $this->esc( $v['name'] ) . '" sheetId="' . ( $k + 1) . '" state="visible" r:id="rId' . ( $k + 2) . '"/>';
				}
				$template = str_replace('{SHEETS}', $s, $template);
				$this->_writeEntry($fh, $cdrec, $cfilename, $template);
				$entries++;
			}
			elseif ( $cfilename === 'docProps/core.xml' ) {
				$template = str_replace('{DATE}', gmdate('Y-m-d\TH:i:s\Z'), $template);
				$this->_writeEntry($fh, $cdrec, $cfilename, $template);
				$entries++;
			} elseif ( $cfilename === 'xl/sharedStrings.xml' ) {
				if (!count($this->SI)) {
					$this->SI[] = 'No Data';
				}
				$si_cnt = count($this->SI);
				$si = '<si><t>'.implode("</t></si>\r\n<si><t>", $this->SI).'</t></si>';
				$template = str_replace(['{CNT}', '{STRINGS}'], [ $si_cnt, $si ], $template );
				$this->_writeEntry($fh, $cdrec, $cfilename, $template);
				$entries++;
			} elseif ( $cfilename === 'xl/worksheets/sheet1.xml' ) {
				foreach ( $this->sheets as $k => $v ) {
					$filename = 'xl/worksheets/sheet'.($k+1).'.xml';
					$xml = $this->_sheetToXML($k, $template);
					$this->_writeEntry($fh, $cdrec, $filename, $xml );
					$entries++;
				}
				$xml = null;
			} elseif ( $cfilename === 'xl/worksheets/_rels/sheet1.xml.rels' ) {
				foreach ( $this->sheets as $k => $v ) {
					if ( count($v['hyperlinks'])) {
						$RH = [];
						$filename = 'xl/worksheets/_rels/sheet' . ( $k + 1 ) . '.xml.rels';
						foreach ( $v['hyperlinks'] as $h ) {
							$RH[] = '<Relationship Id="' . $h['ID'] . '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="' . $this->esc($h['H']) . '" TargetMode="External"/>';
						}
						$xml = str_replace( '{HYPERLINKS}', implode( "\r\n", $RH ), $template );
						$this->_writeEntry( $fh, $cdrec, $filename, $xml );
						$entries++;
					}
				}
				$xml = null;

			} elseif ( $cfilename === '[Content_Types].xml' ) {
				$TYPES = ['<Override PartName="/_rels/.rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'];
				foreach ( $this->sheets as $k => $v) {
					$TYPES[] = '<Override PartName="/xl/worksheets/sheet'.($k+1).
							'.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>';
					if (count( $v['hyperlinks'])) {
						$TYPES[] = '<Override PartName="/xl/worksheets/_rels/sheet'.($k+1).'.xml.rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>';
					}
				}
				$template = str_replace('{TYPES}', implode("\r\n", $TYPES), $template);
				$this->_writeEntry($fh, $cdrec, $cfilename, $template);
				$entries++;
			} elseif ( $cfilename === 'xl/styles.xml' ) {
				$FONTS = ['<fonts count="'.count($this->F).'">'];
				foreach ( $this->F as $f ) {
					$FONTS[] = '<font><name val="'.$this->defaultFont.'"/><family val="2"/>'
						. ( $this->defaultFontSize ? '<sz val="'.$this->defaultFontSize.'"/>' : '' )
						.( $f & self::F_BOLD ? '<b/>' : '')
						.( $f & self::F_ITALIC ? '<i/>' : '')
						.( $f & self::F_UNDERLINE ? '<u/>' : '')
						.( $f & self::F_STRIKE ? '<strike/>' : '')
						.( $f & self::F_HYPERLINK ? '<color rgb="FF0563C1"/><u/>' : '')
						.'</font>';
				}
				$FONTS[] = '</fonts>';
				$XF = ['<cellXfs count="'.count($this->XF).'">'];
				foreach( $this->XF as $xf ) {
					$align = ($xf[2] === self::A_LEFT ? ' applyAlignment="1"><alignment horizontal="left"/>' : '')
						.($xf[2] === self::A_RIGHT ? ' applyAlignment="1"><alignment horizontal="right"/>' : '')
						.($xf[2] === self::A_CENTER ? ' applyAlignment="1"><alignment horizontal="center"/>' : '');
					$XF[] = '<xf numFmtId="'.$xf[0].'" fontId="'.$xf[1].'" fillId="0" borderId="0" xfId="0"'
						.($xf[0] > 0 ? ' applyNumberFormat="1"' : '')
						.($align ? $align . '</xf>' : '/>');

				}
				$XF[] = '</cellXfs>';
				$template = str_replace(['{FONTS}','{XF}'], [implode("\r\n", $FONTS), implode("\r\n", $XF)], $template);
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
		fwrite($fh, pack('V', mb_strlen($cdrec,'8bit')));     // size of central dir
		fwrite($fh, pack('V', $before_cd));         // offset to start of central dir
		fwrite($fh, pack('v', mb_strlen($zipComments,'8bit'))); // .zip file comment length
		fwrite($fh, $zipComments);

		return true;
	}

	protected function _writeEntry($fh, &$cdrec, $cfilename, $data) {
		$zipSignature = "\x50\x4b\x03\x04"; // local file header signature
		$dirSignature = "\x50\x4b\x01\x02"; // central dir header signature

		$e = [];
		$e['uncsize'] = mb_strlen($data, '8bit');

		// if data to compress is too small, just store it
		if($e['uncsize'] < 256){
			$e['comsize'] = $e['uncsize'];
			$e['vneeded'] = 10;
			$e['cmethod'] = 0;
			$zdata = $data;
		} else{ // otherwise, compress it
			$zdata = gzcompress($data);
			$zdata = substr(substr($zdata, 0, - 4 ), 2); // fix crc bug (thanks to Eric Mueller)
			$e['comsize'] = mb_strlen($zdata, '8bit');
			$e['vneeded'] = 10;
			$e['cmethod'] = 8;
		}

		$e['bitflag'] = 0;
		$e['crc_32']  = crc32($data);

		// Convert date and time to DOS Format, and set then
		$lastmod_timeS  = str_pad(decbin(date('s')>=32?date('s')-32:date('s')), 5, '0', STR_PAD_LEFT);
		$lastmod_timeM  = str_pad(decbin(date('i')), 6, '0', STR_PAD_LEFT);
		$lastmod_timeH  = str_pad(decbin(date('H')), 5, '0', STR_PAD_LEFT);
		$lastmod_dateD  = str_pad(decbin(date('d')), 5, '0', STR_PAD_LEFT);
		$lastmod_dateM  = str_pad(decbin(date('m')), 4, '0', STR_PAD_LEFT);
		$lastmod_dateY  = str_pad(decbin(date('Y')-1980), 7, '0', STR_PAD_LEFT);

		# echo "ModTime: $lastmod_timeS-$lastmod_timeM-$lastmod_timeH (".date("s H H").")\n";
		# echo "ModDate: $lastmod_dateD-$lastmod_dateM-$lastmod_dateY (".date("d m Y").")\n";
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
		$e['external_attributes']  = (substr($cfilename, -1) === '/'&&!$zdata)?16:32; // Directory or file name
		$e['comments']             = '';

		$cdrec .= $dirSignature;
		$cdrec .= "\x0\x0";                  // version made by
		$cdrec .= pack('v', $e['vneeded']); // version needed to extract
		$cdrec .= "\x0\x0";                  // general bit flag
		$cdrec .= pack('v', $e['cmethod']); // compression method
		$cdrec .= pack('v', $e['modtime']); // lastmod time
		$cdrec .= pack('v', $e['moddate']); // lastmod date
		$cdrec .= pack('V', $e['crc_32']);  // crc32
		$cdrec .= pack('V', $e['comsize']); // compressed filesize
		$cdrec .= pack('V', $e['uncsize']); // uncompressed filesize
		$cdrec .= pack('v', mb_strlen($cfilename,'8bit')); // file name length
		$cdrec .= pack('v', 0);                // extra field length
		$cdrec .= pack('v', mb_strlen($e['comments'],'8bit')); // file comment length
		$cdrec .= pack('v', 0); // disk number start
		$cdrec .= pack('v', 0); // internal file attributes
		$cdrec .= pack('V', $e['external_attributes']); // internal file attributes
		$cdrec .= pack('V', $e['offset']); // relative offset of local header
		$cdrec .= $cfilename;
		$cdrec .= $e['comments'];
	}

	protected function _sheetToXML($idx, $template) {
		// locale floats fr_FR 1.234,56 -> 1234.56
		$_loc = setlocale(LC_NUMERIC, 0);
		setlocale(LC_NUMERIC,'C');
		$COLS = [];
		$ROWS = [];
		if ( count($this->sheets[$idx]['rows']) ) {
			$COLS[] = '<cols>';
			$CUR_ROW = 0;
			$COL = [];
			foreach( $this->sheets[$idx]['rows'] as $r ) {
				$CUR_ROW++;
				$row = '<row r="'.$CUR_ROW.'">';
				$CUR_COL = 0;
				foreach( $r as $v ) {
					$CUR_COL++;
					if ( !isset($COL[ $CUR_COL ])) {
						$COL[ $CUR_COL ] = 0;
					}
					if ( $v === null || $v === '' ) {
						continue;
					}

					$cname = $this->num2name($CUR_COL) . $CUR_ROW;

					$ct = $cv = null;
					$N = $F = $A = 0;

					if ( is_string($v) ) {

						if ( $v[0] === "\0" ) { // RAW value as string
							$v = substr($v,1);
							$vl = mb_strlen( $v );
						} else {
							if ( strpos( $v, '<' ) !== false ) { // tags?
								if ( strpos( $v, '<b>' ) !== false ) {
									$F += self::F_BOLD;
								}
								if ( strpos( $v, '<i>' ) !== false ) {
									$F += self::F_ITALIC;
								}
								if ( strpos( $v, '<u>' ) !== false ) {
									$F += self::F_UNDERLINE;
								}
								if ( strpos( $v, '<s>' ) !== false ) {
									$F += self::F_STRIKE;
								}
								if ( strpos( $v, '<left>' ) !== false ) {
									$A += self::A_LEFT;
								}
								if ( strpos( $v, '<center>' ) !== false ) {
									$A += self::A_CENTER;
								}
								if ( strpos( $v, '<right>' ) !== false ) {
									$A += self::A_RIGHT;
								}
								if ( preg_match( '/<a href="(https?:\/\/[^"]+)">(.*?)<\/a>/i', $v, $m ) ) {
									$h = explode( '#', $m[1] );
									$this->sheets[ $idx ]['hyperlinks'][] = ['ID' => 'rId' . ( count( $this->sheets[ $idx ]['hyperlinks'] ) + 1 ), 'R' => $cname, 'H' => $h[0], 'L' => isset( $h[1] ) ? $h[1] : ''];
									$F = self::F_HYPERLINK; // Hyperlink
								}
								if ( preg_match( '/<a href="(mailto?:[^"]+)">(.*?)<\/a>/i', $v, $m ) ) {
									$this->sheets[ $idx ]['hyperlinks'][] = ['ID' => 'rId' . ( count( $this->sheets[ $idx ]['hyperlinks'] ) + 1 ), 'R' => $cname, 'H' => $m[1], 'L' => ''];
									$F = self::F_HYPERLINK; // mailto hyperlink
								}
								$v = strip_tags( $v );
							} // tags
							$vl = mb_strlen( $v );
							if ( $v === '0' || preg_match( '/^[-+]?[1-9]\d{0,14}$/', $v ) ) { // Integer as General
								$cv = ltrim( $v, '+' );
								if ( $vl > 10 ) {
									$N = self::N_INT; // [1] 0
								}
							} elseif ( preg_match( '/^[-+]?(0|[1-9]\d*)\.(\d+)$/', $v, $m ) ) {
								$cv = ltrim( $v, '+' );
								if ( strlen( $m[2] ) < 3 ) {
									$N = self::N_DEC;
								}
							} elseif ( preg_match( '/^([-+]?\d+)%$/', $v, $m ) ) {
								$cv = round( $m[1] / 100, 2 );
								$N = self::N_PERCENT_INT; // [9] 0%
							} elseif ( preg_match( '/^([-+]?\d+\.\d+)%$/', $v, $m ) ) {
								$cv = round( $m[1] / 100, 4 );
								$N = self::N_PRECENT_DEC; // [10] 0.00%
							} elseif ( preg_match( '/^(\d\d\d\d)-(\d\d)-(\d\d)$/', $v, $m ) ) {
								$cv = $this->date2excel( $m[1], $m[2], $m[3] );
								$N = self::N_DATE; // [14] mm-dd-yy
							} elseif ( preg_match( '/^(\d\d)\/(\d\d)\/(\d\d\d\d)$/', $v, $m ) ) {
								$cv = $this->date2excel( $m[3], $m[2], $m[1] );
								$N = self::N_DATE; // [14] mm-dd-yy
							} elseif ( preg_match( '/^(\d\d):(\d\d):(\d\d)$/', $v, $m ) ) {
								$cv = $this->date2excel( 0, 0, 0, $m[1], $m[2], $m[3] );
								$N = self::N_TIME; // time
							} elseif ( preg_match( '/^(\d\d\d\d)-(\d\d)-(\d\d) (\d\d):(\d\d):(\d\d)$/', $v, $m ) ) {
								$cv = $this->date2excel( $m[1], $m[2], $m[3], $m[4], $m[5], $m[6] );
								$N = self::N_DATETIME; // [22] m/d/yy h:mm
							} elseif ( preg_match( '/^(\d\d)\/(\d\d)\/(\d\d\d\d) (\d\d):(\d\d):(\d\d)$/', $v, $m ) ) {
								$cv = $this->date2excel( $m[3], $m[2], $m[1], $m[4], $m[5], $m[6] );
								$N = self::N_DATETIME; // [22] m/d/yy h:mm
							} elseif ( preg_match( '/^[0-9+-.]+$/', $v ) ) { // Long ?
								$A = self::A_RIGHT;
							} elseif ( preg_match( '/^https?:\/\/\S+$/i', $v ) ) {
								$h = explode( '#', $v );
								$this->sheets[ $idx ]['hyperlinks'][] = ['ID' => 'rId' . ( count( $this->sheets[ $idx ]['hyperlinks'] ) + 1 ), 'R' => $cname, 'H' => $h[0], 'L' => isset( $h[1] ) ? $h[1] : ''];
								$F = self::F_HYPERLINK; // Hyperlink
							} elseif ( preg_match( "/^[a-zA-Z0-9_\.\-]+@([a-zA-Z0-9][a-zA-Z0-9\-]*\.)+[a-zA-Z]{2,}$/", $v ) ) {
								$this->sheets[ $idx ]['hyperlinks'][] = ['ID' => 'rId' . ( count( $this->sheets[ $idx ]['hyperlinks'] ) + 1 ), 'R' => $cname, 'H' => 'mailto:' . $v, 'L' => ''];
								$F = self::F_HYPERLINK; // Hyperlink
							}
						}
						if ( !$cv) {

							$v = $this->esc( $v );

							if ( mb_strlen( $v ) > 160 ) {
								$ct = 'inlineStr';
								$cv = $v;
							} else {
								$ct = 's'; // shared string
								$cv = false;
								$skey = '~' . $v;
								if ( isset( $this->SI_KEYS[ $skey ] ) ) {
									$cv = $this->SI_KEYS[ $skey ];
								}
								if ( $cv === false ) {
									$this->SI[] = $v;
									$cv = count( $this->SI ) - 1;
									$this->SI_KEYS[ $skey ] = $cv;
								}
							}
						}
					} elseif ( is_int( $v ) ) {
						$vl = mb_strlen( (string) $v );
						$cv = $v;
					} elseif ( is_float( $v ) ) {
						$vl = mb_strlen( (string) $v );
						$cv = $v;
					} elseif ( $v instanceof DateTime ) {
						$vl = 16;
						$cv = $this->date2excel( $v->format('Y'), $v->format('m'), $v->format('d'), $v->format('H'), $v->format('i'), $v->format('s') );
						$N = self::N_DATETIME; // [22] m/d/yy h:mm
					} else {
						continue;
					}

					$COL[ $CUR_COL ] = max( $vl, $COL[ $CUR_COL ] );

					$cs = 0;
					if ( $N + $F + $A > 0 ) {

						if ( isset($this->F_KEYS[ $F ] ) ) {
							$cf = $this->F_KEYS[ $F ];
						} else {
							$cf = count($this->F);
							$this->F_KEYS[$F] = $cf;
							$this->F[] = $F;
						}
						$NFA = 'N' . $N . 'F' . $cf . 'A' . $A;
						if ( isset( $this->XF_KEYS[ $NFA ] ) ) {
							$cs = $this->XF_KEYS[ $NFA ];
						}
						if ( $cs === 0 ) {
							$cs = count( $this->XF );
							$this->XF_KEYS[ $NFA ] = $cs;
							$this->XF[] = [$N, $cf, $A];
						}
					}

					$row .= '<c r="' . $cname . '"'.($ct ? ' t="'.$ct.'"' : '').($cs ? ' s="'.$cs.'"' : '').'>'
					        .($ct === 'inlineStr' ? '<is><t>'.$cv.'</t></is>' : '<v>' . $cv . '</v>')."</c>\r\n";
				}
				$ROWS[] = $row . "</row>\r\n";
			}
			foreach ( $COL as $k => $max ) {
				$COLS[] = '<col min="'.$k.'" max="'.$k.'" width="'.min( $max+1, 60).'" />';
			}
			$COLS[] = '</cols>';
			$REF = 'A1:'.$this->num2name(count($COLS)) . $CUR_ROW;
		} else {
			$ROWS[] = '<row r="1"><c r="A1" t="s"><v>0</v></c></row>';
			$REF = 'A1:A1';
		}
		$HYPERLINKS = [];
		if ( count( $this->sheets[$idx]['hyperlinks']) ) {
			$HYPERLINKS[] = '<hyperlinks>';
			foreach ( $this->sheets[$idx]['hyperlinks'] as $h ) {
				$HYPERLINKS[] = '<hyperlink ref="' . $h['R'] . '" r:id="' . $h['ID'] . '" location="' . $this->esc( $h['L'] ) . '" display="' . $this->esc( $h['H'] . ( $h['L'] ? ' - ' . $h['L'] : '' ) ) . '" />';
			}
			$HYPERLINKS[] = '</hyperlinks>';
		}
		//restore locale
		setlocale(LC_NUMERIC, $_loc);

		return str_replace(['{REF}','{COLS}','{ROWS}','{HYPERLINKS}'],
			[ $REF, implode("\r\n", $COLS), implode("\r\n",$ROWS), implode("\r\n", $HYPERLINKS) ],
			$template );
	}

	public function num2name($num) {
		$numeric = ($num - 1) % 26;
		$letter  = chr( 65 + $numeric );
		$num2    = (int) ( ($num-1) / 26 );
		if ( $num2 > 0 ) {
			return $this->num2name( $num2 ) . $letter;
		}
		return $letter;
	}

	public function date2excel($year, $month, $day, $hours=0, $minutes=0, $seconds=0) {
	    $excelTime = (($hours * 3600) + ($minutes * 60) + $seconds) / 86400;

	    if ( $year === 0 ) {
	        return $excelTime;
        }

		// self::CALENDAR_WINDOWS_1900
		$excel1900isLeapYear = True;
		if (((int)$year === 1900) && ($month <= 2)) { $excel1900isLeapYear = False; }
		$myExcelBaseDate = 2415020;

		//    Julian base date Adjustment
		if ($month > 2) {
			$month -= 3;
		} else {
			$month += 9;
			--$year;
		}
		//    Calculate the Julian Date, then subtract the Excel base date (JD 2415020 = 31-Dec-1899 Giving Excel Date of 0)
		$century = substr($year,0,2);
		$decade = substr($year,2,2);
		$excelDate = floor((146097 * $century) / 4) + floor((1461 * $decade) / 4) + floor((153 * $month + 2) / 5) + $day + 1721119 - $myExcelBaseDate + $excel1900isLeapYear;

		return (float) $excelDate + $excelTime;
	}
	public function setDefaultFont( $name ) {
		$this->defaultFont = $name;
		return $this;
	}
	public function setDefaultFontSize( $size ) {
		$this->defaultFontSize = $size;
		return $this;
	}
	public function esc( $str ) {
		// XML UTF-8: #x9 | #xA | #xD | [#x20-#xD7FF] | [#xE000-#xFFFD] | [#x10000-#x10FFFF]
		// but we use fast version
		return str_replace( ['&', '<', '>', "\x00","\x03","\x0B"], ['&amp;', '&lt;', '&gt;', '', '', ''], $str );
	}
}
