<?php
class ExcelPHP // Version V1.0.0  - Timothy Edwards - 14 Dec 2024
{
	private $debug = false;
	private $file;
	private $sheet_xml;
	private $styles_xml;
	private $encoding = 'UTF-8';
	private $SW; //option for column widths
	private $PR; //Spreadsheet or printout option
	private $tmpDir = 'images'; //default directory for images
	private $Sheetname = [];
	private $Sheetno = [];
	private $Sheetid = [];
	private $Sheetnum;
	private $shared = [];
	private $FSFactor = 13; //Font size conversion factor
	private $CWFactor = 7; //column width conversion factor
	private $Defwidth = 60; //default width of a column
	private $RHFactor = 1.3; //row height conversion factor
	private $stylecount;
	private $DevF; //Default font
	private $DevS; //Default font size
	private $themecol = [];
	private $Chyper = [];
	private $Cindex = array('8' => '000000', 'FFFFFF', 'FF0000', '00FF00', '0000FF', 'FFFF00', 'FF00FF', '00FFFF', '800000', '008000', '000080', '808000', '800080', '008080', 'C0C0C0', '808080', '9999FF', '993366', 'FFFFCC', 'CCFFFF', '660066', 'FF8080', '0066CC', 'CCCCFF', '000080', 'FF00FF', 'FFFF00', '00FFFF', '800080', '800000', '008080', '0000FF', '00CCFF', 'CCFFFF', 'CCFFCC', 'FFFF99', '99CCFF', 'FF99CC', 'CC99FF', 'FFCC99', '3366FF', '33CCCC', '99CC00', 'FFCC00', 'FF9900', 'FF6600', '666699', '969696', '003366', '339966', '003300', '333300', '993300', '993366', '333399', '333333', '000000'); //Array of rgb hex index colours

	
	/**
	 * CONSTRUCTOR
	 * 
	 * @param Boolean $debug Debug mode or not
	 * @param String $encoding selects alternative encoding if required
	 * @param String $tmpDir selects alternative image directory if required
	 * @return void
	 */
	public function __construct($debug_=null, $encoding=null, $tmpDir=null)
	{
		if($debug_ != null) {
			$this->debug = $debug_;
		}
		if ($encoding != null) {
			$this->encoding = $encoding;
		}
		if ($tmpDir != null) {
			$this->tmpDir = $tmpDir;
		}
	}



	/**
	 * Converts the column ref number to column characters
	 * 
	 * @param str - the number ref for a column
	 * @return str - The column chars name
	 */
	private function numtochars($colcount)
	{

		if ($colcount < 27){
			$Acolcount = chr($colcount + 64);
		} else if ($colcount < 677){
			$vchar1 = (int)($colcount/26);
			$vchar2 = $colcount - ($vchar1 * 26);
			if ($vchar2 == 0){
				--$vchar1;
				$vchar2 = 26;
			}
			$Acolcount = chr($vchar1 + 64).chr($vchar2 + 64);
		}
		return $Acolcount;
	}



	/**
	 * Converts the cell reference to its column number equivalent and row number
	 * 
	 * @param str - the cell name
	 * @return array - The cells column and row number
	 */
	private function charstonum($cellname)
	{
		$CharF = preg_replace("/[^A-Z]/", '', $cellname);
		$Fchar1 = ord(substr($CharF,0,1)) - 64;
		$Fchar2 = ord(substr($CharF,1,1)) - 64;
		if (strlen($CharF) == 1){
			$Ncol['char'] = $Fchar1;
		} else {
			$Ncol['char'] = ($Fchar1 * 26) + $Fchar2;
		}
		$Ncol['num'] = preg_replace("/[^0-9]/", '', $cellname);
		return($Ncol);
	}




	/**
	 * READS The Workbook file and Relationships into separated XML files
	 * 
	 * @param var $object The class variable to set as DOMDocument 
	 * @param var $xml The xml file
	 * @param string $encoding The encoding to be used
	 * @return void
	 */
	private function setXmlParts(&$object, $xml, $encoding)
	{
		$object = new DOMDocument();
		$object->encoding = $encoding;
		$object->preserveWhiteSpace = false;
		$object->formatOutput = true;
		$object->loadXML($xml);
		$object->saveXML();
	}


	/**
	 * READS The Workbook, Styles, Shared Strings and Relationships files into separated XML files
	 * and puts the theme colours and Shared Strings into arrays
	 *
	 * @param none
	 * @return void
	 */
	private function readZipPart()
	{
		$zip = new ZipArchive();
		$_xml = 'xl/workbook.xml';
		$_xml_theme = 'xl/theme/theme1.xml';
		$_xml_rels = 'xl/_rels/workbook.xml.rels';
		$_xml_shared = 'xl/sharedStrings.xml';
		
		if (true === $zip->open($this->file)) {
			//Get the main excel workbook file
			if (($index = $zip->locateName($_xml)) !== false) {
				$xml = $zip->getFromIndex($index);
			}
			//Get the styles
			if (($index = $zip->locateName($_xml_theme)) !== false) {
				$xml_theme = $zip->getFromIndex($index);
			}
			//Get the relationships
			if (($index = $zip->locateName($_xml_rels)) !== false) {
				$xml_rels = $zip->getFromIndex($index);
			}
			//Get the shared strings
			if (($index = $zip->locateName($_xml_shared)) !== false) {
				$xml_shared = $zip->getFromIndex($index);
			}
			$zip->close();
		} else die('ERROR - non zip file');

		$enc = mb_detect_encoding($xml);
		$this->setXmlParts($xl_xml, $xml, $enc);
		$this->setXmlParts($theme_xml, $xml_theme, $enc);
		$this->setXmlParts($rels_xml, $xml_rels, $enc);
		if ($xml_shared){
			$this->setXmlParts($shared_xml, $xml_shared, $enc);
		}
		
		if($this->debug) {
			echo "XML File : xl/workbook.xml<br>";
			echo "<textarea style='width:100%; height: 200px;'>";
			echo $xl_xml->saveXML();
			echo "</textarea>";
			echo "<br>XML File : xl/theme/theme1.xml<br>";
			echo "<textarea style='width:100%; height: 200px;'>";
			echo $theme_xml->saveXML();
			echo "</textarea>";
			echo "<br>XML File : xl/_rels/workbook.xml.rels<br>";
			echo "<textarea style='width:100%; height: 200px;'>";
			echo $rels_xml->saveXML();
			echo "</textarea>";
			if ($xml_shared){
				echo "<br>XML File : xl/sharedStrings.xml<br>";
				echo "<textarea style='width:100%; height: 200px;'>";
				echo $shared_xml->saveXML();
				echo "</textarea>";
			}
		}
		
		//Find sheet names and ids
		$reader1 = new XMLReader();
		$reader1->XML($xl_xml->saveXML());
		$sn = 0;
		while ($reader1->read()) {
			if ($reader1->name == 'sheet') {
				$this->Sheetname[$sn] = $reader1->getAttribute("name");
				$this->Sheetno[$sn] = $reader1->getAttribute("sheetId"); // not currently used
				$this->Sheetid[$sn] = $reader1->getAttribute("r:id"); // not currently used
				++$sn;
			}
		}
		$this->Sheetnum = $sn;
		
		//Read theme colours into an array
		$reader3 = new XMLReader();
		$reader3->XML($theme_xml->saveXML());
		$c=0;
		while ($reader3->read()) {
			if ($reader3->name == 'a:sysClr') {
				$wt = $reader3->getAttribute("lastClr");
				if ($c == 0){
					$this->themecol[1] = $wt;
				} else if ($c == 1){
					$this->themecol[0] = $wt;
				}
				++$c;
			}
			if ($reader3->name == 'a:srgbClr') {
				$st = $reader3->getAttribute("val");
				if ($c == 2){
					$this->themecol[3] = $st;
				} else if ($c == 3){
					$this->themecol[2] = $st;
				} else if ($c> 9){
					$tt = $c-10;
					$this->Chyper[$tt] = $st;
				} else {
					$this->themecol[$c] = $st;
				}
				++$c;
			}
		}


		if ($xml_shared){
			//Read shared strings into an array
			$reader2 = new XMLReader();
			$reader2->XML($shared_xml->saveXML());
			$sh = 0;
			$si = $r = $fbold = $fund = $fital = $fsize = $fcol = $fname = $ftext = '';
			while ($reader2->read()) {
				if ($reader2->nodeType == XMLREADER::ELEMENT && $reader2->name == 'si') {
				$si = 'Y';
				}
				if ($reader2->nodeType == XMLREADER::END_ELEMENT && $reader2->name == 'si') {
					$this->shared[$sh] = $ftext;
					$ftext = '';
					++$sh;
					$si = '';
				}
				if ($reader2->nodeType == XMLREADER::ELEMENT && $reader2->name == 'r') {
					$r = 'Y';
				}
				if ($reader2->nodeType == XMLREADER::END_ELEMENT && $reader2->name == 'r') {
					$fbold = $fund = $fital = $fsize = $fcol = $fname = $fstrike = $fscript = '';
					$r = '';
				}

				if (($reader2->nodeType == XMLREADER::ELEMENT && $reader2->name == 't') AND $r == '') {
					$ftext = htmlentities($reader2->expand()->textContent);
				}
				if (($reader2->nodeType == XMLREADER::ELEMENT && $reader2->name == 't') AND $r == 'Y') {
					$ft = htmlentities($reader2->expand()->textContent);
					if ($fbold == '' AND $fund == '' AND $fital == '' AND $fsize == '' AND $fcol == '' AND $fname == ''){
						$ftext .= $ft;
					} else {
						$ftext .= "<span style='".$fname.$fsize.$fcol.$fbold.$fund.$fital.$fstrike.$fscript."'>".$ft."</span>";
					}
				}
				if ($reader2->name == 'vertAlign') {
					$script = $reader2->getAttribute("val");
						if ($script == 'superscript'){
							$fscript = "position: relative; top: -0.6em;";
						} else if ($script == 'subscript'){
							$fscript = "position: relative; bottom: -0.5em;";
						}
				}
				if ($reader2->name == 'strike') {
					$fstrike = " text-decoration:line-through;";
				}
				if ($reader2->name == 'b') {
					$fbold = " font-weight: bold;";
				}
				if ($reader2->name == 'u') {
					if ($reader2->getAttribute("val")){
						$ftype = $reader2->getAttribute("val");
						$fund = " border-bottom: 3px double;";
					} else {
						$fund = " text-decoration: underline;";
					}
				}
				if ($reader2->name == 'i') {
					$fital = " font-style: italic;";
				}
				if ($reader2->name == 'sz') {
					if ($script){
						$fsize = " font-size: ".round($reader2->getAttribute("val")*0.75/$this->FSFactor,2)."rem;";  // Font size for sub and super script
						
					} else {
						$fsize = " font-size: ".round($reader2->getAttribute("val")/$this->FSFactor,2)."rem;";  // Font size
					}
					$script = '';
				}
				if ($reader2->name == 'color') {
					if ($reader2->getAttribute("rgb")){
						$fcol = " color: #".substr($reader2->getAttribute("rgb"),2).";"; //Font colour
					}
					if ($reader2->getAttribute("theme")){
						$Tfcol = $reader2->getAttribute("theme"); // Theme for this font				
						$Ttheme = (int)$reader2->getAttribute("theme");
						if ($reader2->getAttribute("tint")){
							$Ttint = strval(round($reader2->getAttribute("tint"),2));
						} else {
							$Ttint = 0;
						}
						$Trgb = $this->themecol[$Ttheme]; // the rgb theme colour
						if ($Ttint == 0){
							$fcol = " color: #".$Trgb.";";
						} else {
							$fcol = " color: #".$this->calctint($Trgb, $Ttint).";"; // the calculated theme tint
						}
					}
					if ($reader2->getAttribute("indexed")){
						$fcol = " color: #".$this->Cindex[$reader2->getAttribute("indexed")].";"; //Indexed font colour
					}
					
				}
				if ($reader2->name == 'rFont') {
					$FF = $reader2->getAttribute("val");
					if (substr($FF,0,9) == 'Helvetica'){
						$FF = 'Helvetica';
					}
					$fname = " font-family: ".$FF.";"; //Font name
				}
				if ($reader2->name == 'family') {
					$ffam = $reader2->getAttribute("val");
				}
				if ($reader2->name == 'scheme') {
					$fsch = $reader2->getAttribute("val");
				}
			}
		}
	}
	
	
	
	/**
	 * Looks up the Drawing1 XML file and the associated relationships file
	 * 
	 * @param - None
	 * @returns String - 
	 */
	private function drawings($Nsheet)
	{
		$zip = new ZipArchive();
		$_xml_draw = 'xl/drawings/drawing'.$Nsheet.'.xml';
		if (true === $zip->open($this->file)) {
			//Get the images from the xl drawings file
			if (($index = $zip->locateName($_xml_draw)) !== false) {
				$xml_draw = $zip->getFromIndex($index);
			}
			$zip->close();
		}
		$F1text = array();
		// if the drawing.xml file exists parse it
		if (isset($xml_draw)){ 
			$enc = mb_detect_encoding($xml_draw);
			$this->setXmlParts($draw_xml, $xml_draw, $enc);
			if($this->debug) {
				echo "<br>XML File : xl/drawings/drawing".$Nsheet.".xml<br>";
				echo "<textarea style='width:100%; height: 200px;'>";
				echo $draw_xml->saveXML();
				echo "</textarea>";
			}
		}

		$zip = new ZipArchive();
		$_xml_rels_draw = 'xl/drawings/_rels/drawing'.$Nsheet.'.xml.rels';
		if (true === $zip->open($this->file)) {
			//Get the headers from the drawings relationships file
			if (($index = $zip->locateName($_xml_rels_draw)) !== false) {
				$xml_rels_draw = $zip->getFromIndex($index);
			}
			$zip->close();
		}
		
		// if the xl/drawings/drawing.xml.rels file exists parse it
		if (isset($xml_rels_draw)){ 
			$enc = mb_detect_encoding($xml_rels_draw);
			$this->setXmlParts($draw_rels_xml, $xml_rels_draw, $enc);
			if($this->debug) {
				echo "<br>XML File : xl/drawings/_rel/drawing".$Nsheet.".xml.rels<br>";
				echo "<textarea style='width:100%; height: 200px;'>";
				echo $draw_rels_xml->saveXML();
				echo "</textarea>";
			}
		}
		
		if (isset($xml_draw)){ 
			$reader = new XMLReader;
			$reader->XML($draw_xml->saveXML());
			$pictno = 0;
			$xfrm = $im = $nfill = '';
			while ($reader->read()) {
				if ($reader->nodeType == XMLREADER::END_ELEMENT && $reader->name == 'xdr:twoCellAnchor') {
					if ($nfill == 'Y'){
						$Fimage[$pictno] = $Limage[$pictno] = $Imxs[$pictno] = $Imys[$pictno] = $relId[$pictno] = '';
						$nfill = '';
					} else {
						++$pictno;
					}
				}
				if ($reader->nodeType == XMLREADER::END_ELEMENT && $reader->name == 'xdr:from') {
					$Fimage[$pictno] = $ICol.$Irow; //first cell location for image
				}
				if ($reader->nodeType == XMLREADER::END_ELEMENT && $reader->name == 'xdr:to') {
					$Limage[$pictno] = $ICol.$Irow; //last cell location for image
				}
				if ($reader->nodeType == XMLREADER::ELEMENT && $reader->name == 'xdr:col') {
					$temp = 1 + htmlentities($reader->expand()->textContent);
					$ICol = $this->numtochars($temp);
				}
				if ($reader->nodeType == XMLREADER::ELEMENT && $reade->name == 'xdr:colOff') {
					$ICoff = htmlentities($reader->expand()->textContent); //not used at present
				}
				if ($reader->nodeType == XMLREADER::ELEMENT && $reader->name == 'xdr:row') {
					$Irow = 1 + htmlentities($reader->expand()->textContent);
				}
				if ($reader->nodeType == XMLREADER::ELEMENT && $reader->name == 'xdr:rowOff') {
					$IRoff = htmlentities($reader->expand()->textContent);  //not used at present
				}
				if ($reader->nodeType == XMLREADER::ELEMENT && $reader->name == 'a:blip') {
					$relId[$pictno] = $reader->getAttribute("r:embed");  //image name
				}
				if ($reader->nodeType == XMLREADER::ELEMENT && $reader->name == 'a:xfrm') {
					$xfrm = 'Y';
				}
				if (($reader->nodeType == XMLREADER::ELEMENT && $reader->name == 'a:ext') AND $xfrm == 'Y'){
					$Imxs[$pictno] = round($reader->getAttribute("cx")/9000);  //image x size
					$Imys[$pictno] = round($reader->getAttribute("cy")/9000);  //image y size
				}
				if ($reader->nodeType == XMLREADER::END_ELEMENT && $reader->name == 'a:xfrm') {
					$xfrm = '';
				}
				if ($reader->name == 'a:noFill') {
					$nfill = 'Y';
				}
			}
			$ci = 0;
			while ($ci < $pictno){
				$reader2 = new XMLReader;
				$reader2->XML($draw_rels_xml->saveXML());
				while ($reader2->read()) {
					if ($reader2->nodeType == XMLREADER::ELEMENT && $reader2->name=='Relationship') {
						if($reader2->getAttribute("Id") == $relId[$ci]) {
							$link = "xl".substr($reader2->getAttribute('Target'),2);
							break;
						}
					}
				}
				$zip = new ZipArchive();
					$im = null;
					if (true === $zip->open($this->file)) {
						$image[$ci] = $this->createImage($zip->getFromName($link), $relId[$ci], $link, $Nsheet);
					}
				$zip->close();
				++$ci;
			}
				
		}
		return array($Fimage, $Limage, $image, $Imxs, $Imys, $pictno);
	}
				

	/**
	 * Creates an image in the filesystem
	 *  
	 * @param objetc $image - The image object
	 * @param string $relId - The image relationship Id
	 * @param string $name - The image name
	 * @return Array - With image tag definition
	 */
	private function createImage($image, $relId, $name, $Nsheet)
	{
		static $Ccount = 1;
		$fname = '';
		$arr = explode('.', $name);
		$l = count($arr);
		$ext = strtolower($arr[$l-1]);
		
		if (!is_dir($this->tmpDir)){
			mkdir($this->tmpDir, 0755, true);
		}
		if ($ext == 'emf' OR $ext == 'wmf'){
			$ftname = $this->tmpDir.'/'.$relId.'-'.$Nsheet.'.'.$ext;
			$tfile = fopen($ftname, "w");
			fwrite($tfile, $image);
			fclose($tfile);
			$fname = $this->tmpDir.'/'.$relId.$this->HF.'.jpg';
			if ($ext == 'wmf'){ // Note that Imagick will only convert '.wmf' files (NOT '.emf' files)
				$imagick = new Imagick();
				$imagick->setresolution(300, 300);
				$imagick->readImage($ftname);
				$imagick->resizeImage(1000,0,Imagick::FILTER_LANCZOS,1);
				$imagick->setImageFormat('jpg');
				$imagick->writeImage($fname);
			}
			

		} else {
			$im = imagecreatefromstring($image);
			$fname = $this->tmpDir.'/'.$relId.'-'.$Nsheet.'.'.$ext;

			switch ($ext) {
				case 'png':
					// Ensure alpha channel is preserved
					imagesavealpha($im, true);
					// Set alpha blending mode
					imagealphablending($im, false);
					// Output PNG with full alpha channel
					imagepng($im, $fname, 9, PNG_ALL_FILTERS);
					break;
				case 'bmp':
					imagebmp($im, $fname);
					break;
				case 'gif':
					imagegif($im, $fname);
					break;
				case 'jpeg':
				case 'jpg':
					imagejpeg($im, $fname);
					break;
				case 'webp':
					imagewebp($im, $fname);
					break;
				default:
					return null;
			}
			imagedestroy($im);
			$fname = $fname."?id=".time();
		}
		return $fname;
	}



	
	/**
	 * Converts from a ex RGB colour to a HSL colour
	 *  
	 * @param string $rgb - The hex rgb colour
	 * @return string - The HSL colour
	 */
	private function hextohsl($rgb)
	{
		$r = hexdec(substr($rgb,0,2));
		$g = hexdec(substr($rgb,2,2));
		$b = hexdec(substr($rgb,4,2));

		$r /= 255;
		$g /= 255;
		$b /= 255;

		$max = max( $r, $g, $b );
		$min = min( $r, $g, $b );

		$h;
		$s;
		$l = ( $max + $min ) / 2;
		$d = $max - $min;

		if( $d == 0 ){
			$h = $s = 0; // achromatic
		} else {
			$s = $d / ( 1 - abs( 2 * $l - 1 ) );

			switch( $max ){
				case $r:
					$h = 60 * fmod( ( ( $g - $b ) / $d ), 6 ); 
						if ($b > $g) {
						$h += 360;
					}
					break;

				case $g: 
					$h = 60 * ( ( $b - $r ) / $d + 2 ); 
					break;

				case $b: 
					$h = 60 * ( ( $r - $g ) / $d + 4 ); 
					break;
			}			        	        
		}
		// HSL colour values are $h, $s and $l
		$hsl['h'] = $h;
		$hsl['s'] = $s;
		$hsl['l'] = $l;
		return $hsl;
	}
	


	/**
	 * Converts from a HSL colour to a Hex RGB colour
	 *  
	 * @param string $h - The 'h' parameter of the HSL colour
	 * @param string $s - The 's' parameter of the HSL colour
	 * @param string $l - The 'l' parameter of the HSL colour
	 * @return string - The hex rgb colour
	 */
	private function hsltohex($h, $s, $l)
	{
		$c = ( 1 - abs( 2 * $l - 1 ) ) * $s;
		$x = $c * ( 1 - abs( fmod( ( $h / 60 ), 2 ) - 1 ) );
		$m = $l - ( $c / 2 );

		if ( $h < 60 ) {
			$r = $c;
			$g = $x;
			$b = 0;
		} else if ( $h < 120 ) {
			$r = $x;
			$g = $c;
			$b = 0;			
		} else if ( $h < 180 ) {
			$r = 0;
			$g = $c;
			$b = $x;					
		} else if ( $h < 240 ) {
			$r = 0;
			$g = $x;
			$b = $c;
		} else if ( $h < 300 ) {
			$r = $x;
			$g = 0;
			$b = $c;
		} else {
			$r = $c;
			$g = 0;
			$b = $x;
		}

		$r = ( $r + $m ) * 255;
		$g = ( $g + $m ) * 255;
		$b = ( $b + $m ) * 255;
		//convert to hex with leading '0' if necessary
		$r = sprintf("%'02X",$r);
		$g = sprintf("%'02X",$g);
		$b = sprintf("%'02X",$b);
		$rgb = $r.$g.$b;

		return $rgb;
	}
	
	


	/**
	 * Calculates the tint of a theme colour for a particular tint factor
	 *  
	 * @param string $theme - The theme hex rgb colour
	 * @param string $tint - The tint factor
	 * @return string - The tint hex rgb colour
	 */
	private function calctint($theme, $tint)
	{
		//convert the hex colour to HSL colour
		$hsl = $this->hextohsl($theme);

		// HSL colour values are $h, $s and $l
		$l = $hsl['l'];

		//calculate tint of theme colour - Only $l needs changing
		if ($tint < 0){
			$tint = abs($tint);
			$l = $l - $l * $tint; // new value of $l for a -ve tint factor
		} else {
			$l = (1 - $l) * $tint + $l; // new value of $l for a +ve tint factor
		}
		
		//now convert the HSL colour back to rgb
		$rgb = $this->hsltohex($hsl['h'], $hsl['s'], $l);	


		return $rgb;
	}




	/**
	 * Calculates the colorScale colour for a particular value in a range
	 *  
	 * @param array $CScolour - The 2 or 3 colours defining the colourScale conditional formatting
	 * @param string $CSmin - The minimum value in the colourScale conditional formatting group
	 * @param string $CSmax - The maximum value in the colourScale conditional formatting group
	 * @param string $CSave - The average (mean) value in the colourScale conditional formatting group
	 * @param string $Cell - The cell value
	 * @return string - The colourScale conditional formatting hex rgb colour for this cell
	 */
	private function findcolorScale($CScolour, $CSmin, $CSmax, $CSave, $cell)
	{
		if ($cell < $CSmin){
			$cell = $CSmin;
		}
		if ($cell > $CSmax){
			$cell = $CSmax;
		}
		$Numcol = count($CScolour);
		if ($Numcol == 3){
			$hsla = $this->hextohsl($CScolour[0]);
			$hslb = $this->hextohsl($CScolour[1]);
			$hslc = $this->hextohsl($CScolour[2]);
			if ($cell < $CSave){
				if (abs($hsla['h'] - $hslb['h']) > 180){
					if ($hsla['h'] > $hslb['h']){
						$hsla['h'] = $hsla['h'] - 360;
					} else {
						$hslb['h'] = $hslb['h'] - 360;
					}
				}
				$newH = $hsla['h'] + (($hslb['h'] - $hsla['h']) * ($cell - $CSmin) / ($CSave - $CSmin));
				$newS = $hsla['s'] + (($hslb['s'] - $hsla['s']) * ($cell - $CSmin) / ($CSave - $CSmin));
				$newL = $hsla['l'] + (($hslb['l'] - $hsla['l']) * ($cell - $CSmin) / ($CSave - $CSmin));
				if ($newH < 0){
					$newH = $newH + 360;
				}
				$rgb = $this->hsltohex($newH, $newS, $newL);	
			} else {
				if (abs($hslb['h'] - $hslc['h']) > 180){
					if ($hslb['h'] > $hslc['h']){
						$hslb['h'] = $hslb['h'] - 360;
					} else {
						$hslc['h'] = $hslc['h'] - 360;
					}
				}
				$newH = $hslb['h'] + (($hslc['h'] - $hslb['h']) * ($cell - $CSave) / ($CSmax - $CSave));
				$newS = $hslb['s'] + (($hslc['s'] - $hslb['s']) * ($cell - $CSave) / ($CSmax - $CSave));
				$newL = $hslb['l'] + (($hslc['l'] - $hslb['l']) * ($cell - $CSave) / ($CSmax - $CSave));
				if ($newH < 0){
					$newH = $newH + 360;
				}
				$rgb = $this->hsltohex($newH, $newS, $newL);	
			}
		} else {
			$hsla = $this->hextohsl($CScolour[0]);
			$hslb = $this->hextohsl($CScolour[1]);
				if (abs($hsla['h'] - $hslb['h']) > 180){
					if ($hsla['h'] > $hslb['h']){
						$hsla['h'] = $hsla['h'] - 360;
					} else {
						$hslb['h'] = $hslb['h'] - 360;
					}
				}
				$newH = $hsla['h'] + (($hslb['h'] - $hsla['h']) * ($cell - $CSmin) / ($CSmax - $CSmin));
				$newS = $hsla['s'] + (($hslb['s'] - $hsla['s']) * ($cell - $CSmin) / ($CSmax - $CSmin));
				$newL = $hsla['l'] + (($hslb['l'] - $hsla['l']) * ($cell - $CSmin) / ($CSmax - $CSmin));
				if ($newH < 0){
					$newH = $newH + 360;
				}
				$rgb = $this->hsltohex($newH, $newS, $newL);	
		}
		return $rgb;
	}




	/**
	 * Takes the cell border style and colours and returns the border styling
	 * 
	 * @param - string $Bstyle - The Border style.
	 * @param - string $rgb - The RGB colour.
	 * @param - string $theme - The theme colour.
	 * @param - string $tint - The tint of the theme colour.
	 * @param - string $indexed - The indexed colour.
	 * @param - string $side - The cell side.
	 * @return - string - The border formatting
	 */
	private function border($Bstyle, $rgb, $theme, $tint, $indexed, $side)
	{
		if ($rgb){
			$bcol = $rgb;
		} else if ($theme){
			$Trgb = $this->themecol[$theme]; // the rgb theme colour
			if ($tint == 0){
				$bcol = $Trgb;
			} else {
				$bcol = $this->calctint($Trgb, $tint); // the calculated theme tint
			}
		} else if ($indexed){
			$bcol = $indexed;
		}
		if ($bcol ==''){
			$bcol = '000000';
		}
		if ($Bstyle == 'thin'){
			$Borsty = ' border-'.$side.': 1px solid #'.$bcol.';';
		} else if($Bstyle == 'medium'){
			$Borsty = ' border-'.$side.': 2px solid #'.$bcol.';';
		} else if($Bstyle == 'thick'){
			$Borsty = ' border-'.$side.': 3px solid #'.$bcol.';';
		} else if($Bstyle == 'double'){
			$Borsty = ' border-'.$side.': double #'.$bcol.';';
		} else if($Bstyle == 'dotted' OR $Bstyle == 'hair'){
			$Borsty = ' border-'.$side.': 1px dotted #'.$bcol.';';
		} else if($Bstyle == 'dashed' OR $Bstyle == 'dashDot' OR $Bstyle == 'dashDotDot'){
			$Borsty = ' border-'.$side.': 1px dashed #'.$bcol.';';
		} else if($Bstyle == 'mediumDashed' OR $Bstyle == 'mediumDashDot' OR $Bstyle == 'mediumDashDotDot' OR $Bstyle == 'slantDashDot'){
			$Borsty = ' border-'.$side.': 2px dashed #'.$bcol.';';
		}
		return $Borsty; //The border styling and colour
	}



	/**
	 * Looks up the styles in the styles XML file and sets the parameters for all the styles/formats
	 * 
	 * @param - string $Cfile - The name of the XML file to process.
	 * @param - string $Cstart - The number to start the format counting at.
	 * @param - string $Ctype - The type of formating to get.
	 * @return - array - The formatting for the type searched for
	 */
	private function getstyles($Cfile, $Ccount, $Ctype)
	{
		$Ctype2 = $Ctype."s";
		if ($Cfile == 'sheet'){
			$reader = new XMLReader;
			$reader->XML($this->sheet_xml->saveXML());
		} else if ($Cfile == 'styles'){
			$reader = new XMLReader();
			$reader->XML($this->styles_xml->saveXML());

		}
		$type = $fon = $bor = $fil = $Bcol = $script = '';
		while ($reader->read()) {
			if ($reader->nodeType == XMLREADER::ELEMENT && $reader->name == $Ctype2) {
				$type = 'Y';
			}
			
			// Start of font processing
			if ($reader->nodeType == XMLREADER::ELEMENT && $reader->name == 'font') {
				$fon = 'Y';
			}
			// Subscript and Superscript
			if (($reader->name == 'vertAlign') AND ($fon == 'Y') AND ($type == 'Y')){
				$script = $reader->getAttribute("val");
				if ($script == 'superscript'){
					$tff = "position: relative; top: -0.6em;";
				} else if ($script == 'subscript'){
					$tff = "position: relative; bottom: -0.5em;";
				}
				$Cstyle[$Ccount]['fscript'] = $tff;
			}
			if (($reader->name == 'strike') AND ($fon == 'Y') AND ($type == 'Y')) {
				$Cstyle[$Ccount]['fstrike'] = " text-decoration:line-through;";
				if ($reader->getAttribute("val") === '0'){
						$Cstyle[$Ccount]['fstrike'] = '';
				}
			}
			if (($reader->name == 'b') AND ($fon == 'Y') AND ($type == 'Y')) {
				$Cstyle[$Ccount]['fbold'] = " font-weight: bold;";
			}
			if (($reader->name == 'u') AND ($fon == 'Y') AND ($type == 'Y')) {
				$Fonund = '';
				if ($reader->getAttribute("val")){
					$ftype = $reader->getAttribute("val");
					if ($ftype == 'double'){
						$Fonund = " border-bottom: 3px double;";
					} else if ($ftype == 'singleAccounting'){
						$Cstyle[$Ccount]['sacc'] = "Y";
					} else if ($ftype == 'doubleAccounting'){
						$Cstyle[$Ccount]['dacc'] = "Y";
					}
				} else {
					$Fonund = " text-decoration: underline;";
				}
				$Cstyle[$Ccount]['fund'] = $Fonund;
			}
			if (($reader->name == 'i') AND ($fon == 'Y') AND ($type == 'Y')) {
				$Cstyle[$Ccount]['fital'] = " font-style: italic;";
			}
			if (($reader->name == 'sz') AND ($fon == 'Y') AND ($type == 'Y')) {
				if ($script){
					$tfsize = " font-size: ".round($reader->getAttribute("val")*0.75/$this->FSFactor,2)."rem;";  // Font size for sub and super script
					
				} else {
					$tfsize = " font-size: ".round($reader->getAttribute("val")/$this->FSFactor,2)."rem;";  // Font size
				}
				$script = '';
				$Cstyle[$Ccount]['fsize'] = $tfsize;
				if ($Ccount == 0 AND $Ctype == 'font'){
					$this->DevS = $tfsize; // Default Font Size
				}
			}
			if (($reader->name == 'color') AND ($fon == 'Y') AND ($type == 'Y')) {
				if ($reader->getAttribute("rgb")){
					$Foncol = " color: #".substr($reader->getAttribute("rgb"),2).";"; //Font colour
				}
				if ($reader->getAttribute("theme")){
					$Ftheme = $reader->getAttribute("theme"); //Theme for this font
					if ($reader->getAttribute("tint")){
						$Ftint = strval(round($reader->getAttribute("tint"),2));
					} else {
						$Ftint = 0;
					}
					$Trgb = $this->themecol[$Ftheme]; // the rgb theme colour
					if ($Ftint == 0){
						$Foncol = " color: #".$Trgb.";";
					} else {
						$Foncol = " color: #".$this->calctint($Trgb, $Ftint).";"; // the calculated theme tint
					}
				}
				if ($reader->getAttribute("indexed")){
					$Foncol = " color: #".$this->Cindex[$reader->getAttribute("indexed")].";"; //Indexed font colour
				}
				$Cstyle[$Ccount]['fcol'] = $Foncol;
			}
			if (($reader->name == 'name') AND  ($fon == 'Y') AND ($type == 'Y')) {
				$FF = $reader->getAttribute("val");
				if (substr($FF,0,9) == 'Helvetica'){
					$FF = 'Helvetica';
				}
				$Cstyle[$Ccount]['fname'] = " font-family: ".$FF.";"; //Font name
				if ($Ccount == 0 AND $Ctype == 'font'){
					$this->DevF = $Cstyle[$Ccount]['fname'];
				}
			}
			if (($reader->nodeType == XMLREADER::END_ELEMENT && $reader->name == 'font') AND ($type == 'Y')) {
				if (!$Cstyle[$Ccount]['fbold']){
					$Cstyle[$Ccount]['fbold'] = '';
				}
				if (!$Cstyle[$Ccount]['fital']){
					$Cstyle[$Ccount]['fital'] = '';
				}
				$fon = '';
				if ($Ctype == 'font'){
					++$Ccount;
				}
			}
			//End of font format processing
			//----------------------------------------------------------------
			// Start of borders and diagonals format processing
			if (($reader->nodeType == XMLREADER::ELEMENT && $reader->name == 'border') AND ($type == 'Y')){
				$Cstyle[$Ccount]['diagD']  = $reader->getAttribute("diagonalDown");
				$Cstyle[$Ccount]['diagU']  = $reader->getAttribute("diagonalUp");
				$bor = 'Y';
			}
			//Left Border
			if (($reader->name == 'left') AND ($bor == 'Y')) {
				if ($reader->nodeType <> XMLREADER::END_ELEMENT && $reader->name == 'left') {
					if ($reader->getAttribute("style")){
						$bleft = $reader->getAttribute("style"); //Style of left border
					}
					$pos = 'L';
				}
			}
			if (($reader->name == 'color') AND ($pos == 'L')){
				$rgb = $Ctheme = $indexed = '';
				$Ctint = 0;
				$rgb = substr($reader->getAttribute("rgb"),2); // RGB colour of left border
				if ($reader->getAttribute("theme")){
					$Ctheme = $reader->getAttribute("theme"); // Theme colour of left border
					if ($reader->getAttribute("tint")){
						$Ctint = strval(round($reader->getAttribute("tint"),2));
					} else {
						$Ctint = 0;
					}
				}
				if ($reader->getAttribute("indexed")){
					$indexed = $this->Cindex[$reader->getAttribute("indexed")]; //Indexed colour of left border
				}
				$Cstyle[$Ccount]['bleft'] = $this->border($bleft, $rgb, $Ctheme, $Ctint, $indexed, 'left');
				$pos = '';
			}
			// Right Border
			if (($reader->name == 'right') AND ($bor == 'Y')) {
				if ($reader->nodeType <> XMLREADER::END_ELEMENT && $reader->name == 'right') {
					if ($reader->getAttribute("style")){
						$bright = $reader->getAttribute("style"); //Style of right border
					}
					$pos = 'R';
				}
			}
			if (($reader->name == 'color') AND ($pos == 'R')){
				$rgb = $Ctheme = $indexed = '';
				$Ctint = 0;
				$rgb = substr($reader->getAttribute("rgb"),2); //RGB colour of right border
				if ($reader->getAttribute("theme")){
					$Ctheme = (int)$reader->getAttribute("theme"); // Theme colour of right border
					if ($reader->getAttribute("tint")){
						$Ctint = strval(round($reader->getAttribute("tint"),2));
					} else {
						$Ctint = 0;
					}
				}
				if ($reader->getAttribute("indexed")){
					$indexed = $this->Cindex[$reader->getAttribute("indexed")]; //Indexed colour of right border
				}
				$Cstyle[$Ccount]['bright'] = $this->border($bright, $rgb, $Ctheme, $Ctint, $indexed, 'right');
				$pos = '';
			}
			// Top Border
			if (($reader->name == 'top') AND ($bor == 'Y')) {
				if ($reader->nodeType <> XMLREADER::END_ELEMENT && $reader->name == 'top') {
					if ($reader->getAttribute("style")){
						$btop = $reader->getAttribute("style"); //Style of top border
					}
					$pos = 'T';
				}
			}
			if (($reader->name == 'color') AND ($pos == 'T')){
				$rgb = $Ctheme = $indexed = '';
				$Ctint = 0;
				$rgb = substr($reader->getAttribute("rgb"),2); //RGB colour of top border
				if ($reader->getAttribute("theme")){
					$Ctheme = (int)$reader->getAttribute("theme"); // Theme colour of top border
					if ($reader->getAttribute("tint")){
						$Ctint = strval(round($reader->getAttribute("tint"),2));
					} else {
						$Ctint = 0;
					}
				}
				if ($reader->getAttribute("indexed")){
					$indexed = $this->Cindex[$reader->getAttribute("indexed")]; //Indexed colour of top border
				}
				$Cstyle[$Ccount]['btop'] = $this->border($btop, $rgb, $Ctheme, $Ctint, $indexed, 'top');
				$pos = '';
			}
			// Bottom Border
			if (($reader->name == 'bottom') AND ($bor == 'Y')) {
				if ($reader->nodeType <> XMLREADER::END_ELEMENT && $reader->name == 'bottom') {
					if ($reader->getAttribute("style")){
						$bbott = $reader->getAttribute("style"); //Style of bottom border
					}
					$pos = 'B';
				}
			}
			if (($reader->name == 'color') AND $pos == 'B'){
				$rgb = $Ctheme = $indexed = '';
				$Ctint = 0;
				$rgb = substr($reader->getAttribute("rgb"),2);  //RGN colour of bottom border
				if ($reader->getAttribute("theme")){
					$Ctheme = (int)$reader->getAttribute("theme"); // Theme colour of bottom border
					if ($reader->getAttribute("tint")){
						$Ctint = strval(round($reader->getAttribute("tint"),2));
					} else {
						$Ctint = 0;
					}
				}
				if ($reader->getAttribute("indexed")){
					$indexed = $this->Cindex[$reader->getAttribute("indexed")]; //Indexed colour of bottom border
				}
				$Cstyle[$Ccount]['bbott'] = $this->border($bbott, $rgb, $Ctheme, $Ctint, $indexed, 'bottom');
				$pos = '';
			}
			// Cell Diagonal Lines
			if ($reader->name == 'diagonal') {
				if (($reader->nodeType <> XMLREADER::END_ELEMENT && $reader->name == 'diagonal') AND ($bor == 'Y')){
					if ($reader->getAttribute("style")){
						$Dsty = $reader->getAttribute("style"); //Style of diagonal line (not used at present)
					}
					$pos = 'D';
				}
			}
			if (($reader->name == 'color') AND $pos == 'D'){  //Colour of diagonal line
				$coldia = substr($reader->getAttribute("rgb"),2);
				if ($reader->getAttribute("theme")){
					$Dtheme = (int)$reader->getAttribute("theme");
					if ($reader->getAttribute("tint")){
						$Dtint = strval(round($reader->getAttribute("tint"),2));
					} else {
						$Dtint = 0;
					}
					$Trgb = $this->themecol[$Dtheme]; // the rgb theme colour
					if ($Ctint == 0){
						$coldia = $Trgb;
					} else {
						$coldia = $this->calctint($Trgb, $Dtint); // the calculated theme tint
					}
				}
				if ($reader->getAttribute("indexed")){
					$coldia = $this->Cindex[$reader->getAttribute("indexed")]; //Indexed font colour
				}
				if ($coldia == ''){
					$coldia = '000000';
				}
				$Cstyle[$Ccount]['cdiag'] = $coldia;
			}
			if (($reader->nodeType == XMLREADER::END_ELEMENT && $reader->name == 'border') AND $Ctype == 'border'){
				$bor = '';
				++$Ccount;
			}
			// End of borders and diagonal format processing
			//----------------------------------------------------
			// Start of fill formatting styles
			if ($reader->nodeType == XMLREADER::ELEMENT && $reader->name == 'fill') {
				$fil = 'Y';
			}
			if (($Ctype == 'dxf') OR ($Ctype == 'x14:conditionalFormatting')){
				$Bcol = 'bgColor';
			} else {
				$Bcol = 'fgColor';
			}
			if (($reader->nodeType == XMLREADER::ELEMENT && $reader->name == $Bcol) AND ($fil == 'Y')){
				$Frgb = substr($reader->getAttribute("rgb"),2);
				if ($Frgb == ''){
					$Ctheme = (int)$reader->getAttribute("theme");
					if ($reader->getAttribute("tint")){
						$Ctint = strval(round($reader->getAttribute("tint"),2));
					} else {
						$Ctint = 0;
					}
					$Trgb = $this->themecol[$Ctheme]; // the rgb theme colour
					if ($Ctint == 0){
						$Frgb = $Trgb;
					} else {
						$Frgb = $this->calctint($Trgb, $Ctint); // the calculated theme tint
					}
				}
				if ($reader->getAttribute("indexed")){
					$Frgb = $this->Cindex[$reader->getAttribute("indexed")]; //Indexed colour
				}
				$Cstyle[$Ccount]['fill'] = " background-color: #".$Frgb.";";
				$Cstyle[$Ccount]['fillno'] = $Frgb;
			}
			if (($reader->nodeType == XMLREADER::END_ELEMENT && $reader->name == 'fill') AND $Ctype == 'fill'){
				++$Ccount;
				$fil = '';
			}
			// End of Fill styles
			//--------------------------------------------------
			// Start of conditional formatting styles. Uses the Font, Borders and Fill sections above to get the actual format data
			
			if (($reader->nodeType == XMLREADER::END_ELEMENT && $reader->name == 'dxf') AND $Ctype == 'dxf'){
				++$Ccount;
			}
			// End of conditional formatting
			
			if (($reader->nodeType == XMLREADER::END_ELEMENT && $reader->name == 'x14:conditionalFormatting') AND $Ctype == 'x14:conditionalFormatting'){
				++$Ccount;
			}
			// End of conditional formatting
			
			if (($reader->nodeType == XMLREADER::END_ELEMENT && $reader->name == $Ctype2)){
				$Cstyle[0]['num'] = $Ccount;
				$type = '';
			}

		}
		return $Cstyle;
	}




	/**
	 * Looks up the styles in the styles XML file and sets the parameters for all the styles/formats
	 * 
	 * @param - array $Ffirst - The ref for the first cell in a merge.
	 * @param - array $Flast - The ref for the last cell in a merge.
	 * @return - array - The cell formatting
	 */
	private function findstyles($Ffirst, $Flast)
	{
		$zip = new ZipArchive();
		$_xml_styles = 'xl/styles.xml';
		if (true === $zip->open($this->file)) {
			//Get the style references from the word styles file
			if (($index = $zip->locateName($_xml_styles)) !== false) {
				$xml_styles = $zip->getFromIndex($index);
			}
			$zip->close();
		}
		
		$enc = mb_detect_encoding($xml_styles);
		$this->setXmlParts($this->styles_xml, $xml_styles, $enc);
		if ($this->stylecount == 0){
			if($this->debug) {
				echo "<br>XML File : xl/styles.xml<br>";
				echo "<textarea style='width:100%; height: 200px;'>";
				echo $this->styles_xml->saveXML();
				echo "</textarea>";
			}
			$this->stylecount = 1;
		}

		$reader = new XMLReader();
		$reader->XML($this->styles_xml->saveXML());
		$formno = -1;
		$fillno = 0;
		$found = '';
		$Nxref = array('1' => '0', '2' => '0.00', '3' => '#,##0', '4' => '#,##0.00', '5' => '"£"#,##0;\-"£"#,##0', '6' => '"£"#,##0;[Red]\-"£"#,##0', '7' => '"£"#,##0.00;\-"£"#,##0.00', '8' => '"£"#,##0.00;[Red]\-"£"#,##0.00', '9' => '0%', '10' => '0.00%', '11' => '0.00E+00', '12' => '#\ ?/?', '13' => '#\ ??/??', '14' => 'dd/mm/yyyy;@', '15' => 'dd-M-yy;@', '16' => 'dd-M;@', '17' => 'M-yy;@', '18' => 'h:mm\ AM/PM;@', '19' => 'h:mm:ss\ AM/PM;@', '20' => 'hh:mm;@', '21' => 'hh:mm:ss;@', '22' => 'dd/mm/yyyy\ hh:mm;@', '37' => '#,##0;\-#,##0', '38' => '#,##0;[Red]\-#,##0', '39' => '#,##0.00;\-#,##0.00', '40' => '#,##0.00;[Red]\-#,##0.00', '42' => '_-"£"* #,##0_-;\-"£"* #,##0_-;_-"£"* "-"_-;_-@_-', '44' => '_-"£"* #,##0.00_-;\-"£"* #,##0.00_-;_-"£"* "-"??_-;_-@_-', '45' => 'mm:ss;@', '46' => 'ZZZ', '47' => 'mm:ss.0;@');
		while ($reader->read()) {
			// --------------------------------------------------------------------
			//Start of Format Cross References
			if ($reader->nodeType == XMLREADER::ELEMENT && $reader->name == 'cellXfs') {
				$formnum = $reader->getAttribute("count"); // Number of format cross references
				$formno = -1;
			}

			//Get Xrefs for the formats
			if ($reader->name == 'xf') {
				if ($reader->nodeType <> XMLREADER::END_ELEMENT && $reader->name == 'xf') {
					++$formno; //increment the format counter
					$formnumfmt[$formno] = $reader->getAttribute("numFmtId"); //Xref to number format
					$formfontId[$formno] = $reader->getAttribute("fontId"); //Xref to Font number
					$formfillId[$formno] = $reader->getAttribute("fillId"); //Xref to Fill number
					$formborId[$formno] = $reader->getAttribute("borderId"); //Xref to borders
					$formxfId[$formno] = $reader->getAttribute("xfId"); //Xref to cell styles (hyperlink etc)
					if ($reader->getAttribute("applyFont")){
						$formappFo[$formno] = $reader->getAttribute("applyFont"); //Apply a different font from default
					}
					if ($reader->getAttribute("applyAlignment")){
						$formappAl[$formno] = $reader->getAttribute("applyAlignment");//Apply a different text alignment from default
					}
					if ($reader->getAttribute("applyBorder")){
						$formappBo[$formno] = $reader->getAttribute("applyBorder"); //Apply a different border from default
					}
				}
			}
			// Cell alignment for text/numbers
			if ($reader->name == 'alignment') {
				if ($reader->getAttribute("horizontal")){
					$formHalign[$formno] = " text-align: ".$reader->getAttribute("horizontal").";"; //text horizontal alignment (left is default)
				}
				if ($reader->getAttribute("vertical")){
					$Valign = $reader->getAttribute("vertical"); //text vertical alignment (bottom is default)
					if ($Valign == 'top'){
						$formValign[$formno] = " vertical-align:top;";
					} else if ($Valign == 'center'){
						$formValign[$formno] = " vertical-align:middle;";
					} else if ($Valign == 'bottom'){
						$formValign[$formno] = " vertical-align:bottom;";
					}
				}
			}
			
			// End of Format Cross References
			// --------------------------------------------------------------------
			// Start of getting Cell Style References (used mainly for Hyperlinks)
			
			if ($reader->nodeType == XMLREADER::ELEMENT && $reader->name == 'cellStyles') {
				$stylenum = $reader->getAttribute("count"); // Number of cell styles
			}
			if ($reader->name == 'cellStyle') {
				$CxfId = $reader->getAttribute("xfId");
				$Cname[$CxfId] = $reader->getAttribute("name");
			}
			
			// End of Cell Style References
			//---------------------------------------------------------------------------
			// Get the Number Format Styles
			
			if ($reader->nodeType == XMLREADER::ELEMENT && $reader->name == 'numFmts') {
				$numform = $reader->getAttribute("count"); // Number of number formats
			}
			if ($reader->name == 'numFmt') {
				$NumFId = $reader->getAttribute("numFmtId");
				$Nformat[$NumFId] = $reader->getAttribute("formatCode");
			}
			
			// End of Number Format Styler
			
		}
		
		$Font = $this->getstyles('styles', 0, 'font'); //Get all the 'font' formatting styles
		
		$Bord = $this->getstyles('styles', 0, 'border'); //Get all the 'border' formatting styles

		$Fill = $this->getstyles('styles', 0, 'fill'); //Get all the 'fill' formatting styles
		
		$cc = 0;
		while ($cc < $formnum){
			$Cellstyle[$cc]['fill'] = $Fill[$formfillId[$cc]]['fill']; //fill colour
			$Cellstyle[$cc]['bleft'] = $Bord[$formborId[$cc]]['bleft']; //left borders
			$Cellstyle[$cc]['btop'] = $Bord[$formborId[$cc]]['btop'];  //top borders
			
			//creating the diagonal line in cells and getting the appropriate background colour
			if (!isset($Fill[$formfillId[$cc]]['fillno'])){
				$Fill[$formfillId[$cc]]['fillno'] = 'ffffff';
			}
			if ($Bord[$formborId[$cc]]['diagD'] <> 1 AND $Bord[$formborId[$cc]]['diagU'] == 1){
				$Cellstyle[$cc]['bdiag'] = ' background: linear-gradient(to right bottom, #'.$Fill[$formfillId[$cc]]['fillno'].' 0%,#'.$Fill[$formfillId[$cc]]['fillno'].' 48%,#'.$Bord[$formborId[$cc]]['cdiag'].' 50%,#'.$Bord[$formborId[$cc]]['cdiag'].' 51%,#'.$Fill[$formfillId[$cc]]['fillno'].' 52%,#'.$Fill[$formfillId[$cc]]['fillno'].' 100%);';
			}
			if ($Bord[$formborId[$cc]]['diagD'] == 1 AND $Bord[$formborId[$cc]]['diagU'] <> 1){
				$Cellstyle[$cc]['bdiag'] = ' background: linear-gradient(to right top, #'.$Fill[$formfillId[$cc]]['fillno'].' 0%,#'.$Fill[$formfillId[$cc]]['fillno'].' 48%,#'.$Bord[$formborId[$cc]]['cdiag'].' 50%,#'.$Bord[$formborId[$cc]]['cdiag'].' 51%,#'.$Fill[$formfillId[$cc]]['fillno'].' 52%,#'.$Fill[$formfillId[$cc]]['fillno'].' 100%);';
			}
			if ($Bord[$formborId[$cc]]['diagD'] == 1 AND $Bord[$formborId[$cc]]['diagU'] == 1){
				$Cellstyle[$cc]['bdiag'] = ' background: linear-gradient(to right top, #'.$Fill[$formfillId[$cc]]['fillno'].' 0%,#'.$Fill[$formfillId[$cc]]['fillno'].' 48%,#'.$Bord[$formborId[$cc]]['cdiag'].' 50%,#'.$Bord[$formborId[$cc]]['cdiag'].' 51%,#'.$Fill[$formfillId[$cc]]['fillno'].' 52%,#'.$Fill[$formfillId[$cc]]['fillno'].' 100%);';
			}
			
			
			//for merged cells find the right and bottom borders from the details of the last cell of the range
			$aa = sizeof($Ffirst);
			$bb = 0;
			while ($bb < $aa){ 
				if ($Ffirst[$bb] == $cc){ 
					$found = 'Y';
					$Cellstyle[$cc]['bright'] = $Bord[$formborId[$Flast[$bb]]]['bright']; //right borders
					if ($Ccol[$cc] == '' AND !isset($Bord[$formborId[$Flast[$bb]]]['bright'])){
						$Cellstyle[$cc]['bright'] = ' border-right:1px solid LightGray;';
					}
					$Cellstyle[$cc]['bbott'] = $Bord[$formborId[$Flast[$bb]]]['bbott']; //bottom borders
					if ($Ccol[$cc] == '' AND !isset($Bord[$formborId[$Flast[$bb]]]['bbott'])){
						$Cellstyle[$cc]['bbott'] = ' border-bottom:1px solid LightGray;';
					}
					$bb = $aa;
				}
				++$bb;
			} 
			
			//Default right and bottom borders
			if  ($found == ''){
				$Cellstyle[$cc]['bright'] = $Bord[$formborId[$cc]]['bright']; //right borders
				if ($Ccol[$cc] == '' AND !isset($Bord[$formborId[$cc]]['bright'])){
					$Cellstyle[$cc]['bright'] = ' border-right:1px solid LightGray;';
				}
				$Cellstyle[$cc]['bbott'] = $Bord[$formborId[$cc]]['bbott']; //bottom borders
				if ($Ccol[$cc] == '' AND !isset($Bord[$formborId[$cc]]['bbott'])){
					$Cellstyle[$cc]['bbott'] = ' border-bottom:1px solid LightGray;';
				}
			}
			
			
			// Text and Number alignment
			$found = '';
			if (!$formHalign[$cc]){
				$Cellstyle[$cc]['athor'] = " text-align:left;"; //default for text
				$Cellstyle[$cc]['anhor'] = " text-align:right;"; //default for numbers
			} else {
				$Cellstyle[$cc]['athor'] = $formHalign[$cc];  //horizontal text alignment (text)
				$Cellstyle[$cc]['anhor'] = $formHalign[$cc];  //horizontal text alignment (numbers)
			}
			if (!$formValign[$cc]){
				$Cellstyle[$cc]['avert'] = " vertical-align:bottom;";
			} else {
				$Cellstyle[$cc]['avert'] = $formValign[$cc];  //vertical text alignment
			}

			if ($Font[$formfontId[$cc]]['sacc'] == 'Y'){ 
				$Cellstyle[$cc]['bbott'] = ' border-bottom:1px solid #000000;'; //put in Single Accounting line
			}
			if ($Font[$formfontId[$cc]]['dacc'] == 'Y'){
				$Cellstyle[$cc]['bbott'] = ' border-bottom:3px double #000000;'; //put in Double Accounting line
			}
			$Cellstyle[$cc]['fname'] = $Font[$formfontId[$cc]]['fname']; //font name
			$Cellstyle[$cc]['fsize'] = $Font[$formfontId[$cc]]['fsize']; //font size
			$Cellstyle[$cc]['fcol'] = $Font[$formfontId[$cc]]['fcol']; //font colour
			$Cellstyle[$cc]['fbold'] = $Font[$formfontId[$cc]]['fbold']; //font bold
			$Cellstyle[$cc]['fund'] = $Font[$formfontId[$cc]]['fund']; //font underline
			$Cellstyle[$cc]['fital'] = $Font[$formfontId[$cc]]['fital']; //font italics
			$Cellstyle[$cc]['fscript'] = $Font[$formfontId[$cc]]['fscript']; //font superscript/subscript
			$Cellstyle[$cc]['fstrike'] = $Font[$formfontId[$cc]]['fstrike']; //font strikethrough
			
			$Cellstyle[$cc]['hyper'] = $Cname[$formxfId[$cc]]; // cell style - used to indicate cells with a hyperlink
			if (!$Nformat[$formnumfmt[$cc]]){
				$Nformat[$formnumfmt[$cc]] = $Nxref[$formnumfmt[$cc]]; //some common number formats are not always defined in the 'styles' file
			}
			$Cellstyle[$cc]['nform'] = $Nformat[$formnumfmt[$cc]]; // number format style
			++$cc;
		}
		
		// Get all the conditional formatting styles.
		$CF1 = $this->getstyles('styles', 0, 'dxf');
		$dx = 0;
		$Cnum = $CF1[0]['num'];
		while ($dx < $Cnum){
			$Cellstyle[$dx]['Cfill'] = $CF1[$dx]['fill']; // conditional fill colour
			$Cellstyle[$dx]['Cbleft'] = $CF1[$dx]['bleft']; // conditional left borders
			$Cellstyle[$dx]['Cbtop'] = $CF1[$dx]['btop'];  // conditional top borders
			$Cellstyle[$dx]['Cbright'] = $CF1[$dx]['bright'];  // conditional right borders
			$Cellstyle[$dx]['Cbbott'] = $CF1[$dx]['bbott'];  // conditional bottom borders
			$Cellstyle[$dx]['Cfname'] = $CF1[$dx]['fname']; // conditional font name
			$Cellstyle[$dx]['Cfsize'] = $CF1[$dx]['fsize']; // conditional font size
			$Cellstyle[$dx]['Cfcol'] = $CF1[$dx]['fcol']; // conditional font colour
			$Cellstyle[$dx]['Cfbold'] = $CF1[$dx]['fbold']; // conditional font bold
			$Cellstyle[$dx]['Cfund'] = $CF1[$dx]['fund']; // conditional font underline
			$Cellstyle[$dx]['Cfital'] = $CF1[$dx]['fital']; // conditional font italics
			$Cellstyle[$dx]['Cfscript'] = $CF1[$dx]['fscript']; // conditional font superscript/subscript
			$Cellstyle[$dx]['Cfstrike'] = $CF1[$dx]['fstrike']; // conditional font strikethrough
			++$dx;
		}
		$Cellstyle[0]['dxf'] = $dx;
		return $Cellstyle;
	}
	
	/**
	 * CONVERT DECIMAL TO A FRACTION
	 *  
	 * @param string $dec - The decimal number
	 * @param string $n - The number which the denominator must be below
	 * @return array- The numerator and denominator
	 */
	private function float2rat($n, $dec) 
	{
		if ($n <> 0 AND $n <> ''){
			$frac = array();
			$tolerance = 1.e-6;
			$h1=1; $h2=0;
			$k1=0; $k2=1;
			$b = 1/$n;
			do {
				$kk = $k1;
				$hh = $h1;
				$b = 1/$b;
				$a = floor($b);
				$aux = $h1; $h1 = $a*$h1+$h2; $h2 = $aux;
				$aux = $k1; $k1 = $a*$k1+$k2; $k2 = $aux;
				$b = $b-$a;
			} while (abs($n-$h1/$k1) > $n*$tolerance AND $k1 < $dec);
			$frac['num'] = $hh;
			$frac['den'] = $kk;
		} else {
			$frac['num'] = '';
			$frac['den'] = '';
		}

		return $frac;
}


	/**
	 * CONVERT PROCESS CURRENCY FORMATTING
	 *  
	 * @param string $Ncode - The currency formatting
	 * @return array- The details of the formatting
	 */
	private function currency($Ncode) 
	{
		$details = array();
		if (strpos($Ncode,'"')){ //find location of currency unit in double inverted commas
			$S1 = strpos($Ncode,'"'); 
			$S1a = strrpos($Ncode,'"');
			$Spos = $S1 + 1;
			$Len = $S1a - $Spos;
			$details['unit'] = substr($Ncode,$Spos,$Len);
		} else if (strpos($Ncode,'$')){ //find location of currency unit preceded by a '$'
			if (strpos($Ncode,'0[$')){ // find trailing currency
				$details['pos'] = 'T';
				$cp = strpos($Ncode,'0[$')+ 1;
			} else if (strpos($Ncode,'0\ [$')){ // find trailing currency (euros)
				$details['pos'] = 'T';
				$cp = strpos($Ncode,'0\ [$')+ 3;
			} else { // find leading currency units
				$details['pos'] = 'L';
				$cp = strpos($Ncode,'[$');
			}
			$cstr = substr($Ncode,$cp);
			$S2a = strpos($cstr,'-');
			$Spos = 2;
			$Len = $S2a - $Spos;
			$details['unit'] = substr($cstr,$Spos,$Len);
		} else {
			$details['unit'] = ''; //If no currency unit
		}
		$Sloc = strpos($Ncode,'#') + 1;
		$details['sep'] = substr($Ncode,$Sloc,1); //type of separator
		$Dloc = strpos($Ncode,'#0') + 2;
		$details['point'] = substr($Ncode,$Dloc,1); //type of decimal point
		if (strpos($Ncode,'Red')){
			$details['red'] = 'Y';
		} else {
			$details['red'] = '';
		}
		if (strpos($Ncode,'-[')){
			$details['minus'] = '-';
		} else if (strpos($Ncode,'-"')){
			$details['minus'] = '-';
		} else if (strpos($Ncode,'-#')){
			$details['minus'] = '-';
		} else {
			$details['minus'] = '';
		}
		return $details;
	}



	/**
	 * PROCESS HEADER/FOOTER TEXT
	 *  
	 * @param string $Stext - The text of a section of a header/footer
	 * @return string - The HTML code of a section
	 */
	private function HFsect($Stext)
	{
		$parts = explode('*',$Stext);
		$Npart = sizeof($parts);
		$FU = $FE = $FK = '';
		$FF = $FB = $FI = '';
		$n = 0;
		while ($n < $Npart){
			$parts[$n] = nl2br($parts[$n]);
			if (substr($parts[$n],1,4) == 'quot'){ //font details
				$Stpos = strpos(substr($parts[$n],1),',');
				if ($Stpos > 10){
					$Flen = $Stpos - 5;
					 $FF = " font-family: ".substr($parts[$n],6,$Flen).";";
				}
				if (strpos($parts[$n],'Bold')){
					$FB = " font-weight: bold;";
				} else {
					$FB = '';
				}
				if (strpos($parts[$n],'Regular')){
					$FB = '';
				}

				if (strpos($parts[$n],'Italic')){
					$FI = " font-style: italic;";
				} else {
					$FI = '';
				}
				$SS = substr($parts[$n],6);
				$IV = strpos($SS,'quot');
				if ($IV + 5 < strlen($SS)){
					$text4 = substr($SS,$IV+5);
				}
			} else if (substr($parts[$n],0,1) == 'U'){ //font single underling
				if ($FU == ''){
					$FU = " text-decoration: underline;";
					$FE = '';
				} else {
					$FU = '';
				}
				$text1 = substr($parts[$n],1);
			} else if (substr($parts[$n],0,1) == 'E'){ //font double underlining
				if ($FE == ''){
					$FE = " border-bottom: 3px double;";
					$FU = '';
				} else {
					$FE = '';
				}
				$text1 = substr($parts[$n],1);
			} else if (substr($parts[$n],0,1) == 'K'){ // font colour info
				$Fcol = substr($parts[$n],1,6);
				if (substr($Fcol,2,1) == '+' OR substr($Fcol,2,1) == '-'){ //find colour from theme and tint
					$pole = '';
					$theme = substr($Fcol,1,1);
					if (substr($Fcol,2,1) == '-'){
						$pole = '-';
					}
					$tt = substr($Fcol,4,2);
					if ($tt == '00'){
						$tint = '0';
					} else if ($tt >= '03' AND $tt <= '05'){
						$tint = '0.05';
					} else if ($tt >= '08' AND $tt <= '10'){
						$tint = '0.1';
					} else if ($tt >= '13' AND $tt <= '15'){
						$tint = '0.15';
					} else if ($tt >= '23' AND $tt <= '25'){
						$tint = '0.25';
					} else if ($tt >= '33' AND $tt <= '35'){
						$tint = '0.35';
					} else if ($tt >= '38' AND $tt <= '40'){
						$tint = '0.4';
					} else if ($tt >= '48' AND $tt <= '50'){
						$tint = '0.5';
					} else if ($tt >= '58' AND $tt <= '60'){
						$tint = '0.6';
					} else if ($tt >= '73' AND $tt <= '75'){
						$tint = '0.75';
					} else if ($tt >= '78' AND $tt <= '80'){
						$tint = '0.8';
					} else if ($tt >= '88' AND $tt <= '90'){
						$tint = '0.9';
					}
					$tint = $pole.$tint;
					$Trgb = $this->themecol[$theme]; // the rgb theme colour
					if ($tint == 0){
						$Fcol = $Trgb;
					} else {
						$Fcol = $this->calctint($Trgb, $tint); // the calculated theme tint
					}
				}
				$FK = " color: #".$Fcol.";"; //Font colour
				$text1 = substr($parts[$n],7);
			} else if (is_numeric(substr($parts[$n],0,2)) AND $n <> 0) { //font size info
				$FS = " font-size: ".round(intval(substr($parts[$n],0,2))/$this->FSFactor,2)."rem;";
				$text2 = substr($parts[$n],2);
			} else { //if no font formatting
				$text3 = $parts[$n];
			}
			if ($text1 <> ''){
				if ($FS == ''){
					$FS = $this->DevS; //default font size
				}
				if ($FF == ''){
					$FF = $this->DevF; //default font name
				}
				$Htext = nl2br($Htext);
				$Htext .= "<span style='".$FF.$FB.$FI.$FU.$FE.$FK.$FS."'>".$text1."</span>";
				$text1 = '';
			}
			if ($text4 <> ''){
				if ($FS == ''){
					$FS = $this->DevS; //default font size
				}
				if ($FF == ''){
					$FF = $this->DevF; //default font name
				}
				$Htext .= "<span style='".$FF.$FB.$FI.$FU.$FE.$FK.$FS."'>".$text4."</span>";
				$text4 = '';
			}
			if ($text2 <> ''){
				if ($FS == ''){
					$FS = $this->DevS; //default font size
				}
				if ($FF == ''){
					$FF = $this->DevF; //default font name
				}
				$Htext .= "<span style='".$FF.$FB.$FI.$FU.$FE.$FK.$FS."'>".$text2."</span>";
				$text2 = '';
			}
			if ($text3 <> ''){
				if ($FS == ''){
					$FS = $this->DevS; //default font size
				}
				if ($FF == ''){
					$FF = $this->DevF; //default font name
				}
				$Htext .= "<span style='".$FF.$FS."'>".$text3."</span>";
				$text3 = '';
			}
			++$n;
		}
		return $Htext;
	}
	
	
	/**
	 * PROCESS HEADERS AND FOOTERS
	 *  
	 * @param string $HFtext - The Header or Footer XML test
	 * @return  - The header/footer html text
	 */
	private function HeadFoot($HFtext)
	{
		$LHF = $CHF = $RHF = '&nbsp;';
		$HFtext = preg_replace('/&amp;/', '*', $HFtext);
		$Lpos = strpos($HFtext,'*L');
		$Cpos = strpos($HFtext,'*C');
		$Rpos = strpos($HFtext,'*R');
		if ($Lpos){ //Find the left section text if it exists
			if ($Cpos){
				$len = $Cpos - $Lpos - 2;
				$Ltext = substr($HFtext,$Lpos+2,$len);
			} else if ($Rpos){
				$len = $Rpos - $Lpos - 2;
				$Ltext = substr($HFtext,$Lpos+2,£len);
			} else {
				$Ltext = substr($HFtext,$Lpos+2);
			}
			$LHF = $this->HFsect($Ltext);
		}
		if ($Cpos){ //Find the centre section text if it exists
			if ($Rpos){
				$len = $Rpos - $Cpos - 2;
				$Ctext = substr($HFtext,$Cpos+2,$len);
			} else {
				$Ctext = substr($HFtext,$Cpos+2);
			}
			$CHF = $this->HFsect($Ctext);
		}
		if ($Rpos){ //Find the right section text if it exists
			$Rtext = substr($HFtext,$Rpos+2);
			$RHF = $this->HFsect($Rtext);
		}
		
		if (($LHF AND $CHF AND $RHF) OR ($LHF AND $CHF) OR ($CHF AND $RHF)){
			$text = "<table width='100%'><tr><td width='33%' style='text-align:left; '>".$LHF."</td><td width='33%' style='text-align:centre; '>".$CHF."</td><td width='33%' style='text-align:right; '>".$RHF."</td></tr></table>";
		} else if ($LHF AND !$CHF AND $RHF){
			$text = "<table width='100%'><tr><td width='50%' style='text-align:left; '>".$LHF."</td><td width='50%' style='text-align:right; '>".$RHF."</td></tr></table>";
		} else if ($LHF){
			$text = "<table width='100%'><tr><td style='text-align:left; '>".$LHF."</td></tr></table>";
		} else if ($CHF){
			$text = "<table width='100%'><tr><td style='text-align:center; '>".$LHF."</td></tr></table>";
		} else if ($RHF){
			$text = "<table width='100%'><tr><td style='text-align:right; '>".$LHF."</td></tr></table>";
		} 
		return $text;
	}
	
	

	/**
	 * CALCULATE STANDARD DEVIATION
	 *  
	 * @param array $my_arr - The array
	 * @return string - The Standard Deviation of the array
	 */
	private function std_deviation($my_arr)
	{
	   $no_element = count($my_arr);
	   $var = 0.0;
	   $avg = array_sum($my_arr)/$no_element;
	   foreach($my_arr as $i)
	   {
		  $var += pow(($i - $avg), 2);
	   }
	   return (float)sqrt($var/$no_element);
	}




	
	/**
	 * PROCESS A SHEET
	 *  
	 * @param string $content - The XML node content
	 * @return string - The HTML code of the table row
	 */
	private function checkSheet($Nsheet)
	{
		$zip = new ZipArchive();
		$_xml_sheet = 'xl/worksheets/sheet'.$Nsheet.'.xml';
		if (true === $zip->open($this->file)) {
			//Get the sheet file
			if (($index = $zip->locateName($_xml_sheet)) !== false) {
				$xml_sheet = $zip->getFromIndex($index);
			}
			$zip->close();
		}

		$enc = mb_detect_encoding($xml_sheet);
		$this->setXmlParts($this->sheet_xml, $xml_sheet, $enc);
		if($this->debug) {
			echo "<br>XML File : xl/worksheets/sheet".$Nsheet.".xml<br>";
			echo "<textarea style='width:100%; height: 200px;'>";
			echo $this->sheet_xml->saveXML();
			echo "</textarea>";
		}

		$reader = new XMLReader;
		$reader->XML($this->sheet_xml->saveXML());
		$mergeno = $confor = 0;
		$Mfound = '';
		$cell = array();  //cell contents - a number or a reference to the string
		$Sdata = array();
		$Ddata = array();
		$Ffirst = array();
		$Flast = array();
		$Cwidth = array();
		$crange = array();
		$tst = $temp = -1;
		$text = $Ccount = '';
		$CF = $Xcount = $Tpriority = $dB = $lc = 0;
		
		$cs = '';
		$fr = $ctype = '';
		while ($reader->read()) {
			if ($reader->name == 'dimension') {
				$range = $reader->getAttribute("ref"); //defined cell range of spreadsheet
			}
			if ($reader->name == 'col') {
				$c1 = $reader->getAttribute("min"); //column number
				$c2 = $reader->getAttribute("max");
				$cw = $reader->getAttribute("width"); // defined column width
				$Cwidth[$c1] = (int)($cw * $this->CWFactor);
			}

			if ($reader->name == 'row') {
				if ($reader->getAttribute("spans")){
					$rr = $reader->getAttribute("r"); //row number
					if ($fr == ''){
						$rf = $rr;
						$fr = 'Y';
					}
				}
				if ($reader->getAttribute("ht")){
					$rh = $reader->getAttribute("ht"); //defined row height
					$Rhight[$rr] = (int)($rh * $this->RHFactor);
				}
			}
			if ($reader->name == 'c') {
				if ($reader->nodeType <> XMLREADER::END_ELEMENT) {
				++$tst;
					$cellno[$tst] = $reader->getAttribute("r"); //cell number
					$inv[$cellno[$tst]] = $tst;
					if ($reader->getAttribute("t")){
						$datasource = $reader->getAttribute("t"); //get reference to shared string text
					} else {
						$datasource = '';
					}
					$Ddata[$cellno[$tst]] = $datasource;
					if ($reader->getAttribute("s")){
						$stylesource = $reader->getAttribute("s"); //get reference to cell styling
					} else {
						$stylesource = 0;
					}
					$Sdata[$cellno[$tst]] = $stylesource;
				}
			}
			if ($reader->nodeType == XMLREADER::END_ELEMENT && $reader->name == 'row') {
				if ($tt > $temp){ // Finding the last column used
					$cd = $this->charstonum($cellno[$tt]); 
					if($cd['char'] > $lc){
						$lc = $cd['char'];
					}
					$temp = $tt;
				}
			}			
			if ($reader->nodeType == XMLREADER::ELEMENT && $reader->name == 'v') {
				$cell[$tst] = htmlentities($reader->expand()->textContent); // get a number or a reference to the text held in shared strings (referenced by attribute 't')
				$lr = $rr; //to find the last used row in the spreadsheet
				$tt = $tst;
			}
			if ($reader->nodeType == XMLREADER::ELEMENT && $reader->name == 'mergeCells') {
				$mergenum = $reader->getAttribute("count"); //number of merged cell ranges
			}
			
			if ($reader->name == 'mergeCell') {
				$Cmerge[$mergeno] = $reader->getAttribute("ref"); //merged cell range
				++$mergeno;
			}
			if ($reader->nodeType == XMLREADER::ELEMENT && $reader->name == 'oddHeader') {
				$Thead = " ".htmlentities($reader->expand()->textContent);
			}
			if ($reader->nodeType == XMLREADER::ELEMENT && $reader->name == 'oddFooter') {
				$Tfoot = " ".htmlentities($reader->expand()->textContent);
			}
			// Conditional formatting
			if ($reader->nodeType == XMLREADER::ELEMENT && $reader->name == 'conditionalFormatting') {
				$crange[$confor] = $reader->getAttribute("sqref"); // Conditional Formatting cell range
			}
			if ($reader->nodeType == XMLREADER::ELEMENT && $reader->name == 'xm:sqref') {
				if ($dfstype[$confor] == 'dataBar'){
					$dBref[$dB] = htmlentities($reader->expand()->textContent);
				}
				$crange[$confor] = htmlentities($reader->expand()->textContent); // Referenced Conditional Formatting cell range
			}
			if (($reader->nodeType == XMLREADER::ELEMENT && $reader->name == 'cfRule') OR ($reader->nodeType == XMLREADER::ELEMENT && $reader->name == 'x14:cfRule')){
				$priority = $reader->getAttribute("priority"); // Conditional Format priority
				if ($priority < $Tpriority OR $Tpriority == 0){
					$dfstype[$confor] = $reader->getAttribute("type"); // type of condition
					if ($reader->getAttribute("dxfId")){
						$dfsref[$confor] = $reader->getAttribute("dxfId"); // Conditional Format reference
					}
					if ($reader->getAttribute("dxfId") === '0'){
						$dfsref[$confor] = 0; // Conditional Format reference
					}
					$dfsop[$confor] = $reader->getAttribute("operator"); // Conditional Format type
					if ($reader->getAttribute("text")){
						$dfstext[$confor] = $reader->getAttribute("text"); // for text
					}
					$dfpcent[$confor] = $reader->getAttribute("percent"); // for top10 formatting
					$dfbott[$confor] = $reader->getAttribute("bottom"); // for top10 formatting
					$dfrank[$confor] = $reader->getAttribute("rank"); // for top10 formatting
					$dfUave[$confor] = 'T';
					if ($reader->getAttribute("aboveAverage") === '0'){
						$dfUave[$confor] = 'B'; // for over/under average formatting
					}
					$dfEave[$confor] = $reader->getAttribute("equalAverage"); // for over/under average formatting
					$dfSdev[$confor] = $reader->getAttribute("stdDev"); // for over/under average 					
					$Tpriority = $priority;
				} else {
					$Ccount = 'Y';
				}
				$CF = 0;
				$csa = $csb = 0;
				$cs = 'Y';
			}
			if (($reader->nodeType == XMLREADER::ELEMENT && $reader->name == 'formula') AND  $Ccount == '' AND $CF == 0) {
				$form1 = htmlentities($reader->expand()->textContent);
				if (substr($form1,0,1) == '$'){
					$form1 = str_replace('$','', $form1);
					if ($Ddata[$form1] == ''){
						$Cform1[$confor] = $cell[$inv[$form1]];
					} else {
						$Cform1[$confor] = $this->shared[$cell[$inv[$form1]]];
					}
				} else {
					$Cform1[$confor] = $form1;
				}
				$CF = '1';
			}
			if (($reader->nodeType == XMLREADER::ELEMENT && $reader->name == 'formula') AND  $Ccount == '' AND $CF == 1) {
				$form2 = htmlentities($reader->expand()->textContent);
				if (substr($form2,0,1) == '$'){
					$form2 = str_replace('$','', $form2);
					if ($Ddata[$form2] == ''){
						$Cform2[$confor] = $cell[$inv[$form2]];
					} else {
						$Cform2[$confor] = $this->shared[$cell[$inv[$form2]]];
					}
				} else {
					$Cform2[$confor] = $form2;
				}
			}
			if ($reader->name == 'cfvo'){
				$cfvo[$confor][$csa] = $reader->getAttribute("type");
				if ($cfvo[$confor][$csa] == 'num' OR $cfvo[$confor][$csa] == 'percent' OR $cfvo[$confor][$csa] == 'percentile'){
					$form2 = $reader->getAttribute("val"); //min/max for database
					if (substr($form2,0,1) == '$'){
						$form2 = str_replace('$','', $form2);
						$cfvoval[$confor][$csa] = $cell[$inv[$form2]];
					} else {
						$cfvoval[$confor][$csa] = $form2;
					}
					
				}
				++$csa;
			}
			if (($reader->name == 'color') AND ($cs == 'Y')){
				if ($reader->getAttribute("rgb")){
					$CScolour[$confor][$csb] = substr($reader->getAttribute("rgb"),2);
				} else {
					$Ctheme = (int)$reader->getAttribute("theme");
					if ($reader->getAttribute("tint")){
						$Ctint = strval(round($reader->getAttribute("tint"),2));
					} else {
						$Ctint = 0;
					}
					$Trgb = $this->themecol[$Ctheme]; // the rgb theme colour
					if ($Ctint == 0){
						$CScolour[$confor][$csb] = $Trgb;
					} else {
						$CScolour[$confor][$csb] = $this->calctint($Trgb, $Ctint); // the calculated theme tint
					}
				}
				++$csb;
			}
			if ($reader->nodeType == XMLREADER::END_ELEMENT && $reader->name == 'cfRule'){
				$cs = '';
			}
			if ($reader->nodeType == XMLREADER::ELEMENT && $reader->name == 'xm:f') {
				$Xtemp = htmlentities($reader->expand()->textContent); // get value link to conditional formatting
				$Xtemp = str_replace('$','', $Xtemp);
				if ($Ddata[$Xtemp] == ''){
					$Xvalue = $Cform1[$confor] = $cell[$inv[$Xtemp]];
				} else {
					$Xvalue = $dfstext[$confor] = $this->shared[$cell[$inv[$Xtemp]]];
				}
			}
			if ($reader->nodeType == XMLREADER::ELEMENT && $reader->name == 'x14:dataBar') {
				$dBBord[$dB] = 0;
				$dBGrad[$dB] = 1;
				if ($reader->getAttribute("border")){
					$dBBord[$dB] = $reader->getAttribute("border");
				}
				if ($reader->getAttribute("gradient") === '0'){
					$dBGrad[$dB] = $reader->getAttribute("gradient");
				}
			}
			if (($reader->name == 'x14:borderColor') AND ($dfstype[$confor] == 'dataBar')){
				if ($reader->getAttribute("rgb")){
					$dBBcol[$dB] = substr($reader->getAttribute("rgb"),2);
				} else {
					$Ctheme = (int)$reader->getAttribute("theme");
					if ($reader->getAttribute("tint")){
						$Ctint = strval(round($reader->getAttribute("tint"),2));
					} else {
						$Ctint = 0;
					}
					$Trgb = $this->themecol[$Ctheme]; // the rgb theme colour
					if ($Ctint == 0){
						$dBBcol[$dB] = $Trgb;
					} else {
						$dBBcol[$dB] = $this->calctint($Trgb, $Ctint); // the calculated theme tint
					}
				}
			}
			if (($reader->name == 'x14:negativeFillColor') AND ($dfstype[$confor] == 'dataBar')){
				if ($reader->getAttribute("rgb")){
					$dBNFcol[$dB] = substr($reader->getAttribute("rgb"),2);
				} else {
					$Ctheme = (int)$reader->getAttribute("theme");
					if ($reader->getAttribute("tint")){
						$Ctint = strval(round($reader->getAttribute("tint"),2));
					} else {
						$Ctint = 0;
					}
					$Trgb = $this->themecol[$Ctheme]; // the rgb theme colour
					if ($Ctint == 0){
						$dBNFcol[$dB] = $Trgb;
					} else {
						$dBNFcol[$dB] = $this->calctint($Trgb, $Ctint); // the calculated theme tint
					}
				}
			}
			if (($reader->name == 'x14:negativeBorderColor') AND ($dfstype[$confor] == 'dataBar')){
				if ($reader->getAttribute("rgb")){
					$dBNBcol[$dB] = substr($reader->getAttribute("rgb"),2);
				} else {
					$Ctheme = (int)$reader->getAttribute("theme");
					if ($reader->getAttribute("tint")){
						$Ctint = strval(round($reader->getAttribute("tint"),2));
					} else {
						$Ctint = 0;
					}
					$Trgb = $this->themecol[$Ctheme]; // the rgb theme colour
					if ($Ctint == 0){
						$dBNBcol[$dB] = $Trgb;
					} else {
						$dBNBcol[$dB] = $this->calctint($Trgb, $Ctint); // the calculated theme tint
					}
				}
			}
			if (($reader->name == 'x14:axisColor') AND ($dfstype[$confor] == 'dataBar')){
				if ($reader->getAttribute("rgb")){
					$dBAcol[$dB] = substr($reader->getAttribute("rgb"),2);
				} else {
					$Ctheme = (int)$reader->getAttribute("theme");
					if ($reader->getAttribute("tint")){
						$Ctint = strval(round($reader->getAttribute("tint"),2));
					} else {
						$Ctint = 0;
					}
					$Trgb = $this->themecol[$Ctheme]; // the rgb theme colour
					if ($Ctint == 0){
						$dBAcol[$dB] = $Trgb;
					} else {
						$dBAcol[$dB] = $this->calctint($Trgb, $Ctint); // the calculated theme tint
					}
				}
			}
			if ($reader->nodeType == XMLREADER::END_ELEMENT && $reader->name == 'conditionalFormatting') {
				$Tpriority = 0;
				$Ccount = '';
				++$confor;
				$end1 = $confor;
			}
			if ($reader->nodeType == XMLREADER::END_ELEMENT && $reader->name == 'x14:conditionalFormatting'){
				$Tpriority = 0;
				$Ccount = '';
				++$Xcount;
				if ($dfstype[$confor] == 'dataBar'){
					++$dB;
				}
				++$confor;
			}

		}
		if ($tst > 0 ){ //Check to see that the sheet is not blank
			$text = "<table style='".$defstyle." border-collapse: separate; border-spacing: 0px; margin-left: auto; margin-right: auto;'>"; // start the table if there is data in the spreadsheet
			//Get first and last cell in the spreadsheet and their cell/row numbers
			$this->charstonum($cellname);
			$sheetsize = explode(':', $range);
			$ab = $this->charstonum($sheetsize[0]);
			$Cfirst = $ab['char'];
			$Rfirst = $ab['num'];
			if ($Rfirst < $rf){ //check to see that there aren't some blank rows above the first occupies cell that are in the range noted in the worksheet
				$Rfirst = $rf;
			}
			$yz = $this->charstonum($sheetsize[1]);
			$Clast = $yz['char'];
			$Rlast = $yz['num'];
			if ($Clast > $lc){ //check to see that there aren't some blank columns beyond the last occupied cell that are in the range noted in the worksheet
				$Clast = $lc;
			}
			if ($Rlast > $lr){ //check to see that there aren't some blank rows beyond the last occupied cell that are in the range noted in the worksheet
				$Rlast = $lr + 1;
			}

			//Details of merged cells needed for findstyles to determine the right and bottom borders of merged cells
			$b = 0;
			while($b < $mergeno){
				$mm = explode(":",$Cmerge[$b]);
				$Mfirst[$b] = $mm[0];

				$Ffirst[$b] = $Sdata[$mm[0]];
				$Mlast[$b] = $mm[1];
				$Flast[$b] = $Sdata[$mm[1]];
				
				$Ixr[$b] = '';
				++$b;
			}
			
			//get the position of images from the drawing xml sheet and add it to the merge array
			$mmc = $b;
			$array = $this->drawings($Nsheet);
			$Fimage = $array[0];
			$Limage = $array[1];
			$Iname = $array[2];
			$Imxs = $array[3];
			$Imys  = $array[4];
			$pictno = $array[5];
			$ii = 0;
			if ($pictno > 0){
				$mergeno = $b + $pictno;
				while ($b < $mergeno){
					$Mfirst[$b] = $Fimage[$ii];
					$Ffirst[$b] = $Sdata[$Mfirst[$b]];
					$ab = $this->charstonum($Mfirst[$b]); //check to see if the top left of the image is outwith the spreadsheet range
					if($ab['char'] < $Cfirst){
						$Cfirst = $ab['char'];
					}
					if($ab['num'] < $Rfirst){
						$Rfirst = $ab['num'];
					}
					$Mlast[$b] = $Limage[$ii];
					$Flast[$b] = $Sdata[$Mlast[$b]];
					$ab = $this->charstonum($Mlast[$b]); //check to see if the bottom right of the image is outwith the spreadsheet range
					if($ab['char'] > $Clast){
						$Clast = $ab['char'];
					}
					if($ab['num'] > $Rlast){
						$Rlast = $ab['num'];
					}
					$Ixr[$b] = 'Y';
					++$b;
					++$ii;
				}
			}
			
			$Cellstyle = $this->findstyles($Ffirst,$Flast); // Sends the first and last cell of merges/images and gets the style parameters from the styles XML file
			if ($Xcount > 0){
				$Ccount = $Cellstyle[0]['dxf'];
				// Get the additional 'text' conditional formatting
				$CF2 = $this->getstyles('sheet', $Ccount, 'x14:conditionalFormatting');
				$d = 0;
				while ($d < $Xcount){
					$Cellstyle[$Ccount]['Cfscript'] = $CF2[$Ccount]['fscript'];
					$Cellstyle[$Ccount]['Cfstrike'] = $CF2[$Ccount]['fstrike'];
					$Cellstyle[$Ccount]['Cfbold'] = $CF2[$Ccount]['fbold'];
					$Cellstyle[$Ccount]['Cfund'] = $CF2[$Ccount]['fund'];
					$Cellstyle[$Ccount]['Cfital'] = $CF2[$Ccount]['fital'];
					$Cellstyle[$Ccount]['Cfsize'] = $CF2[$Ccount]['fsize'];
					$Cellstyle[$Ccount]['Cfcol'] = $CF2[$Ccount]['fcol'];
					$Cellstyle[$Ccount]['Cfname'] = $CF2[$Ccount]['fname'];
					$Cellstyle[$Ccount]['Cbleft'] = $CF2[$Ccount]['bleft'];
					$Cellstyle[$Ccount]['Cbright'] = $CF2[$Ccount]['bright'];
					$Cellstyle[$Ccount]['Cbtop'] = $CF2[$Ccount]['btop'];
					$Cellstyle[$Ccount]['Cbbott'] = $CF2[$Ccount]['bbott'];
					$Cellstyle[$Ccount]['Cfill'] = $CF2[$Ccount]['fill'];
					
					$dfsref[$end1] = $Ccount;
					++$end1;
					++$Ccount;
					++$d;
				}
			}

			if ($Thead){
				$Sinfo['head'] = $this->HeadFoot($Thead);				
			}
			if ($Tfoot){
				$Sinfo['foot'] = $this->HeadFoot($Tfoot);
			}
			
			
			
			// start processing and inclusion of conditional formatting
			$cond = 0;
			$Atemp = array();
			while ($cond < $confor){
				if ($dfstype[$cond] == 'duplicateValues' OR $dfstype[$cond] == 'uniqueValues' OR $dfstype[$cond] == 'colorScale' OR $dfstype[$cond] == 'dataBar' OR $dfstype[$cond] == 'top10' OR $dfstype[$cond] == 'aboveAverage'){
					$consize = explode(':', $crange[$cond]);
					$ab = $this->charstonum($consize[0]);
					$CCfirst = $ab['char'];
					$CRfirst = $ab['num'];
					$yz = $this->charstonum($consize[1]);
					$CClast = $yz['char'];
					$CRlast = $yz['num'];
					// start of finding any duplicated cell values in areas of duplicate conditional formatting
					if ($dfstype[$cond] == 'duplicateValues' OR $dfstype[$cond] == 'uniqueValues'){
						$a = 0;
						$duparr = array();
						while ($CCfirst <= $CClast){
							while ($CRfirst <= $CRlast){
								$tcl = $this->numtochars($CCfirst).$CRfirst;
								if ($Ddata[$tcl] == ''){
									$duparr[$a] = $cell[$inv[$tcl]];
								} else {
									$duparr[$a] = $this->shared[$cell[$inv[$tcl]]];
								}
								++$a;
								++$CRfirst;
							}
							++$CCfirst;
						}
						foreach (array_count_values($duparr) as $value => $count) {
							if ($count > 1) {
								$dupl[$cond][] = $value;
							}
						}
					}
					// end of finding any duplicated cell values in areas of duplicate conditional formatting
					// Start of finding values etc for colorScale, dataBar, top10 and aboveAverage conditional formatting
					if ($dfstype[$cond] == 'colorScale' OR $dfstype[$cond] == 'dataBar' OR $dfstype[$cond] == 'top10' OR $dfstype[$cond] == 'aboveAverage'){
						$csv = 0;
						$CSsum[$cond] = $CSmin[$cond] = $CSmax[$cond] = 0;
						$tt = 0;
						$AAtot = 0;
						while ($CCfirst <= $CClast){
							while ($CRfirst <= $CRlast){
								$tcl = $this->numtochars($CCfirst).$CRfirst;
								$Atemp[$cond][$tt] = $cell[$inv[$tcl]];
								++$tt;
								if ($cell[$inv[$tcl]]){
									$CSsum[$cond] = $CSsum[$cond] + $cell[$inv[$tcl]];
									if ($csv == 0){
										$CSmax[$cond] = $CSmin[$cond] = $cell[$inv[$tcl]];
									} else if ($cell[$inv[$tcl]] < $CSmin[$cond]){
										$CSmin[$cond] = $cell[$inv[$tcl]];
									} else if ($cell[$inv[$tcl]] > $CSmax[$cond]){
										$CSmax[$cond] = $cell[$inv[$tcl]];
									}
									$AAtot += $cell[$inv[$tcl]];
									++$csv;
								}
								++$CRfirst;
							}
							++$CCfirst;
						}
						sort($Atemp[$cond]);
						// find value for aboveAverage cond formatting
						if ($dfstype[$cond] == 'aboveAverage'){
							if ($dfSdev[$cond]){
								$sdev = $this->std_deviation($Atemp[$cond]);
								if ($dfSdev[$cond] == 1){
								$Aave[$cond] = ($AAtot / $tt) + $sdev;
								$Bave[$cond] = ($AAtot / $tt) - $sdev;
								} else if ($dfSdev[$cond] == 2){
								$Aave[$cond] = ($AAtot / $tt) + (2 * $sdev);
								$Bave[$cond] = ($AAtot / $tt) - (2 * $sdev);
								} else if ($dfSdev[$cond] == 3){
								$Aave[$cond] = ($AAtot / $tt) + (3 * $sdev);
								$Bave[$cond] = ($AAtot / $tt) - (3 * $sdev);
								}
							} else {
								$Aave[$cond] = $Bave[$cond] = $AAtot / $tt;
							}
							
						} 
						// find value for top10 cond formatting
						if ($dfstype[$cond] == 'top10'){
							if ($dfpcent[$cond] == 1){
								$drank = intval($tt * $dfrank[$cond] / 100);
							} else {
								$drank = $dfrank[$cond];
							}
							if ($dfbott[$cond] == 1){ // bottom ranking
								$dr = $drank - 1;
								$t10[$cond] = $Atemp[$cond][$dr];
								$ttype[$cond] = "B";
							} else { //top ranking
								$dr = $tt - $drank;
								$t10[$cond] = $Atemp[$cond][$dr];
								$ttype[$cond] = "T";
							}
						}
						//find min, middle and max values for colorscale formatting
						if ($dfstype[$cond] == 'colorScale'){
							//determine displayed min value
							if ($cfvo[$cond][0] == 'percentile'){
								$pindex = $cfvoval[$cond][0] / 100 * ($tt + 1);
								$a1 = floor($pindex)-1;
								$a2 = ceil($pindex)-1;
								$a3 = $pindex -1 - $a1;
								if ($pindex == ($a1 + 1)){
									$CSmin1[$cond] = $Atemp[$cond][$a1];
								} else {
									$CSmin1[$cond] = $Atemp[$cond][$a1] + ($a3 * ($Atemp[$cond][$a2] - $Atemp[$cond][$a1]));
								}
							} else if ($cfvo[$cond][0] == 'percent'){
								$CSmin1[$cond] = ($cfvoval[$cond][0]/100 * ($CSmax[$cond] - $CSmin[$cond])) + $CSmin[$cond];
							} else if ($cfvo[$cond][0] == 'num'){
								$CSmin1[$cond] = $cfvoval[$cond][0];
							} else {
								$CSmin1[$cond] = $CSmin[$cond];
							}
							if ($cfvo[$cond][2]){
								$tmp = 2;
								//determine middle colour value for a 3 colour colorScale
								if ($cfvo[$cond][1] == 'percentile'){
									$pindex = $cfvoval[$cond][1] / 100 * ($tt + 1);
									$a1 = floor($pindex)-1;
									$a2 = ceil($pindex)-1;
									$a3 = $pindex -1 - $a1;
									if ($pindex == ($a1 + 1)){
										$CSave[$cond] = $Atemp[$cond][$a1];
									} else {
										$CSave[$cond] = $Atemp[$cond][$a1] + ($a3 * ($Atemp[$cond][$a2] - $Atemp[$cond][$a1]));
									}
								} else if ($cfvo[$cond][1] == 'percent'){
									$CSave[$cond] = ($cfvoval[$cond][1]/100 * ($CSmax[$cond] - $CSmin[$cond])) + $CSmin[$cond];
								} else if ($cfvo[$cond][1] == 'num'){
									$CSave[$cond] = $cfvoval[$cond][1];
								}
							} else {
								$tmp = 1;
							}
							// determine displayed max value
							if ($cfvo[$cond][$tmp] == 'percentile'){
								$pindex = $cfvoval[$cond][$tmp] / 100 * ($tt + 1);
								$a1 = floor($pindex)-1;
								$a2 = ceil($pindex)-1;
								$a3 = $pindex -1 - $a1;
								if ($pindex == ($a1 + 1)){
									$CSmax1[$cond] = $Atemp[$cond][$a1];
								} else {
									$CSmax1[$cond] = $Atemp[$cond][$a1] + ($a3 * ($Atemp[$cond][$a2] - $Atemp[$cond][$a1]));
								}
							} else if ($cfvo[$cond][$tmp] == 'percent'){
								$CSmax1[$cond] = ($cfvoval[$cond][$tmp]/100 * ($CSmax[$cond] - $CSmin[$cond])) + $CSmin[$cond];
							} else if ($cfvo[$cond][$tmp] == 'num'){
								$CSmax1[$cond] = $cfvoval[$cond][$tmp];
							} else {
								$CSmax1[$cond] = $CSmax[$cond];
							}
							
						}
						//Find the min and max etc. values of the displayed dataBar
						if ($dfstype[$cond] == 'dataBar' AND $CScolour[$cond][0] <> ''){
							if ($cfvoval[$cond][0]){
								if ($cfvo[$cond][0] == 'percent'){
									$Range = $CSmax[$cond] - $CSmin[$cond];
									$dbmin[$cond] = (($cfvoval[$cond][0] / 100) * $Range) + $CSmin[$cond];
								} else if ($cfvo[$cond][0] == 'percentile'){
									$pindex = $cfvoval[$cond][0] / 100 * (count($Atemp[$cond]) + 1);
									$a1 = floor($pindex)-1;
									$a2 = ceil($pindex)-1;
									$a3 = $pindex -1 - $a1;
									if ($pindex == ($a1 + 1)){
										$dbmin[$cond] = $Atemp[$cond][$a1];
									} else {
										$dbmin[$cond] = $Atemp[$cond][$a1] + ($a3 * ($Atemp[$cond][$a2] - $Atemp[$cond][$a1]));
									}
								} else if ($cfvo[$cond][0] == 'num'){
									$dbmin[$cond] = $cfvoval[$cond][0];
								}
							} else {
								if ($CSmin[$cond] < 0){
									$dbmin[$cond] = $CSmin[$cond];
								} else {
									$dbmin[$cond] = 0;
								}
							}
							if ($cfvoval[$cond][1]){
								if ($cfvo[$cond][1] == 'percent'){
									$Range = $CSmax[$cond] - $CSmin[$cond];
									$dbmax[$cond] = (($cfvoval[$cond][1] / 100) * $Range) + $CSmin[$cond];
								} else if ($cfvo[$cond][1] == 'percentile'){
									$pindex = $cfvoval[$cond][1] / 100 * (count($Atemp[$cond]) + 1);
									$a1 = floor($pindex)-1;
									$a2 = ceil($pindex)-1;
									$a3 = $pindex-1 - $a1;
									if ($pindex == ($a1 + 1)){
										$dbmax[$cond] = $Atemp[$cond][$a1];
									} else {
										$dbmax[$cond] = $Atemp[$cond][$a1] + ($a3 * ($Atemp[$cond][$a2] - $Atemp[$cond][$a1]));
									}
								} else {
									$dbmax[$cond] = $cfvoval[$cond][1];
								}
							} else {
								if ($CSmax[$cond] > 0){
									$dbmax[$cond] = $CSmax[$cond];
								} else {
									$dbmax[$cond] = 0;
								}
							}
							$dbdiff[$cond] = $dbmax[$cond] - $dbmin[$cond];
							if ($dbmin[$cond] < 0){
								$dbnegR[$cond] = (abs($dbmin[$cond]) / $dbdiff[$cond]) * 100;
							} else {
								$dbnegR[$cond] = 0;
							}
							if ($dbmax[$cond] > 0){
								$dbposR[$cond] = 100 - $dbnegR[$cond];
							}
							
						}
					}
					// End of finding values etc for colorScale and dataBar conditional formatting
				}
				++$cond;
			}

			// start of finding cells which need conditional formatting
			$cond = 0;
			while ($cond < $confor){
				$consize = explode(':', $crange[$cond]);
				$ab = $this->charstonum($consize[0]);
				$CCfirst = $ab['char'];
				$CRfirst = $ab['num'];
				$yz = $this->charstonum($consize[1]);
				$CClast = $yz['char'];
				$CRlast = $yz['num'];
				while ($CCfirst <= $CClast){
					while ($CRfirst <= $CRlast){
						$cl = $this->numtochars($CCfirst).$CRfirst;
						if ($dfsop[$cond] == 'greaterThan'){
							if ($cell[$inv[$cl]] > $Cform1[$cond]){
								$cfound = 'Y';
							}
						} else if ($dfsop[$cond] == 'greaterThanOrEqual'){
							if ($cell[$inv[$cl]] >= $Cform1[$cond]){
								$cfound = 'Y';
							}
						} else if ($dfsop[$cond] == 'lessThan'){
							if ($cell[$inv[$cl]] < $Cform1[$cond]){
								$cfound = 'Y';
							}
						} else if ($dfsop[$cond] == 'lessThanOrEqual'){
							if ($cell[$inv[$cl]] <= $Cform1[$cond]){
								$cfound = 'Y';
							}
						} else if ($dfsop[$cond] == 'between'){
							if (($cell[$inv[$cl]] >= $Cform1[$cond]) AND ($cell[$inv[$cl]] <= $Cform2[$cond])){
								$cfound = 'Y';
							}
						} else if ($dfsop[$cond] == 'notBetween'){
							if (($cell[$inv[$cl]] < $Cform1[$cond]) OR ($cell[$inv[$cl]] > $Cform2[$cond])){
								$cfound = 'Y';
							}
						} else if ($dfsop[$cond] == 'equal'){
							if ($cell[$inv[$cl]] == $Cform1[$cond]){
								$cfound = 'Y';
							}
						} else if ($dfsop[$cond] == 'notEqual'){
							if ($cell[$inv[$cl]] <> $Cform1[$cond]){
								$cfound = 'Y';
							}
						} else if ($dfsop[$cond] == 'containsText'){
							$stext = " ".$this->shared[$cell[$inv[$cl]]]." ";
							if (stripos($stext,$dfstext[$cond])){
								$cfound = 'Y';
							}
						} else if ($dfsop[$cond] == 'notContains'){
							$stext = " ".$this->shared[$cell[$inv[$cl]]]." ";
							if (!stripos($stext,$dfstext[$cond])){
								$cfound = 'Y';
							}
						} else if ($dfsop[$cond] == 'beginsWith'){
							$temp = strlen($dfstext[$cond]);
							if (substr($this->shared[$cell[$inv[$cl]]],0,$temp) == $dfstext[$cond]){
								$cfound = 'Y';
							}
						} else if ($dfsop[$cond] == 'endsWith'){
							$temp = strlen($dfstext[$cond]) * -1;
							if (substr($this->shared[$cell[$inv[$cl]]],$temp) == $dfstext[$cond]){
								$cfound = 'Y';
							}
						} else if ($dfstype[$cond] == 'duplicateValues' OR $dfstype[$cond] == 'uniqueValues'){
							if ($Ddata[$cl] == ''){
								$ctemp = $cell[$inv[$cl]];
							} else {
								$ctemp = $this->shared[$cell[$inv[$cl]]];
							}
							$q = 0;
							$tcfound = '';
							while ($dupl[$cond][$q]){
								if ($dupl[$cond][$q] == $ctemp){
									$tcfound = 'Y';
									}
								++$q;
							}
							if ($dfstype[$cond] == 'duplicateValues'){
								if ($tcfound == 'Y'){
									$cfound = 'Y';
								}
							} else {
								if ($tcfound == ''){
									$cfound = 'Y';
								}
							}
						} else if ($dfstype[$cond] == 'aboveAverage'){
							if ($dfUave[$cond] == 'B'){
								if ($dfEave[$cond] == '1'){
									if ($cell[$inv[$cl]] <= $Bave[$cond]){
										$cfound = 'Y';
									}
								} else {
									if ($cell[$inv[$cl]] < $Bave[$cond]){
										$cfound = 'Y';
									}
								}
							} else {
								if ($dfEave[$cond] == '1'){
									if ($cell[$inv[$cl]] >= $Aave[$cond]){
										$cfound = 'Y';
									}
								} else {
									if ($cell[$inv[$cl]] > $Aave[$cond]){
										$cfound = 'Y';
									}
								}
							}
						} else if ($dfstype[$cond] == 'top10'){
							if ($ttype[$cond] == 'B' AND  $cell[$inv[$cl]] <= $t10[$cond]){
								$cfound = 'Y';
							}
							if ($ttype[$cond] == 'T' AND $cell[$inv[$cl]] >= $t10[$cond]){
								$cfound = 'Y';
							}
						} else if ($dfstype[$cond] == 'colorScale'){
							$Css[$cl]['Cfill'] = " background-color: #".$this->findcolorScale($CScolour[$cond], $CSmin1[$cond], $CSmax1[$cond], $CSave[$cond], $cell[$inv[$cl]]).";";
						} else if ($dfstype[$cond] == 'dataBar' AND $CScolour[$cond][0] <> ''){
							$dbt = 0;
							while ($dbt < $dB){
								if ($crange[$cond] ==  $dBref[$dbt]){
									$Dfound = $dbt;
								}
								++$dbt;
							} 
							if ($Rhight[$CRfirst]){
								if ($Rhight[$CRfirst] > 20){
									$dbheight = $Rhight[$CRfirst] - 5;
								} 
							} else {
									$dbheight = 14;
								}
							if ($cell[$inv[$cl]] >= 0){
								if ($dbmin[$cond] < 0 ){
									$dbval = ($cell[$inv[$cl]] / $dbmax[$cond]) * $dbposR[$cond];
									
								} else {
									$dbval = (($cell[$inv[$cl]] - $dbmin[$cond]) / ($dbmax[$cond] - $dbmin[$cond])) * $dbposR[$cond];
								}
								if ($dbval > $dbposR[$cond]){
									$dbval = $dbposR[$cond];
								}
								if ($dbval < 0){
									$dbval = 0;
								}
								$Css[$cl]['dBfill'] = "<div style = 'width: ".$dbval."%; margin-top: 1px; height: ".$dbheight."px;";
								if ($dBGrad[$Dfound] == 1){
									$Css[$cl]['dBfill'] .= " background: linear-gradient(to right, #".$CScolour[$cond][0]." 0%, #FFFFFF 100%);";
								} else {
									$Css[$cl]['dBfill'] .= " background: #".$CScolour[$cond][0].";";
								}
								if ($dBBord[$Dfound] == 1 AND $dbval <> 0){
									$Css[$cl]['dBfill'] .= " border: 1px solid #".$dBBcol[$Dfound].";";
								} else {
									$Css[$cl]['dBfill'] .= " border-top: 1px solid #".$CScolour[$cond][0].";";
									$Css[$cl]['dBfill'] .= " border-bottom: 1px solid #".$CScolour[$cond][0].";";
								}
								if ($dbnegR[$cond] <> 0){
									$Css[$cl]['dBfill'] .= " margin-left: ".$dbnegR[$cond]."%; border-left: 1px dashed #".$dBAcol[$Dfound].";";
								}
								$Css[$cl]['dBfill'] .= " '></div>";
							}
							if ($cell[$inv[$cl]] < 0){
								if ($dbmax[$cond] < 0 ){
									$dbval2 = ($cell[$inv[$cl]] / $dbmin[$cond]) * $dbnegR[$cond];
								} else {
									$dbval2 = (($cell[$inv[$cl]] - $dbmin[$cond]) / (0 - $dbmin[$cond])) * $dbnegR[$cond];
								}
								if ($dbval2 > 100){
									$dbval2 = 100;
								}
								if ($dbval2 < 0){
									$dbval2 = 0;
								}
								$dbval = $dbnegR[$cond] - $dbval2;
								$Css[$cl]['dBfill'] = "<div style = 'width: ".$dbval."%; margin-top: 1px; height: ".$dbheight."px;";
							if ($dBGrad[$Dfound] == 1){
									$Css[$cl]['dBfill'] .= " background: linear-gradient(to right, #FFFFFF 0%, #".$dBNFcol[$Dfound]." 100%);";
								} else {
									$Css[$cl]['dBfill'] .= " background: #".$dBNFcol[$Dfound].";";
								}
								if ($dBBord[$Dfound] == 1 AND $dbval <> 0){
									$Css[$cl]['dBfill'] .= " border: 1px solid #".$dBNBcol[$Dfound].";";
								} else {
									$Css[$cl]['dBfill'] .= " border-top: 1px solid #".$dBNFcol[$Dfound].";";
									$Css[$cl]['dBfill'] .= " border-bottom: 1px solid #".$dBNFcol[$Dfound].";";
								}
								if ($dbval <> 0){
									$Css[$cl]['dBfill'] .= " margin-left: ".$dbval2."%;";
								}
								if ($dbmax[$cond] > 0){
									$Css[$cl]['dBfill'] .= "  border-right: 1px dashed #".$dBAcol[$Dfound].";";
								}
								$Css[$cl]['dBfill'] .= " '></div>";
							}
							
						}
						if ($cfound == 'Y'){
							if ($Cellstyle[$dfsref[$cond]]['Cfill']){
								$Css[$cl]['Cfill'] = $Cellstyle[$dfsref[$cond]]['Cfill'];
							}
							if($Cellstyle[$dfsref[$cond]]['Cbleft']){
								$Css[$cl]['Cbleft'] = $Cellstyle[$dfsref[$cond]]['Cbleft'];
							}
							if ($Cellstyle[$dfsref[$cond]]['Cbtop']){
								$Css[$cl]['Cbtop'] = $Cellstyle[$dfsref[$cond]]['Cbtop'];
							}
							if ($Cellstyle[$dfsref[$cond]]['Cbtop']){
								$Css[$cl]['Cbright'] = $Cellstyle[$dfsref[$cond]]['Cbright'];
							}
							if ($Cellstyle[$dfsref[$cond]]['Cbbott']){
								$Css[$cl]['Cbbott'] = $Cellstyle[$dfsref[$cond]]['Cbbott'];
							}
							if ($Cellstyle[$dfsref[$cond]]['Cfname']){
								$Css[$cl]['Cfname'] = $Cellstyle[$dfsref[$cond]]['Cfname'];
							}
							if ($Cellstyle[$dfsref[$cond]]['Cfsize']){
								$Css[$cl]['Cfsize'] = $Cellstyle[$dfsref[$cond]]['Cfsize'];
							}
							if ($Cellstyle[$dfsref[$cond]]['Cfcol']){
								$Css[$cl]['Cfcol'] = $Cellstyle[$dfsref[$cond]]['Cfcol'];
							}
							if ($Cellstyle[$dfsref[$cond]]['Cfbold']){
								$Css[$cl]['Cfbold'] = $Cellstyle[$dfsref[$cond]]['Cfbold'];
							}
							if ($Cellstyle[$dfsref[$cond]]['Cfund']){
								$Css[$cl]['Cfund'] = $Cellstyle[$dfsref[$cond]]['Cfund'];
							}
							if ($Cellstyle[$dfsref[$cond]]['Cfital']){
								$Css[$cl]['Cfital'] = $Cellstyle[$dfsref[$cond]]['Cfital'];
							}
							if ($Cellstyle[$dfsref[$cond]]['Cfscript']){
								$Css[$cl]['Cfscript'] = $Cellstyle[$dfsref[$cond]]['Cfscript'];
							}
							if ($Cellstyle[$dfsref[$cond]]['Cfstrike']){
								$Css[$cl]['Cfstrike'] = $Cellstyle[$dfsref[$cond]]['Cfstrike'];
							}
							$cfound = '';
						}
						++$CRfirst;
					}
					++$CCfirst;
				}
				++$cond;
			}

			//Start of processing data for each cell entry in the sheet
			$cc = 0;
			while ($cc <= $tst){ 
				//Start of processing number formatting
				if (($cell[$cc] >= 0 OR $cell[$cc] < 0) AND $cell[$cc] <> '' AND $Ddata[$cellno[$cc]] == ''){ 
					$temp = $temp2 = '';
					$Ncode = " ".$Cellstyle[$Sdata[$cellno[$cc]]]['nform'];
					if ($Cellstyle[$Sdata[$cellno[$cc]]]['nform'] == 'ZZZ'){
						$tnint = (int)$cell[$cc];
						$trem = $cell[$cc] - $tnint;
						$noh = ($tnint * 24) + (int)($trem * 24);
						$tom = (($trem * 24) - (int)($trem * 24)) * 60;
						$nom = (int)$tom;
						$nos = round(($tom - $nom) * 60);
						$cell[$cc] = $noh.":".$nom.":".$nos;
					} else if (substr($Cellstyle[$Sdata[$cellno[$cc]]]['nform'],-2,2) == ';@' OR strpos($Cellstyle[$Sdata[$cellno[$cc]]]['nform'],'mm')){
						// Start of apply date and time formatting to the relevant cell contents
						if (substr($Cellstyle[$Sdata[$cellno[$cc]]]['nform'],-2,2) == ';@' ){
							$temp = " ".substr($Cellstyle[$Sdata[$cellno[$cc]]]['nform'],0,-2);
						} else {
							$temp = " ".$Cellstyle[$Sdata[$cellno[$cc]]]['nform'];
						}
						if (strpos($temp,'h:m') OR strpos($temp,'m:s')){
							$patterns3 = array('/hh/', '/h]/', '/h/', '/:mm/', '/mm:/', '/ss/', '/0/');
							if (strpos($temp,'AM/PM')){
								$replace3 = array('f', 'g', 'g', ':i', 'i:', 's', 'v');
								$temp = preg_replace($patterns3, $replace3, $temp);
								$temp = preg_replace('/f/', 'h', $temp);
							} else {
								$replace3 = array('H', 'G', 'G', ':i', 'i:', 's', 'v');
								$temp = preg_replace($patterns3, $replace3, $temp);
							}
						}
						$patterns = array('/yyyy/', '/yy/','/mmmmm/','/mmmm/', '/mmm/', '/mm/', '/m/', '/dddd/', '/ddd/', '/dd/', '/d/');
						$replace = array('Y', 'y', 'M', 'F', 'M', 'o', 'n', 'l', 'D', 'e', 'j');
						$patterns2 = array('/e/', '/o/');
						$replace2 = array('d', 'm');
						$temp2 = preg_replace($patterns, $replace, $temp);
						$temp2 = str_replace("\\","",$temp2);
						$temp2 = preg_replace($patterns2, $replace2, $temp2);
						if (strpos($temp2,']')){
							$tpos = strpos($temp2,']');
							$temp2 = substr($temp2,$tpos+1);
						} else {
							$temp2 = substr($temp2,1);
							if (substr($temp2,0,1) == '['){
								$temp2 = substr($temp2,1);
							}
						}
						$nint = (int)$cell[$cc];
						$tdate=date_create("1899-12-30");
						date_add($tdate,date_interval_create_from_date_string($nint." days"));
						$ndec = $cell[$cc] - $nint;
						$MM = '';
						if ($ndec <> 0){
							$nsecs = floor($ndec * 24 * 3600);
							$tdate->add(new \DateInterval('PT'.$nsecs.'S'));
							$Msec = ($ndec * 24 * 3600) - $nsecs;
							if ($Msec <> 0){
								$Mmili = round($Msec * 10);
							}
							if (strpos($temp2,'v')){
								$MM = ".".$Mmili;
								$temp2 = str_replace(".v","",$temp2);
							}
						}
						$apm = '';
						if (strpos($temp2,'AM/PM')){
							$temp2 = str_replace("AM/PM","",$temp2);
							if ($nsecs < 12 * 3600){
								$apm = ' AM';
							} else {
								$apm = ' PM';
							}
						}
						$cell[$cc] = date_format($tdate,$temp2).$MM.$apm;
						// End of date/time cell contents formatting
					} else if (substr($Ncode,1,2) == '#\\'){ //Number formatting for fractions
						$Fint = (int)$cell[$cc];
						$Fdec = $cell[$cc] - $Fint;
						if (substr($Ncode,-1,1) == '2'){
							$DD = round($Fdec*2);
							$DI = 2;
						} else if (substr($Ncode,-1,1) == '4'){
							$DD = round($Fdec*4);
							$DI = 4;
						} else if (substr($Ncode,-1,1) == '8'){
							$DD = round($Fdec*8);
							$DI = 8;
						} else if (substr($Ncode,-2,2) == '16'){
							$DD = round($Fdec*16);
							$DI = 16;
						} else if (substr($Ncode,-2,2) == '10'){
							$DD = round($Fdec*10);
							$DI = 10;
						} else if (substr($Ncode,-3,3) == '100'){
							$DD = round($Fdec*100);
							$DI = 100;
						} else if (substr($Ncode,-4,4) == '/???'){
							$fract = $this->float2rat($Fdec, 1000);
							$DD = $fract['num'];
							$DI = $fract['den'];
						} else if (substr($Ncode,-3,3) == '/??'){
							$fract = $this->float2rat($Fdec, 100);
							$DD = $fract['num'];
							$DI = $fract['den'];
						} else if (substr($Ncode,-2,2) == '/?'){
							$fract = $this->float2rat($Fdec, 10);
							$DD = $fract['num'];
							$DI = $fract['den'];
						}
						if ($DD == 0){
							$cell[$cc] = $Fint."&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;";
						} else {
							$cell[$cc] = $Fint." ".$DD."/".$DI;
						}
					} else if (strpos($Ncode,'0')){
						$ff = strpos($Ncode,'#');
						if ($ff){
							$fff = substr($Ncode,$ff);
							$num0 = substr_count($fff,0);
						} else {
							$num0 = substr_count($Ncode,0);
						}					
						$Ndec = $num0 - 1; //number of decimal places
						if (substr($Ncode,1,1) == '0' OR substr($Ncode,1,1) == '#'){ //Find all number formats that start with a '0' or a '#'
							if (substr($Ncode,-1,1) == '%'){
								$temp = $cell[$cc] * 100; // %age number formats
								$Dpoint = substr($Ncode,2,1); //type of decimal point
								$cell[$cc] = number_format($temp,$Ndec,$Dpoint)."%"; //display %age numbers
							} else if (substr($Ncode,-1,1) == '0'){ //When alternatives for negative numbers and no currency unit
								if (strpos($Ncode,';')){
									$Dneg = explode(';',$Ncode);
									$num0 = substr_count($Dneg[0],0);
									$Ndec1 = $num0 - 1; //number of decimal places
									if ($cell[$cc] < 0){
										$Dneg[1] = " ".$Dneg[1];
										$curr = $this->currency($Dneg[1]);
										$cell[$cc] = abs($cell[$cc]);
									} else {
										$curr = $this->currency($Dneg[0]);
									}
									$minus = $curr['minus'];
									$sep = $curr['sep'];
									$Dpoint = $curr['point'];
									$red = $curr['red'];
									if ($red == 'Y'){
										$cell[$cc] = "<span style='color:red'>".$minus.number_format($cell[$cc],$Ndec1,$Dpoint,$sep)."</span>"; //Display with no currency unit
										
									} else {
										$cell[$cc] = $minus.number_format($cell[$cc],$Ndec1,$Dpoint,$sep); //Display currency with no currency unit
									}
										
									
								} else if (substr($Ncode,1,1) == '#'){ //plain number with thousands separator
									$sep = substr($Ncode,2,1); //type of separator
									$Dpoint = substr($Ncode,6,1); //type of decimal point
									$cell[$cc] = number_format($cell[$cc],$Ndec,$Dpoint,$sep); //Display numbers with a thousands separator
								} else if (strpos($Ncode,'E')){
									$Edec = $Ndec - 2;  // number of decimals
									$EE = "%.".$Edec."E";
									$cell[$cc] = sprintf($EE,$cell[$cc]); //display numbers in scientific notation 
								} else { 
									$Dpoint = substr($Ncode,2,1); //type of decimal point
									$cell[$cc] = number_format($cell[$cc],$Ndec,$Dpoint,''); //Display numbers without a thousands separator
								}
							} else { // for trailing currency symbols
								$minus = '';
								if (strpos($Ncode,';')){
									$Dneg = explode(';',$Ncode);
									$ff = strpos($Dneg[0],'[');
									$fff = substr($Dneg[0],0,$ff);
									$num0 = substr_count($fff,0);
									$Ndec1 = $num0 - 1; //number of decimal places
									if ($cell[$cc] < 0){
										$Dneg[1] = " ".$Dneg[1];
										$curr = $this->currency($Dneg[1]);
										$cell[$cc] = abs($cell[$cc]);
									} else {
										$curr = $this->currency($Dneg[0]);
									}
									$minus = $curr['minus'];
								} else {
									$ff = strpos($Ncode,'[');
									$fff = substr($Ncode,0,$ff);
									$num0 = substr_count($fff,0);
									$Ndec1 = $num0 - 1; //number of decimal places
									$curr = $this->currency($Ncode);
									if ($cell[$cc] < 0){
										$cell[$cc] = abs($cell[$cc]);
										$minus = '-';
									}
								}
								$sep = $curr['sep'];
								$Dpoint = $curr['point'];
								$red = $curr['red'];
								if ($red == 'Y'){
									$cell[$cc] = "<span style='color:red'>".$minus.number_format($cell[$cc],$Ndec1,$Dpoint,$sep)." ".$curr['unit']."</span>"; //Display currency with a trailing currency unit
									
								} else {
									$cell[$cc] = $minus.number_format($cell[$cc],$Ndec1,$Dpoint,$sep)." ".$curr['unit']; //Display currency with a trailing currency unit
								}

							}
						} else { //for leading currency symbols and all Accounting formats
							if (strpos($Ncode,';')){ //when there is an alternative formatting for -ve values
								$min = '';
								$Dneg = explode(';',$Ncode);
								$ff = strpos($Dneg[0],'#');
								$fff = substr($Dneg[0],$ff);
								$num0 = substr_count($fff,0);
								$Ndec1 = $num0 - 1; //number of decimal places
								$account = '';
								if (substr($Dneg[0],1,1) == '_'){ //this for Accounting alignment
									$account = 'Y';
									$Dneg[0] = substr($Dneg[0],0,2)."_".substr($Dneg[0],3);
									$curr = $this->currency($Dneg[0]);
									if ($cell[$cc] < 0){
										$cell[$cc] = abs($cell[$cc]);
										$min = "-";
									}
								} else if ($cell[$cc] < 0){
									$curr = $this->currency($Dneg[1]);
									$cell[$cc] = abs($cell[$cc]);
								} else {
									$curr = $this->currency($Dneg[0]);
								}
								$sep = $curr['sep'];
								$Dpoint = $curr['point'];
								$red = $curr['red'];
								$minus = $curr['minus'];
								if ($red == 'Y'){
									$cell[$cc] = "<span style='color:red'>".$minus.$curr['unit'].number_format($cell[$cc],$Ndec1,$Dpoint,$sep)."</span>"; //Display currency with a leading currency unit
									
								} else {
									if ($account == 'Y'){ //For Accounting alignment
										$clead = $ctrail = '';
										if ($curr['pos'] == 'T'){
											$ctrail = $curr['unit'];
										} else {
											$clead = $curr['unit'];
										}
										if ($cell[$cc] === '0'){
											$cell[$cc] = "<div style='float:left;'>&nbsp;".$min.$minus.$clead."</div><div style='float:right;'>-&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;".$ctrail."&nbsp;</div>"; // In accounting format '0' is replaced by a '-'
										} else {
											$cell[$cc] = "<div style='float:left;'>&nbsp;".$min.$minus.$clead."</div><div style='float:right;'>".number_format($cell[$cc],$Ndec1,$Dpoint,$sep)."&nbsp;".$ctrail."&nbsp;</div>"; //Display accountancy format with a leading or trailing currency unit
										}
									} else { //for normal alignment
										$cell[$cc] = $min.$minus.$curr['unit'].number_format($cell[$cc],$Ndec1,$Dpoint,$sep); //Display currency with a leading currency unit
									}
								}
								
							} else { //when there is no alternative formatting for -ve values
								$curr = $this->currency($Ncode);
								//Find leading currency unit
								$sep = $curr['sep'];
								$Dpoint = $curr['point'];
								if ($cell[$cc] < 0){
									$cell[$cc] = abs($cell[$cc]);
									$cell[$cc] = "-".$curr['unit'].number_format($cell[$cc],$Ndec,$Dpoint,$sep); //Display negative currency with a leading currency unit,							
								} else {
									$cell[$cc] = $curr['unit'].number_format($cell[$cc],$Ndec,$Dpoint,$sep); //Display currency with a leading currency unit
								}
							}
						}
					} else {
						if ($cell[$cc] < 10){
							$cell[$cc] = round($cell[$cc],9);
						} else if ($cell[$cc] < 100){
							$cell[$cc] = round($cell[$cc],8);
						} else if ($cell[$cc] < 1000){
							$cell[$cc] = round($cell[$cc],7);
						} else if ($cell[$cc] < 10000){
							$cell[$cc] = round($cell[$cc],6);
						} else if ($cell[$cc] < 100000){
							$cell[$cc] = round($cell[$cc],5);
						} else if ($cell[$cc] < 1000000){
							$cell[$cc] = round($cell[$cc],4);
						} else if ($cell[$cc] < 10000000){
							$cell[$cc] = round($cell[$cc],3);
						} else if ($cell[$cc] < 100000000){
							$cell[$cc] = round($cell[$cc],2);
						} else if ($cell[$cc] < 1000000000){
							$cell[$cc] = round($cell[$cc],1);
						} else {
							$cell[$cc] = round($cell[$cc],0);
						}
					}
				}
				// End of Number Format processing
				
				if ($Cellstyle[$Sdata[$cellno[$cc]]]['hyper'] == 'Hyperlink'){
					$test[$cc] = "<a href='".$this->shared[$cell[$cc]]."'>".$this->shared[$cell[$cc]]."</a>";
					$this->shared[$cell[$cc]] = $test[$cc]; //Adds in the hyperlink code for a hyperlink
				}
				
				// get text/number formatting
				if ($Css[$cellno[$cc]]['Cfname']){
					$Tfname = $Css[$cellno[$cc]]['Cfname'];
				} else {
					$Tfname = $Cellstyle[$Sdata[$cellno[$cc]]]['fname'];
				}
				if ($Css[$cellno[$cc]]['Cfsize']){
					$Tfsize = $Css[$cellno[$cc]]['Cfsize'];
				} else {
					$Tfsize = $Cellstyle[$Sdata[$cellno[$cc]]]['fsize'];
				}
				if ($Css[$cellno[$cc]]['Cfcol']){
					$Tfcol = $Css[$cellno[$cc]]['Cfcol'];
				} else {
					$Tfcol = $Cellstyle[$Sdata[$cellno[$cc]]]['fcol'];
				}
				if ($Css[$cellno[$cc]]['Cfbold']){
					$Tfbold = $Css[$cellno[$cc]]['Cfbold'];
				} else {
					$Tfbold = $Cellstyle[$Sdata[$cellno[$cc]]]['fbold'];
				}
				if ($Css[$cellno[$cc]]['Cfund']){
					$Tfund = $Css[$cellno[$cc]]['Cfund'];
				} else {
					$Tfund = $Cellstyle[$Sdata[$cellno[$cc]]]['fund'];
				}
				if ($Css[$cellno[$cc]]['Cfital']){
					$Tfital = $Css[$cellno[$cc]]['Cfital'];
				} else {
					$Tfital = $Cellstyle[$Sdata[$cellno[$cc]]]['fital'];
				}
				if ($Css[$cellno[$cc]]['Cfscript']){
					$Tfscript = $Css[$cellno[$cc]]['Cfscript'];
				} else {
					$Tfscript = $Cellstyle[$Sdata[$cellno[$cc]]]['fscript'];
				}
				if ($Css[$cellno[$cc]]['Cfstrike']){
					$Tfstrike = $Css[$cellno[$cc]]['Cfstrike'];
				} else {
					$Tfstrike = $Cellstyle[$Sdata[$cellno[$cc]]]['fstrike'];
				}
				$fortext = $Tfname.$Tfsize.$Tfcol.$Tfbold.$Tfund.$Tfital.$Tfscript.$Tfstrike;
				
				// get cell formatting
				if ($Css[$cellno[$cc]]['Cfill']){
					$Tfill = $Css[$cellno[$cc]]['Cfill'];
				} else {
					$Tfill = $Cellstyle[$Sdata[$cellno[$cc]]]['fill'];
				}
				if ($Css[$cellno[$cc]]['Cbleft']){
					$Tbleft = $Css[$cellno[$cc]]['Cbleft'];
				} else {
					$Tbleft = $Cellstyle[$Sdata[$cellno[$cc]]]['bleft'];
				}
				if ($Css[$cellno[$cc]]['Cbtop']){
					$Tbtop = $Css[$cellno[$cc]]['Cbtop'];
				} else {
					$Tbtop = $Cellstyle[$Sdata[$cellno[$cc]]]['btop'];
				}
				if ($Css[$cellno[$cc]]['Cbright']){
					$Tbright = $Css[$cellno[$cc]]['Cbright'];
				} else {
					$Tbright = $Cellstyle[$Sdata[$cellno[$cc]]]['bright'];
				}
				if ($Css[$cellno[$cc]]['Cbbott']){
					$Tbbott = $Css[$cellno[$cc]]['Cbbott'];
				} else {
					$Tbbott = $Cellstyle[$Sdata[$cellno[$cc]]]['bbott'];
				}
				$forcell = $Tbleft.$Tbright.$Tbtop.$Tbbott.$Tfill.$Cellstyle[$Sdata[$cellno[$cc]]]['avert'].$Cellstyle[$Sdata[$cellno[$cc]]]['bdiag']; // get common formatting
				$forcellN[$cellno[$cc]] = $forcell.$Cellstyle[$Sdata[$cellno[$cc]]]['anhor']; //cell formatting for numbers
				$forcellT[$cellno[$cc]] = $forcell.$Cellstyle[$Sdata[$cellno[$cc]]]['athor']; //cell formatting for text
				
				if ($Ddata[$cellno[$cc]] == ''){
					if ($Css[$cellno[$cc]]['dBfill'] == ''){
						$Wdata[$cellno[$cc]] = " style='".$forcellN[$cellno[$cc]]."'><span style='".$fortext."'>".$cell[$cc]."</span></td>"; //get text and formatting for numbers
					} else {
						$Wdata[$cellno[$cc]] = " style='".$forcellN[$cellno[$cc]]."'><span style='".$fortext."'>".$Css[$cellno[$cc]]['dBfill']."<div style='position: relative; bottom: 15px; margin-bottom: -15px;'>".$cell[$cc]."</div></span></td>"; //get text and formatting for numbers						
					}
				} else {
					$Wdata[$cellno[$cc]] = " style='".$forcellT[$cellno[$cc]]."'><span style='".$fortext."'>".$this->shared[$cell[$cc]]."</span></td>"; //get text and formatting for strings (come from 'Shared Strings')
				}
				++$cc;
			}
			
			//Start of finding size of merged cells
			$tt=0;
			while ($tt < $mergeno){ // find width and height of merged cell ranges
				$FC[$tt] = preg_replace("/[^A-Z]/", '', $Mfirst[$tt]);
				$FN[$tt] = preg_replace("/[^0-9]/", '', $Mfirst[$tt]);
				$LC[$tt] = preg_replace("/[^A-Z]/", '', $Mlast[$tt]);
				$LN[$tt] = preg_replace("/[^0-9]/", '', $Mlast[$tt]);
				if (strlen($FC[$tt]) == 1){
					$FCnum[$tt] = ord($FC[$tt]) - 64; // number equiv of chars where A=1
				} else if (strlen($FC[$tt]) == 2){	
					$Vchar1 = ord(substr($FC[$tt],0,1)) - 64; //number equiv of chars where A=1
					$Vchar2 = ord(substr($FC[$tt],1,1)) - 64;
					$FCnum[$tt] = (($Vchar1 * 26) + $Vchar2);
				}
				if (strlen($LC[$tt]) == 1){
					$LCnum[$tt] = ord($LC[$tt]) - 64; // number equiv of chars where A=1
				} else if (strlen($LC[$tt]) == 2){	
					$Vchar3 = ord(substr($LC[$tt],0,1)) - 64;
					$Vchar4 = ord(substr($LC[$tt],1,1)) - 64;
					$LCnum[$tt] = (($Vchar3 * 26) + $Vchar4);
				}
				$Cdiff[$tt] = 1 + ($LCnum[$tt] - $FCnum[$tt]);
				$Ndiff[$tt] = 1 + ((int)$LN[$tt] - (int)$FN[$tt]); //height of merged cells
				++$tt;
			}
			// End of finding size of merged cells.
			
			$defstyle = " border-right:1px solid LightGray; border-bottom:1px solid LightGray;";
			$CRcom = $Cellstyle[$Sdata[$cellno[0]]]['fname'].$Cellstyle[0]['fsize']." background-color: LightGray; text-align:center;"; // Common formatting for all row and column references
			$TLcell = $CRcom."border:1px solid Gray; min-width:20px;"; //cell stying for top left corner
			$TRcell = $CRcom."border-top:1px solid Gray; border-right:1px solid Gray; border-bottom:1px solid Gray;"; //cell stying for column letters above top row
			$LCcell = $CRcom."border-left:1px solid Gray; border-bottom:1px solid Gray; border-right:1px solid Gray; vertical-align:bottom;"; //cell stying for row numbers on left
			$Rmerge = 0;
			$rowcount = $Rfirst - 1;
			while ($rowcount <= $Rlast){
				$colcount = $Cfirst;
				if ($rowcount < $Rfirst){
					if ($this->PR == 'P'){
						$text .= "<td></td>";				
					} else {
						$text .= "<td style='".$TLcell."'>&nbsp;</td>"; //Top left corner
					}
				} else {
					if ($this->PR == 'P'){
						$text .= "<td style='border-right: 1px solid LightGray'></td>";				
					} else {
						if ($this->SW == 'O' AND isset($Rhight[$rowcount])){
							$text .= "<tr><td style='".$LCcell." height:".$Rhight[$rowcount]."px;'>".$rowcount."</td>"; //row numbers with defined height row
						} else {
							$text .= "<tr><td style='".$LCcell."'>".$rowcount."</td>"; //row numbers with no row defined height
						}
					}
				}
				while ($colcount <= $Clast){
					$Acolcount = $this->numtochars($colcount);
					$a = 0;
					$Mfound = '';
					$Mfirst = '';
					while ($a < $mergeno){
						if ($colcount >= $FCnum[$a] AND $colcount <= $LCnum[$a] AND $rowcount >= $FN[$a] AND $rowcount <= $LN[$a]){
							if ($colcount == $FCnum[$a] AND $rowcount == $FN[$a]){
								$Mfirst = 'Y';
								$Mno = $a;
								if ($Ixr[$a] == 'Y')
									$PP = $a - $mmc;
							} else {
								$Mfound = 'Y';
								$a = $mergeno;
							}
						}
						++$a;
					}
					if ($Mfound == ''){
						$tcell = $Acolcount.$rowcount;
						if ($Mfirst == 'Y'){
							if ($Cdiff[$Mno] > 1){
								$mtext = " colspan ='".$Cdiff[$Mno]."'";
							}
							if ($Ndiff[$Mno] > 1){
								$mtext .= " rowspan ='".$Ndiff[$Mno]."'";
							}
						} else {
							$mtext = '';
						}
						if ($rowcount < $Rfirst){
							if ($this->SW == 'O'){
								if (!$Cwidth[$colcount]){
									$Cwidth[$colcount] = $this->Defwidth;
								}
								if ($this->PR == 'P'){
									$text .= "<td style='min-width:".$Cwidth[$colcount]."px; max-width:".$Cwidth[$colcount]."px; border-bottom: 1px solid LightGray'>&nbsp;</td>";
								} else {
									$text .= "<td style='".$TRcell." min-width:".$Cwidth[$colcount]."px; max-width:".$Cwidth[$colcount]."px;'>".$Acolcount."</td>"; // column references when column width is defined
								}
							} else {
								if ($this->PR == 'P'){
									$text .= "<td style='border-bottom: 1px solid LightGray'>&nbsp;</td>";
								} else {
									$text .= "<td style='".$TRcell."'>&nbsp;".$Acolcount."&nbsp;</td>"; // column references when column width is not defined
								}
							}
						} else if ($PP <> ''){
							if (!$Sdata[$tcell]){
								$text .= "<td".$mtext." style='".$defstyle."'><image src='".$Iname[$PP]."'  style='width:".$Imxs[$PP]."px; height:".$Imys[$PP]."px; padding:5px 5px 5px 5px;' />"; //for images when the cell(s) have no defined border
							} else {
								$tform = $Cellstyle[$Sdata[$tcell]]['bleft'].$Cellstyle[$Sdata[$tcell]]['bright'].$Cellstyle[$Sdata[$tcell]]['btop'].$Cellstyle[$Sdata[$tcell]]['bbott'];
								$text .= "<td".$mtext." style='".$$tform."'><image src='".$Iname[$PP]."'  style='width:".$Imxs[$PP]."px; height:".$Imys[$PP]."px; padding:5px 5px 5px 5px;' />"; //for images when cells have a defined border
								$tform = '';
							}
							$PP = '';
						//checks whether in the cell is blank or not
						} else if (!isset($Wdata[$tcell])){ 
							if (!$Sdata[$tcell]){
								$text .= "<td".$mtext." style='".$defstyle."'>&nbsp;</td>"; //blank cells with no formatting
							} else {
								$text .= "<td".$mtext." style='".$forcellT[$Sdata[$tcell]]."'>&nbsp;</td>"; //for blank cells with formatting
							}
						} else {
							$text .= "<td".$mtext.$Wdata[$tcell]; //for cells with text/numbers
						}
					}
					++$colcount;
				}
				++$rowcount;
				$text .= "</tr> \n";
			}
			$text .= "</table>";
		}
		$Sinfo['text'] = $text;
		return $Sinfo;
	}
	

			





	/**
	 * READS THE GIVEN XLSX FILE INTO HTML FORMAT
	 *  
	 * @param String $filename - The XLSX file name
	 * @return String - With HTML code of the XLSX file
	 */
	public function readDocument($filename,$options)
	{
		$tdate=date_create("1899-12-30");
		date_add($tdate,date_interval_create_from_date_string("45529 days"));
		$tdate->add(new \DateInterval('P5000D'));
		$tdate->add(new \DateInterval('PT5000S'));
		
		if (!file_exists($filename)){
			exit("Cannot find file : ".$filename." ");
		}
		$this->file = $filename; // makes the filename available throughout the class
		$Optlen = strlen($options);
		if ($Optlen == 0){
			$SS = 'A';
		}
		if ($Optlen > 0){
			$SS = substr($options,0,1); // An 'A' displays all spreadsheets. A number selects that particular one.
		} 
		if ($Optlen > 1){
			$this->SW = substr($options,1,1); // left blank or an 'A' the column widths are undefined (auto). An 'O' (default) causes them to try and replicate the original Excel column widths
		} else {
			$this->SW = 'O';
		}
		if ($Optlen > 2){
			$this->PR = substr($options,2,1); // if the 3rd option character is left blank or is a 'P', the display will be like Excel printout mode with headers/footers and no row or column references. A 'S' will put it into standard spreadsheet mode with row and column references and sheet name and no headers/footers.
		} else {
			$this->PR = 'P';
		}
	
	
		$this->readZipPart(); // Makes the document and relationships file available throughout the class
		$this->stylecount = 0;

		//look at each sheet in turn
		$text = "<div style=' text-align:center;'>";
		if ($SS == 'A'){
			$sc = 1;
			while ($sc <= $this->Sheetnum){
				$Sinfo = $this->checkSheet($sc);
				$Stext = $Sinfo['text'];
				if (strlen($Stext) > 0){ //check to see if the sheet is populated
					if (strlen($Sinfo['head']) > 40 AND $this->PR == 'P'){
						$text .= $Sinfo['head']; //adds the header is it exists
					}
					if ($this->PR == 'S'){
						$text .= "<h2>Sheet name - '".$this->Sheetname[$sc-1]."'</h2>";
					}
					$text .= $Stext;
					if (strlen($Sinfo['foot']) > 40 AND $this->PR == 'P'){
						$text .= $Sinfo['foot']; //adds the footer is it exists
					}
					$text .= "<br>&nbsp;<br>";
				}
				++$sc;
			}
		} else {
			$Sinfo = $this->checkSheet($SS);
			$Stext .= $Sinfo['text'];
			if (strlen($Stext) > 0){ //check to see if the sheet is populated
				if (strlen($Sinfo['head']) > 40 AND $this->PR == 'P'){
					$text .= $Sinfo['head']; //adds the header is it exists
				}						
				if ($this->PR == 'S'){
					$text .= "<h2>Sheet name - '".$this->Sheetname[$SS-1]."'</h2>";
				}
				$text .= $Stext;
				if (strlen($Sinfo['foot']) > 40 AND $this->PR == 'P'){
					$text .= $Sinfo['foot']; //adds the footer is it exists
				}
				$text .= "<br>&nbsp;<br>";
			} else {
				$text .= "<h2>Sheet ".$SS." of this Excel spreadsheet does not exist.</h2>";
			}
		}
		
		$Stext .= "</div>";
		return mb_convert_encoding($text, $this->encoding); // Output the generated HTML text of the DOCX document
	}
}
	
