<?php

namespace avadim\FastExcelReader;

/**
 * Class Excel
 *
 * @package avadim\FastExcelReader
 */
class Excel
{
    public const EXCEL_2007_MAX_ROW = 1048576;
    public const EXCEL_2007_MAX_COL = 16384;

    public const KEYS_ORIGINAL = 0;
    public const KEYS_FIRST_ROW = 1;
    public const KEYS_ROW_ZERO_BASED = 2;
    public const KEYS_COL_ZERO_BASED = 4;
    public const KEYS_ZERO_BASED = 6;
    public const KEYS_ROW_ONE_BASED = 8;
    public const KEYS_COL_ONE_BASED = 16;
    public const KEYS_ONE_BASED = 24;
    public const KEYS_RELATIVE = 32;
    public const KEYS_SWAP = 64;

    protected $file;

    /** @var Reader */
    protected $xmlReader;

    protected $fileList = [];

    protected $relations = [];

    protected $sharedStrings = [];

    protected $styles = [];

    /** @var Sheet[] */
    protected $sheets = [];

    protected $defaultSheetId;

    protected $dateFormat = null;

    protected $date1904 = false;
    protected $timezone;


    /**
     * Excel constructor
     *
     * @param string|null $file
     */
    public function __construct($file = null)
    {
        if ($file) {
            $this->file = $file;
            $this->_prepare($file);
        }
        $this->timezone = date_default_timezone_get();
    }

    /**
     * @param string $file
     */
    protected function _prepare($file)
    {
        $this->xmlReader = new Reader($file);
        $this->fileList = $this->xmlReader->fileList();
        foreach ($this->fileList as $fileName) {
            if (strpos($fileName, 'xl/drawings/drawing') === 0) {
                $this->relations['drawings'][] = $fileName;
            }
            elseif (strpos($fileName, 'xl/media/') === 0) {
                $this->relations['media'][] = $fileName;
            }
            elseif (strpos($fileName, 'xl/theme/') === 0) {
                $this->relations['theme'][] = $fileName;
            }
        }

        $innerFile = 'xl/_rels/workbook.xml.rels';
        $this->xmlReader->openZip($innerFile);
        while ($this->xmlReader->read()) {
            if ($this->xmlReader->nodeType === \XMLReader::ELEMENT && $this->xmlReader->name === 'Relationship') {
                $type = basename($this->xmlReader->getAttribute('Type'));
                if ($type) {
                    $this->relations[$type][$this->xmlReader->getAttribute('Id')] = 'xl/' . $this->xmlReader->getAttribute('Target');
                }
            }
        }
        $this->xmlReader->close();

        if (isset($this->relations['worksheet'])) {
            $this->_loadSheets();
        }
        if (isset($this->relations['sharedStrings'])) {
            $this->_loadSharedStrings(reset($this->relations['sharedStrings']));
        }
        if (isset($this->relations['styles'])) {
            $this->_loadStyles(reset($this->relations['styles']));
        }

        if ($this->sheets) {
            // set current sheet
            $this->selectFirstSheet();
        }
    }

    /**
     * @param string|null $innerFile
     */
    protected function _loadSheets($innerFile = null)
    {
        if (!$innerFile) {
            $innerFile = 'xl/workbook.xml';
        }
        $this->xmlReader->openZip($innerFile);
        $sheetCnt = count($this->relations['worksheet']);
        while ($this->xmlReader->read()) {
            if ($this->xmlReader->nodeType === \XMLReader::ELEMENT && $this->xmlReader->name === 'workbookPr') {
                $date1904 = (string)$this->xmlReader->getAttribute('date1904');
                if ($date1904 === '1' || $date1904 === 'true') {
                    $this->date1904 = true;
                }
            }
            if ($this->xmlReader->nodeType === \XMLReader::ELEMENT && $this->xmlReader->name === 'sheet') {
                $rId = $this->xmlReader->getAttribute('r:id');
                $sheetId = $this->xmlReader->getAttribute('sheetId');
                $path = $this->relations['worksheet'][$rId];
                if ($path) {
                    $sheetName = $this->xmlReader->getAttribute('name');
                    $this->sheets[$sheetId] = static::createSheet($sheetName, $sheetId, $this->file, $this->relations['worksheet'][$rId]);
                    $this->sheets[$sheetId]->excel = $this;
                }
                /*
                if (--$sheetCnt < 1) {
                    break;
                }
                */
            }
        }
        $this->xmlReader->close();
    }

    /**
     * @param string|null $innerFile
     */
    protected function _loadSharedStrings($innerFile = null)
    {
        if (!$innerFile) {
            $innerFile = 'xl/sharedStrings.xml';
        }
        $this->xmlReader->openZip($innerFile);
        while ($this->xmlReader->read()) {
            if ($this->xmlReader->nodeType === \XMLReader::ELEMENT && $this->xmlReader->name === 'si' && $node = $this->xmlReader->expand()) {
                $this->sharedStrings[] = $node->textContent;
            }
        }
        $this->xmlReader->close();
    }

    /**
     * @param string|null $innerFile
     */
    protected function _loadStyles($innerFile = null)
    {
        if (!$innerFile) {
            $innerFile = 'xl/styles.xml';
        }
        $this->xmlReader->openZip($innerFile);
        $styleType = '';
        while ($this->xmlReader->read()) {
            if ($this->xmlReader->nodeType === \XMLReader::ELEMENT) {
                if ($this->xmlReader->name === 'cellStyleXfs' || $this->xmlReader->name === 'cellXfs') {
                    $styleType = $this->xmlReader->name;
                    continue;
                }
                if ($this->xmlReader->name === 'numFmt') {
                    $numFmtId = (int)$this->xmlReader->getAttribute('numFmtId');
                    $formatCode = $this->xmlReader->getAttribute('formatCode');
                    $numFmts[$numFmtId] = $formatCode;
                }
                elseif ($this->xmlReader->name === 'xf') {
                    $numFmtId = (int)$this->xmlReader->getAttribute('numFmtId');
                    if (isset($numFmts[$numFmtId])) {
                        $format = $numFmts[$numFmtId];
                        if (strpos($format, 'M') !== false || strpos($format, 'm') !== false) {
                            $this->styles[$styleType][] = ['format' => $numFmts[$numFmtId], 'formatType' => 'd'];
                        }
                        else {
                            $this->styles[$styleType][] = ['format' => $numFmts[$numFmtId]];
                        }
                    }
                    elseif (($numFmtId >= 14 && $numFmtId <= 22) || ($numFmtId >= 45 && $numFmtId <= 47)) {
                            $this->styles[$styleType][] = ['formatType' => 'd'];
                    }
                    else {
                        $this->styles[$styleType][] = null;
                    }
                }
            }
        }
        $this->xmlReader->close();
    }

    protected function _loadStyleNumFmts($root, $tagName)
    {
        foreach ($root->childNodes as $child) {
            if ($child->nodeType == XML_ELEMENT_NODE) {
                $numFmtId = $child->getAttribute('numFmtId');
                $formatCode = $child->getAttribute('formatCode');
                if ($numFmtId !== '' && $formatCode !== '') {
                    $node = [
                        'format-num-id' => (int)$numFmtId,
                        'format-pattern' => $formatCode,
                    ];
                    $this->styles['_'][$tagName][$node['format-num-id']] = $node;
                }
            }
        }
    }

    protected function _loadStyleFonts($root, $tagName)
    {
        foreach ($root->childNodes as $font) {
            $node = [];
            foreach ($font->childNodes as $fontStyle) {
                if ($fontStyle->nodeName === 'b') {
                    $node['font-style-bold'] = 1;
                }
                elseif ($fontStyle->nodeName === 'u') {
                    $node['font-style-underline'] = ($fontStyle->getAttribute('formatCode') === 'double' ? 2 : 1);
                }
                elseif ($fontStyle->nodeName === 'i') {
                    $node['font-style-italic'] = 1;
                }
                elseif ($fontStyle->nodeName === 'strike') {
                    $node['font-style-strike'] = 1;
                }
                elseif (($fontStyle->nodeType == XML_ELEMENT_NODE && $v = $fontStyle->getAttribute('val')) !== '') {
                    if ($fontStyle->nodeName === 'sz') {
                        $name = 'font-size';
                    }
                    else {
                        $name = 'font-' . $fontStyle->nodeName;
                    }
                    $node[$name] = $v;
                }
            }
            $this->styles['_'][$tagName][] = $node;
        }
    }

    protected function _loadStyleFills($root, $tagName)
    {
        foreach ($root->childNodes as $fill) {
            $node = [];
            foreach ($fill->childNodes as $patternFill) {
                if (($patternFill->nodeType == XML_ELEMENT_NODE && $v = $patternFill->getAttribute('patternType')) !== '') {
                    $node['fill-pattern'] = $v;
                }
                foreach ($patternFill->childNodes as $child) {
                    if ($child->nodeName === 'fgColor') {
                        $node['fill-color'] = '#' . substr($child->getAttribute('rgb'), 2);
                    }
                }
            }
            $this->styles['_'][$tagName][] = $node;
        }
    }

    protected function _loadStyleBorders($root, $tagName)
    {
        foreach ($root->childNodes as $border) {
            $node = [];
            foreach ($border->childNodes as $side) {
                if (($side->nodeType == XML_ELEMENT_NODE && $v = $side->getAttribute('style')) !== '') {
                    $node['border-' . $side->nodeName . '-style'] = $v;
                }
                else {
                    $node['border-' . $side->nodeName . '-style'] = null;
                }
                foreach ($side->childNodes as $child) {
                    if ($child->nodeName === 'color') {
                        $node['border-' . $side->nodeName . '-color'] = '#' . substr($child->getAttribute('rgb'), 2);
                    }
                }
            }
            $this->styles['_'][$tagName][] = $node;
        }
    }

    protected function _loadStyleCellXfs($root, $tagName)
    {
        $attributes = ['numFmtId', 'fontId', 'fillId', 'borderId', 'xfId'];
        foreach ($root->childNodes as $xf) {
            $node = [];
            foreach ($attributes as $attribute) {
                if (($xf->nodeType == XML_ELEMENT_NODE && $v = $xf->getAttribute($attribute)) !== '') {
                    if (substr($attribute, -2) === 'Id') {
                        $node[$attribute] = (int)$v;
                    }
                    else {
                        $node[$attribute] = $v;
                    }
                }
            }
            foreach ($xf->childNodes as $child) {
                if ($child->nodeName === 'alignment') {
                    if ($child->nodeType == XML_ELEMENT_NODE && $v = $child->getAttribute('horizontal')) {
                        $node['format']['format-align-horizontal'] = $v;
                    }
                    if ($child->nodeType == XML_ELEMENT_NODE && $v = $child->getAttribute('vertical')) {
                        $node['format']['format-align-vertical'] = $v;
                    }
                    if ($child->nodeType == XML_ELEMENT_NODE && ($v = $child->getAttribute('wrapText'))/* && ($v === '1')*/) {
                        $node['format']['format-wrap-text'] = (intval($v)) ? true : false;
                    }
                }
            }
            $this->styles['_'][$tagName][] = $node;
        }
    }

    /**
     * @param string|null $innerFile
     */
    protected function _loadCompleteStyles(string $innerFile = null)
    {
        if (!$innerFile) {
            $innerFile = 'xl/styles.xml';
        }
        $this->xmlReader->openZip($innerFile);
        while ($this->xmlReader->read()) {
            if ($this->xmlReader->nodeType === \XMLReader::ELEMENT) {
                switch ($this->xmlReader->name) {
                    case 'numFmts':
                        $this->_loadStyleNumFmts($this->xmlReader->expand(), 'numFmts');
                        break;
                    case 'fonts':
                        $this->_loadStyleFonts($this->xmlReader->expand(), 'fonts');
                        break;
                    case 'fills':
                        $this->_loadStyleFills($this->xmlReader->expand(), 'fills');
                        break;
                    case 'borders':
                        $this->_loadStyleBorders($this->xmlReader->expand(), 'borders');
                        break;
                    case 'cellStyleXfs':
                        $this->_loadStyleCellXfs($this->xmlReader->expand(), 'cellStyleXfs');
                        break;
                    case 'cellXfs':
                        $this->_loadStyleCellXfs($this->xmlReader->expand(), 'cellXfs');
                        break;
                    default:
                        //
                }
            }
        }
        $this->xmlReader->close();
    }

    /**
     * Open XLSX file
     *
     * @param string $file
     *
     * @return Excel
     */
    public static function open(string $file)
    {
        return new self($file);
    }

    /**
     * @param string $sheetName
     * @param $sheetId
     * @param $file
     * @param $path
     *
     * @return Sheet
     */
    public static function createSheet($sheetName, $sheetId, $file, $path)
    {
        return new Sheet($sheetName, $sheetId, $file, $path);
    }

    /**
     * Converts an alphabetic column index to a numeric
     *
     * @param string $colLetter
     *
     * @return int
     */
    public static function colNum(string $colLetter)
    {
        static $colNumbers = [];

        if (isset($colNumbers[$colLetter])) {
            return $colNumbers[$colLetter];
        }
        // Strip cell reference down to just letters
        $letters = preg_replace('/[^A-Z]/', '', strtoupper($colLetter));

        if (strlen($letters) >= 3 && $letters > 'XFD') {
            return self::EXCEL_2007_MAX_COL;
        }
        // Iterate through each letter, starting at the back to increment the value
        for ($index = 0, $i = 0; $letters !== ''; $letters = substr($letters, 0, -1), $i++) {
            $index += (ord(substr($letters, -1)) - 64) * (26 ** $i);
        }

        $colNumbers[$colLetter] = ($index <= self::EXCEL_2007_MAX_COL) ? (int)$index: self::EXCEL_2007_MAX_COL;

        return $colNumbers[$colLetter];
    }

    /**
     * Convert column number to letter
     *
     * @param int $colNumber ONE based
     *
     * @return string
     */
    public static function colLetter(int $colNumber)
    {
        static $colLetters = ['',
            'A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z',
            'AA', 'AB', 'AC', 'AD', 'AE', 'AF', 'AG', 'AH', 'AI', 'AJ', 'AK', 'AL', 'AM', 'AN', 'AO', 'AP', 'AQ', 'AR', 'AS', 'AT', 'AU', 'AV', 'AW', 'AX', 'AY', 'AZ',
        ];

        if (isset($colLetters[$colNumber])) {
            return $colLetters[$colNumber];
        }

        if ($colNumber > 0 && $colNumber <= self::EXCEL_2007_MAX_COL) {
            $num = $colNumber - 1;
            for ($letter = ''; $num >= 0; $num = (int)($num / 26) - 1) {
                $letter = chr($num % 26 + 0x41) . $letter;
            }
            $colLetters[$colNumber] = $letter;

            return $letter;
        }

        return '';
    }

    /**
     * @param $excelDateTime
     *
     * @return int
     */
    public function timestamp($excelDateTime)
    {
        if (is_numeric($excelDateTime)) {
            $d = floor($excelDateTime);
            $t = $excelDateTime - $d;
            if ($this->date1904) {
                $d += 1462; // days since 1904
            }

            // Adjust for Excel erroneously treating 1900 as a leap year.
            if ($d <= 59) {
                $d++;
            }
            $t = (abs($d) > 0) ? ($d - 25569) * 86400 + round($t * 86400) : round($t * 86400);
        }
        else {
            if ($this->timezone !== 'UTC') {
                date_default_timezone_set('UTC');
            }
            $t = strtotime($excelDateTime);
            if ($this->timezone !== 'UTC') {
                date_default_timezone_set($this->timezone);
            }
        }

        return (int)$t;
    }

    /**
     * @param string $dateFormat
     *
     * @return $this
     */
    public function setDateFormat(string $dateFormat)
    {
        $this->dateFormat = $dateFormat;

        return $this;
    }

    /**
     * @return string|null
     */
    public function getDateFormat()
    {
        return $this->dateFormat;
    }

    public function formatDate($value, $format = null)
    {
        if ($format || $this->dateFormat) {
            return gmdate($format ?: $this->dateFormat, $value);
        }

        return $value;
    }

    /**
     * Returns style array by style Idx
     *
     * @param $styleIdx
     *
     * @return array
     */
    public function styleByIdx($styleIdx)
    {
        return $this->styles['cellXfs'][$styleIdx] ?? [];
    }

    /**
     * Returns string array by index
     *
     * @param $stringId
     *
     * @return string|null
     */
    public function sharedString($stringId)
    {
        return $this->sharedStrings[$stringId] ?? null;
    }

    /**
     * Returns names array of all sheets
     *
     * @return array
     */
    public function getSheetNames()
    {
        $result = [];
        foreach ($this->sheets as $sheetId => $sheet) {
            $result[$sheetId] = $sheet->name();
        }
        return $result;
    }

    /**
     * Returns current or specified sheet
     *
     * @param string|null $name
     *
     * @return Sheet|null
     */
    public function sheet($name = null)
    {
        $resultId = null;
        if (!$name) {
            $resultId = $this->defaultSheetId;
        }
        else {
            foreach ($this->sheets as $sheetId => $sheet) {
                if ($sheet->isName($name)) {
                    $resultId = $sheetId;
                    break;
                }
            }
        }
        if ($resultId && isset($this->sheets[$resultId])) {
            return $this->sheets[$resultId];
        }

        return null;
    }

    /**
     * Returns a sheet by name
     *
     * @param string $name
     * @param string|null $areaRange
     * @param bool|null $firstRowKeys
     *
     * @return Sheet
     */
    public function getSheet(string $name, string $areaRange = null, $firstRowKeys = false)
    {
        foreach ($this->sheets as $sheet) {
            if ($sheet->isName($name)) {
                if ($areaRange) {
                    $sheet->setReadArea($areaRange, $firstRowKeys);
                }

                return $sheet;
            }
        }
        throw new Exception('Sheet name "' . $name . '" not found');
    }

    /**
     * Returns a sheet by ID
     *
     * @param int $sheetId
     * @param string|null $areaRange
     * @param bool|null $firstRowKeys
     *
     * @return Sheet
     */
    public function getSheetById(int $sheetId, string $areaRange = null, $firstRowKeys = false)
    {
        if (!isset($this->sheets[$sheetId])) {
            throw new Exception('Sheet ID "' . $sheetId . '" not found');
        }
        if ($areaRange) {
            $this->sheets[$sheetId]->setReadArea($areaRange, $firstRowKeys);
        }

        return $this->sheets[$sheetId];
    }

    /**
     * Returns the first sheet as default
     *
     * @param string|null $areaRange
     * @param bool|null $firstRowKeys
     *
     * @return Sheet
     */
    public function getFirstSheet(string $areaRange = null, $firstRowKeys = false)
    {
        $sheetId = array_key_first($this->sheets);
        $sheet = $this->sheets[$sheetId];
        if ($areaRange) {
            $sheet->setReadArea($areaRange, $firstRowKeys);
        }

        return $sheet;
    }

    /**
     * Selects default sheet by name
     *
     * @param string $name
     * @param string|null $areaRange
     * @param bool|null $firstRowKeys
     *
     * @return Sheet
     */
    public function selectSheet(string $name, string $areaRange = null, $firstRowKeys = false)
    {
        $sheet = $this->getSheet($name, $areaRange, $firstRowKeys);
        $this->defaultSheetId = $sheet->id();

        return $sheet;
    }

    /**
     * Selects default sheet by ID
     *
     * @param int $sheetId
     * @param string|null $areaRange
     * @param bool|null $firstRowKeys
     *
     * @return Sheet
     */
    public function selectSheetById(int $sheetId, string $areaRange = null, $firstRowKeys = false)
    {
        $sheet = $this->getSheetById($sheetId, $areaRange, $firstRowKeys);
        $this->defaultSheetId = $sheet->id();

        return $sheet;
    }

    /**
     * Selects the first sheet as default
     *
     * @param string|null $areaRange
     * @param bool|null $firstRowKeys
     *
     * @return Sheet
     */
    public function selectFirstSheet(string $areaRange = null, $firstRowKeys = false)
    {
        $sheet = $this->getFirstSheet($areaRange, $firstRowKeys);
        $this->defaultSheetId = $sheet->id();

        return $sheet;
    }

    /**
     * @param string $areaRange
     * @param bool|null $firstRowKeys
     *
     * @return Sheet
     */
    public function setReadArea(string $areaRange, $firstRowKeys = false)
    {
        return $this->sheets[$this->defaultSheetId]->setReadArea($areaRange, $firstRowKeys);
    }

    /**
     * Reads cell values and passes them to a callback function
     *
     * @param callback $callback
     * @param int|null $resultMode
     */
    public function readCallback(callable $callback, int $resultMode = null, $styleIdxInclude = null)
    {
        $this->sheets[$this->defaultSheetId]->readCallback($callback, $resultMode);
    }

    /**
     * Returns cell values as a two-dimensional array from default sheet [row][col]
     *
     *  readRows()
     *  readRows(true)
     *  readRows(false, Excel::KEYS_ZERO_BASED)
     *  readRows(Excel::KEYS_ZERO_BASED | Excel::KEYS_RELATIVE)
     *
     * @param array|bool|int|null $columnKeys
     * @param int|null $resultMode
     * @param bool|null $styleIdxInclude
     *
     * @return array
     */
    public function readRows($columnKeys = [], int $resultMode = null, $styleIdxInclude = null)
    {
        return $this->sheets[$this->defaultSheetId]->readRows($columnKeys, $resultMode, $styleIdxInclude);
    }

    /**
     * Returns cell values and styles as a two-dimensional array from default sheet [row][col]
     *
     * @param array|bool|int|null $columnKeys
     * @param int|null $resultMode
     *
     * @return array
     */
    public function readRowsWithStyles($columnKeys = [], int $resultMode = null)
    {
        return $this->sheets[$this->defaultSheetId]->readRowsWithStyles($columnKeys, $resultMode);
    }

    /**
     * Returns cell values as a two-dimensional array from default sheet [col][row]
     *
     * @param array|bool|int|null $columnKeys
     * @param int|null $resultMode
     *
     * @return array
     */
    public function readColumns($columnKeys = null, int $resultMode = null)
    {
        return $this->sheets[$this->defaultSheetId]->readColumns($columnKeys, $resultMode);
    }

    /**
     * Returns cell values and styles as a two-dimensional array from default sheet [col][row]
     *
     * @param array|bool|int|null $columnKeys
     * @param int|null $resultMode
     *
     * @return array
     */
    public function readColumnsWithStyles($columnKeys = null, int $resultMode = null)
    {
        return $this->sheets[$this->defaultSheetId]->readColumnsWithStyles($columnKeys, $resultMode);
    }

    /**
     * Returns the values of all cells as array
     *
     * @return array
     */
    public function readCells()
    {
        return $this->sheets[$this->defaultSheetId]->readCells();
    }

    /**
     * Returns the values and styles of all cells as array
     *
     * @return array
     */
    public function readCellsWithStyles()
    {
        return $this->sheets[$this->defaultSheetId]->readCellsWithStyles();
    }

    public function innerFileList()
    {
        return $this->fileList;
    }

    /**
     * Returns TRUE if the workbook contains an any draw objects (not images only)
     *
     * @return bool
     */
    public function hasDrawings()
    {
        return !empty($this->relations['drawings']);
    }

    /**
     * Returns TRUE if any sheet contains an image object
     *
     * @return bool
     */
    public function hasImages()
    {
        if ($this->hasDrawings()) {
            foreach ($this->sheets as $sheet) {
                if ($sheet->countImages()) {
                    return true;
                }
            }
        }

        return false;
    }

    /**
     * @return int
     */
    public function countImages()
    {
        $result = 0;
        if ($this->hasDrawings()) {
            foreach ($this->sheets as $sheet) {
                $result += $sheet->countImages();
            }
        }

        return $result;
    }

    /**
     * @return array
     */
    public function getImageList()
    {
        $result = [];
        if ($this->hasDrawings()) {
            foreach ($this->sheets as $sheet) {
                $result[$sheet->name()] = $sheet->getImageList();
            }
        }

        return $result;
    }

    public function readStyles()
    {
        if (!isset($this->styles['_'])) {
            $this->styles['_'] = [];
            $this->_loadCompleteStyles();
        }

        return $this->styles['_'];
    }

    public function readSizes()
    {
        //$this->getReader();
        while ($this->xmlReader->read()) {
            if ($this->xmlReader->nodeType === \xmlReader::ELEMENT && $this->xmlReader->name === 'sheetFormatPr') {
                $this->sizes['defaultColWidth'] = floatval($this->xmlReader->getAttribute('defaultColWidth'));
                $this->sizes['defaultRowHeight'] = floatval($this->xmlReader->getAttribute('defaultRowHeight'));
            }
            if ($this->xmlReader->nodeType === \xmlReader::ELEMENT && $this->xmlReader->name === 'col') {
                $minCol = intval($this->xmlReader->getAttribute('min'));
                $maxCol = intval($this->xmlReader->getAttribute('max'));
                if ($minCol == $maxCol)
                    $this->sizes['cols'][$minCol] = floatval($this->xmlReader->getAttribute('width'));
                else if ($minCol < $maxCol) {
                    for ($i = $minCol; $i <= $maxCol; $i++) {
                        $this->sizes['cols'][$i] = floatval($this->xmlReader->getAttribute('width'));
                    }
                }
            }
            if ($this->xmlReader->nodeType === \xmlReader::ELEMENT && $this->xmlReader->name === 'row') {
                $index = intval($this->xmlReader->getAttribute('r'));
                $this->sizes['rows'][$index] = floatval($this->xmlReader->getAttribute('ht'));
            }
            if ($this->xmlReader->nodeType === \xmlReader::ELEMENT && $this->xmlReader->name === 'mergeCell') {
                $this->sizes['mergeCell'][] = $this->xmlReader->getAttribute('ref');
            }
        }
        return $this->sizes;
    }

    /**
     * @param int $styleIdx
     * @param bool|null $flat
     *
     * @return array
     */
    public function getCompleteStyleByIdx(int $styleIdx, $flat = false)
    {
        static $completedStyles = [];

        if (![$this->file]) {
            return [];
        }

        if (!isset($completedStyles[$this->file][$styleIdx])) {
            if ($styleIdx !== 0) {
                $result = $this->getCompleteStyleByIdx(0);
            }
            else {
                $result = [];
            }
            $styles = $this->readStyles();
            if (isset($styles['cellXfs'][$styleIdx])) {
                $result = array_replace_recursive($result, $styles['cellXfs'][$styleIdx]);
            }

            if (isset($result['xfId']) && isset($styles['cellStyleXfs'][$result['xfId']])) {
                if ($styleIdx === 0 || ($styleIdx > 0 && $result['xfId'])) {
                    $result = array_replace_recursive($result, $styles['cellStyleXfs'][$result['xfId']]);
                }
                unset($result['xfId']);
            }

            if (isset($result['numFmtId']) && isset($styles['numFmts'][$result['numFmtId']])) {
                if (isset($result['format'])) {
                    $result['format'] = array_replace_recursive($result['format'], $styles['numFmts'][$result['numFmtId']]);
                }
                else {
                    $result['format'] = $styles['numFmts'][$result['numFmtId']];
                }
                unset($result['numFmtId']);
            }

            if (isset($result['fontId']) && isset($styles['fonts'][$result['fontId']])) {
                if (isset($result['font'])) {
                    $result['font'] = array_replace_recursive($result['font'], $styles['fonts'][$result['fontId']]);
                }
                else {
                    $result['font'] = $styles['fonts'][$result['fontId']];
                }
                unset($result['fontId']);
            }

            if (isset($result['fillId']) && isset($styles['fills'][$result['fillId']])) {
                if (isset($result['fill'])) {
                    $result['fill'] = array_replace_recursive($result['fill'], $styles['fills'][$result['fillId']]);
                }
                else {
                    $result['fill'] = $styles['fills'][$result['fillId']];
                }
                unset($result['fillId']);
            }

            if (isset($result['borderId']) && isset($styles['borders'][$result['borderId']])) {
                if (isset($result['border'])) {
                    $result['border'] = array_replace_recursive($result['border'], $styles['borders'][$result['borderId']]);
                }
                else {
                    $result['border'] = $styles['borders'][$result['borderId']];
                }
                unset($result['borderId']);
            }

            $completedStyles[$this->file][$styleIdx] = $result;
        }
        else {
            $result = $completedStyles[$this->file][$styleIdx];
        }

        if ($flat && $result) {
            $result = array_merge(...array_values($result));
        }

        return $result;
    }
}

// EOF