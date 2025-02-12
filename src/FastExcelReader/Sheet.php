<?php

namespace avadim\FastExcelReader;

class Sheet
{
    public Excel $excel;

    protected string $zipFilename;

    protected string $sheetId;

    protected string $name;

    protected string $path;

    protected ?string $dimension = null;

    protected array $area = [];

    protected array $props = [];

    /** @var Reader */
    protected Reader $xmlReader;


    public function __construct($sheetName, $sheetId, $file, $path)
    {
        $this->name = $sheetName;
        $this->sheetId = $sheetId;
        $this->zipFilename = $file;
        $this->path = $path;

        $this->area = [
            'row_min' => 1,
            'col_min' => 1,
            'row_max' => Excel::EXCEL_2007_MAX_ROW,
            'col_max' => Excel::EXCEL_2007_MAX_COL,
            'first_row' => false,
        ];
    }

    /**
     * @param $cell
     * @param $styleIdx
     *
     * @return mixed
     */
    protected function _cellValue($cell, &$styleIdx = null)
    {
        // Determine data type
        $dataType = (string)$cell->getAttribute('t');
        $cellValue = null;
        foreach($cell->childNodes as $node) {
            if ($node->nodeName === 'v') {
                $cellValue = $node->nodeValue;
                break;
            }
        }

        $format = null;
        // Value is a shared string
        if ($dataType === 's' && is_numeric($cellValue) && ($str = $this->excel->sharedString((int)$cellValue))) {
            $cellValue = $str;
        }
        $styleIdx = (int)$cell->getAttribute('s');
        if ( $dataType === '' || $dataType === 'n'  || $dataType === 's' ) { // number or data as string
            if ($styleIdx > 0 && ($style = $this->excel->styleByIdx($styleIdx))) {
                $format = $style['format'] ?? null;
                if (isset($style['formatType'])) {
                    $dataType = $style['formatType'];
                }
            }
        }

        $value = '';

        switch ( $dataType ) {
            case 'b':
                // Value is boolean
                $value = (bool)$cellValue;
                break;

            case 'inlineStr':
                // Value is rich text inline
                $value = $cell->textContent;
                break;

            case 'e':
                // Value is an error message
                $value = (string)$cellValue;
                break;

            case 'd':
                // Value is a date and non-empty
                if (!empty($cellValue)) {
                    $value = $this->excel->formatDate($this->excel->timestamp($cellValue));
                }
                break;

            default:
                // Value is a string
                $value = (string) $cellValue;

                // Check for numeric values
                if (is_numeric($value)) {
                    /** @noinspection TypeUnsafeComparisonInspection */
                    if ($value == (int)$value) {
                        $value = (int)$value;
                    }
                    /** @noinspection TypeUnsafeComparisonInspection */
                    elseif ($value == (float)$value) {
                        $value = (float)$value;
                    }
                }
        }

        return $value;
    }

    /**
     * @return string
     */
    public function id(): string
    {
        return $this->sheetId;
    }

    /**
     * @return string
     */
    public function name(): string
    {
        return $this->name;
    }

    /**
     * @param string $name
     *
     * @return bool
     */
    public function isName(string $name): bool
    {
        return strcasecmp($this->name, $name) === 0;
    }

    /**
     * @param string|null $file
     *
     * @return Reader
     */
    protected function getReader(string $file = null): Reader
    {
        if (empty($this->xmlReader)) {
            if (!$file) {
                $file = $this->zipFilename;
            }
            $this->xmlReader = new Reader($file);
        }

        return $this->xmlReader;
    }

    public function dimension(): ?string
    {
        if ($this->dimension === null) {
            $xmlReader = $this->getReader();
            $xmlReader->openZip($this->path);
            if ($xmlReader->seekOpenTag('dimension')) {
                $this->dimension = (string)$xmlReader->getAttribute('ref');
            }

        }
        return $this->dimension;
    }

    /**
     * Count rows by dimension value
     *
     * @return int
     */
    public function countRows(): int
    {
        $areaRange = $this->dimension();
        if ($areaRange && preg_match('/^([A-Za-z]+)(\d+)(:([A-Za-z]+)(\d+))?$/', $areaRange, $matches)) {
            return (int)$matches[5] - (int)$matches[2] + 1;
        }

        return 0;
    }

    /**
     * Count columns by dimension value
     *
     * @return int
     */
    public function countColumns(): int
    {
        $areaRange = $this->dimension();
        if ($areaRange && preg_match('/^([A-Za-z]+)(\d+)(:([A-Za-z]+)(\d+))?$/', $areaRange, $matches)) {
            return Excel::colNum($matches[4]) - Excel::colNum($matches[1]) + 1;
        }

        return 0;
    }

    /**
     * Count columns by dimension value, alias of countColumns()
     *
     * @return int
     */
    public function countCols(): int
    {
        return $this->countColumns();
    }

    /**
     * @param $dateFormat
     *
     * @return $this
     */
    public function setDateFormat($dateFormat): Sheet
    {
        $this->excel->setDateFormat($dateFormat);

        return $this;
    }

    protected static function _areaRange(string $areaRange)
    {
        $area = [];
        if (preg_match('/^([A-Za-z]+)(\d+)(:([A-Za-z]+)(\d+))?$/', $areaRange, $matches)) {
            $area['col_min'] = Excel::colNum($matches[1]);
            $area['row_min'] = (int)$matches[2];
            if (empty($matches[3])) {
                $area['col_max'] = Excel::EXCEL_2007_MAX_COL;
                $area['row_max'] = Excel::EXCEL_2007_MAX_ROW;
            }
            else {
                $area['col_max'] = Excel::colNum($matches[4]);
                $area['row_max'] = (int)$matches[5];
            }
        }
        elseif (preg_match('/^([A-Za-z]+)(:([A-Za-z]+))?$/', $areaRange, $matches)) {
            $area['col_min'] = Excel::colNum($matches[1]);
            if (empty($matches[2])) {
                $area['col_max'] = Excel::EXCEL_2007_MAX_COL;
            }
            else {
                $area['col_max'] = Excel::colNum($matches[3]);
            }
            $area['row_min'] = 1;
            $area['row_max'] = Excel::EXCEL_2007_MAX_ROW;
        }

        return $area;
    }

    /**
     * setReadArea('C3:AZ28') - set top left and right bottom of read area
     * setReadArea('C3') - set top left only
     *
     * @param string $areaRange
     * @param bool|null $firstRowKeys
     *
     * @return $this
     */
    public function setReadArea(string $areaRange, ?bool $firstRowKeys = false): Sheet
    {
        $area = self::_areaRange($areaRange);
        if ($area && isset($area['row_max'])) {
            $this->area = $area;
            $this->area['first_row'] = $firstRowKeys;

            return $this;
        }
        throw new Exception('Wrong address or range "' . $areaRange . '"');
    }

    /**
     * setReadArea('C:AZ') - set left and right columns of read area
     * setReadArea('C') - set left column only
     *
     * @param string $columnsRange
     * @param bool|null $firstRowKeys
     *
     * @return $this
     */
    public function setReadAreaColumns(string $columnsRange, ?bool $firstRowKeys = false): Sheet
    {
        $area = self::_areaRange($columnsRange);
        if ($area) {
            $this->area = $area;
            $this->area['first_row'] = $firstRowKeys;

            return $this;
        }
        throw new Exception('Wrong address or range "' . $columnsRange . '"');
    }

    /**
     * Returns cell values as a two-dimensional array
     *      [1 => ['A' => _value_A1_], ['B' => _value_B1_]],
     *      [2 => ['A' => _value_A2_], ['B' => _value_B2_]]
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
    public function readRows($columnKeys = [], int $resultMode = null, ?bool $styleIdxInclude = null): array
    {
        $data = [];
        $this->readCallback(static function($row, $col, $val) use (&$columnKeys, &$data) {
            if (isset($columnKeys[$col])) {
                $data[$row][$columnKeys[$col]] = $val;
            }
            else {
                $data[$row][$col] = $val;
            }
        }, $columnKeys, $resultMode, $styleIdxInclude);

        if ($data && ($resultMode & Excel::KEYS_SWAP)) {
            $newData = [];
            $rowKeys = array_keys($data);
            $len = count($rowKeys);
            foreach (array_keys(reset($data)) as $colKey) {
                $rowValues = array_column($data, $colKey);
                if ($len - count($rowValues)) {
                    $rowValues = array_pad($rowValues, $len, null);
                }
                $newData[$colKey] = array_combine($rowKeys, $rowValues);
            }
            return $newData;
        }

        return $data;
    }

    /**
     * Returns values and styles of cells as array ['v' => _value_, 's' => _styles_]
     *
     * @param array|bool|int|null $columnKeys
     * @param int|null $resultMode
     *
     * @return array
     */
    public function readRowsWithStyles($columnKeys = [], int $resultMode = null): array
    {
        $data = $this->readRows($columnKeys, $resultMode, true);

        foreach ($data as $row => $rowData) {
            foreach ($rowData as $col => $cellData) {
                if (isset($cellData['s'])) {
                    $data[$row][$col]['s'] = $this->excel->getCompleteStyleByIdx($cellData['s']);
                }
            }
        }

        return $data;
    }

    /**
     * Returns cell values as a two-dimensional array from default sheet [col][row]
     *      ['A' => [1 => _value_A1_], [2 => _value_A2_]],
     *      ['B' => [1 => _value_B1_], [2 => _value_B2_]]
     *
     * @param array|bool|int|null $columnKeys
     * @param int|null $resultMode
     * @param bool|null $styleIdxInclude
     *
     * @return array
     */
    public function readColumns($columnKeys = null, int $resultMode = null, ?bool $styleIdxInclude = null): array
    {
        if (is_int($columnKeys) && $columnKeys > 1 && $resultMode === null) {
            $resultMode = $columnKeys | Excel::KEYS_RELATIVE;
            $columnKeys = $columnKeys & Excel::KEYS_FIRST_ROW;
        }
        else {
            $resultMode = $resultMode | Excel::KEYS_RELATIVE;
        }

        return $this->readRows($columnKeys, $resultMode | Excel::KEYS_SWAP);
    }

    /**
     * Returns values and styles of cells as array ['v' => _value_, 's' => _styles_]
     *
     * @param array|bool|int|null $columnKeys
     * @param int|null $resultMode
     *
     * @return array
     */
    public function readColumnsWithStyles($columnKeys = null, int $resultMode = null): array
    {
        $data = $this->readColumns($columnKeys, $resultMode, true);

        foreach ($data as $col => $colData) {
            foreach ($colData as $row => $cellData) {
                if (isset($cellData['s'])) {
                    $data[$col][$row]['s'] = $this->excel->getCompleteStyleByIdx($cellData['s']);
                }
            }
        }

        return $data;
    }

    /**
     * Returns values and styles of cells as array
     *
     * @param bool|null $styleIdxInclude
     *
     * @return array
     */
    public function readCells(?bool $styleIdxInclude = null): array
    {
        $data = [];
        $this->readCallback(static function($row, $col, $val) use (&$data) {
            $data[$col . $row] = $val;
        }, [], null, $styleIdxInclude);

        return $data;
    }

    /**
     * Returns values and styles of cells as array ['v' => _value_, 's' => _styles_]
     *
     * @return array
     */
    public function readCellsWithStyles(): array
    {
        $data = $this->readCells(true);
        foreach ($data as $cell => $cellData) {
            if (isset($cellData['s'])) {
                $data[$cell]['s'] = $this->excel->getCompleteStyleByIdx($cellData['s']);
            }
        }

        return $data;
    }

    /**
     * Reads cell values and passes them to a callback function
     *
     * @param callback $callback Callback function($row, $col, $value)
     * @param array|bool|int|null $columnKeys
     * @param int|null $resultMode
     * @param bool|null $styleIdxInclude
     */
    public function readCallback(callable $callback, $columnKeys = [], int $resultMode = null, ?bool $styleIdxInclude = null)
    {
        foreach ($this->nextRow($columnKeys, $resultMode, $styleIdxInclude) as $row => $rowData) {
            foreach ($rowData as $col => $val) {
                $needBreak = $callback($row, $col, $val);
                if ($needBreak) {
                    return;
                }
            }
        }
    }

    /**
     * Read cell values row by row, returns either an array of values or an array of arrays
     *
     *      nextRow(..., ...) : <rowNum> => [<colNum1> => <value1>, <colNum2> => <value2>, ...]
     *      nextRow(..., ..., true) : <rowNum> => [<colNum1> => ['v' => <value1>, 's' => <style1>], <colNum2> => ['v' => <value2>, 's' => <style2>], ...]
     *
     * @param array|bool|int|null $columnKeys
     * @param int|null $resultMode
     * @param bool|null $styleIdxInclude
     * @param int|null $rowLimit
     *
     * @return \Generator|null
     */
    public function nextRow($columnKeys = [], int $resultMode = null, ?bool $styleIdxInclude = null, int $rowLimit = 0): ?\Generator
    {
        // <dimension ref="A1:C1"/>
        // sometimes sheets doesn't contain this tag
        if ($this->dimension === null) {
            $this->dimension();
        }

        $readArea = $this->area;
        if (!empty($columnKeys) && is_array($columnKeys)) {
            $firstRowKeys = is_int($resultMode) && ($resultMode & Excel::KEYS_FIRST_ROW);
            $columnKeys = array_combine(array_map('strtoupper', array_keys($columnKeys)), array_values($columnKeys));
        }
        elseif ($columnKeys === true) {
            $firstRowKeys = true;
            $columnKeys = [];
        }
        elseif ($resultMode & Excel::KEYS_FIRST_ROW) {
            $firstRowKeys = true;
        }
        else {
            $firstRowKeys = !empty($readArea['first_row']);
        }

        if ($columnKeys && ($resultMode & Excel::KEYS_FIRST_ROW)) {
            foreach ($this->nextRow([], 0, null, 1) as $firstRowData) {
                $columnKeys = array_merge($firstRowData, $columnKeys);
                break;
            }
        }

        $xmlReader = $this->getReader();
        $xmlReader->openZip($this->path);

        $rowData = [];
        $rowNum = 0;
        $rowOffset = $colOffset = null;
        $row = -1;
        $rowCnt = -1;
        if ($xmlReader->seekOpenTag('sheetData')) {
            while ($xmlReader->read()) {
                if ($rowLimit > 0 && $rowCnt >= $rowLimit) {
                    break;
                }
                if ($xmlReader->nodeType === \XMLReader::END_ELEMENT && $xmlReader->name === 'sheetData') {
                    break;
                }

                if ($xmlReader->nodeType === \XMLReader::END_ELEMENT && $xmlReader->name === 'row' && $rowNum >= $readArea['row_min'] && $rowNum <= $readArea['row_max']) {
                    if ($rowCnt === 0 && $firstRowKeys) {
                        if (!$columnKeys) {
                            if ($styleIdxInclude) {
                                $columnKeys = array_combine(array_keys($rowData), array_column($rowData, 'v'));
                            }
                            else {
                                $columnKeys = $rowData;
                            }
                        }
                    }
                    else {
                        $row = $rowNum - $rowOffset;
                        yield $row => $rowData;
                    }
                    $rowData = [];
                    continue;
                }

                if ($xmlReader->nodeType === \XMLReader::ELEMENT) {
                    if ($xmlReader->name === 'row') { // <row ...> - tag row begins
                        $rowNum = (int)$xmlReader->getAttribute('r');

                        if ($rowNum > $readArea['row_max']) {
                            break;
                        }
                        if ($rowNum < $readArea['row_min']) {
                            continue;
                        }

                        $rowCnt += 1;
                        if ($rowOffset === null) {
                            $rowOffset = 0;
                            if (is_int($resultMode) && $resultMode) {
                                if ($resultMode & Excel::KEYS_ROW_ZERO_BASED) {
                                    $rowOffset = $rowNum + ($firstRowKeys ? 1 : 0);
                                }
                                elseif ($resultMode & Excel::KEYS_ROW_ONE_BASED) {
                                    $rowOffset = $rowNum - 1 + ($firstRowKeys ? 1 : 0);
                                }
                            }
                        }
                    } // <row ...> - tag row end

                    elseif ($xmlReader->name === 'c') { // <c ...> - tag cell begins
                        $addr = $xmlReader->getAttribute('r');
                        if ($addr && preg_match('/^([A-Za-z]+)(\d+)$/', $addr, $m)) {
                            //
                            if ($m[2] < $readArea['row_min'] || $m[2] > $readArea['row_max']) {
                                continue;
                            }
                            $colLetter = $m[1];
                            $colNum = Excel::colNum($colLetter);

                            if ($colNum >= $readArea['col_min'] && $colNum <= $readArea['col_max']) {
                                if ($colOffset === null) {
                                    $colOffset = $colNum - 1;
                                    if (is_int($resultMode) && ($resultMode & Excel::KEYS_COL_ZERO_BASED)) {
                                        $colOffset += 1;
                                    }
                                }
                                if ($resultMode) {
                                    if (!($resultMode & (Excel::KEYS_COL_ZERO_BASED | Excel::KEYS_COL_ONE_BASED))) {
                                        $col = $colLetter;
                                    }
                                    else {
                                        $col = $colNum - $colOffset;
                                    }
                                }
                                else {
                                    $col = $colLetter;
                                }
                                $cell = $xmlReader->expand();
                                if (is_array($columnKeys) && isset($columnKeys[$colLetter])) {
                                    $col = $columnKeys[$colLetter];
                                }
                                $value = $this->_cellValue($cell, $styleIdx);
                                if ($styleIdxInclude) {
                                    $rowData[$col] = ['v' => $value, 's' => $styleIdx];
                                }
                                else {
                                    $rowData[$col] = $value;
                                }
                            }
                        }
                    } // <c ...> - tag cell end
                }
            }
        }
        if ($row > -1 && $rowData) {
            yield $row => $rowData;
        }

        $xmlReader->close();

        return null;
    }

    /**
     * @return string|null
     */
    protected function drawingFilename(): ?string
    {
        $findName = str_replace('/worksheets/sheet', '/drawings/drawing', $this->path);

        return in_array($findName, $this->excel->innerFileList(), true) ? $findName : null;
    }

    /**
     * @param $xmlName
     *
     * @return array
     */
    protected function extractDrawingInfo($xmlName): array
    {
        $drawings = [
            'xml' => $xmlName,
            'rel' => dirname($xmlName) . '/_rels/' . basename($xmlName) . '.rels',
        ];
        $contents = file_get_contents('zip://' . $this->zipFilename . '#' . $xmlName);
        if (preg_match_all('#<xdr:twoCellAnchor[^>]*>(.*)</xdr:twoCellAnchor#siU', $contents, $anchors)) {
            foreach ($anchors[1] as $twoCellAnchor) {
                $drawing = [];
                if (preg_match('#<xdr:pic>(.*)</xdr:pic>#siU', $twoCellAnchor, $pic)) {
                    if (preg_match('#<a:blip\s(.*)r:embed="(.+)"#siU', $twoCellAnchor, $m)) {
                        $drawing['rId'] = $m[2];
                    }
                    if ($drawing && preg_match('#<xdr:cNvPr(.*)\sname="(.*)">#siU', $pic[1], $m)) {
                        $drawing['name'] = $m[2];
                    }
                }
                if ($drawing) {
                    if (preg_match('#<xdr:from[^>]*>(.*)</xdr:from#siU', $twoCellAnchor, $m)) {
                        if (preg_match('#<xdr:col>(.*)</xdr:col#siU', $m[1], $m1)) {
                            $drawing['colIdx'] = (int)$m1[1];
                            $drawing['col'] = Excel::colLetter($drawing['colIdx'] + 1);
                        }
                        if (preg_match('#<xdr:row>(.*)</xdr:row#siU', $m[1], $m1)) {
                            $drawing['rowIdx'] = (int)$m1[1];
                            $drawing['row'] = (string)($drawing['rowIdx'] + 1);
                        }
                    }
                    $drawings['media'][$drawing['rId']] = $drawing;
                    if (isset($drawing['col'], $drawing['row'])) {
                        $drawing['cell'] = $drawing['col'] . $drawing['row'];
                    }
                }
            }
        }

        if (!empty($drawings['media'])) {
            $contents = file_get_contents('zip://' . $this->zipFilename . '#' . $drawings['rel']);
            if (preg_match_all('#<Relationship\s([^>]+)>#siU', $contents, $rel)) {
                foreach ($rel[1] as $str) {
                    if (preg_match('#Id="(\w+)#', $str, $m1) && preg_match('#Target="([^"]+)#', $str, $m2)) {
                        $rId = $m1[1];
                        if (isset($drawings['media'][$rId])) {
                            $drawings['media'][$rId]['target'] = str_replace('../', 'xl/', $m2[1]);
                        }
                    }
                }
            }
        }

        $result = [
            'xml' => $drawings['xml'],
            'rel' => $drawings['rel'],
        ];
        foreach ($drawings['media'] as $media) {
            if (isset($media['target'])) {
                $addr = $media['col'] . $media['row'];
                if (!isset($media['name'])) {
                    $media['name'] = $addr;
                }
                $result['images'][$addr] = $media;
                $result['rows'][$media['row']][] = $addr;
            }
        }

        return $result;
    }

    /**
     * @return bool
     */
    public function hasDrawings(): bool
    {
        return (bool)$this->drawingFilename();
    }

    /**
     * @return int
     */
    public function countImages(): int
    {
        $result = 0;
        if ($this->hasDrawings()) {
            if (!isset($this->props['drawings'])) {
                if ($xmlName = $this->drawingFilename()) {
                    $this->props['drawings'] = $this->extractDrawingInfo($xmlName);
                }
                else {
                    $this->props['drawings'] = [];
                }
            }
            if (!empty($this->props['drawings']['images'])) {
                $result = count($this->props['drawings']['images']);
            }
        }

        return $result;
    }

    /**
     * @return array
     */
    public function getImageList(): array
    {
        $result = [];
        if ($this->countImages()) {
            foreach ($this->props['drawings']['images'] as $addr => $image) {
                $result[$addr] = [
                    'image_name' => $image['name'],
                    'file_name' => basename($image['target']),
                ];
            }
        }

        return $result;
    }

    /**
     * @return array
     */
    public function getImageListByRow($row): array
    {
        $result = [];
        if ($this->countImages()) {
            if (isset($this->props['drawings']['rows'][$row])) {
                foreach ($this->props['drawings']['rows'][$row] as $addr) {
                    $result[$addr] = [
                        'image_name' => $this->props['drawings']['images'][$addr]['name'],
                        'file_name' => basename($this->props['drawings']['images'][$addr]['target']),
                    ];
                }
            }
        }

        return $result;
    }

    /**
     * Returns TRUE if the cell contains an image
     *
     * @param string $cell
     *
     * @return bool
     */
    public function hasImage(string $cell): bool
    {
        if ($this->countImages()) {

            return isset($this->props['drawings']['images'][strtoupper($cell)]);
        }

        return false;
    }

    /**
     * Returns full path of an image from the cell (if exists) or null
     *
     * @param string $cell
     *
     * @return string|null
     */
    public function imageEntryFullPath(string $cell): ?string
    {
        if ($this->countImages()) {
            $cell = strtoupper($cell);
            if (isset($this->props['drawings']['images'][$cell])) {

                return 'zip://' . $this->zipFilename . '#' . $this->props['drawings']['images'][$cell]['target'];
            }
        }

        return null;
    }

    /**
     * Returns the MIME type for an image from the cell as determined by using information from the magic.mime file
     * Requires fileinfo extension
     *
     * @param string $cell
     *
     * @return string|null
     */
    public function getImageMimeType(string $cell): ?string
    {
        if (function_exists('mime_content_type') && ($path = $this->imageEntryFullPath($cell))) {
            return mime_content_type($path);
        }

        return null;
    }

    /**
     * Returns the name for an image from the cell as it defines in XLSX
     *
     * @param string $cell
     *
     * @return string|null
     */
    public function getImageName(string $cell): ?string
    {
        if ($this->countImages()) {
            $cell = strtoupper($cell);
            if (isset($this->props['drawings']['images'][$cell])) {

                return $this->props['drawings']['images'][$cell]['name'];
            }
        }

        return null;
    }

    /**
     * Returns an image from the cell as a blob (if exists) or null
     *
     * @param string $cell
     *
     * @return string|null
     */
    public function getImageBlob(string $cell): ?string
    {
        if ($path = $this->imageEntryFullPath($cell)) {
            return file_get_contents($path);
        }

        return null;
    }

    /**
     * Writes an image from the cell to the specified filename
     *
     * @param string $cell
     * @param string|null $filename
     *
     * @return string|null
     */
    public function saveImage(string $cell, ?string $filename = null): ?string
    {
        if ($contents = $this->getImageBlob($cell)) {
            if (!$filename) {
                $filename = basename($this->props['drawings']['images'][strtoupper($cell)]['target']);
            }
            if (file_put_contents($filename, $contents)) {
                return realpath($filename);
            }
        }

        return null;
    }

    /**
     * Writes an image from the cell to the specified directory
     *
     * @param string $cell
     * @param string $dirname
     *
     * @return string|null
     */
    public function saveImageTo(string $cell, string $dirname): ?string
    {
        $filename = basename($this->props['drawings']['images'][strtoupper($cell)]['target']);

        return $this->saveImage($cell, str_replace(['\\', '/'], '', $dirname) . DIRECTORY_SEPARATOR . $filename);
    }
}