<?php

declare(strict_types=1);

namespace Mnb\PHPExcel\Reader;

use Mnb\PHPExcel\Reader\State\FormulaResult;
use Mnb\PHPExcel\Support\ErrorCode;
use Mnb\PHPExcel\Support\LocaleNormalizer;
use Mnb\PHPExcel\Support\MnbExcelException;
use Mnb\PHPExcel\Support\Xml\XmlReader;
use Mnb\PHPExcel\Support\Zip\ZipArchive;

/** Forward-only OpenDocument Spreadsheet (.ods) reader. */
final class OdsReader implements IterableReaderInterface, FormatAwareReaderInterface, SheetNamesReaderInterface
{
    public function format(): string
    {
        return 'ods';
    }

    /** @return list<list<mixed>> */
    public function readSheet(string $path, int|string $sheet = 1, array $options = []): array
    {
        return array_values(iterator_to_array($this->iterateSheet($path, $sheet, $options), true));
    }

    /** @return \Generator<int,list<mixed>> */
    public function iterateSheet(string $path, int|string $sheet = 1, array $options = []): iterable
    {
        $this->ensureExtensions();
        $realPath = $this->assertOds($path, $options);
        $projection = ColumnProjection::fromOptions($options);
        $selectedIndex = is_int($sheet) || ctype_digit((string) $sheet) ? max(1, (int) $sheet) : null;
        $selectedName = is_string($sheet) && !ctype_digit($sheet) ? $sheet : null;
        $reader = new XMLReader();
        if (!@$reader->open($this->zipUri($realPath, 'content.xml'), null, LIBXML_NONET | LIBXML_COMPACT)) {
            throw MnbExcelException::withCode('Unable to open ODS content.xml.', ErrorCode::FILE_OPEN_FAILED, ['path' => $path]);
        }

        $sheetIndex = 0;
        $insideSelected = false;
        $selectedFound = false;
        $tableDepth = null;
        $sourceRow = 0;
        $delivered = 0;
        $startRow = max(1, (int) ($options['start_row'] ?? 1));
        $endRow = isset($options['end_row']) ? max(1, (int) $options['end_row']) : null;
        $sourceLimitRows = isset($options['source_limit_rows']) ? max(0, (int) $options['source_limit_rows']) : null;
        $maxRows = isset($options['max_source_rows']) ? max(0, (int) $options['max_source_rows']) : null;
        $maxRepeatedRows = max(1, (int) ($options['max_repeated_rows'] ?? 1000000));

        try {
            while ($reader->read()) {
                if ($reader->nodeType === XMLReader::ELEMENT && $reader->localName === 'table') {
                    $sheetIndex++;
                    $name = $this->attributeByLocalName($reader, 'name') ?? 'Sheet' . $sheetIndex;
                    $insideSelected = ($selectedIndex !== null && $sheetIndex === $selectedIndex)
                        || ($selectedName !== null && $name === $selectedName);
                    if ($insideSelected) {
                        $selectedFound = true;
                        $tableDepth = $reader->depth;
                    }
                    continue;
                }
                if ($reader->nodeType === XMLReader::END_ELEMENT && $reader->localName === 'table' && $insideSelected && $reader->depth === $tableDepth) {
                    break;
                }
                if (!$insideSelected || $reader->nodeType !== XMLReader::ELEMENT || $reader->localName !== 'table-row') {
                    continue;
                }

                $repeat = max(1, (int) ($this->attributeByLocalName($reader, 'number-rows-repeated') ?? 1));
                if ($repeat > $maxRepeatedRows) {
                    throw MnbExcelException::withCode('ODS repeated-row limit exceeded.', ErrorCode::FILE_READ_FAILED, ['repeat' => $repeat, 'max_repeated_rows' => $maxRepeatedRows]);
                }
                $rowXml = $reader->readOuterXml();
                $row = $this->parseRow($rowXml, $options, $projection);

                for ($copy = 0; $copy < $repeat; $copy++) {
                    $sourceRow++;
                    if ($maxRows !== null && $sourceRow > $maxRows) {
                        throw MnbExcelException::withCode('ODS row limit exceeded.', ErrorCode::FILE_READ_FAILED, ['rows' => $sourceRow, 'max_source_rows' => $maxRows]);
                    }
                    if ($sourceRow < $startRow) {
                        continue;
                    }
                    if ($endRow !== null && $sourceRow > $endRow) {
                        break 2;
                    }
                    if ($sourceLimitRows !== null && $delivered >= $sourceLimitRows) {
                        break 2;
                    }
                    yield $sourceRow - 1 => $row;
                    $delivered++;
                }
            }
        } finally {
            $reader->close();
        }

        if (!$selectedFound) {
            throw new MnbExcelException('ODS sheet does not exist: ' . (string) $sheet);
        }
    }

    /** @return list<string> */
    public function sheetNames(string $path, array $options = []): array
    {
        $this->ensureExtensions();
        $realPath = $this->assertOds($path, $options);
        $reader = new XMLReader();
        if (!@$reader->open($this->zipUri($realPath, 'content.xml'), null, LIBXML_NONET | LIBXML_COMPACT)) {
            throw MnbExcelException::withCode('Unable to open ODS content.xml.', ErrorCode::FILE_OPEN_FAILED, ['path' => $path]);
        }
        $names = [];
        try {
            while ($reader->read()) {
                if ($reader->nodeType === XMLReader::ELEMENT && $reader->localName === 'table') {
                    $names[] = $this->attributeByLocalName($reader, 'name') ?? 'Sheet' . (count($names) + 1);
                }
            }
        } finally {
            $reader->close();
        }
        return $names;
    }

    /** @return array<string,list<list<mixed>>> */
    public function readWorkbook(string $path, array $options = []): array
    {
        $workbook = [];
        foreach ($this->sheetNames($path, $options) as $name) {
            $workbook[$name] = $this->readSheet($path, $name, $options);
        }
        return $workbook;
    }

    /** @return list<mixed> */
    private function parseRow(string $xml, array $options, ColumnProjection $projection): array
    {
        $reader = new XMLReader();
        if (!@$reader->XML($xml, null, LIBXML_NONET | LIBXML_COMPACT)) {
            throw MnbExcelException::withCode('Unable to parse ODS row.', ErrorCode::FILE_READ_FAILED);
        }
        $row = [];
        $columnNumber = 1;
        $maxRepeatedColumns = max(1, (int) ($options['max_repeated_columns'] ?? 16384));
        try {
            while ($reader->read()) {
                if ($reader->nodeType !== XMLReader::ELEMENT || !in_array($reader->localName, ['table-cell', 'covered-table-cell'], true)) {
                    continue;
                }
                $repeat = max(1, (int) ($this->attributeByLocalName($reader, 'number-columns-repeated') ?? 1));
                if ($repeat > $maxRepeatedColumns) {
                    throw MnbExcelException::withCode('ODS repeated-column limit exceeded.', ErrorCode::FILE_READ_FAILED, ['repeat' => $repeat, 'max_repeated_columns' => $maxRepeatedColumns]);
                }

                $selected = !$projection->active();
                if (!$selected) {
                    for ($offset = 0; $offset < $repeat; $offset++) {
                        if ($projection->includesIndex($columnNumber + $offset)) {
                            $selected = true;
                            break;
                        }
                    }
                }
                $value = null;
                if ($selected) {
                    $value = $reader->localName === 'covered-table-cell' ? null : $this->cellValue($reader, $options);
                }

                for ($offset = 0; $offset < $repeat; $offset++, $columnNumber++) {
                    if ($projection->active() && !$projection->includesIndex($columnNumber)) {
                        continue;
                    }
                    if ($projection->active() && $projection->compact()) {
                        $row[] = $value;
                    } else {
                        $row[$columnNumber - 1] = $value;
                    }
                }
            }
        } finally {
            $reader->close();
        }

        if ($row !== [] && !array_is_list($row)) {
            ksort($row);
            $max = max(array_keys($row));
            $row = array_replace(array_fill(0, $max + 1, null), $row);
        }
        if ((bool) ($options['trim_trailing_empty_cells'] ?? true)) {
            while ($row !== [] && end($row) === null) {
                array_pop($row);
            }
        }
        return $row;
    }

    private function cellValue(XMLReader $reader, array $options): mixed
    {
        $valueType = strtolower((string) ($this->attributeByLocalName($reader, 'value-type') ?? 'string'));
        $formula = $this->attributeByLocalName($reader, 'formula');
        $cached = match ($valueType) {
            'float', 'currency', 'percentage' => LocaleNormalizer::parseCanonicalNumber((string) ($this->attributeByLocalName($reader, 'value') ?? '0'), $options),
            'boolean' => strtolower((string) ($this->attributeByLocalName($reader, 'boolean-value') ?? 'false')) === 'true',
            'date' => $this->formatDate((string) ($this->attributeByLocalName($reader, 'date-value') ?? ''), $options),
            'time' => (string) ($this->attributeByLocalName($reader, 'time-value') ?? ''),
            default => $this->readCellText($reader),
        };

        if ($formula === null || $formula === '') {
            return $cached;
        }
        $normalizedFormula = preg_replace('/^of:=/i', '=', $formula) ?? $formula;
        $mode = strtolower((string) ($options['formula_cells'] ?? 'formula'));
        return match ($mode) {
            'cached_value' => $cached,
            'both' => new FormulaResult($normalizedFormula, $cached, $valueType, ['source' => 'ods']),
            default => $normalizedFormula,
        };
    }

    private function readCellText(XMLReader $reader): string
    {
        if ($reader->isEmptyElement) {
            return '';
        }
        $depth = $reader->depth;
        $text = '';
        while ($reader->read()) {
            if ($reader->nodeType === XMLReader::END_ELEMENT && $reader->depth === $depth) {
                break;
            }
            if (in_array($reader->nodeType, [XMLReader::TEXT, XMLReader::CDATA, XMLReader::SIGNIFICANT_WHITESPACE], true)) {
                $text .= $reader->value;
            }
            if ($reader->nodeType === XMLReader::ELEMENT && $reader->localName === 'line-break') {
                $text .= "\n";
            }
            if ($reader->nodeType === XMLReader::ELEMENT && $reader->localName === 'tab') {
                $text .= "\t";
            }
        }
        return $text;
    }

    private function formatDate(string $value, array $options): mixed
    {
        if ($value === '') {
            return null;
        }
        try {
            $date = new \DateTimeImmutable($value);
        } catch (\Throwable) {
            return $value;
        }
        if ((bool) ($options['return_datetime'] ?? false)) {
            return $date;
        }
        return $date->format(str_contains($value, 'T') ? (string) ($options['datetime_format'] ?? 'Y-m-d H:i:s') : (string) ($options['date_format'] ?? 'Y-m-d'));
    }

    private function attributeByLocalName(XMLReader $reader, string $localName): ?string
    {
        if (!$reader->hasAttributes || !$reader->moveToFirstAttribute()) {
            return null;
        }
        $value = null;
        do {
            if ($reader->localName === $localName) {
                $value = $reader->value;
                break;
            }
        } while ($reader->moveToNextAttribute());
        $reader->moveToElement();
        return $value;
    }

    /** @param array<string,mixed> $options */
    private function assertOds(string $path, array $options): string
    {
        $realPath = realpath($path);
        if ($realPath === false || !is_file($path)) {
            throw MnbExcelException::withCode('ODS file not found: ' . $path, ErrorCode::FILE_NOT_FOUND, ['path' => $path]);
        }
        $size = filesize($realPath);
        $max = isset($options['max_file_bytes']) ? max(0, (int) $options['max_file_bytes']) : null;
        if ($max !== null && $size !== false && $size > $max) {
            throw MnbExcelException::withCode('ODS file exceeds max_file_bytes.', ErrorCode::FILE_READ_FAILED, ['size_bytes' => $size, 'max_file_bytes' => $max]);
        }
        $zip = new ZipArchive();
        if ($zip->open($realPath) !== true) {
            throw MnbExcelException::withCode('Unable to open ODS package.', ErrorCode::FILE_OPEN_FAILED, ['path' => $path]);
        }
        try {
            $hasContent = $zip->locateName('content.xml') !== false;
            $mimetype = $zip->getFromName('mimetype');
        } finally {
            $zip->close();
        }
        if (!$hasContent || (is_string($mimetype) && trim($mimetype) !== '' && trim($mimetype) !== 'application/vnd.oasis.opendocument.spreadsheet')) {
            throw MnbExcelException::withCode('The file is not a valid ODS spreadsheet.', ErrorCode::UNSUPPORTED_FORMAT, ['path' => $path]);
        }
        return $realPath;
    }

    private function zipUri(string $realPath, string $entry): string
    {
        return 'zip://' . str_replace('\\', '/', $realPath) . '#' . $entry;
    }

    private function ensureExtensions(): void
    {
        if (!class_exists(ZipArchive::class)) {
            throw MnbExcelException::withCode('ext-zip is required to read ODS files.', ErrorCode::EXTENSION_MISSING);
        }
        if (!class_exists(XMLReader::class)) {
            throw MnbExcelException::withCode('ext-xmlreader is required to read ODS files.', ErrorCode::EXTENSION_MISSING);
        }
    }
}
