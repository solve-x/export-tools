<?php

namespace SolveX\ExportTools;

use Box\Spout\Common\Type;
use Box\Spout\Reader\ReaderFactory;
use Box\Spout\Writer\Style\StyleBuilder;
use Box\Spout\Writer\WriterFactory;

/**
 * Converts CSV files to XLSX.
 */
class CsvToExcelConverter
{
    /**
     * Converts given CSV file to the given Excel (XLSX) file.
     *
     * Conversion assumes that tab (\t) is used as the field delimiter and newline (\n)
     * is the end-of-line character. Fields can be enclosed with double quotes (").
     *
     * On Linux this uses a small external program written in C, included in binary form in csv2xlsx/ subfolder.
     * Program executable is statically linked against its dependencies (libxlsxwriter, zlib and libcsv)
     * to ensure portability.
     *
     * On Windows this uses the Spout (Box\Spout) PHP library.
     *
     * @param string $pathCsv Full path to the input file in CSV format.
     * @param string $pathExcel Full path to the output Excel file.
     * @return void
     */
    public function convert($pathCsv, $pathExcel)
    {
        $isWindows = (strtoupper(substr(PHP_OS, 0, 3)) === 'WIN');

        if ($isWindows) {
            $this->convertWindows($pathCsv, $pathExcel);
        } else {
            $this->convertLinux($pathCsv, $pathExcel);
        }
    }

    protected function convertLinux($pathCsv, $pathExcel)
    {
        $executable = __DIR__ . '/csv2xlsx/csv2xlsx';
        $convertCommand = $executable . " '$pathCsv' '$pathExcel'";
        exec($convertCommand);
    }

    protected function convertWindows($pathCsv, $pathExcel)
    {
        ini_set('max_execution_time', 0);
        $reader = ReaderFactory::create(Type::CSV);

        $defaultStyle = (new StyleBuilder())
            ->setFontName('Arial')
            ->setFontSize(12)
            ->build();

        $headerStyle = (new StyleBuilder())
            ->setFontBold()
            ->setFontSize(12)
            ->build();

        $reader->setFieldDelimiter("\t");
        $reader->setFieldEnclosure('"');
        $reader->setEndOfLineCharacter("\n");
        $reader->open($pathCsv);

        $writer = WriterFactory::create(Type::XLSX);
        $writer->setDefaultRowStyle($defaultStyle);
        $writer->openToFile($pathExcel);

        $firstRow = true;
        foreach ($reader->getSheetIterator() as $sheet) {
            foreach ($sheet->getRowIterator() as $row) {
                if ($firstRow) {
                    $writer->addRowWithStyle($row, $headerStyle);
                    $firstRow = false;
                    continue;
                }
                $writer->addRow($row);
            }
        }

        $writer->close();
        $reader->close();
    }
}
