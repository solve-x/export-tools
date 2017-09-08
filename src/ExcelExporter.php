<?php

namespace App\Support;

use Illuminate\Database\ConnectionInterface;

/**
 * ExcelExporter provides a convenient and fast way to export database tables (or any SELECT)
 * into an Excel file. First, data is exported into a temporary CSV file (MySQL's built-in feature),
 * then a C-language program (CsvToExcelConverter) is used to convert that CSV to Excel.
 */
class ExcelExporter
{
    /**
     * @var ConnectionInterface
     */
    protected $db;

    /**
     * @var CsvToExcelConverter
     */
    protected $csvToExcelConverter;

    /**
     * ExcelExporter constructor.
     *
     * @param ConnectionInterface $db
     * @param CsvToExcelConverter $csvToExcelConverter
     */
    public function __construct(ConnectionInterface $db, CsvToExcelConverter $csvToExcelConverter = null)
    {
        $this->db = $db;
        $this->csvToExcelConverter = $csvToExcelConverter ?: new CsvToExcelConverter();
    }

    /**
     * Stores results of $sql (SELECT statement) into a temporary Excel file.
     *
     * @param string $sql
     * @param array $columns Names of columns (header).
     * @return string Path to the temporary Excel file.
     */
    public function export($sql, $columns)
    {
        $path = $this->getTemporaryCsvPath();
        $this->db->statement(
            $this->prepareIntoOutfileSqlStatement($sql, $columns, $path)
        );
        $excelPath = $path . '.xlsx';
        $this->csvToExcelConverter->convert($path, $excelPath);
        return $excelPath;
    }

    /**
     * Note: MySQL needs permission to write to the system's temporary directory (see secure-file-priv) !
     * This method returns the path to a file (non-existing) in the temporary directory.
     *
     * @return string
     */
    protected function getTemporaryCsvPath()
    {
        $filename = 'tmp_' . bin2hex(random_bytes(5)) . '.csv';
        $path = sys_get_temp_dir() . '/' . $filename;
        return $path;
    }

    /**
     * Note: MySQL "SELECT INTO OUTFILE" statement requires special syntax in order to add
     * column headers.
     *
     * @param string $sql
     * @param array $columns
     * @param string $path
     * @return string
     */
    protected function prepareIntoOutfileSqlStatement($sql, $columns, $path)
    {
        $select = $this->getSelectStatement($columns);
        $csvOptions = $this->getCsvOptions();

        return "
            (
                $select
            )
            UNION ALL
            (
                $sql
                INTO OUTFILE '$path'
                $csvOptions
            )
        ";
    }

    protected function getSelectStatement(array $columns)
    {
        $columns = $this->wrapWithSingleQuotes($columns);
        $columns = implode(',', $columns);
        return "SELECT $columns";
    }

    protected function wrapWithSingleQuotes(array $strings)
    {
        $f = function ($string) {
            return "'$string'";
        };

        return array_map($f, $strings);
    }

    protected function getCsvOptions()
    {
        return "
            FIELDS ENCLOSED BY '\"'
            ESCAPED BY '\"'
            TERMINATED BY '\t'
            LINES TERMINATED BY '\\n'
        ";
    }
}
