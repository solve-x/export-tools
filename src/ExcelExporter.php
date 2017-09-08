<?php

namespace SolveX\ExportTools;

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
        $path = $this->exportCSV($sql, $columns);
        $excelPath = str_replace('.csv', '.xlsx', $path);
        $this->csvToExcelConverter->convert($path, $excelPath);
        unlink($path);
        return $excelPath;
    }

    /**
     * Stores results of $sql (SELECT statement) into a temporary csv file.
     *
     * @param  string $sql     SELECT statement
     * @param  array  $columns Names of headers
     * @return string          Path to the temporary csv file
     */
    public function exportCSV($sql, $columns)
    {
        $path = $this->getTemporaryCsvPath();
        $this->db->statement(
            $this->prepareIntoOutfileSqlStatement($sql, $columns, $path)
        );
        return $path;
    }

    /**
     * Generates a random file name in a temporary directory where MySQL has write permissions (see secure_file_priv).
     *
     * @link https://dev.mysql.com/doc/refman/5.7/en/server-system-variables.html#sysvar_secure_file_priv
     * @return string
     */
    protected function getTemporaryCsvPath()
    {
        $filename = 'tmp_' . bin2hex(random_bytes(5)) . '.csv';
        $record = $this->db->selectOne('SELECT @@secure_file_priv AS secure_file');
        $directory = $record->secure_file === null ? sys_get_temp_dir() : $record->secure_file;
        return $directory . DIRECTORY_SEPARATOR . $filename;
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
