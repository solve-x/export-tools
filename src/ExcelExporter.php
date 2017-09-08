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
     * Query execution is delegated to a component implementing
     * Illuminate\Database\ConnectionInterface interface. This interface
     * is provided by Laravel (but can be used without it, see
     * https://github.com/illuminate/database).
     *
     * @var ConnectionInterface
     */
    protected $db;

    /**
     * Conversion between CSV and Excel is delegated to another class.
     *
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
     * Stores results of $sql (a SELECT statement) into a temporary Excel file.
     *
     * @param string $sql
     * @param array $columns Names of columns (header).
     * @return string Path to the temporary Excel file.
     */
    public function export($sql, $columns)
    {
        $pathCsv = $this->exportCsv($sql, $columns);
        $pathExcel = str_replace('.csv', '.xlsx', $pathCsv);
        $this->csvToExcelConverter->convert($pathCsv, $pathExcel);
        $this->tryRemoveFile($pathCsv);
        return $pathExcel;
    }

    /**
     * Stores results of $sql (a SELECT statement) into a temporary csv file.
     *
     * @param  string $sql
     * @param  array $columns Names of columns (header).
     * @return string Path to the temporary csv file.
     */
    public function exportCsv($sql, $columns)
    {
        $path = $this->getTemporaryCsvPath();
        $this->db->statement(
            $this->prepareIntoOutfileSqlStatement($sql, $columns, $path)
        );
        return $path;
    }

    /**
     * Generates a random filename in a temporary directory where MySQL should have write permissions (see secure_file_priv).
     * Important: make sure that the PHP process using ExcelExporter has write permission to the same
     * directory as well!
     *
     * @link https://dev.mysql.com/doc/refman/5.7/en/server-system-variables.html#sysvar_secure_file_priv
     * @return string
     */
    protected function getTemporaryCsvPath()
    {
        $filename = 'tmp_' . bin2hex(random_bytes(5)) . '.csv';
        $record = $this->db->selectOne('SELECT @@secure_file_priv AS secure_file');
        $directory = $record->secure_file ?: sys_get_temp_dir();
        return $directory . DIRECTORY_SEPARATOR . $filename;
    }

    /**
     * Tries to remove the given file. This is used to remove temporary CSV file
     * created by MySQL when using "SELECT INTO OUTFILE" statement.
     * Since MySQL created that file (with "mysql" user as the owner for example),
     * this process (PHP process under Apache or similar) might not be allowed to remove it.
     * We continue anyway (you should adjust permissions or regularly remove temporary files).
     *
     * @param string $path
     */
    protected function tryRemoveFile($path)
    {
        try {
            unlink($path);
        } catch (\ErrorException $e) {
            // We are not allowed to delete this file ("Operation not permitted"),
            // proceed anyway.
            // NOTE: unlink doesn't actually throw exceptions, it generates errors.
            // But Laravel and other frameworks convert this error into an
            // ErrorException.
        }
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

    /**
     * Prepare a select statement from column names (headers).
     *
     * @param array $columns
     * @return string
     */
    protected function getSelectStatement(array $columns)
    {
        $columns = $this->wrapWithSingleQuotes($columns);
        $columns = implode(',', $columns);
        return "SELECT $columns";
    }

    /**
     * Wrap each string in an array of strings with single quotes.
     *
     * @param array $strings
     * @return array
     */
    protected function wrapWithSingleQuotes(array $strings)
    {
        $f = function ($string) {
            return "'$string'";
        };

        return array_map($f, $strings);
    }

    /**
     * CSV export options supported by MySQL.
     *
     * @return string
     */
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
