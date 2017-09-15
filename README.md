# ExportTools

## Fast MySQL-to-Excel exporter

This library uses MySQL's `SELECT INTO OUTFILE` feature to export selected rows into a csv file
and then converts that csv into xlsx via a C-language program (Linux only). This is much faster than other pure PHP
solutions.

Example usage:
```php
<?php

use SolveX\ExportTools\ExcelExporter;

$db = app('db'); // (Illuminate\Database\ConnectionInterface (Laravel)
$excelExporter = new ExcelExporter($db);

$sql = "SELECT first_name, last_name, phone, email FROM `customers` WHERE id = 1";
$columns = ['First Name', 'Last Name', 'Phone', 'Email'];

$path = $excelExporter->export($sql, $columns);
```
