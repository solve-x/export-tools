# ExportTools

Description

```php
<?php

use SolveX\ExportTools\ExcelExporter;

$db = app('DB'); // Laravel
$excelExporter = new ExcelExporter($db);

$sql = "SELECT first_name, last_name, phone, email FROM `customers` WHERE id = 1";
$columns = ['First Name', 'Last Name', 'Phone', 'Email'];

$path = $excelExporter->export($sql, $columns); // string Path to the temporary Excel file.


```
