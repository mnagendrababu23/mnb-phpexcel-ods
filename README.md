# MNB PHPExcel ODS

<<<<<<< HEAD
Native streaming ODS reader module for MNB PHPExcel.
Documentation URL: https://mnbphpexcel.space/getting-started/installation
This package is generated from the MNB PHPExcel monorepo. Do not copy source files between modules manually.
## MNB PHPExcel Assistant

Generate MNB PHPExcel code using our dedicated ChatGPT assistant:

[Open MNB PHPExcel AI Assistant](https://chatgpt.com/g/g-6a6e31d80350819194b68853d41c1561-mnb-phpexcel-assistant)
## Install
=======
Independent native OpenDocument Spreadsheet reader. Requires core, `ext-libxml`, and `ext-zlib`.
>>>>>>> 92854cd (Release v2.0.0)

```bash
composer require mnb/mnb-phpexcel-ods:^2.0
```

```php
use Mnb\PHPExcel\Format\Ods;

$rows = Ods::read('report.ods')
    ->sheet('Data')
    ->withHeaderRow()
    ->toArray();
```

Version 2.0 provides native ODS reading and row streaming. ODS writing is not yet part of this package.
