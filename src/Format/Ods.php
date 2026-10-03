<?php

declare(strict_types=1);

namespace Mnb\PHPExcel\Format;

use Mnb\PHPExcel\Cloud\CloudAccess;
use Mnb\PHPExcel\Reader\OdsReader;
use Mnb\PHPExcel\Reader\Options\ReaderOptions;
use Mnb\PHPExcel\Reader\ReadSession;

final class Ods
{
    use CloudAccess;
    /** @param array<string,mixed>|ReaderOptions $options */
    public static function read(string $path, array|ReaderOptions $options = []): ReadSession
    {
        return new ReadSession($path, new OdsReader(), $options);
    }
}
