<?php

require_once(implode(DIRECTORY_SEPARATOR, [__DIR__, '..', 'src', 'SimpleXLSXGen.php']));

$xlsx = new \Shuchkin\SimpleXLSXGen();

$xlsx->addSheet([
    ["foo", "bar"],
    ["something", '<style nf="DD.MM.YYYY">2024-01-01</style>'],
]);

$xlsx->saveAs(basename(__FILE__, ".php") . '.xlsx');
