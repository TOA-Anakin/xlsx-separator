<?php

require __DIR__ . '/../vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Spreadsheet;

$inputFileName = __DIR__ . '/../input/INDIHU_PROD_DATA_22_6_2023_DN.xlsx';
$outputDirPath = __DIR__ . '/../output/';

foreach (glob("$outputDirPath*.xlsx") as $filename) {
    unlink($filename);
}

$spreadsheet = IOFactory::load($inputFileName);
$worksheet = $spreadsheet->getActiveSheet();

$header = $worksheet->rangeToArray('A1:' . $worksheet->getHighestColumn() . '1', null, true, true, true)[1];
$userData = [];

foreach ($worksheet->getRowIterator(2) as $row) {
    $cellIterator = $row->getCellIterator();
    $cellIterator->setIterateOnlyExistingCells(false);

    $rowData = [];
    foreach ($cellIterator as $cell) {
        $rowData[] = $cell->getValue();
    }

    $user = $rowData[4];
    if ($user) {
        if (!isset($userData[$user])) {
            $userData[$user] = [];
        }
        $userData[$user][] = $rowData;
    }
}

$newFilesCount = 0;
foreach ($userData as $user => $data) {
    $newSpreadsheet = new Spreadsheet();
    $newWorksheet = $newSpreadsheet->getActiveSheet();

    $newWorksheet->fromArray($header, null, 'A1');

    $newWorksheet->fromArray($data, null, 'A2');

    $writer = IOFactory::createWriter($newSpreadsheet, 'Xlsx');
    $emailName = str_replace('@', '[[at]]', $user);
    $outputFileName = "$emailName.xlsx";
    $writer->save($outputDirPath . $outputFileName);

    $newFilesCount++;
}

echo "Files have been created successfully: $newFilesCount\n";
