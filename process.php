<?php
require 'vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Csv;
use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Shared\Date;

if ($_SERVER['REQUEST_METHOD'] == 'POST' && isset($_FILES['file'])) {
    $file = $_FILES['file']['tmp_name'];

    // Load the uploaded file
    $spreadsheet = IOFactory::load($file);
    $sheet = $spreadsheet->getActiveSheet();

    // Create a new Spreadsheet
    $newSpreadsheet = new Spreadsheet();
    $newSheet = $newSpreadsheet->getActiveSheet();

    // Set the sheet to right-to-left
    $newSheet->setRightToLeft(true);

    // Column titles
    $titles = [
        'الرقم', 'العميل', 'الصنف', 'المادة', 'الكمية', 'التاريخ', 'الباركود',
        'جوال العميل', 'ارتفاع الظهرية', 'ارتفاع البوكس', 'ارتفاع الارجل', 
        'ارتفاع المرتبة', 'الاكواب الاضافية', 'الالوان', 'ملاحظات إضافية', 
        'المستخدم', 'السعر الإجمالي', 'الباقي', 'موعد التسليم المقترح', 
        'المدينة', 'المنطقة', 'المدينة-الحي', 'حالة الطلب', 'العنوان', 
        'معرف القطعة', 'المصدر', 'نوع التوصيل'
    ];

    // Set column titles
    foreach ($titles as $col => $title) {
        $newSheet->setCellValueByColumnAndRow($col + 1, 1, $title);
        $newSheet->getStyleByColumnAndRow($col + 1, 1)->getFont()->setBold(true);
    }

    // Read the data from the old file and insert it into the new file
    $highestRow = $sheet->getHighestRow();
    $mapping = [
        1 => 2, 2 => 3, 3 => 4, 4 => 5, 5 => 6, 6 => 8, 7 => 9,
        8 => 10, 9 => 12, 10 => 13, 11 => 14, 12 => 15, 13 => 16, 
        14 => 17, 15 => 18, 16 => 19, 17 => 20, 18 => null, 19 => 22, 
        20 => null, 21 => 23, 22 => 24, 23 => null, 24 => null, 25 => null, 
        26 => null, 27 => 25
    ];

    for ($row = 2; $row <= $highestRow; $row++) {
        foreach ($mapping as $newCol => $oldCol) {
            if ($oldCol !== null) {
                $value = $sheet->getCellByColumnAndRow($oldCol, $row)->getValue();
                if ($newCol == 6 || $newCol == 19) {
                    if (Date::isDateTime($sheet->getCellByColumnAndRow($oldCol, $row))) {
                        $value = Date::excelToDateTimeObject($value);
                        $newSheet->setCellValueByColumnAndRow($newCol, $row, $value->format('d/m/Y'));
                        $newSheet->getStyleByColumnAndRow($newCol, $row)->getNumberFormat()->setFormatCode('dd/mm/yyyy');
                    }
                } else {
                    $newSheet->setCellValueByColumnAndRow($newCol, $row, $value);
                }
            }
        }
    }

    // Adjust column width to fit the content
    foreach (range('A', $newSheet->getHighestColumn()) as $columnID) {
        $newSheet->getColumnDimension($columnID)->setWidth(15); // Set a fixed width for all columns
    }

    $newSheet->getColumnDimension('F')->setWidth(20); // Set a specific width for column 6
    $newSheet->getColumnDimension('S')->setWidth(20); // Set a specific width for column 19

    // Save the new file as CSV
    $writer = new Csv($newSpreadsheet);
    $outputFile = 'C:/Users/lena9/OneDrive/Desktop/InHouse/export_program/exports/new_file.csv';
    $writer->save($outputFile);

    // Download the new file
    header('Content-Description: File Transfer');
    header('Content-Type: application/octet-stream');
    header('Content-Disposition: attachment; filename="' . basename($outputFile) . '"');
    header('Expires: 0');
    header('Cache-Control: must-revalidate');
    header('Pragma: public');
    header('Content-Length: ' . filesize($outputFile));
    readfile($outputFile);
    exit;
}
?>