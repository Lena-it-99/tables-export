<?php
require 'vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Style\NumberFormat;

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
        'جوال العميل', 'ارتفاع الظهرية', 'ارتفاع البوكس', 'ارتفاع الأرجل', 
        'ارتفاع المرتبة', 'الأكواب الإضافية', 'الألوان', 'ملاحظات اضافية', 
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
                $newSheet->setCellValueByColumnAndRow($newCol, $row, $value);

                // Set date format for columns 6 and 19
                if ($newCol == 6 || $newCol == 19) {
                    $newSheet->getStyleByColumnAndRow($newCol, $row)->getNumberFormat()->setFormatCode(NumberFormat::FORMAT_DATE_DDMMYYYY);
                }
            }
        }
    }

    // Set column widths to fit content
    $newSheet->getColumnDimension('F')->setAutoSize(true); // Column 6 (F)
    $newSheet->getColumnDimension('S')->setAutoSize(true); // Column 19 (S)

    // Save the new file
    $outputDir = 'C:\\Users\\lena9\\OneDrive\\Desktop\\InHouse\\برنامج تحويل الجداول\\exports';
    if (!file_exists($outputDir)) {
        mkdir($outputDir, 0777, true);
    }
    $outputFile = $outputDir . '\\new_file.xlsx';
    $writer = new Xlsx($newSpreadsheet);
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