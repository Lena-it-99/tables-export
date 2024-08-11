<?php
require 'vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Shared\Date;

if ($_SERVER['REQUEST_METHOD'] == 'POST' && isset($_FILES['file'])) {
    $file = $_FILES['file']['tmp_name'];

    // اتصال بقاعدة البيانات
    $servername = "localhost";
    $username = "root"; // اسم المستخدم الافتراضي لـ XAMPP
    $password = ""; // كلمة المرور الافتراضية لـ XAMPP
    $dbname = "orders_db";

    $conn = new mysqli($servername, $username, $password, $dbname);

    // تحقق من الاتصال
    if ($conn->connect_error) {
        die("Connection failed: " . $conn->connect_error);
    }

    // تحميل الملف المرفوع
    $spreadsheet = IOFactory::load($file);
    $sheet = $spreadsheet->getActiveSheet();

    // قراءة البيانات من الملف القديم وإدخالها في الجدول الجديد
    $highestRow = $sheet->getHighestRow();
    $mapping = [
        1 => 2, 2 => 3, 3 => 4, 4 => 5, 5 => 6, 6 => 8, 7 => 9,
        8 => 10, 9 => 12, 10 => 13, 11 => 14, 12 => 15, 13 => 16, 
        14 => 17, 15 => 18, 16 => 19, 17 => 20, 18 => null, 19 => 22, 
        20 => null, 21 => 23, 22 => 24, 23 => null, 24 => null, 25 => null, 
        26 => null, 27 => 25
    ];

    for ($row = 2; $row <= $highestRow; $row++) {
        $data = [];
        foreach ($mapping as $newCol => $oldCol) {
            if ($oldCol !== null) {
                $value = $sheet->getCellByColumnAndRow($oldCol, $row)->getValue();
                if ($newCol == 6 || $newCol == 19) {
                    if (Date::isDateTime($sheet->getCellByColumnAndRow($oldCol, $row))) {
                        $value = Date::excelToDateTimeObject($value)->format('Y-m-d');
                    }
                }
                $data[$newCol] = $value;
            } else {
                $data[$newCol] = null;
            }
        }

        $sql = "INSERT INTO orders (order_number, client_name, item, material, quantity, date, barcode, client_phone, back_height, box_height, leg_height, mattress_height, additional_cups, colors, additional_notes, user, total_price, remaining, proposed_delivery_date, city, region, city_district, order_status, address, piece_id, source, delivery_type)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)";

        $stmt = $conn->prepare($sql);
        $stmt->bind_param("ssssisissiiisdsdssisssssss", ...array_values($data));
        $stmt->execute();
    }

    $stmt->close();
    $conn->close();
    echo "Data has been inserted successfully.";
}
?>