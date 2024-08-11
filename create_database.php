<?php
$servername = "localhost";
$username = "root"; // اسم المستخدم الافتراضي لـ XAMPP
$password = ""; // كلمة المرور الافتراضية لـ XAMPP
$dbname = "orders_db";

// إنشاء اتصال بقاعدة البيانات
$conn = new mysqli($servername, $username, $password);

// التحقق من الاتصال
if ($conn->connect_error) {
    die("Connection failed: " . $conn->connect_error);
}

// إنشاء قاعدة البيانات
$sql = "CREATE DATABASE IF NOT EXISTS $dbname";
if ($conn->query($sql) === TRUE) {
    echo "Database created successfully\n";
} else {
    echo "Error creating database: " . $conn->error . "\n";
}

// تحديد قاعدة البيانات النشطة
$conn->select_db($dbname);

// إنشاء جدول الطلبات
$sql = "CREATE TABLE IF NOT EXISTS orders (
    id INT(11) AUTO_INCREMENT PRIMARY KEY,
    order_number VARCHAR(255),
    client_name VARCHAR(255),
    item VARCHAR(255),
    material VARCHAR(255),
    quantity INT,
    date DATE,
    barcode VARCHAR(255),
    client_phone VARCHAR(255),
    back_height INT,
    box_height INT,
    leg_height INT,
    mattress_height INT,
    additional_cups INT,
    colors VARCHAR(255),
    additional_notes TEXT,
    user VARCHAR(255),
    total_price DECIMAL(10,2),
    remaining DECIMAL(10,2),
    proposed_delivery_date DATE,
    city VARCHAR(255),
    region VARCHAR(255),
    city_district VARCHAR(255),
    order_status VARCHAR(255),
    address TEXT,
    piece_id VARCHAR(255),
    source VARCHAR(255),
    delivery_type VARCHAR(255)
)";

if ($conn->query($sql) === TRUE) {
    echo "Table created successfully\n";
} else {
    echo "Error creating table: " . $conn->error . "\n";
}

// إغلاق الاتصال بقاعدة البيانات
$conn->close();
?>