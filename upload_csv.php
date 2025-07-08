<?php
require 'vendor/autoload.php';
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

// Database
$host = 'localhost'; $dbname = 'image_to_excel_db'; $username = 'root'; $password = '';

header('Content-Type: application/json');
$data = json_decode(file_get_contents('php://input'), true);

if (!$data || !isset($data['filename']) || !isset($data['text'])) {
    echo json_encode(['success' => false, 'message' => 'Invalid input']);
    exit;
}

$filename = preg_replace('/[^a-zA-Z0-9_\-\.]/', '_', $data['filename']);
$text = $data['text'];

try {
    $pdo = new PDO("mysql:host=$host;dbname=$dbname;charset=utf8mb4", $username, $password);
    $pdo->setAttribute(PDO::ATTR_ERRMODE, PDO::ERRMODE_EXCEPTION);
    $stmt = $pdo->prepare("INSERT INTO uploads (filename, extracted_text) VALUES (?, ?)");
    $stmt->execute([$filename, $text]);

    $spreadsheet = new Spreadsheet();
    $sheet = $spreadsheet->getActiveSheet();
    $lines = preg_split('/\r\n|\r|\n/', $text);
    foreach ($lines as $rowNum => $line) {
        if (!$line) continue;
        $columns = explode("\t", $line);
        foreach ($columns as $colNum => $value) {
            $colLetter = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::stringFromColumnIndex($colNum + 1);
            $sheet->setCellValue($colLetter . ($rowNum + 1), $value);
        }
    }
    if (!is_dir(__DIR__ . '/excel_files')) mkdir(__DIR__ . '/excel_files', 0777, true);
    $excelFilename = 'output_' . time() . '.xlsx';
    $savePath = __DIR__ . '/excel_files/' . $excelFilename;
    $writer = new Xlsx($spreadsheet);
    $writer->save($savePath);

    echo json_encode(['success' => true, 'file_url' => 'excel_files/' . $excelFilename]);
} catch (Exception $e) {
    echo json_encode(['success' => false, 'message' => $e->getMessage()]);
}
