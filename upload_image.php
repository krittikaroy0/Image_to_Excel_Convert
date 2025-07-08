<?php
require 'vendor/autoload.php';
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

// Database config
$host = 'localhost';
$dbname = 'image_to_excel_db';
$username = 'root';
$password = '';

// Set JSON header
header('Content-Type: application/json');

// Read JSON input
$data = json_decode(file_get_contents('php://input'), true);

// Validate input
if (!$data || !isset($data['filename']) || !isset($data['text'])) {
    echo json_encode(['success' => false, 'message' => 'Invalid input']);
    exit;
}

$filename = preg_replace('/[^a-zA-Z0-9_\-\.]/', '_', $data['filename']);
$text = $data['text'];

try {
    // Connect DB
    $pdo = new PDO("mysql:host=$host;dbname=$dbname;charset=utf8mb4", $username, $password);
    $pdo->setAttribute(PDO::ATTR_ERRMODE, PDO::ERRMODE_EXCEPTION);

    // Insert into uploads table
    $stmt = $pdo->prepare("INSERT INTO uploads (filename, extracted_text) VALUES (?, ?)");
    $stmt->execute([$filename, $text]);

    // Create Excel
    $spreadsheet = new Spreadsheet();
    $sheet = $spreadsheet->getActiveSheet();

    // Split text lines and columns (tab separated)
    $lines = preg_split('/\r\n|\r|\n/', $text);

    foreach ($lines as $rowNum => $line) {
        if (trim($line) === '') continue;
        $columns = explode("\t", $line);
        foreach ($columns as $colNum => $value) {
            $colLetter = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::stringFromColumnIndex($colNum + 1);
            $sheet->setCellValue($colLetter . ($rowNum + 1), $value);
        }
    }

    // Ensure output dir exists
    $outputDir = __DIR__ . '/excel_files';
    if (!is_dir($outputDir)) mkdir($outputDir, 0777, true);

    $excelFilename = 'output_' . time() . '.xlsx';
    $savePath = $outputDir . '/' . $excelFilename;

    $writer = new Xlsx($spreadsheet);
    $writer->save($savePath);

    // Return JSON with download URL (relative path)
    echo json_encode([
        'success' => true,
        'file_url' => 'excel_files/' . $excelFilename
    ]);
} catch (Exception $e) {
    echo json_encode([
        'success' => false,
        'message' => $e->getMessage()
    ]);
}
