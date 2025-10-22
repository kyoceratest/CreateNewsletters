<?php
declare(strict_types=1);

// Basic CORS + JSON headers (no behavior change)
header('Content-Type: application/json; charset=utf-8');
header('Access-Control-Allow-Origin: *');
header('Access-Control-Allow-Methods: POST, OPTIONS');
header('Access-Control-Allow-Headers: Content-Type');

// Preflight support
if (($_SERVER['REQUEST_METHOD'] ?? '') === 'OPTIONS') {
    http_response_code(204);
    exit;
}

if (($_SERVER['REQUEST_METHOD'] ?? '') !== 'POST') {
    http_response_code(405);
    echo json_encode(['error' => 'Method not allowed'], JSON_UNESCAPED_UNICODE | JSON_UNESCAPED_SLASHES);
    exit;
}

// Read and validate payload
$input = file_get_contents('php://input') ?: '';
$data = json_decode($input, true);
if (!is_array($data) || !isset($data['fileName']) || !isset($data['content'])) {
    http_response_code(400);
    echo json_encode(['error' => 'Missing fileName or content'], JSON_UNESCAPED_UNICODE | JSON_UNESCAPED_SLASHES);
    exit;
}

$fileName = (string)$data['fileName'];
$content  = (string)$data['content'];

// Sanitize filename (keep alnum, underscore, hyphen)
$fileName = preg_replace('/[^a-zA-Z0-9_\-]/', '', $fileName) ?? '';
if ($fileName === '') {
    $fileName = 'newsletter_' . date('Y-m-d_H-i-s');
}

// Ensure .html extension
if (!str_ends_with($fileName, '.html')) {
    $fileName .= '.html';
}

// Target directory (unchanged)
$targetDir = 'C:\\Newsletter\\';

// Ensure directory exists
if (!is_dir($targetDir) && !mkdir($targetDir, 0777, true)) {
    http_response_code(500);
    echo json_encode(['error' => 'Could not create directory'], JSON_UNESCAPED_UNICODE | JSON_UNESCAPED_SLASHES);
    exit;
}

$filePath = $targetDir . $fileName;

// Save file
if (file_put_contents($filePath, $content) !== false) {
    echo json_encode([
        'success'  => true,
        'message'  => 'File saved successfully',
        'fileName' => $fileName,
        'filePath' => $filePath,
    ], JSON_UNESCAPED_UNICODE | JSON_UNESCAPED_SLASHES);
    exit;
}

http_response_code(500);
echo json_encode(['error' => 'Could not save file'], JSON_UNESCAPED_UNICODE | JSON_UNESCAPED_SLASHES);
