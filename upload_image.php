<?php
declare(strict_types=1);

header('Content-Type: application/json; charset=utf-8');
header('Access-Control-Allow-Origin: *');
header('Access-Control-Allow-Methods: POST, OPTIONS');
header('Access-Control-Allow-Headers: Content-Type');

if (($_SERVER['REQUEST_METHOD'] ?? '') === 'OPTIONS') {
    http_response_code(204);
    exit;
}

if (($_SERVER['REQUEST_METHOD'] ?? '') !== 'POST') {
    http_response_code(405);
    echo json_encode(['error' => 'Method not allowed']);
    exit;
}

if (!isset($_FILES['image']) || !is_array($_FILES['image'])) {
    http_response_code(400);
    echo json_encode(['error' => 'Missing image']);
    exit;
}

$err = (int)($_FILES['image']['error'] ?? UPLOAD_ERR_NO_FILE);
if ($err !== UPLOAD_ERR_OK) {
    http_response_code(400);
    echo json_encode(['error' => 'Upload error', 'code' => $err]);
    exit;
}

$tmpPath = (string)($_FILES['image']['tmp_name'] ?? '');
$origName = (string)($_FILES['image']['name'] ?? '');
if ($tmpPath === '' || !is_uploaded_file($tmpPath)) {
    http_response_code(400);
    echo json_encode(['error' => 'Invalid upload']);
    exit;
}

// Detect MIME using finfo
$finfo = new finfo(FILEINFO_MIME_TYPE);
$mime  = $finfo->file($tmpPath) ?: '';
$allowed = [
    'image/jpeg' => '.jpg',
    'image/png'  => '.png',
    'image/webp' => '.webp',
];
if (!isset($allowed[$mime])) {
    http_response_code(400);
    echo json_encode(['error' => 'Unsupported image type', 'mime' => $mime]);
    exit;
}

// Ensure Image/ directory exists
$targetDir = __DIR__ . DIRECTORY_SEPARATOR . 'Image' . DIRECTORY_SEPARATOR;
if (!is_dir($targetDir) && !mkdir($targetDir, 0777, true)) {
    http_response_code(500);
    echo json_encode(['error' => 'Failed to create Image directory']);
    exit;
}

// Determine requested filename and overwrite policy
$requested = (string)($_POST['filename'] ?? $origName);
$overwrite = isset($_POST['overwrite']) && (string)$_POST['overwrite'] === '1';

// Sanitize requested filename
$requested = preg_replace('/[\\\/:*?"<>|]/', '', $requested) ?: 'image';
$requested = preg_replace('/\s+/', '_', $requested);
$requested = preg_replace('/\.+/', '.', $requested); // collapse multiple dots

// Split stem/ext from requested; enforce extension by MIME
$extByMime = $allowed[$mime];
$stem = $requested;
$dotPos = strrpos($requested, '.');
if ($dotPos !== false) {
    $stem = substr($requested, 0, $dotPos);
}
$stem = preg_replace('/[^a-zA-Z0-9_\-]/', '', $stem) ?: 'image';
$stem = substr($stem, 0, 40);
$finalName = $stem . $extByMime;
$dest = $targetDir . $finalName;

if (!$overwrite) {
    // If not overwriting, ensure uniqueness
    if (file_exists($dest)) {
        $ts = date('Ymd_His');
        $rand = bin2hex(random_bytes(3));
        $finalName = $stem . '_' . $ts . '_' . $rand . $extByMime;
        $dest = $targetDir . $finalName;
    }
} else {
    // Overwrite requested: remove existing file if present
    if (file_exists($dest)) {
        @unlink($dest); // best-effort
    }
}

if (!move_uploaded_file($tmpPath, $dest)) {
    http_response_code(500);
    echo json_encode(['error' => 'Failed to save file']);
    exit;
}

$url = 'Image/' . $finalName;
echo json_encode([
    'success' => true,
    'url' => $url,
]);
