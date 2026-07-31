<?php
// Connection to the live InfinityFree MySQL database.
// This file must be uploaded to InfinityFree and run FROM their server —
// it will not work if run from your own computer, since InfinityFree
// blocks external/remote MySQL connections on the free plan.

$host = "sql108.infinityfree.com";
$db   = "if0_42494692_songlist";
$user = "if0_42494692";
$pass = "Tel82apo";

try {
    $pdo = new PDO(
        "mysql:host={$host};dbname={$db};charset=utf8mb4",
        $user,
        $pass,
        [
            PDO::ATTR_ERRMODE            => PDO::ERRMODE_EXCEPTION,
            PDO::ATTR_DEFAULT_FETCH_MODE => PDO::FETCH_ASSOC,
        ]
    );
} catch (PDOException $e) {
    http_response_code(500);
    header('Content-Type: application/json');
    echo json_encode(["error" => "Database connection failed"]);
    exit;
}
