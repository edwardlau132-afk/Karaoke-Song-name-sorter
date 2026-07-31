<?php
// Search endpoint: reads filter values from the query string, queries the
// selected "songlist" table, and returns matching rows as JSON.
//
// Example request:
//   search.php?Singer_歌手=Jacky&source=new

header('Content-Type: application/json; charset=utf-8');
require 'db_conn.php';

// Whitelist of selectable tables. NEVER build the table name directly from
// user input — always go through this map, so a request can only ever
// target one of these two exact table names.
$tableMap = [
    'old' => 'songlist',
    'new' => 'songlist2',
];

$source = (isset($_GET['source']) && isset($tableMap[$_GET['source']])) ? $_GET['source'] : 'old';
$table  = $tableMap[$source];

// Maps the field names used by the front-end (script.js) to the actual
// column names in the songlist table.
$fieldMap = [
    'Number_編號'          => 'Number',
    'Song_Title_歌名'       => 'SongTitle',
    'Singer_歌手'           => 'Singer',
    'PinYin_国語拼音'       => 'PinYin',
    'Cantonese_粤語拼音'    => 'Cantonese',
    'Word_Number_字數統計'  => 'Words',
];

$where  = [];
$params = [];

foreach ($fieldMap as $frontKey => $column) {
    if (isset($_GET[$frontKey]) && trim($_GET[$frontKey]) !== '') {
        $placeholder = ':' . $column;
        $where[] = "`{$column}` LIKE {$placeholder}";
        $params[$placeholder] = '%' . trim($_GET[$frontKey]) . '%';
    }
}

$sql = "SELECT `Number`, `SongTitle`, `Singer`, `PinYin`, `Cantonese`, `Words` FROM `{$table}`";
if ($where) {
    $sql .= " WHERE " . implode(' AND ', $where);
}
$sql .= " LIMIT 300";

try {
    $stmt = $pdo->prepare($sql);
    $stmt->execute($params);
    $rows = $stmt->fetchAll();
} catch (PDOException $e) {
    http_response_code(500);
    echo json_encode(["error" => "Query failed"]);
    exit;
}

// Translate DB column names back to the front-end field keys expected by script.js
$out = array_map(function ($row) use ($fieldMap) {
    $obj = [];
    foreach ($fieldMap as $frontKey => $column) {
        $obj[$frontKey] = $row[$column] ?? '';
    }
    return $obj;
}, $rows);

echo json_encode($out, JSON_UNESCAPED_UNICODE);
