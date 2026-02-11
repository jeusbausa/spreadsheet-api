<?php

declare(strict_types=1);

require __DIR__ . "/vendor/autoload.php";

use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Spreadsheet;

header("Access-Control-Allow-Origin: *");
header("Access-Control-Allow-Methods: POST, OPTIONS");
header("Access-Control-Allow-Headers: Content-Type, Authorization");

if ($_SERVER["REQUEST_METHOD"] === "OPTIONS") {
    http_response_code(200);
    exit;
}

if ($_SERVER["REQUEST_METHOD"] !== "POST") {
    header("Content-Type: application/json");
    http_response_code(405);
    echo json_encode(["message" => "Method Not Allowed"]);
    exit;
}

$baseUrl = getenv("APP_URL") ?: "http://localhost:3000";

$clusterIds = isset($_GET["clusterIds"]) ? $_GET["clusterIds"] : null;

$url = "{$baseUrl}/api/export?clusterIds={$clusterIds}";

$ch = curl_init($url);

curl_setopt_array($ch, [
    CURLOPT_POST => true,
    CURLOPT_RETURNTRANSFER => true,
    CURLOPT_HTTPHEADER => [
        "Content-Type: application/json",
        "Accept: application/json"
    ]
]);

$response = curl_exec($ch);
$status   = curl_getinfo($ch, CURLINFO_HTTP_CODE);
curl_close($ch);

if ($status !== 200 || !$response) {
    header("Content-Type: application/json");
    echo json_encode(["message" => "API error"]);
    exit;
}

$apiResponse = json_decode($response, true);

if (!isset($apiResponse["data"]["clusters"])) {
    header("Content-Type: application/json");
    echo json_encode(["message" => "Invalid response structure"]);
    exit;
}

$templateGroups = $apiResponse["data"]["clusters"];

$finalSpreadsheet = new Spreadsheet();
$finalSpreadsheet->removeSheetByIndex(0);

$templateConfig = [
    "template_2"  => ["path" => "/templates/excel/4-ccl-p2.xlsx",  "spacing" => 10],
    "template_10" => ["path" => "/templates/excel/2-ccl-p10.xlsx", "spacing" => 18],
    "template_15" => ["path" => "/templates/excel/2-ccl-p15.xlsx", "spacing" => 22],
    "template_20" => ["path" => "/templates/excel/2-ccl-p20.xlsx", "spacing" => 27],
];

foreach ($templateGroups as $templateKey => $groups) {

    if (!isset($templateConfig[$templateKey])) {
        continue;
    }

    $spreadsheet = IOFactory::load(__DIR__ . $templateConfig[$templateKey]["path"]);
    $templateSheet = $spreadsheet->getSheet(0);
    $templateTitle = $templateSheet->getTitle();
    $spacing = $templateConfig[$templateKey]["spacing"];

    $count = count($groups);

    for ($i = 1; $i <= $count; $i++) {

        $newSheet = $spreadsheet->duplicateWorksheetByTitle($templateTitle);
        $rowCluster = 4;

        foreach ($groups["group_" . $i] as $_i => $cluster) {

            $staffName = strtoupper(
                $cluster["assignedStaff"]["firstName"] . " " .
                    $cluster["assignedStaff"]["lastName"]
            );

            $clusterCode = strtoupper(
                $cluster["assignedStaff"]["codeName"] . "-" .
                    $cluster["code"]
            );

            $clientStartRow = $rowCluster + 3;

            $newSheet->setCellValue("C{$rowCluster}", $staffName);
            $newSheet->setCellValue("G{$rowCluster}", date_format(date_create($cluster["dateOfRelease"]), "F j, Y"));
            $newSheet->setCellValue("K{$rowCluster}", date("F j, Y"));

            $newSheet->setCellValue("C" . ($rowCluster + 1), $clusterCode);
            $newSheet->setCellValue("G" . ($rowCluster + 1), date_format(date_create($cluster["dateOfFirstPayment"]), "F j, Y"));
            $newSheet->setCellValue("K" . ($rowCluster + 1), str_replace("WEEKS_", "", $cluster["loanTerm"]) . " Weeks");

            $row = $clientStartRow;

            foreach ($cluster["clients"] ?? [] as $index => $client) {

                $newSheet->setCellValue("A{$row}", $index + 1);
                $newSheet->setCellValue(
                    "B{$row}",
                    $client["client"]["firstName"] . " " .
                        $client["client"]["middleName"] . " " .
                        $client["client"]["lastName"]
                );
                $newSheet->setCellValue("C{$row}", $client["loanReceivable"] ?? 0);
                $newSheet->setCellValue("D{$row}", $client["skCumulative"] ?? 0);
                $newSheet->setCellValue("E{$row}", $client["pastDue"] ?? 0);
                $newSheet->setCellValue("G{$row}", $client["weeklyInstallment"] ?? 0);

                $row++;
            }

            $rowCluster += $spacing;
        }
    }

    $spreadsheet->removeSheetByIndex(0);

    foreach ($spreadsheet->getAllSheets() as $sheet) {

        $originalTitle = $sheet->getTitle();
        $newTitle = $originalTitle;
        $counter = 1;

        while ($finalSpreadsheet->sheetNameExists($newTitle)) {
            $newTitle = $originalTitle . " " . $counter;
            $counter++;
        }

        $sheet->setTitle($newTitle);
        $finalSpreadsheet->addExternalSheet($sheet);
    }
}

if ($finalSpreadsheet->getSheetCount() === 0) {
    $finalSpreadsheet->createSheet()->setTitle("EMPTY");
}

header("Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
header("Content-Disposition: attachment; filename='export.xlsx'");
header("Cache-Control: max-age=0");

$writer = IOFactory::createWriter($finalSpreadsheet, "Xlsx");
$writer->save("php://output");
exit;
