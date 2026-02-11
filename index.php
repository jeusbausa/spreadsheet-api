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

$ch = curl_init("http://localhost:3000/api/export");

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


foreach ($templateGroups as $templateKey => $groups) {

    if ($templateKey === "template_2") {

        $spreadsheet = IOFactory::load(__DIR__ . "/templates/excel/4-ccl-p2.xlsx");
        $templateSheet = $spreadsheet->getSheet(0);
        $templateTitle = $templateSheet->getTitle();

        $count = count($groups);

        for ($i = 1; $i <= $count; $i++) {

            $newSheet = $spreadsheet->duplicateWorksheetByTitle($templateTitle);
            $rowCluster = 4;

            foreach ($groups["group_" . $i] as $_i => $cluster) {

                $staffName = strtoupper($cluster["assignedStaff"]["firstName"] . " " . $cluster["assignedStaff"]["lastName"]);
                $clusterCode = strtoupper($cluster["assignedStaff"]["codeName"] . "-" . $cluster["code"]);

                $clientStartRow = $rowCluster + 3;

                $newSheet->setCellValue("C{$rowCluster}", $staffName);
                $newSheet->SetCellValue("G{$rowCluster}", date_format(date_create($cluster["dateOfRelease"]), "F j, Y"));
                $newSheet->SetCellValue("K{$rowCluster}", date("F j, Y"));

                $newSheet->setCellValue("C" . ($rowCluster + 1), $clusterCode);
                $newSheet->SetCellValue("G" . ($rowCluster + 1), date_format(date_create($cluster["dateOfFirstPayment"]), "F j, Y"));
                $newSheet->SetCellValue("K" . ($rowCluster + 1), str_replace("WEEKS_", "", $cluster["loanTerm"]) . " Weeks");

                $row = $clientStartRow;

                foreach ($cluster["clients"] ?? [] as $index => $client) {

                    $newSheet->setCellValue("A{$row}", $index + 1);
                    $newSheet->setCellValue("B{$row}", ($client["client"]["firstName"] . " " . $client["client"]["middleName"] . " " . $client["client"]["lastName"]));
                    $newSheet->setCellValue("C{$row}", $client["loanReceivable"] ?? 0);
                    $newSheet->setCellValue("D{$row}", $client["skCumulative"] ?? 0);
                    $newSheet->setCellValue("E{$row}", $client["pastDue"] ?? 0);
                    $newSheet->setCellValue("G{$row}", $client["weeklyInstallment"] ?? 0);

                    $row++;
                }

                $rowCluster += 10;
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

    if ($templateKey === "template_10") {

        $spreadsheet = IOFactory::load(__DIR__ . "/templates/excel/2-ccl-p10.xlsx");
        $templateSheet = $spreadsheet->getSheet(0);
        $templateTitle = $templateSheet->getTitle();

        $count = count($groups);

        for ($i = 1; $i <= $count; $i++) {

            $newSheet = $spreadsheet->duplicateWorksheetByTitle($templateTitle);
            $rowCluster = 4;

            foreach ($groups["group_" . $i] as $_i => $cluster) {

                $staffName = strtoupper($cluster["assignedStaff"]["firstName"] . " " . $cluster["assignedStaff"]["lastName"]);
                $clusterCode = strtoupper($cluster["assignedStaff"]["codeName"] . "-" . $cluster["code"]);

                $clientStartRow = $rowCluster + 3;

                $newSheet->setCellValue("C{$rowCluster}", $staffName);
                $newSheet->SetCellValue("G{$rowCluster}", date_format(date_create($cluster["dateOfRelease"]), "F j, Y"));
                $newSheet->SetCellValue("K{$rowCluster}", date("F j, Y"));

                $newSheet->setCellValue("C" . ($rowCluster + 1), $clusterCode);
                $newSheet->SetCellValue("G" . ($rowCluster + 1), date_format(date_create($cluster["dateOfFirstPayment"]), "F j, Y"));
                $newSheet->SetCellValue("K" . ($rowCluster + 1), str_replace("WEEKS_", "", $cluster["loanTerm"]) . " Weeks");


                $row = $clientStartRow;

                foreach ($cluster["clients"] ?? [] as $index => $client) {

                    $newSheet->setCellValue("A{$row}", $index + 1);
                    $newSheet->setCellValue("B{$row}", ($client["client"]["firstName"] . " " . $client["client"]["middleName"] . " " . $client["client"]["lastName"]));
                    $newSheet->setCellValue("C{$row}", $client["loanReceivable"] ?? 0);
                    $newSheet->setCellValue("D{$row}", $client["skCumulative"] ?? 0);
                    $newSheet->setCellValue("E{$row}", $client["pastDue"] ?? 0);
                    $newSheet->setCellValue("G{$row}", $client["weeklyInstallment"] ?? 0);

                    $row++;
                }

                $rowCluster += 18;
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

    if ($templateKey === "template_15") {

        $spreadsheet = IOFactory::load(__DIR__ . "/templates/excel/2-ccl-p15.xlsx");
        $templateSheet = $spreadsheet->getSheet(0);
        $templateTitle = $templateSheet->getTitle();

        $count = count($groups);

        for ($i = 1; $i <= $count; $i++) {

            $newSheet = $spreadsheet->duplicateWorksheetByTitle($templateTitle);
            $rowCluster = 4;

            foreach ($groups["group_" . $i] as $_i => $cluster) {

                $staffName = strtoupper($cluster["assignedStaff"]["firstName"] . " " . $cluster["assignedStaff"]["lastName"]);
                $clusterCode = strtoupper($cluster["assignedStaff"]["codeName"] . "-" . $cluster["code"]);

                $clientStartRow = $rowCluster + 3;

                $newSheet->setCellValue("C{$rowCluster}", $staffName);
                $newSheet->SetCellValue("G{$rowCluster}", date_format(date_create($cluster["dateOfRelease"]), "F j, Y"));
                $newSheet->SetCellValue("K{$rowCluster}", date("F j, Y"));

                $newSheet->setCellValue("C" . ($rowCluster + 1), $clusterCode);
                $newSheet->SetCellValue("G" . ($rowCluster + 1), date_format(date_create($cluster["dateOfFirstPayment"]), "F j, Y"));
                $newSheet->SetCellValue("K" . ($rowCluster + 1), str_replace("WEEKS_", "", $cluster["loanTerm"]) . " Weeks");


                $row = $clientStartRow;

                foreach ($cluster["clients"] ?? [] as $index => $client) {

                    $newSheet->setCellValue("A{$row}", $index + 1);
                    $newSheet->setCellValue("B{$row}", ($client["client"]["firstName"] . " " . $client["client"]["middleName"] . " " . $client["client"]["lastName"]));
                    $newSheet->setCellValue("C{$row}", $client["loanReceivable"] ?? 0);
                    $newSheet->setCellValue("D{$row}", $client["skCumulative"] ?? 0);
                    $newSheet->setCellValue("E{$row}", $client["pastDue"] ?? 0);
                    $newSheet->setCellValue("G{$row}", $client["weeklyInstallment"] ?? 0);

                    $row++;
                }

                $rowCluster += 22;
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

    if ($templateKey === "template_20") {

        $spreadsheet = IOFactory::load(__DIR__ . "/templates/excel/2-ccl-p20.xlsx");
        $templateSheet = $spreadsheet->getSheet(0);
        $templateTitle = $templateSheet->getTitle();

        $count = count($groups);

        for ($i = 1; $i <= $count; $i++) {

            $newSheet = $spreadsheet->duplicateWorksheetByTitle($templateTitle);
            $rowCluster = 4;

            foreach ($groups["group_" . $i] as $_i => $cluster) {

                $staffName = strtoupper($cluster["assignedStaff"]["firstName"] . " " . $cluster["assignedStaff"]["lastName"]);
                $clusterCode = strtoupper($cluster["assignedStaff"]["codeName"] . "-" . $cluster["code"]);

                $clientStartRow = $rowCluster + 3;

                $newSheet->setCellValue("C{$rowCluster}", $staffName);
                $newSheet->SetCellValue("G{$rowCluster}", date_format(date_create($cluster["dateOfRelease"]), "F j, Y"));
                $newSheet->SetCellValue("K{$rowCluster}", date("F j, Y"));

                $newSheet->setCellValue("C" . ($rowCluster + 1), $clusterCode);
                $newSheet->SetCellValue("G" . ($rowCluster + 1), date_format(date_create($cluster["dateOfFirstPayment"]), "F j, Y"));
                $newSheet->SetCellValue("K" . ($rowCluster + 1), str_replace("WEEKS_", "", $cluster["loanTerm"]) . " Weeks");


                $row = $clientStartRow;

                foreach ($cluster["clients"] ?? [] as $index => $client) {

                    $newSheet->setCellValue("A{$row}", $index + 1);
                    $newSheet->setCellValue("B{$row}", ($client["client"]["firstName"] . " " . $client["client"]["middleName"] . " " . $client["client"]["lastName"]));
                    $newSheet->setCellValue("C{$row}", $client["loanReceivable"] ?? 0);
                    $newSheet->setCellValue("D{$row}", $client["skCumulative"] ?? 0);
                    $newSheet->setCellValue("E{$row}", $client["pastDue"] ?? 0);
                    $newSheet->setCellValue("G{$row}", $client["weeklyInstallment"] ?? 0);

                    $row++;
                }

                $rowCluster += 27;
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
