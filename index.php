<?php
require_once(__DIR__ . '/vendor/autoload.php');
use PhpOffice\PhpSpreadsheet\Spreadsheet; 
use PhpOffice\PhpSpreadsheet\Writer\Xlsx; 
# Create a new Xls Reader
$reader = new \PhpOffice\PhpSpreadsheet\Reader\Xlsx();

// Tell the reader to only read the data. Ignore formatting etc.
$reader->setReadDataOnly(true);

// Read the spreadsheet file.
$spreadsheet = $reader->load(__DIR__ . '/VehicleRouteMovement.xlsx');

$vehicleroutemovement = $spreadsheet->getSheet(2);
$vehiclecost = $spreadsheet->getSheet(3);

$vehiclemovementdata = $vehicleroutemovement->rangeToArray('A2:G98');
$vehiclecostdata = $vehiclecost->rangeToArray('A2:C16');
$arraymaintenancebybranch=[];
$arraybranch=[];
// $taken = array();

foreach($vehiclemovementdata as $key => $item) {
   
    if(!in_array($item[6], $arraybranch)&&$item[6]!="") {
        $arraybranch[] = $item[6];
    } 
}
// echo json_encode($arraybranch);

foreach($arraybranch as $branch){
    $currentbranch=$branch;
    $totalmaintenancecost=0;
    $totalmaintenancebycar=0;
    foreach($vehiclecostdata as $vehicle){
        $currentcar=$vehicle[0];
        // echo $currentcar;
        $totalcostcar=0;
        $totalmilage=0;
        foreach($vehiclemovementdata as $currentmovement){
            if($currentmovement[2]==$currentcar&&$currentmovement[6]==$currentbranch){
                $totalmilage=$currentmovement[4]+$totalmilage;
            }
        }
        // echo $totalmilage;
        $totalmaintenancebycar=$totalmilage/$vehicle[1]*$vehicle[2];
        // echo $totalmaintenancecost. ' current car ' . $currentcar;
        $totalmaintenancecost=$totalmaintenancecost+$totalmaintenancebycar;
    }
    $arraymaintenancebybranch[]=array($currentbranch,$totalmaintenancecost);

    // $arraymaintenancebybranch[$currentbranch]=$totalmaintenancecost;
}
$writer = new \PhpOffice\PhpSpreadsheet\Writer\Xlsx($spreadsheet);
$result=$spreadsheet->getSheet(1);
$result->getStyle('B1:B18')->getNumberFormat()->setFormatCode('#,##0.00');

$result->fromArray(
    $arraymaintenancebybranch,
    null,
    'A2'
);
$writer->save("VehicleRouteMovement.xlsx");
echo json_encode($arraymaintenancebybranch);

// output the data to the console, so you can see what there is.
// // echo json_encode($vehiclemovementdata);
// die(print_r($vehiclemovementdata, true));
?>