<?php
session_start();

ini_set('memory_limit', '800M');
ini_set('max_execution_time', 600);
ini_set("display_errors", 1);
error_reporting(E_ALL);

require_once 'phpexcel/PHPExcel/IOFactory.php';

include_once('Functions/DBConnection.php');
include_once('Functions/ProcessFile.php');
include_once('Functions/ValidateFields.php');
include_once('Functions/LookupData.php');

$processFunc    = new ProcessFile();
$validateFunc   = new ValidateFields();
$lookupData     = new LookupData();
$conn           = new DBConnection();

/* normalize POST */
$_POST = array_change_key_case($_POST, CASE_LOWER);

/* 1. GET TOKEN (VERY IMPORTANT) */
$token = $_SESSION['current_flow_token'];

if (!isset($_SESSION['flows'][$token])) {
    echo "<script>
            alert('Your session has expired. Please log in again to continue.');
            window.location.href='http://testweb.intercommerce.com.ph/login.asp';
          </script>";
    exit;
}

/* 2. LOAD DATA FOR THIS TAB ONLY */
$flow = $_SESSION['flows'][$token];

$csncod         = $flow['csncod'];
$locTin         = $flow['loctin'];
$zoneCode       = $flow['zonecode'];
$ptopsTin       = $flow['ptopstin'];
$enterpriseType = $flow['enterprisetype'];
$compNam        = $flow['compnam'];
$userID         = $flow['userid'];
$lstexporter    = $flow['lstexporter'];
$locbroktin     = $flow['locbroktin'];
$loccod         = $flow['loccod'];
$allaccids      = $flow['allaccids'];
$mod_cod        = $flow['mod_cod'];
$mod_cod2       = $flow['mod_cod2'];
$cltcode        = $flow['cltcode'];
$redirection    = $flow['redirection'];

// note: mag-add nalang ng validation sa $accountType 
// pag inimplement na din itong excel uploading sa forwarder account
$accountType = "exporter";

// If csncode Missing then go back to step 1 page
if (!isset($flow['csncod']) || empty($flow['csncod'])) {
    echo "<script>
            alert('Session expired. Please reselect Forwarder Name.');
            window.location.href='http://testweb.intercommerce.com.ph/webcws/'.$redirection.'.asp';
          </script>";
    exit;
}

//CHECK FILE FORMAT/TYPE
$fileFormat = $processFunc->__checkFileFormat($_FILES["file"]["type"]);

//UPLOAD FILE
if($fileFormat) {

    $uploadInDirectory = $processFunc->__uploadFileInDirectory($_FILES['file']['name']); 

    if(!$uploadInDirectory){
		
        echo "<script>
                alert('Failed uploading the excel file. Please try again');
                window.location.href='index.php?token=$token';
            </script>";

    }

} else {
    //INVALID FILE 
    echo "<script>
            alert('Invalid File Type. Upload Excel File.');
            window.location.href='index.php?token=$token';
        </script>";

}

//GET FILE
$excelDetails = $processFunc->__getPHPExcelDetails($_FILES['file']['name']);

    //START PHPEXCEL
    $objReader      = PHPExcel_IOFactory::createReader($excelDetails["type"]);
    $objReader->setReadDataOnly(true);
    $objPHPExcel    = $objReader->load($excelDetails["inputFile"]);

    $objWorksheet   = $objPHPExcel->getSheetByName('General');
    $objWorksheet2  = $objPHPExcel->getSheetByName('CONTAINER SEAL NO');
    $objWorksheet3  = $objPHPExcel->getSheetByName('Items');
    $objWorksheet4  = $objPHPExcel->getSheetByName('Financial');

    $errorCounter = 0; //REQUIRED TO BE CORRECTED
    $errorCounter1 = 0; //PROCEED EVENTHOUGH NOT CORRECTED
    $errorLists = array();

    //SCAN EXCEL FILE 
    if($objWorksheet){

        $highestRow     = $objWorksheet->getHighestRow();
        $highestColumn  = $objWorksheet->getHighestColumn();
        $headingsArray  = $objWorksheet->rangeToArray('A1:'.$highestColumn.'1',null, true, true, true);
        
        $headingsArray  = $headingsArray[1];
        $r = -1;

        if(strtoupper($highestColumn) != 'L') {

            echo "<script>
                    alert('File content is not compatible. Please check the General sheet');
                    window.location.href='index.php?token=$token';
                </script>";

            unlink($excelDetails['inputFile']);
            die();
        }

        if ($highestRow > 2) {

            echo "<script>
                    alert('Only one record is allowed in General sheet.');
                    window.location.href='index.php?token=$token';
                </script>";

            unlink($excelDetails['inputFile']);
            die();
        }

        $checkItemsA2 = $objPHPExcel->getActiveSheet()->getCell('A2')->getValue();
        if($checkItemsA2 == NULL || $checkItemsA2 == '') {

            echo "<script>
                alert('Cannot proceed. Please check file.');
                window.location.href='index.php?token=$token';
            </script>";

            unlink($excelDetails['inputFile']);
            die();
            
        }else{

            for ($row = 2; $row <= $highestRow; ++$row) {

                $dataRow = $objWorksheet->rangeToArray('A'.$row.':'.$highestColumn.$row, null, true, true, true);

                //CHECK IF ALL CELLS ARE EMPTY
                if( $validateFunc->isEmptyRow(reset($dataRow)) ) 
                { 
                    continue; //skip empty row
                } 

                    ++$r;
                    
                    foreach($headingsArray as $columnKey => $columnHeading) {

                        //VALUES
                        $Consignee              =   $validateFunc->trim_val($dataRow[$row]['A']);
                        $Address                =   $validateFunc->trim_val($dataRow[$row]['B']);
                        $Port                   =   $validateFunc->trim_val($dataRow[$row]['C']);
                        $PurposeOfExportation   =   $validateFunc->trim_val($dataRow[$row]['D']);
                        $ManifestNo             =   $validateFunc->trim_val($dataRow[$row]['E']);
                        $BillOfLading           =   $validateFunc->trim_val($dataRow[$row]['F']);
                        $VesselAircraft         =   $validateFunc->trim_val2($dataRow[$row]['G']);
                        $LocationOfGoods        =   $validateFunc->trim_val2($dataRow[$row]['H']);
                        $ProvinceOfOrigin       =   $validateFunc->trim_val2($dataRow[$row]['I']);
                        $CountryOfDestination   =   $validateFunc->trim_val2($dataRow[$row]['J']);
                        $PortOfLoading          =   $validateFunc->trim_val($dataRow[$row]['K']);
                        $PortOfDeparture        =   $validateFunc->trim_val($dataRow[$row]['L']);
                    }

                /* VALIDATE FIELD VALUES */

                    //Consignee
                    if( $validateFunc->max_length($Consignee, 70) ) 
                    { 
                        $consignee[] = $row - 1; 
                        $errorCounter1++; 
                    }
                    if( !empty($Consignee) &&  ($validateFunc->match_char($Consignee)) == 0 ) 
                    { 
                        $consigneeMatch[] = $row - 1; 
                        $errorCounter++; 
                    }

                    if( !empty($Consignee) ) {

                        $checkConsigneeExists = $validateFunc->__checkValidConsignee($conn, $Consignee, $cltcode);

                        if( !$checkConsigneeExists )
                        {
                            $checkConsignee[] = $row - 1;
                            $errorCounter++;
                        }

                    } else {

                        $checkConsignee[] = $row - 1;
                        $errorCounter++;
                    }

                    //Address
                    if( $validateFunc->max_length($Address, 105) ) 
                    { 
                        $address[] = $row - 1; $errorCounter1++; 
                    }
                    if( !empty($Address) && ($validateFunc->match_char($Address)) == 0 )
                    {
                        $addressMatch[] = $row - 1; 
                        $errorCounter++; 
                    }

                    if( !empty($Address) && !empty($checkConsigneeExists) ) {

                        if ( !$validateFunc->__checkValidAddress($Address, $checkConsigneeExists) )
                        {
                            $checkAddress[] = $row - 1;
                            $errorCounter++;
                        }
                    }

                    //Port (Office of Clearance)
                    if( !empty($Port) )
                    { 
                        //CHECK PORT IF EXISTS
                        $checkOfficeOfClearanceExists = $validateFunc->__checkValidPortOfDeparture($Port); 

                        if( !$checkOfficeOfClearanceExists ) 
                        {
                            $checkOfficeOfClearance[] = $row - 1; 
                            $errorCounter++; 
                        } else {

                            // Case-sensitive check: DB match is case-insensitive by default,
                            // so confirm the Excel value's casing exactly matches offClrCod
                            // as stored in DmOffClr.
                            if ( strcmp($Port, $checkOfficeOfClearanceExists['offClrCod']) !== 0 )
                            {
                                $checkOfficeOfClearanceCase[] = $row - 1;
                                $errorCounter++;
                            }

                            // Container Seal Number Validation Based on Mode of Transport
                            if ($checkOfficeOfClearanceExists['offClrMode'] === "BY AIR")
                            {
                                $checkModeOfTransportation[] = $row - 1; 
                            }
                        }
                    } else {
                        $checkOfficeOfClearance[] = $row - 1; 
                        $errorCounter++; 
                    }

                    //PurposeOfExportation
                    if( !empty($PurposeOfExportation) )
                    { 
                        //CHECK PurposeOfExportation IF EXISTS
                        $checkPurposeOfExportationExists = $validateFunc->__checkValidPurposeOfExportation($PurposeOfExportation); //MODIFY

                        if( !$checkPurposeOfExportationExists ) 
                        {
                            $checkPurposeOfExportation[] = $row - 1; 
                            $errorCounter++; 
                        }
                    } else {
                        $checkPurposeOfExportation[] = $row - 1; 
                        $errorCounter++; 
                    }

                    //ManifestNo
                    if( !empty($ManifestNo) && ($validateFunc->match_manifestFormat($ManifestNo)) == 0 )
                    { 
                        $manifestNoMatch[] = $row - 1; 
                        $errorCounter++; 
                    }

                    //BillOfLading
                    if( $validateFunc->max_length($BillOfLading, 52) ) 
                    { 
                        $billOfLading[] = $row - 1; $errorCounter1++; 
                    }
                    if( !empty($BillOfLading) && ($validateFunc->match_alphanum($BillOfLading)) == 0 )
                    { 
                        $billOfLadingMatch[] = $row - 1; 
                        $errorCounter++; 
                    }

                    /*
                    |--------------------------------------------------------------------------
                    | Additional validation (same as Classic ASP)
                    |--------------------------------------------------------------------------
                    */
                    $billOfLadingInvalid = [];

                    $requiredBrokers = [
                        "200615811",
                        "200615811000",
                        "215722696",
                        "215722696000",
                        "225879904",
                        "225879904000",
                        "204867435",
                        "204867435000",
                        "432899304",
                        "432899304000",
                        "738464204",
                        "738464204000"
                    ];

                    if (in_array(trim($locbroktin), $requiredBrokers))
                    {
                        $BillOfLading = trim($BillOfLading);

                        // Required
                        if ($BillOfLading == "")
                        {
                            $billOfLadingInvalid[] = $row - 1;
                            $errorCounter++;
                        }
                        else
                        {
                            // Reject repeated digits (7-10 digits)
                            if (
                                preg_match('/^(\d)\1{6,9}$/', $BillOfLading) ||
                                $BillOfLading == "1111111116"
                            )
                            {
                                $billOfLadingInvalid[] = $row - 1;
                                $errorCounter++;
                            }
                            // Must be exactly 10 digits
                            else if (!preg_match('/^\d{10}$/', $BillOfLading))
                            {
                                $billOfLadingInvalid[] = $row - 1;
                                $errorCounter++;
                            }
                            else
                            {
                                // AWB check digit validation
                                $first9 = substr($BillOfLading, 0, 9);
                                $checkDigit = substr($BillOfLading, 9, 1);

                                if (($first9 % 7) != $checkDigit)
                                {
                                    $billOfLadingInvalid[] = $row - 1;
                                    $errorCounter++;
                                }
                            }
                        }
                    }
                    
                    //VesselAircraft
                    if( $validateFunc->max_length($VesselAircraft, 27) ) 
                    { 
                        $vesselAircraft[] = $row - 1; $errorCounter1++; 
                    }
                    if( !empty($VesselAircraft) && ($validateFunc->match_char($VesselAircraft)) == 0 )
                    { 
                        $vesselAircraftMatch[] = $row - 1; 
                        $errorCounter++; 
                    }

                    // LocationOfGoods
                    if( !empty($LocationOfGoods) ) {

                        //CHECK LocationOfGoods IF EXISTS
                        $checkLocationOfGoodsExists = $validateFunc->__checkValidLocationOfGoods($LocationOfGoods); //MODIFY

                        if( !$checkLocationOfGoodsExists )
                        {
                            $checkLocationOfGoods[] = $row - 1;
                            $errorCounter++;
                        }

                    }else{

                        $checkLocationOfGoods[] = $row - 1;
                        $errorCounter++;
                    }

                    // ProvinceOfOrigin
                    if( !empty($ProvinceOfOrigin) ) {

                        //CHECK ProvinceOfOrigin IF EXISTS
                        $checkProvinceOfOriginExists = $validateFunc->__checkValidProvinceOfOrigin($ProvinceOfOrigin); //MODIFY

                        if( !$checkProvinceOfOriginExists )
                        {
                            $checkProvinceOfOrigin[] = $row - 1;
                            $errorCounter++;
                        }

                    }else{

                        $checkProvinceOfOrigin[] = $row - 1;
                        $errorCounter++;
                    }

                    // CountryOfDestination
                    if( !empty($CountryOfDestination) ) {

                        //CHECK CountryOfDestination IF EXISTS
                        $checkCountryOfDestinationExists = $validateFunc->__checkValidCountryOfDestination($CountryOfDestination); //MODIFY

                        if( !$checkCountryOfDestinationExists )
                        {
                            $checkCountryOfDestination[] = $row - 1;
                            $errorCounter++;
                        }

                    }else{

                        $checkCountryOfDestination[] = $row - 1;
                        $errorCounter++;
                    }

                    // PortOfLoading
                    if( !empty($PortOfLoading) ) {

                        //CHECK PortOfLoading IF EXISTS
                        $checkPortOfLoadingExists = $validateFunc->__checkValidPortOfLoading($PortOfLoading); //MODIFY

                        if( !$checkPortOfLoadingExists )
                        {
                            $checkPortOfLoading[] = $row - 1;
                            $errorCounter++;
                        }

                    }else{

                        $checkPortOfLoading[] = $row - 1;
                        $errorCounter++;
                    }

                    // PortOfDeparture
                    if( !empty($PortOfDeparture) ) {

                        //CHECK PortOfDeparture IF EXISTS
                        $checkPortOfDepartureExists = $validateFunc->__checkValidPortOfDeparture($PortOfDeparture); //MODIFY

                        if( !$checkPortOfDepartureExists )
                        {
                            $checkPortOfDeparture[] = $row - 1;
                            $errorCounter++;
                        }

                    }else{

                        $checkPortOfDeparture[] = $row - 1;
                        $errorCounter++;
                    } 
            }
        }
    }

    if($objWorksheet2){

        $highestRow     = $objWorksheet2->getHighestRow();
        $highestColumn  = $objWorksheet2->getHighestColumn();
        $headingsArray  = $objWorksheet2->rangeToArray('A1:'.$highestColumn.'1',null, true, true, true);
        
        $headingsArray  = $headingsArray[1];
        $r = -1;

        if(strtoupper($highestColumn) != 'C') {

            echo "<script>
                    alert('File content is not compatible. Please check the Container Seal No sheet');
                    window.location.href='index.php?token=$token';
                </script>";

            unlink($excelDetails['inputFile']);
            die();
        }

        $checkItemsA2 = $objPHPExcel->getActiveSheet()->getCell('A2')->getValue();
        if($checkItemsA2 == NULL || $checkItemsA2 == '') {

            echo "<script>
                    alert('Cannot proceed. Please check file.');
                    window.location.href='index.php?token=$token';
                </script>";

            unlink($excelDetails['inputFile']);
            die();
            
        }else{

            for ($row = 2; $row <= $highestRow; ++$row) {

                $dataRow = $objWorksheet2->rangeToArray('A'.$row.':'.$highestColumn.$row, null, true, true, true);

                //CHECK IF ALL CELLS ARE EMPTY
                if( $validateFunc->isEmptyRow(reset($dataRow)) ) 
                { 
                    continue; //skip empty row
                } 

                    ++$r;
                    
                    foreach($headingsArray as $columnKey => $columnHeading) {

                        //VALUES
                        $ContainerNumber           =   $validateFunc->trim_val($dataRow[$row]['A']);
                        $SealNumber                =   $validateFunc->trim_val($dataRow[$row]['B']);
                        $ContainerSize             =   $validateFunc->trim_val($dataRow[$row]['C']);
                    }

                /* VALIDATE FIELD VALUES */

                    
                    // If the mode of transport is by air then validate it
                    if (!empty($checkModeOfTransportation))
                    {
                        if (
                            !empty($ContainerNumber) ||
                            !empty($SealNumber) ||
                            !empty($ContainerSize)
                        )
                        {
                            $containerDetailsNotAllowed[] = $row - 1;
                            $errorCounter++;
                        }
                    }
                    else 
                    {
                        //ContainerNumber
                        if( $validateFunc->max_length($ContainerNumber, 100) ) 
                        { 
                            $containerNumber[] = $row - 1; 
                            $errorCounter1++; 
                        }
                        if( !empty($ContainerNumber) &&  ($validateFunc->match_char($ContainerNumber)) == 0 ) 
                        { 
                            $containerNumberMatch[] = $row - 1; 
                            $errorCounter++; 
                        }

                        //SealNumber
                        if( $validateFunc->max_length($SealNumber, 100) ) 
                        { 
                            $sealNumber[] = $row - 1; $errorCounter1++; 
                        }
                        if( !empty($SealNumber) && ($validateFunc->match_char($SealNumber)) == 0 )
                        {
                            $sealNumberMatch[] = $row - 1; 
                            $errorCounter++; 
                        }

                        //ContainerSize
                        if( !empty($ContainerSize) )
                        { 
                            //CHECK ContainerSize IF EXISTS
                            $checkContainerSizeExists = $validateFunc->__checkValidContainerSize($ContainerSize); //MODIFY

                            if( !$checkContainerSizeExists ) 
                            {
                                $checkContainerSize[] = $row - 1; 
                                $errorCounter++; 
                            }
                        } else {
                            $checkContainerSize[] = $row - 1; 
                            $errorCounter++; 
                        }
                    }
            }
        }
    }

    if($objWorksheet3){

        $highestRow     = $objWorksheet3->getHighestRow();
        $highestColumn  = $objWorksheet3->getHighestColumn();
        $headingsArray  = $objWorksheet3->rangeToArray('A1:'.$highestColumn.'1',null, true, true, true);
        
        $headingsArray  = $headingsArray[1];
        $r = -1;

        if(strtoupper($highestColumn) != 'K') {

            echo "<script>
                    alert('File content is not compatible.  Please check the Items sheet');
                    window.location.href='index.php?token=$token';
                </script>";

            unlink($excelDetails['inputFile']);
            die();
        }

        $checkItemsA2 = $objPHPExcel->getActiveSheet()->getCell('A2')->getValue();
        if($checkItemsA2 == NULL || $checkItemsA2 == '') {

            echo "<script>
                    alert('Cannot proceed. Please check file.');
                    window.location.href='index.php?token=$token';
                </script>";

            unlink($excelDetails['inputFile']);
            die();
            
        }else{

            for ($row = 2; $row <= $highestRow; ++$row) {

                $dataRow = $objWorksheet3->rangeToArray('A'.$row.':'.$highestColumn.$row, null, true, true, true);

                //CHECK IF ALL CELLS ARE EMPTY
                if( $validateFunc->isEmptyRow(reset($dataRow)) ) 
                { 
                    continue; //skip empty row
                } 

                    ++$r;
                    
                    foreach($headingsArray as $columnKey => $columnHeading) {

                        //VALUES
                        
                        $ItemCode              =   $validateFunc->trim_val($dataRow[$row]['A']);
                        $MarksAndNumber        =   $validateFunc->trim_val($dataRow[$row]['B']);
                        $NumberOfPackage       =   $validateFunc->trim_val($dataRow[$row]['C']);
                        $PackageCode           =   $validateFunc->trim_val($dataRow[$row]['D']);
                        $InvoiceNumber         =   $validateFunc->trim_val($dataRow[$row]['E']);
                        $SuplementaryValue     =   $validateFunc->trim_val($dataRow[$row]['F']);
                        $ProcedureCode         =   $validateFunc->trim_val($dataRow[$row]['G']);
                        $ExtendedCode          =   $validateFunc->trim_val($dataRow[$row]['H']);
                        $ItemGrossWeight       =   $validateFunc->trim_val($dataRow[$row]['I']);
                        $ItemNetWeight         =   $validateFunc->trim_val($dataRow[$row]['J']);
                        $ItemInvoiceValue      =   $validateFunc->trim_val($dataRow[$row]['K']);
                    }

                /* VALIDATE FIELD VALUES */

                    //ItemCode
                    if( $validateFunc->max_length($ItemCode, 300) ) 
                    { 
                        $itemCode[] = $row - 1; 
                        $errorCounter1++; 
                    }
                    if( !empty($ItemCode) ) 
                    {
                        if ( ($validateFunc->match_char($ItemCode)) == 0 ) 
                        {
                            $itemCodeMatch[] = $row - 1; 
                            $errorCounter++; 
                        }

                        $checkItemCodeExists = $validateFunc->__checkItemCode($ItemCode, $allaccids, $accountType);
                        if( empty($checkItemCodeExists) )
                        {
                            $checkItemCode[] = $row - 1;
                            $errorCounter++;
                        }
                        if (!empty($checkItemCodeExists['uom_cod1']) && empty($SuplementaryValue))
                        {
                            $checkSuplementaryValue[] = $row - 1;
                            $errorCounter++;
                        }
                        if (empty($checkItemCodeExists['uom_cod1']) && !empty($SuplementaryValue))
                        {
                            $checkSuplementaryValue1[] = $row - 1;
                            $errorCounter++;
                        }

                    } else {
                        $itemCodeMatch[] = $row - 1; 
                        $errorCounter++; 
                    }

                    

                    //MarksAndNumber
                    if( $validateFunc->max_length($MarksAndNumber, 70) ) 
                    { 
                        $marksAndNumber[] = $row - 1; $errorCounter1++; 
                    }
                    if( !empty($MarksAndNumber) ) 
                    {
                        if ( ($validateFunc->match_char($MarksAndNumber)) == 0 ) 
                        {
                            $marksAndNumberMatch[] = $row - 1; 
                            $errorCounter++; 
                        }
                    } else {
                        $marksAndNumberMatch[] = $row - 1; 
                        $errorCounter++; 
                    }

                    //NumberOfPackage
                    if( $validateFunc->max_length($NumberOfPackage, 6) ) 
                    { 
                        $numberOfPackage[] = $row - 1; $errorCounter1++; 
                    }
                    if( !empty($NumberOfPackage) && ($validateFunc->match_numbers($NumberOfPackage)) == 0 )
                    {
                        $numberOfPackageMatch[] = $row - 1; 
                        $errorCounter++; 
                    }

                    //PackageCode
                    if( !empty($PackageCode) && $validateFunc->match_packcodeFormat($PackageCode) ) {

                        //CHECK PACKAGE CODE IF EXISTS
                        $checkPackCodeExists = $validateFunc->__checkValidPackCode($PackageCode); //MODIFY

                        if( !$checkPackCodeExists )
                        {
                            $checkPackCode[] = $row - 1;
                            $errorCounter++;
                        }

                    }else{

                        $checkPackCode[] = $row - 1;
                        $errorCounter++;
                    }

                    //InvoiceNumber
                    if( $validateFunc->max_length($InvoiceNumber, 300) ) 
                    { 
                        $invoiceNumber[] = $row - 1; $errorCounter1++; 
                    }
                    if( !empty($InvoiceNumber) && ($validateFunc->match_char($InvoiceNumber)) == 0 )
                    {
                        $invoiceNumberMatch[] = $row - 1; 
                        $errorCounter++; 
                    }

                    //SuplementaryValue
                    if( !empty($SuplementaryValue) && ($validateFunc->match_numbers($SuplementaryValue)) == 0 )
                    {
                        $suplementaryValueMatch[] = $row - 1; 
                        $errorCounter++; 
                    }

                    //procedureCode
                    if( $validateFunc->max_length($ProcedureCode, 6) ) 
                    { 
                        $cOO[] = $row - 1; $errorCounter1++; 
                    }
                    if( !empty($ProcedureCode) && ($validateFunc->match_numbers($ProcedureCode)) == 0 )
                    {
                        $cOOMatch[] = $row - 1; 
                        $errorCounter++; 
                    }

                    //ExtendedCode
                    if( !empty($ExtendedCode) && ($ExtendedCode) !== "000" )
                    {
                        $checkExtCode[] = $row - 1; 
                        $errorCounter++; 
                    }

                    //ItemGrossWeight
                    if( !empty($ItemGrossWeight) && ($validateFunc->match_amount($ItemGrossWeight)) == 0 )
                    {
                        $itemGrossWeightMatch[] = $row - 1; 
                        $errorCounter++; 
                    }

                    //ItemNetWeight
                    if( !empty($ItemNetWeight) && ($validateFunc->match_amount($ItemNetWeight)) == 0 )
                    {
                        $itemNetWeightMatch[] = $row - 1; 
                        $errorCounter++; 
                    }

                    //ItemInvoiceValue
                    if( !empty($ItemInvoiceValue) && ($validateFunc->match_amount($ItemInvoiceValue)) == 0 )
                    {
                        $itemInvoiceValueMatch[] = $row - 1; 
                        $errorCounter++; 
                    }
            }
        }
    }

    if($objWorksheet4){

        $highestRow     = $objWorksheet4->getHighestRow();
        $highestColumn  = $objWorksheet4->getHighestColumn();
        $headingsArray  = $objWorksheet4->rangeToArray('A1:'.$highestColumn.'1',null, true, true, true);
        
        $headingsArray  = $headingsArray[1];
        $r = -1;

        if(strtoupper($highestColumn) != 'B') {

            echo "<script>
                    alert('1.File content is not compatible. Please check the Financial sheet');
                    window.location.href='index.php?token=$token';
                </script>";

            unlink($excelDetails['inputFile']);
            die();
        }

        if ($highestRow > 2) {

            echo "<script>
                    alert('Only one record is allowed in Financial sheet.');
                    window.location.href='index.php?token=$token';
                </script>";

            unlink($excelDetails['inputFile']);
            die();
        }

        $checkItemsA2 = $objPHPExcel->getActiveSheet()->getCell('A2')->getValue();
        if($checkItemsA2 == NULL || $checkItemsA2 == '') {

            echo "<script>
                alert('Cannot proceed. Please check file.');
                window.location.href='index.php?token=$token';
            </script>";

            unlink($excelDetails['inputFile']);
            die();
            
        }else{

            for ($row = 2; $row <= $highestRow; ++$row) {

                $dataRow = $objWorksheet4->rangeToArray('A'.$row.':'.$highestColumn.$row, null, true, true, true);

                //CHECK IF ALL CELLS ARE EMPTY
                if( $validateFunc->isEmptyRow(reset($dataRow)) ) 
                { 
                    continue; //skip empty row
                } 

                    ++$r;
                    
                    foreach($headingsArray as $columnKey => $columnHeading) {

                        //VALUES
                        $TermsOfDelivery =   $validateFunc->trim_val($dataRow[$row]['A']);
                        $TermsOfPayment  =   $validateFunc->trim_val($dataRow[$row]['B']);
                    }

                /* VALIDATE FIELD VALUES */

                    //TermsOfDelivery
                    if( !empty($TermsOfDelivery) ) {

                        //CHECK TermsOfDelivery IF EXISTS
                        $checkTermsOfDelivery = $validateFunc->__checkValidTermsOfDelivery($TermsOfDelivery); //MODIFY

                        if( !$checkTermsOfDelivery )
                        {
                            $termsOfDelivery[] = $row - 1; 
                            $errorCounter1++; 
                        }

                    }else{

                        $termsOfDelivery[] = $row - 1; 
                        $errorCounter1++; 
                    }

                    //TermsOfPayment
                    if( !empty($TermsOfPayment) ) {

                        //CHECK TermsOfPayment IF EXISTS
                        $checkTermsOfPayment = $validateFunc->__checkValidTermsOfPayment($TermsOfPayment); //MODIFY

                        if( !$checkTermsOfPayment )
                        {
                            $termsOfPayment[] = $row - 1; 
                            $errorCounter1++; 
                        }

                    }else{

                        $termsOfPayment[] = $row - 1; 
                        $errorCounter1++; 
                    }
            }
        }
    }

    if ($errorCounter > 0 || $errorCounter1 > 0) {
        
        /* TEST */
        // die(var_dump($validateFunc->match_char($Consignee)));

        /* ERROR MESSAGES */
        // General Sheet Validation
            //Consignee
            if(!empty($consignee)){
                $errorLists[] = array(
                                    "ErrMsg" => "Exceeds the max characters allowed (70)",
                                    "Column" => "Consignee",
                                    "Rows" => implode(", " ,$consignee)
                                );
            }
            if(!empty($consigneeMatch)){
                $errorLists[] = array(
                                    "ErrMsg" => "Only accept letters, numbers and few special characters (-_.,:;#$%()*/) - Required",
                                    "Column" => "Consignee",
                                    "Rows" => implode(", " ,$consigneeMatch)
                                );
            }

            if(!empty($checkConsignee)){
                $errorLists[] = array(
                                    "ErrMsg" => "Invalid Consignee/Buyer, please check Buyer Lookup",
                                    "Column" => "Consignee",
                                    "Rows" => implode(", " ,$checkConsignee)
                                );
            }
            
            //Address
            if(!empty($address)){
                $errorLists[] = array(
                                    "ErrMsg" => "Exceeds the max characters allowed (105)",
                                    "Column" => "Address",
                                    "Rows" => implode(", " ,$address)
                                );
            }
            if(!empty($addressMatch)){
                $errorLists[] = array(
                                    "ErrMsg" => "Only accept letters, numbers and few special characters (-_.,:;#$%()*/) - Required",
                                    "Column" => "Address",
                                    "Rows" => implode(", " ,$addressMatch)
                                );
            }
            if(!empty($checkAddress)){
                $errorLists[] = array(
                                    "ErrMsg" => "Address does not match the registered address for this Consignee",
                                    "Column" => "Address",
                                    "Rows" => implode(", " ,$checkAddress)
                                );
            }
            
            //Port
            if(!empty($checkOfficeOfClearance)){
                $errorLists[] = array(
                                    "ErrMsg" => "Invalid Port",
                                    "Column" => "Port",
                                    "Rows" => implode(", " ,$checkOfficeOfClearance)
                                );
            }

            if(!empty($checkOfficeOfClearanceCase)){
                $errorLists[] = array(
                                    "ErrMsg" => "Port code casing does not match master data",
                                    "Column" => "Port",
                                    "Rows" => implode(", " ,$checkOfficeOfClearanceCase)
                                );
            }
            
            //PurposeOfExportation
            if(!empty($checkPurposeOfExportation)){
                $errorLists[] = array(
                                    "ErrMsg" => "Invalid Purpose of Exportation",
                                    "Column" => "Purpose Of Exportation",
                                    "Rows" => implode(", " ,$checkPurposeOfExportation)
                                );
            }
            
            //ManifestNo
            if(!empty($manifestNoMatch)){
                $errorLists[] = array(
                                    "ErrMsg" => "Invalid Manifest Number! Format: NNNMMMM-YY",
                                    "Column" => "Manifest Number",
                                    "Rows" => implode(", " ,$manifestNoMatch)
                                );
            }
            
            //BillOfLading
            if(!empty($billOfLading)){
                $errorLists[] = array(
                                    "ErrMsg" => "Exceeds the max characters allowed (52)",
                                    "Column" => "Bill of Lading",
                                    "Rows" => implode(", " ,$billOfLading)
                                );
            }            
            if(!empty($billOfLadingMatch)){
                $errorLists[] = array(
                                    "ErrMsg" => "Only accept letters and numbers - Required",
                                    "Column" => "Bill of Lading",
                                    "Rows" => implode(", " ,$billOfLadingMatch)
                                );
            }            
            if(!empty($billOfLadingInvalid)){
                $errorLists[] = array(
                                    "ErrMsg" => "Please enter a valid bill of lading/airway bill.",
                                    "Column" => "Bill of Lading",
                                    "Rows" => implode(", " ,$billOfLadingInvalid)
                                );
            }            
            
            //VesselAircraft
            if(!empty($vesselAircraft)){
                $errorLists[] = array(
                                    "ErrMsg" => "Exceeds the max characters allowed (27)",
                                    "Column" => "Vessel / Aircraft",
                                    "Rows" => implode(", " ,$vesselAircraft)
                                );
            }
            if(!empty($vesselAircraftMatch)){
                $errorLists[] = array(
                                    "ErrMsg" => "Only accept letters, numbers and few special characters (-_.,:;#$%()*/) - Required",
                                    "Column" => "Vessel / Aircraft",
                                    "Rows" => implode(", " ,$vesselAircraftMatch)
                                );
            }

            //LocationOfGoods
            if(!empty($checkLocationOfGoods)){
                $errorLists[] = array(
                                    "ErrMsg" => "Invalid Location of Goods",
                                    "Column" => "Location of Goods",
                                    "Rows" => implode(", " ,$checkLocationOfGoods)
                                );
            }

            //ProvinceOfOrigin
            if(!empty($checkProvinceOfOrigin)){
                $errorLists[] = array(
                                    "ErrMsg" => "Invalid Province of Origin",
                                    "Column" => "Province of Origin",
                                    "Rows" => implode(", " ,$checkProvinceOfOrigin)
                                );
            }

            //CountryOfDestination
            if(!empty($checkCountryOfDestination)){
                $errorLists[] = array(
                                    "ErrMsg" => "Invalid Country of Destination",
                                    "Column" => "Country of Destination",
                                    "Rows" => implode(", " ,$checkCountryOfDestination)
                                );
            }

            //PortOfLoading
            if(!empty($checkPortOfLoading)){
                $errorLists[] = array(
                                    "ErrMsg" => "Invalid Port of Loading",
                                    "Column" => "Port of Loading",
                                    "Rows" => implode(", " ,$checkPortOfLoading)
                                );
            }

            //PortOfDeparture
            if(!empty($checkPortOfDeparture)){
                $errorLists[] = array(
                                    "ErrMsg" => "Invalid Port of Departure",
                                    "Column" => "Port of Departure",
                                    "Rows" => implode(", " ,$checkPortOfDeparture)
                                );
            }

        // Container Seal No Sheet Validation
            // If the mode of transport is by sea then validate it
            if (empty($checkModeOfTransportation))
            {
                if(!empty($containerNumber)){
                    $errorLists[] = array(
                                        "ErrMsg" => "Exceeds the max characters allowed (100)",
                                        "Column" => "Container Number",
                                        "Rows" => implode(", " ,$containerNumber)
                                    );
                }
                if(!empty($containerNumberMatch)){
                    $errorLists[] = array(
                                        "ErrMsg" => "Only accept letters, numbers and few special characters (-_.,:;#$%()*/) - Required",
                                        "Column" => "Container Number",
                                        "Rows" => implode(", " ,$containerNumberMatch)
                                    );
                }

                if(!empty($sealNumber)){
                    $errorLists[] = array(
                                        "ErrMsg" => "Exceeds the max characters allowed (100)",
                                        "Column" => "Seal Number",
                                        "Rows" => implode(", " ,$sealNumber)
                                    );
                }
                if(!empty($sealNumberMatch)){
                    $errorLists[] = array(
                                        "ErrMsg" => "Only accept letters, numbers and few special characters (-_.,:;#$%()*/) - Required",
                                        "Column" => "Seal Number",
                                        "Rows" => implode(", " ,$sealNumberMatch)
                                    );
                }

                if(!empty($checkContainerSize)){
                    $errorLists[] = array(
                                        "ErrMsg" => "Invalid Container Size",
                                        "Column" => "Container Size",
                                        "Rows" => implode(", " ,$checkContainerSize)
                                    );
                }
            } else {
                if(!empty($containerDetailsNotAllowed)){
                    $errorLists[] = array(
                                        "ErrMsg" => "Container details are not allowed for this mode of transport.",
                                        "Column" => "Container Number, Seal Number, Container Size",
                                        "Rows" => implode(", " ,$containerDetailsNotAllowed)
                                    );
                }
                
            }

        // Items Sheet Validation
            // ItemCode
            if(!empty($itemCode)){
                $errorLists[] = array(
                                    "ErrMsg" => "Exceeds the max characters allowed (300)",
                                    "Column" => "Item Code",
                                    "Rows" => implode(", " ,$itemCode)
                                );
            }
            if(!empty($itemCodeMatch)){
                $errorLists[] = array(
                                    "ErrMsg" => "Only accept letters, numbers and few special characters (-_.,:;#$%()*/) - Required",
                                    "Column" => "Item Code",
                                    "Rows" => implode(", " ,$itemCodeMatch)
                                );
            }
            if(!empty($checkItemCode)){
                $errorLists[] = array(
                                    "ErrMsg" => "invalid Item Code, please check Exportables Lookup",
                                    "Column" => "Item Code",
                                    "Rows" => implode(", " ,$checkItemCode)
                                );
            }
            
            // MarksAndNumber
            if(!empty($marksAndNumber)){
                $errorLists[] = array(
                                    "ErrMsg" => "Exceeds the max characters allowed (70)",
                                    "Column" => "Marks and Number",
                                    "Rows" => implode(", " ,$marksAndNumber)
                                );
            }
            if(!empty($marksAndNumberMatch)){
                $errorLists[] = array(
                                    "ErrMsg" => "Only accept letters, numbers and few special characters (-_.,:;#$%()*/) - Required",
                                    "Column" => "Marks and Number",
                                    "Rows" => implode(", " ,$marksAndNumberMatch)
                                );
            }
            
            // NumberOfPackage
            if(!empty($numberOfPackage)){
                $errorLists[] = array(
                                    "ErrMsg" => "Exceeds the max characters allowed (10)",
                                    "Column" => "Number of Package",
                                    "Rows" => implode(", " ,$NumberOfPackage)
                                );
            }
            if(!empty($numberOfPackageMatch)){
                $errorLists[] = array(
                                    "ErrMsg" => "Only accept letters and numbers - Required",
                                    "Column" => "Number of Package",
                                    "Rows" => implode(", " ,$numberOfPackageMatch)
                                );
            }

            // PackageCode
            if(!empty($checkPackCode)){
                $errorLists[] = array(
                                    "ErrMsg" => "Invalid Package Code",
                                    "Column" => "Package Code",
                                    "Rows" => implode(", " ,$checkPackCode)
                                );
            }
            
            // InvoiceNumber
            if(!empty($invoiceNumber)){
                $errorLists[] = array(
                                    "ErrMsg" => "Exceeds the max characters allowed (300)",
                                    "Column" => "Invoice Number",
                                    "Rows" => implode(", " ,$invoiceNumber)
                                );
            }
            if(!empty($invoiceNumberMatch)){
                $errorLists[] = array(
                                    "ErrMsg" => "Only accept letters and numbers - Required",
                                    "Column" => "Invoice Number",
                                    "Rows" => implode(", " ,$invoiceNumberMatch)
                                );
            }
            
            // SuplementaryValue
            if(!empty($suplementaryValueMatch)){
                $errorLists[] = array(
                                    "ErrMsg" => "Only accept whole numbers",
                                    "Column" => "Supplementary Value",
                                    "Rows" => implode(", " ,$suplementaryValueMatch)
                                );
            }
            if(!empty($checkSuplementaryValue)){
                $errorLists[] = array(
                                    "ErrMsg" => "Supplementary Value required for the following item",
                                    "Column" => "Supplementary Value",
                                    "Rows" => implode(", " ,$checkSuplementaryValue)
                                );
            }
            if(!empty($checkSuplementaryValue1)){
                $errorLists[] = array(
                                    "ErrMsg" => "Supplementary Value is not allowed for the following item",
                                    "Column" => "Supplementary Value",
                                    "Rows" => implode(", " ,$checkSuplementaryValue1)
                                );
            }
            
            // procedureCode
            if(!empty($cOO)){
                $errorLists[] = array(
                                    "ErrMsg" => "Exceeds the max characters allowed (6)",
                                    "Column" => "Procedure Code",
                                    "Rows" => implode(", " ,$cOO)
                                );
            }
            if(!empty($cOOMatch)){
                $errorLists[] = array(
                                    "ErrMsg" => "Only accept letters and numbers - Required",
                                    "Column" => "Procedure Code",
                                    "Rows" => implode(", " ,$cOOMatch)
                                );
            }
            
            // ExtendedCode
            if(!empty($checkExtCode)){
                $errorLists[] = array(
                                    "ErrMsg" => "Invalid code for the Extended Code - Required",
                                    "Column" => "EXTENDED CODE",
                                    "Rows" => implode(", " ,$checkExtCode)
                                );
            }

            //ItemGrossWeight
            if(!empty($itemGrossWeightMatch)){ 
                $errorLists[] = array(
                                    "ErrMsg" => "Invalid entry (e.g. 0.00) - Required",
                                    "Column" => "ITEM GROSS WEIGHT",
                                    "Rows" => implode(", " ,$itemGrossWeightMatch)
                                );
            }

            //ItemNetWeight
            if(!empty($itemNetWeightMatch)){ 
                $errorLists[] = array(
                                    "ErrMsg" => "Invalid entry (e.g. 0.00) - Required",
                                    "Column" => "ITEM NET WEIGHT",
                                    "Rows" => implode(", " ,$itemNetWeightMatch)
                                );
            }

            //AIRBILLNO
            if(!empty($airbillNo)){
                $errorLists[] = array(
                                    "ErrMsg" => "Exceeds the max characters allowed (26)",
                                    "Column" => "AIRBILL/BL NUMBER",
                                    "Rows" => implode(", " ,$airbillNo)
                                );
            }
            if(!empty($airbillNoMatch)){
                $errorLists[] = array(
                                    "ErrMsg" => "Only accept letters and numbers - Required",
                                    "Column" => "AIRBILL/BL NUMBER",
                                    "Rows" => implode(", " ,$airbillNoMatch)
                                );
            }

            //ItemInvoiceValue
            if(!empty($itemInvoiceValueMatch)){ 
                $errorLists[] = array(
                                    "ErrMsg" => "Invalid entry (e.g. 1000.00) - Required",
                                    "Column" => "INVOICE NUMBER",
                                    "Rows" => implode(", " ,$itemInvoiceValueMatch)
                                );
            }

        // Financial Sheet Validation 

            //TermsOfDelivery
            if(!empty($termsOfDelivery)){ 
                $errorLists[] = array(
                                    "ErrMsg" => "Invalid Term of Delivery - Required",
                                    "Column" => "TERMS OF DELIVERY",
                                    "Rows" => implode(", " ,$termsOfDelivery)
                                );
            }  

            //TermsOfPayment
            if(!empty($termsOfPayment)){ 
                $errorLists[] = array(
                                    "ErrMsg" => "Invalid Term of Payment - Required",
                                    "Column" => "TERMS OF DELIVERY",
                                    "Rows" => implode(", " ,$termsOfPayment)
                                );
            }

        //

            $_SESSION['errormsg'] = $errorLists;
            $_SESSION['required'] = $errorCounter;
            $_SESSION['proceed']  = $errorCounter1;

            echo "<script>
                    window.location.href='index.php?msg=error&token=$token';
                </script>";
            die();
    
        
    }else{

        // die("NO ERROR TO DISPLAY");
        
        // START PROCESSING //
        $objReader      = PHPExcel_IOFactory::createReader($excelDetails["type"]);
        $objReader->setReadDataOnly(true);
        $objPHPExcel    = $objReader->load($excelDetails["inputFile"]);

        $objWorksheet   = $objPHPExcel->getSheetByName('General');
        $objWorksheet2  = $objPHPExcel->getSheetByName('CONTAINER SEAL NO');
        $objWorksheet3  = $objPHPExcel->getSheetByName('Items');
        $objWorksheet4  = $objPHPExcel->getSheetByName('Financial');

        $applNo    = $validateFunc->generateApplNo($conn, $csncod);

        if(!$objWorksheet || !$objWorksheet2 || !$objWorksheet3 || !$objWorksheet4){
            echo "<script>
                    alert('Cannot proceed. Please check file.');
                    window.location.href='index.php?token=$token';
            </script>";
            die();
        }

        // ------------ START General WORKSHEET (MAPPING and DB INSERTION) ------------ //
        if($objWorksheet){
            $highestColumn2 = $objWorksheet->getHighestColumn();

            // Get the first data row (row 2)
            $dataRow2 = $objWorksheet->rangeToArray('A2:'.$highestColumn2.'2', null, true, true, true);

            // Check if column A has a value
            if (isset($dataRow2[2]['A']) && $dataRow2[2]['A'] != '') {

                $master_val = strtoupper($dataRow2[2]['A']);
                
                $forwarder = $lookupData->getForwarders($conn, $lstexporter);
                $importer  = $lookupData->getImporters($conn, $loccod);
                $exchRate  = $lookupData->getExchangeRate($conn, 'USD');

                // ----------------- CONSIGNEE ADDRESS: pull from BUYER master data ----------------- //
                // Re-fetch the validated buyer record (same lookup used during validation)
                // rather than re-using the free-typed $Address column from the Excel file.
                $buyerRecord = $validateFunc->__checkValidConsignee($conn, $Consignee, $cltcode);

                if ($buyerRecord) {

                    // Use BUYER's own pre-split address lines directly — do NOT
                    // re-concatenate + str_split, since that breaks mid-word.
                    $conadr1 = isset($buyerRecord['Expadr1']) ? $buyerRecord['Expadr1'] : '';
                    $conadr2 = isset($buyerRecord['Expadr2']) ? $buyerRecord['Expadr2'] : '';
                    $conadr3 = isset($buyerRecord['Expadr3']) ? $buyerRecord['Expadr3'] : '';

                    // BUYER has 4 lines, ConAdr only has 3 columns — fold Expadr4 into
                    // line 3 with a space, matching how the ASP page just concatenates
                    // Expadr1.." ".Expadr2.." ".Expadr3.." ".Expadr4 for display.
                    if (!empty($buyerRecord['Expadr4'])) {
                        $conadr3 = trim($conadr3 . ' ' . $buyerRecord['Expadr4']);
                    }

                } else {

                    // Fallback: shouldn't happen since validation already rejected
                    // rows where the buyer isn't found — guard anyway.
                    $chunksAddress = str_split($Address, 35);
                    $conadr1 = isset($chunksAddress[0]) ? $chunksAddress[0] : '';
                    $conadr2 = isset($chunksAddress[1]) ? $chunksAddress[1] : '';
                    $conadr3 = isset($chunksAddress[2]) ? $chunksAddress[2] : '';
                }
                // ----------------- END CONSIGNEE ADDRESS ----------------- //
                
                $modeofTransport = $lookupData->getModeofTransport($conn, $Port);
                
                // Prepare and execute insert
                $insert_master = "INSERT INTO tblEXPAPL_Master (Applno, ConName, ConAdr1, ConAdr2, ConAdr3, OffClear, Manifest, Waybill, DECTIN, DECname, DecAdr1, DecAdr2, DecAdr3, Cexp, Cdest, Vessel, ExpCode, ExpName, ExpAdr1, ExpAdr2, RegOfc, mdec, mdec2, Exhrate, PortofLoad, PortofDept, ProvofOrig, CreationDate, Stat, ConTIN, IAN, LGoods, Purpose, cltcode, SenderID, modeOfTransport, isExcelFileAppl) 
                                VALUES (:applno, :conname, :conadr1, :conadr2, :conadr3, :offclear, :manifest, :waybill, :dectin, :decname, :decadr1, :decadr2, :decadr3, :cexp, :cdest, :vessel, :expcode, :expname, :expadr1, :expadr2, :regofc, :mdec, :mdec2, :exhrate, :portofload, :portofdept, :provoforig, :creationdate, :stat, :contin, :ian, :lgoods, :purpose, :cltcode, :senderid, :modeoftransport, :isExcelFileAppl)";
                
                try {
                    $stmt3 = $conn->connectIPPEZA()->prepare($insert_master);
                    $stmt3->execute([
                        ':applno'           => $applNo,
                        ':conname'          => $Consignee, 
                        ':conadr1'          => $conadr1, 
                        ':conadr2'          => $conadr2, 
                        ':conadr3'          => $conadr3,
                        ':offclear'         => $Port,
                        ':manifest'         => $ManifestNo,
                        ':waybill'          => $BillOfLading,
                        ':dectin'           => $locbroktin,
                        ':decname'          => $lstexporter,
                        ':decadr1'          => $forwarder['For_adr1'], 
                        ':decadr2'          => $forwarder['For_adr2'],
                        ':decadr3'          => $forwarder['For_adr3'],
                        ':cexp'             => 'PH',
                        ':cdest'            => $CountryOfDestination,
                        ':vessel'           => $VesselAircraft,
                        ':expcode'          => $loccod,
                        ':expname'          => $compNam,
                        ':expadr1'          => $importer['address1'],
                        ':expadr2'          => $importer['address2'],
                        ':regofc'           => $importer['zonecode'],
                        ':mdec'             => $mod_cod,
                        ':mdec2'            => $mod_cod2,
                        ':exhrate'          => $exchRate['rat_exc'],
                        ':portofload'       => $PortOfLoading,
                        ':portofdept'       => $PortOfDeparture,
                        ':provoforig'       => $ProvinceOfOrigin,
                        ':creationdate'     => date('Y-m-d H:i:s'),
                        ':stat'             => 'I',
                        ':contin'           => $locTin,
                        ':ian'              => 'isPTOPS',
                        ':lgoods'           => $LocationOfGoods,
                        ':purpose'          => $PurposeOfExportation,
                        ':cltcode'          => $cltcode,
                        ':senderid'         => $userID,
                        ':modeoftransport'  => $modeofTransport['offClrMode'],
                        ':isExcelFileAppl'  => 1
                    ]);
                } catch (PDOException $e3) {
                    echo "ERROR: " . $e3->getMessage();
                    die();
                }
            }
        }
        // ------------ END General WORKSHEET ------------ //

        // ------------ START Container Seal No WORKSHEET (MAPPING and DB INSERTION) ------------ //
        if($objWorksheet2){

            $highestRow2    = $objWorksheet2->getHighestRow();
            $highestColumn2 = $objWorksheet2->getHighestColumn();
            $headingsArray2 = $objWorksheet2->rangeToArray('A1:'.$highestColumn2.'1',null, true, true, true);
            $headingsArray2 = $headingsArray2[1];

           
            $namedDataArray2 = array();

            for ($row2 = 2; $row2 <= $highestRow2; ++$row2) {

                $dataRow2   = $objWorksheet2->rangeToArray('A'.$row2.':'.$highestColumn2.$row2,null, true, true, true);

                $ContainerNumber = strtoupper(trim($dataRow2[$row2]['A']));
                $SealNumber      = strtoupper(trim($dataRow2[$row2]['B']));
                $ContainerSize   = strtoupper(trim($dataRow2[$row2]['C']));

                try {

                    // Prepare and execute insert
                    $stmt = $conn->connectIPPEZA()->prepare("INSERT INTO tblEXPAPL_ContPEZA (Applno, Container, Seal, ContainerSize, ModeOfShipment) VALUES (:applno, :container, :seal, :containerSize, :modeOfShipment)");

                    $stmt->execute([
                        ':applno'        => $applNo,
                        ':container'     => $ContainerNumber,
                        ':seal'          => $SealNumber,
                        ':containerSize' => $ContainerSize,
                        ':modeOfShipment'=> "FCL"
                    ]);

                } catch (PDOException $e) {
                    echo "ERROR: " . $e3->getMessage();
                    die();
                }
            }

        }else {

            echo "<script>
                    alert('Cannot proceed, could not find Container Seal Nos worksheet. Please check file.');
                    window.location.href='index.php?token=$token';
                 </script>";
            die();

        }
        // ------------ END Container Seal No WORKSHEET ------------ //

        // ------------ START Items WORKSHEET (MAPPING and DB INSERTION) ------------ //
        if($objWorksheet3){
            
            $highestRow3     = $objWorksheet3->getHighestDataRow();
            $highestColumn3  = $objWorksheet3->getHighestColumn();
            $headingsArray3  = $objWorksheet3->rangeToArray('A1:'.$highestColumn3.'1',null, true, true, true);
            $headingsArray3  = $headingsArray3[1];

            $r = -1;
            $namedDataArray3 = array();

            for ($row3 = 2; $row3 <= $highestRow3; ++$row3) {

                $dataRow3   = $objWorksheet3->rangeToArray('A'.$row3.':'.$highestColumn3.$row3,null, true, true, true);

                if (trim($dataRow3[$row3]['A']) == '') {
                    continue;
                }
                
                ++$r;
                
                foreach($headingsArray3 as $columnKey => $columnHeading) {
                    
                    $ItemCode              =   $validateFunc->trim_val($dataRow3[$row3]['A']);
                    
                    $checkItemCodeExists = $validateFunc->__checkItemCode($ItemCode, $allaccids, $accountType);
                    
                    $isRegulated = "";
                    if ($checkItemCodeExists['status'] == "M") {
                        $isRegulated = "True";
                    } else if ($checkItemCodeExists['status'] == "A")  {
                        $isRegulated = "False"; 
                    }

                    $Regulated              = $isRegulated;
                    $goodsdesc1             = $checkItemCodeExists['commodityDesc'];
                    $HSCode                 = $checkItemCodeExists['HsCode'];
                    $HSCode_Tar             = $checkItemCodeExists['HsCode_Tar'];
                    $PTOPS_ROWID            = $checkItemCodeExists['PTOPS_ROWID'];
                    $ecai_no_list           = $checkItemCodeExists['ecai_no'];
                    
                    $MarksAndNumber        =   $validateFunc->trim_val($dataRow3[$row3]['B']);
                    $NumberOfPackage       =   $validateFunc->trim_val($dataRow3[$row3]['C']);
                    $PackageCode           =   $validateFunc->trim_val($dataRow3[$row3]['D']);
                    $InvoiceNumber         =   $validateFunc->trim_val($dataRow3[$row3]['E']);
                    $SuplementaryValue     =   $validateFunc->trim_val($dataRow3[$row3]['F']);
                    $ProcedureCode         =   $validateFunc->trim_val($dataRow3[$row3]['G']);
                    $ExtendedCode          =   $validateFunc->trim_val($dataRow3[$row3]['H']);
                    $ItemGrossWeight       =   $validateFunc->trim_val($dataRow3[$row3]['I']);
                    $ItemNetWeight         =   $validateFunc->trim_val($dataRow3[$row3]['J']);
                    $ItemInvoiceValue      =   $validateFunc->trim_val($dataRow3[$row3]['K']);

                }

                /*DEFAULT VALUES*/
                $quo_cod                = 'NNNNN';
                $quo_dsc                = 'NOT RELATED, NO RSTRCTN/CNDTN/RYLTS/ARRNGMNTS';
                $ValMethodNum           = '1';
                $ValMethodDesc          = 'TRANSACTION VALUE';
                $Ocharges               = '0';
                $IFreight               = '0';
                $InvCurr                = 'USD';
                $Pref                   = 'NONE';   
                $ProcDesc               = '1000';
                $CoCode                 = 'PH';

                $ItemGrossWeight    = number_format((float)str_replace(',', '', $ItemGrossWeight), 2, '.', '');
                $ItemNetWeight      = number_format((float)str_replace(',', '', $ItemNetWeight), 2, '.', '');
                $ItemInvoiceValue   = number_format((float)str_replace(',', '', $ItemInvoiceValue), 2, '.', '');
                
                $ItemCode               = strtoupper($ItemCode);
                $MarksAndNumber         = strtoupper($MarksAndNumber);
                $NumberOfPackage        = strtoupper($NumberOfPackage);
                $PackageCode            = strtoupper($PackageCode);
                $InvoiceNumber          = strtoupper($InvoiceNumber);
                $SuplementaryValue      = strtoupper($SuplementaryValue);
                $ProcedureCode          = strtoupper($ProcedureCode);
                $ExtendedCode           = strtoupper($ExtendedCode);
                $ItemGrossWeight        = strtoupper($ItemGrossWeight);
                $ItemNetWeight          = strtoupper($ItemNetWeight);
                $ItemInvoiceValue       = strtoupper($ItemInvoiceValue);
                
                // ----------------- START GET ITEMNO ----------------- //
                $itemNo = '';

                $getItemNo = $validateFunc->__getItemNo($applNo); //MODIFY

                if( !empty($getItemNo) ) 
                {
                    $itemNo = $getItemNo + 1 ;
                }
                else
                {
                    $itemNo = 1;
                }
                
                // ----------------- END GET ITEMNO ----------------- //

                $insert_sql1 = "";

                //INSERT STATEMENT
                $insert_sql1 = "INSERT INTO TBLEXPAPL_DETAIL (ApplNo, ItemNo, itemcode, Marks1, NoPack, PackCode, InvNo, SupVal1, [Procedure], ExtCode, ItemGWeight, ItemNWeight, InvValue, quo_cod, quo_dsc, ValMethodNum, ValMethodDesc, Ocharges, IFreight, InvCurr, Pref, ProcDesc, CoCode, Regulated, goodsdesc1, HSCode, HSCode_Tar, PTOPS_ROWID, ecai_no_list)
                                VALUES (:applno, :itemNo, :itemcode, :marks1, :nopack, :packcode, :invno, :supval1, :procedure, :extcode, :itemgrossweight, :itemnetweight, :iteminvoicevalue, :quo_cod, :quo_dsc, :valMethodNum, :valMethodDesc, :Ocharges, :IFreight, :InvCurr, :Pref, :ProcDesc, :CoCode, :Regulated, :goodsdesc1, :HSCode, :HSCode_Tar, :PTOPS_ROWID, :ecai_no_list)";

                //EXECUTE QUERY TO INSERT
                try{

                    $sqlExecute1  = $insert_sql1;
                    $stmt1 = $conn->connectIPPEZA()->prepare($sqlExecute1);
                    $stmt1->execute([
                        ':applno'          => $applNo,
                        ':itemNo'          => $itemNo, 
                        ':itemcode'        => $ItemCode,
                        ':marks1'          => $MarksAndNumber,
                        ':nopack'          => $NumberOfPackage,
                        ':packcode'        => $PackageCode,
                        ':invno'           => $InvoiceNumber,
                        ':supval1'         => $SuplementaryValue,
                        ':procedure'       => $ProcedureCode,
                        ':extcode'         => $ExtendedCode,
                        ':itemgrossweight' => $ItemGrossWeight,
                        ':itemnetweight'   => $ItemNetWeight,
                        ':iteminvoicevalue'=> $ItemInvoiceValue,
                        ':quo_cod'               => $quo_cod,
                        ':quo_dsc'               => $quo_dsc,
                        ':valMethodNum'          => $ValMethodNum,
                        ':valMethodDesc'         => $ValMethodDesc,
                        ':Ocharges'              => $Ocharges,
                        ':IFreight'              => $IFreight,
                        ':InvCurr'               => $InvCurr,
                        ':Pref'                  => $Pref,
                        ':ProcDesc'              => $ProcDesc,
                        ':CoCode'                => $CoCode,
                        ':Regulated'             => $Regulated,
                        ':goodsdesc1'            => $goodsdesc1,
                        ':HSCode'                => $HSCode,
                        ':HSCode_Tar'            => $HSCode_Tar,
                        ':PTOPS_ROWID'           => $PTOPS_ROWID,
                        ':ecai_no_list'          => $ecai_no_list,
                    ]);
                                    
                } catch (PDOException $e1) {
                    echo "ERROR : " . $e1->getMessage();
                    die();
                }
            
            }
        
        }else{
    
            echo "<script>
                    alert('Cannot proceed, could not find Items/AdditionalCTN worksheet. Please check file.');
                    window.location.href='index.php?token=$token';
                </script>";
            die();
        }
        // ------------ END Items WORKSHEET ------------ //

        // ------------ START Financial WORKSHEET (MAPPING and DB INSERTION) ------------ //
        if($objWorksheet4){
            $highestColumn4 = $objWorksheet4->getHighestColumn();

            // Get the first data row (row 2)
            $dataRow4 = $objWorksheet4->rangeToArray('A2:'.$highestColumn4.'2', null, true, true, true);

            // Check if column A has a value
            if (isset($dataRow4[2]['A']) && $dataRow4[2]['A'] != '') {

                /*DEFAULT VALUES*/
                $BankCode       = "998";
                $BranchCode     = "N.A.";
                $CustomVal      = "300.00";
                $CustCurr       = "USD";
                $WharCurr       = "PHP";
                $ArrasCurr      = "PHP";
                $WOBankCharge   = "0";
                $Forex          = "0";
                $BRN            = "000000000-0000000";


                // Prepare and execute insert
                $insert_financial = "INSERT INTO tblEXPAPL_FIN (Applno, Tdelivery, Tpayment, BankCode, BranchCode, BankRef, CustomVal, CustCurr, WharCurr, ArrasCurr, WOBankCharge, Forex) 
                                VALUES (:applno, :tdelivery, :tpayment, :bankcode, :branchcode, :bankref, :customval, :custcurr, :wharcurr, :arrascurr, :wobankcharge, :forex)";
                
                try {
                    $stmt4 = $conn->connectIPPEZA()->prepare($insert_financial);
                    $stmt4->execute([
                        ':applno'       => $applNo,
                        ':tdelivery'    => $TermsOfDelivery, 
                        ':tpayment'     => $TermsOfPayment, 
                        ':bankcode'     => $BankCode, 
                        ':branchcode'   => $BranchCode, 
                        ':bankref'      => $BRN, 
                        ':customval'    => $CustomVal, 
                        ':custcurr'     => $CustCurr, 
                        ':wharcurr'     => $WharCurr, 
                        ':arrascurr'    => $ArrasCurr, 
                        ':wobankcharge' => $WOBankCharge, 
                        ':forex'        => $Forex, 
                    ]);
                } catch (PDOException $e3) {
                    echo "ERROR: " . $e3->getMessage();
                    die();
                }
            }
        }
        // ------------ END Financial WORKSHEET ------------ //


        // ------------ DB UPDATE - TBLEXPAPL_MASTER ------------ //

        $totalItems = 0;
        $totalPacks = 0;

        $totalCount = $validateFunc->__getTotalItems($applNo);

        if( isset($totalCount['totalItems']) && !empty($totalCount['totalItems']) )
        {
            $totalItems = $totalCount['totalItems'];
        }

        if( isset($totalCount['totalPacks']) && !empty($totalCount['totalPacks']) )
        {
            $totalPacks = number_format($totalCount['totalPacks']);
            $totalPacks = (int)preg_replace('/[^\d]/', '', $totalPacks);
        }

        //die(var_dump($totalPacks));

        $updateQuery = "UPDATE TBLEXPAPL_MASTER SET ItemCon = '$totalItems', Items = '$totalItems', Packs = '$totalPacks' WHERE ApplNo = '$applNo'";

        try{

            $sqlUpdate  = $updateQuery;
            $stmtUpdate = $conn->connectIPPEZA()->prepare($sqlUpdate);
            $stmtUpdate->execute();
                            
        } catch (PDOException $a) {
            echo "ERROR : " . $a->getMessage();
            die();
        }
    
        // ------------ END DB UPDATE - TBLEXPAPL_MASTER ------------ //

        // ------------ REMOVE UPLOADED EXCEL FILE ------------ //

        $getFilename = $processFunc->__getPHPExcelDetails($_FILES['file']['name']);
        unlink($getFilename['inputFile']);

        // ------------ END REMOVE UPLOADED EXCEL FILE ------------ //
        
        // ------------ REMOVE USED SESSION FLOW ------------ //

        unset($_SESSION['flows'][$token]);
        
        // ------------ END REMOVE USED SESSION FLOW ------------ //

        echo "<script>
                window.location.href='index.php?redirection=$redirection&msg=success&applno=$applNo';
            </script>";


    }