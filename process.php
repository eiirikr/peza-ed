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


    $isBulkClient = (strtoupper(trim($cltcode)) === "BEAEROBV");

    if ($isBulkClient) {

        // ================================================================
        // BULK MODE (BEAEROBV): single 'Appl' sheet, one row = one application
        // ================================================================


    $objWorksheetAppl = $objPHPExcel->getSheetByName('Appl');

    if (!$objWorksheetAppl) {
        echo "<script>
                alert('Cannot proceed. Could not find the Appl worksheet. Please check file.');
                window.location.href='index.php?token=$token';
            </script>";
        unlink($excelDetails['inputFile']);
        die();
    }

    $highestColumnAppl = $objWorksheetAppl->getHighestColumn();
    $highestRowAppl     = $objWorksheetAppl->getHighestRow();

    if (strtoupper($highestColumnAppl) != 'AE') {
        echo "<script>
                alert('File content is not compatible. Please check the Appl sheet.');
                window.location.href='index.php?token=$token';
            </script>";
        unlink($excelDetails['inputFile']);
        die();
    }

    $checkA2 = $objWorksheetAppl->getCell('A2')->getValue();
    if ($checkA2 == NULL || $checkA2 == '') {
        echo "<script>
                alert('Cannot proceed. Please check file.');
                window.location.href='index.php?token=$token';
            </script>";
        unlink($excelDetails['inputFile']);
        die();
    }

    /*
    |--------------------------------------------------------------------------
    | Item lookup
    |--------------------------------------------------------------------------
    */
    function __lookupItemByPTOPSRowID($conn, $ptopsRowId, $allAccIDs) {

        $accIds = array_filter(array_map('trim', explode(',', $allAccIDs)));
        if (empty($accIds) || $ptopsRowId === '') {
            return null;
        }

        $placeholders = implode(',', array_fill(0, count($accIds), '?'));

        $sql = "SELECT TOP 1
                    t.PTOPS_ROWID   AS PTOPS_ROWID,
                    t.HSCode        AS HsCode,
                    t.HSCode_tar    AS HsCode_Tar,
                    t.commoditydesc AS commodityDesc,
                    t.commoditycode AS commodityCode,
                    t.Status        AS status,
                    t.ecai_no       AS ecai_no,
                    g.uom_cod1      AS uom_cod1
                FROM dbo.tblExItem t
                LEFT OUTER JOIN PEZA.dbo.GBTARTAB g
                    ON t.HSCode = g.hs6_cod + g.tar_pr1
                    AND t.HSCode_Tar = g.tar_pr2
                WHERE t.PTOPS_ROWID = ?
                    AND t.accreditation_id IN ($placeholders)";

        $params = array_merge([$ptopsRowId], $accIds);

        $stmt = $conn->connectPEZAexpPTOPS()->prepare($sql);
        $stmt->execute($params);
        $result = $stmt->fetch(PDO::FETCH_ASSOC);

        return $result ? $result : null;
    }

    $r = -1;
    $applicationRowsAppl = array(); // row numbers that hold a real application

    for ($row = 2; $row <= $highestRowAppl; ++$row) {

        $dataRow = $objWorksheetAppl->rangeToArray('A'.$row.':AE'.$row, null, true, true, true);

        if (trim($dataRow[$row]['A']) == '') {
            continue; // skip fully blank row
        }

        ++$r;
        $applicationRowsAppl[] = $row;

        $Consignee            = strtoupper($validateFunc->trim_val($dataRow[$row]['A']));
        $ConAdr1              = $validateFunc->trim_val($dataRow[$row]['B']);
        $ConAdr2              = $validateFunc->trim_val($dataRow[$row]['C']);
        $ConAdr3              = $validateFunc->trim_val($dataRow[$row]['D']);
        $Port                 = strtoupper(trim($validateFunc->trim_val($dataRow[$row]['E'])));
        $PurposeOfExportation = strtoupper(trim($validateFunc->trim_val($dataRow[$row]['F'])));
        $ManifestNo           = strtoupper($validateFunc->trim_val($dataRow[$row]['G']));
        $BillOfLading         = $validateFunc->trim_val($dataRow[$row]['H']);
        $VesselAircraft       = $validateFunc->trim_val2($dataRow[$row]['I']);
        $LocationOfGoods      = strtoupper(trim($validateFunc->trim_val2($dataRow[$row]['J'])));
        $ProvinceOfOrigin     = $validateFunc->trim_val2($dataRow[$row]['K']);
        $CountryOfDestination = strtoupper(trim($validateFunc->trim_val2($dataRow[$row]['L'])));
        $PortOfLoading        = strtoupper(trim($validateFunc->trim_val($dataRow[$row]['M'])));
        $PortOfDeparture      = strtoupper(trim($validateFunc->trim_val($dataRow[$row]['N'])));
        $ContainerNumber      = strtoupper(trim($validateFunc->trim_val($dataRow[$row]['O'])));
        $SealNumber           = strtoupper(trim($validateFunc->trim_val($dataRow[$row]['P'])));
        $ContainerSize        = strtoupper(trim($validateFunc->trim_val($dataRow[$row]['Q'])));
        $ExItemID             = trim($validateFunc->trim_val($dataRow[$row]['R']));
        $Marks1                = strtoupper($validateFunc->trim_val($dataRow[$row]['S']));
        $Marks2                = strtoupper($validateFunc->trim_val($dataRow[$row]['T']));
        $NumberOfPackage      = $validateFunc->trim_val($dataRow[$row]['U']);
        $PackageCode          = $validateFunc->trim_val($dataRow[$row]['V']);
        $InvoiceNumber        = $validateFunc->trim_val($dataRow[$row]['W']);
        $SuplementaryValue    = $validateFunc->trim_val($dataRow[$row]['X']);
        $ProcedureCode        = $validateFunc->trim_val($dataRow[$row]['Y']);
        $ExtendedCode         = $validateFunc->trim_val($dataRow[$row]['Z']);
        $ItemGrossWeight      = $validateFunc->trim_val($dataRow[$row]['AA']);
        $ItemNetWeight        = $validateFunc->trim_val($dataRow[$row]['AB']);
        $ItemInvoiceValue     = $validateFunc->trim_val($dataRow[$row]['AC']);
        $TermsOfDelivery      = strtoupper($validateFunc->trim_val($dataRow[$row]['AD']));
        $TermsOfPayment       = strtoupper($validateFunc->trim_val($dataRow[$row]['AE']));

        /* ============================ VALIDATE ============================ */

        // Consignee
        if( $validateFunc->max_length($Consignee, 70) )
        {
            $consignee[] = $row - 1;
            $errorCounter1++;
        }
        if( !empty($Consignee) && ($validateFunc->match_char($Consignee)) == 0 )
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

        if( $validateFunc->max_length($ConAdr1, 35) ) { $conAdr1Len[] = $row - 1; $errorCounter1++; }
        if( $validateFunc->max_length($ConAdr2, 35) ) { $conAdr2Len[] = $row - 1; $errorCounter1++; }
        if( $validateFunc->max_length($ConAdr3, 35) ) { $conAdr3Len[] = $row - 1; $errorCounter1++; }

        // Port (Office of Clearance)
        if( !empty($Port) )
        {
            $checkOfficeOfClearanceExists = $validateFunc->__checkValidPortOfDeparture($Port);

            if( !$checkOfficeOfClearanceExists )
            {
                $checkOfficeOfClearance[] = $row - 1;
                $errorCounter++;
            } else {
                if ($checkOfficeOfClearanceExists['offClrMode'] === "BY AIR")
                {
                    $checkModeOfTransportation[] = $row - 1;
                }
            }
        } else {
            $checkOfficeOfClearance[] = $row - 1;
            $errorCounter++;
        }

        // PurposeOfExportation
        if( !empty($PurposeOfExportation) )
        {
            $checkPurposeOfExportationExists = $validateFunc->__checkValidPurposeOfExportation($PurposeOfExportation);
            if( !$checkPurposeOfExportationExists )
            {
                $checkPurposeOfExportation[] = $row - 1;
                $errorCounter++;
            }
        } else {
            $checkPurposeOfExportation[] = $row - 1;
            $errorCounter++;
        }

        // ManifestNo
        if( !empty($ManifestNo) && ($validateFunc->match_manifestFormat($ManifestNo)) == 0 )
        {
            $manifestNoMatch[] = $row - 1;
            $errorCounter++;
        }

        // BillOfLading
        if( empty($BillOfLading) )
        {
            $billOfLadingRequired[] = $row - 1;
            $errorCounter++;
        }
        else if( $validateFunc->max_length($BillOfLading, 26) )
        {
            $billOfLading[] = $row - 1;
            $errorCounter++;
        }
        if( !empty($BillOfLading) && ($validateFunc->match_alphanum($BillOfLading)) == 0 )
        {
            $billOfLadingMatch[] = $row - 1;
            $errorCounter++;
        }

        // Same broker-specific AWB check-digit validation as the general flow
        $billOfLadingInvalid = isset($billOfLadingInvalid) ? $billOfLadingInvalid : [];

        $requiredBrokers = [
            "200615811", "200615811000", "215722696", "215722696000",
            "225879904", "225879904000", "204867435", "204867435000",
            "432899304", "432899304000", "738464204", "738464204000"
        ];

        if (in_array(trim($locbroktin), $requiredBrokers))
        {
            $BillOfLading = trim($BillOfLading);

            if ($BillOfLading == "")
            {
                $billOfLadingInvalid[] = $row - 1;
                $errorCounter++;
            }
            else
            {
                if (
                    preg_match('/^(\d)\1{6,9}$/', $BillOfLading) ||
                    $BillOfLading == "1111111116"
                )
                {
                    $billOfLadingInvalid[] = $row - 1;
                    $errorCounter++;
                }
                else if (!preg_match('/^\d{10}$/', $BillOfLading))
                {
                    $billOfLadingInvalid[] = $row - 1;
                    $errorCounter++;
                }
                else
                {
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

        // VesselAircraft
        if( empty($VesselAircraft) )
        {
            $vesselAircraftRequired[] = $row - 1;
            $errorCounter++;
        }
        else if( $validateFunc->max_length($VesselAircraft, 27) )
        {
            $vesselAircraft[] = $row - 1;
            $errorCounter++;
        }
        if( !empty($VesselAircraft) && !preg_match('/^[A-Za-z0-9 ]+$/', $VesselAircraft) )
        {
            $vesselAircraftMatch[] = $row - 1;
            $errorCounter++;
        }

        // LocationOfGoods
        if( !empty($LocationOfGoods) ) {
            $checkLocationOfGoodsExists = $validateFunc->__checkValidLocationOfGoods($LocationOfGoods);
            if( !$checkLocationOfGoodsExists )
            {
                $checkLocationOfGoods[] = $row - 1;
                $errorCounter++;
            }
        } else {
            $checkLocationOfGoods[] = $row - 1;
            $errorCounter++;
        }

        // ProvinceOfOrigin
        if( !empty($ProvinceOfOrigin) ) {
            $checkProvinceOfOriginExists = $validateFunc->__checkValidProvinceOfOrigin($ProvinceOfOrigin);
            if( !$checkProvinceOfOriginExists )
            {
                $checkProvinceOfOrigin[] = $row - 1;
                $errorCounter++;
            }
        } else {
            $checkProvinceOfOrigin[] = $row - 1;
            $errorCounter++;
        }

        // CountryOfDestination
        if( !empty($CountryOfDestination) ) {
            $checkCountryOfDestinationExists = $validateFunc->__checkValidCountryOfDestination($CountryOfDestination);
            if( !$checkCountryOfDestinationExists )
            {
                $checkCountryOfDestination[] = $row - 1;
                $errorCounter++;
            }
        } else {
            $checkCountryOfDestination[] = $row - 1;
            $errorCounter++;
        }

        // PortOfLoading
        if( !empty($PortOfLoading) ) {
            $checkPortOfLoadingExists = $validateFunc->__checkValidPortOfLoading($PortOfLoading);
            if( !$checkPortOfLoadingExists )
            {
                $checkPortOfLoading[] = $row - 1;
                $errorCounter++;
            }
        } else {
            $checkPortOfLoading[] = $row - 1;
            $errorCounter++;
        }

        // PortOfDeparture
        if( !empty($PortOfDeparture) ) {
            $checkPortOfDepartureExists = $validateFunc->__checkValidPortOfDeparture($PortOfDeparture);
            if( !$checkPortOfDepartureExists )
            {
                $checkPortOfDeparture[] = $row - 1;
                $errorCounter++;
            }
        } else {
            $checkPortOfDeparture[] = $row - 1;
            $errorCounter++;
        }

        // Container fields -- gated per-row by THIS row's own mode of transport
        $thisRowIsByAir = !empty($checkModeOfTransportation) && in_array($row - 1, $checkModeOfTransportation);

        if ($thisRowIsByAir)
        {
            if (!empty($ContainerNumber) || !empty($SealNumber) || !empty($ContainerSize))
            {
                $containerDetailsNotAllowed[] = $row - 1;
                $errorCounter++;
            }
        }
        else
        {
            if( empty($ContainerNumber) )
            {
                $containerNumberRequired[] = $row - 1;
                $errorCounter++;
            }
            else if( $validateFunc->max_length($ContainerNumber, 100) )
            {
                $containerNumber[] = $row - 1;
                $errorCounter1++;
            }
            else if( ($validateFunc->match_alphanum($ContainerNumber)) == 0 )
            {
                $containerNumberMatch[] = $row - 1;
                $errorCounter++;
            }

            if( empty($SealNumber) )
            {
                $sealNumberRequired[] = $row - 1;
                $errorCounter++;
            }
            else if( $validateFunc->max_length($SealNumber, 100) )
            {
                $sealNumber[] = $row - 1;
                $errorCounter1++;
            }
            else if( ($validateFunc->match_alphanum($SealNumber)) == 0 )
            {
                $sealNumberMatch[] = $row - 1;
                $errorCounter++;
            }

            if( empty($ContainerSize) )
            {
                $containerSizeRequired[] = $row - 1;
                $errorCounter++;
            }
            else
            {
                $checkContainerSizeExists = $validateFunc->__checkValidContainerSize($ContainerSize);
                if( !$checkContainerSizeExists )
                {
                    $checkContainerSize[] = $row - 1;
                    $errorCounter++;
                }
            }
        }

        $checkItemCodeExists = null;
        if( empty($ExItemID) )
        {
            $checkItemCode[] = $row - 1;
            $errorCounter++;
        }
        else if( ($validateFunc->match_numbers($ExItemID)) == 0 )
        {
            $checkItemCode[] = $row - 1;
            $errorCounter++;
        }
        else
        {
            $checkItemCodeExists = __lookupItemByPTOPSRowID($conn, $ExItemID, $allaccids);
            if( empty($checkItemCodeExists) )
            {
                $checkItemCode[] = $row - 1;
                $errorCounter++;
            }
            else
            {
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
            }
        }

        // Marks and Numbers
        if( empty($Marks1) )
        {
            $marksAndNumberMatch[] = $row - 1;
            $errorCounter++;
        }
        else if( $validateFunc->max_length($Marks1, 35) )
        {
            $marksAndNumber[] = $row - 1;
            $errorCounter1++;
        }
        if( !empty($Marks1) && !preg_match('/^[A-Za-z0-9 ]+$/', $Marks1) )
        {
            $marksAndNumberSpecialChar[] = $row - 1;
            $errorCounter++;
        }
        if( $validateFunc->max_length($Marks2, 35) )
        {
            $marks2Len[] = $row - 1;
            $errorCounter1++;
        }
        if( !empty($Marks2) && !preg_match('/^[A-Za-z0-9 ]+$/', $Marks2) )
        {
            $marks2Match[] = $row - 1;
            $errorCounter++;
        }

        // NumberOfPackage
        if( $validateFunc->max_length($NumberOfPackage, 10) )
        {
            $numberOfPackage[] = $row - 1;
            $errorCounter1++;
        }

        if ( $NumberOfPackage === '' )
        {
            $numberOfPackageMatch[] = $row - 1;
            $errorCounter++;
        }
        else if ( ($validateFunc->match_numbers($NumberOfPackage)) == 0 )
        {
            $numberOfPackageMatch[] = $row - 1;
            $errorCounter++;
        }
        else if ( (int)$NumberOfPackage <= 0 )
        {
            $numberOfPackageZero[] = $row - 1;
            $errorCounter++;
        }
        else if ( (float)$NumberOfPackage > 2147483647 )
        {
            $numberOfPackageTooLarge[] = $row - 1;
            $errorCounter++;
        }

        // PackageCode
        if( !empty($PackageCode) && $validateFunc->match_packcodeFormat($PackageCode) ) {
            $checkPackCodeExists = $validateFunc->__checkValidPackCode($PackageCode);
            if( !$checkPackCodeExists )
            {
                $checkPackCode[] = $row - 1;
                $errorCounter++;
            }
        } else {
            $checkPackCode[] = $row - 1;
            $errorCounter++;
        }

        // InvoiceNumber
        if( empty($InvoiceNumber) )
        {
            $invoiceNumberRequired[] = $row - 1;
            $errorCounter++;
        }
        else
        {
            if( $validateFunc->max_length($InvoiceNumber, 300) )
            {
                $invoiceNumber[] = $row - 1;
                $errorCounter1++;
            }
            if( ($validateFunc->match_char($InvoiceNumber)) == 0 )
            {
                $invoiceNumberMatch[] = $row - 1;
                $errorCounter++;
            }
        }

        // SuplementaryValue
        if( !empty($SuplementaryValue) )
        {
            if( $validateFunc->max_length($SuplementaryValue, 15) )
            {
                $suplementaryValueLength[] = $row - 1;
                $errorCounter++;
            }
            else if( ($validateFunc->match_numbers($SuplementaryValue)) == 0 )
            {
                $suplementaryValueMatch[] = $row - 1;
                $errorCounter++;
            }
        }

        // ProcedureCode
        if( !empty($ProcedureCode) ) {
            $checkProcedureCodeExists = $validateFunc->__checkValidNatlCode($ProcedureCode);
            if( !$checkProcedureCodeExists )
            {
                $checkProcedureCode[] = $row - 1;
                $errorCounter++;
            }
        } else {
            $checkProcedureCode[] = $row - 1;
            $errorCounter++;
        }

        // ExtendedCode
        if( !empty($ExtendedCode) ) {
            $checkExtendedCodeExists = $validateFunc->__checkValidExtCode($ExtendedCode);
            if( !$checkExtendedCodeExists )
            {
                $checkExtCode[] = $row - 1;
                $errorCounter++;
            }
        } else {
            $checkExtCode[] = $row - 1;
            $errorCounter++;
        }

        // ItemGrossWeight
        if( empty($ItemGrossWeight) )
        {
            $itemGrossWeightRequired[] = $row - 1;
            $errorCounter++;
        }
        else
        {
            if( $validateFunc->max_length($ItemGrossWeight, 10) )
            {
                $itemGrossWeightLength[] = $row - 1;
                $errorCounter++;
            }
            else if( !$validateFunc->match_weightFormat($ItemGrossWeight) || (float)$ItemGrossWeight <= 0 )
            {
                $itemGrossWeightMatch[] = $row - 1;
                $errorCounter++;
            }
        }

        // ItemNetWeight
        if( empty($ItemNetWeight) )
        {
            $itemNetWeightRequired[] = $row - 1;
            $errorCounter++;
        }
        else
        {
            if( $validateFunc->max_length($ItemNetWeight, 10) )
            {
                $itemNetWeightLength[] = $row - 1;
                $errorCounter++;
            }
            else if( !$validateFunc->match_weightFormat($ItemNetWeight) || (float)$ItemNetWeight <= 0 )
            {
                $itemNetWeightMatch[] = $row - 1;
                $errorCounter++;
            }
            else if( !empty($ItemGrossWeight) && $validateFunc->match_weightFormat($ItemGrossWeight) && (float)$ItemNetWeight > (float)$ItemGrossWeight )
            {
                $itemNetExceedsGross[] = $row - 1;
                $errorCounter++;
            }
        }

        // ItemInvoiceValue
        if( empty($ItemInvoiceValue) )
        {
            $itemInvoiceValueRequired[] = $row - 1;
            $errorCounter++;
        }
        else if( $validateFunc->max_length($ItemInvoiceValue, 11) )
        {
            $itemInvoiceValueLength[] = $row - 1;
            $errorCounter++;
        }
        else if( !$validateFunc->match_weightFormat($ItemInvoiceValue) || (float)$ItemInvoiceValue <= 0 )
        {
            $itemInvoiceValueMatch[] = $row - 1;
            $errorCounter++;
        }

        // TermsOfDelivery
        if( !empty($TermsOfDelivery) ) {
            $checkTermsOfDelivery = $validateFunc->__checkValidTermsOfDelivery($TermsOfDelivery);
            if( !$checkTermsOfDelivery )
            {
                $termsOfDelivery[] = $row - 1;
                $errorCounter++;
            }
        } else {
            $termsOfDelivery[] = $row - 1;
            $errorCounter++;
        }

        // TermsOfPayment
        if( !empty($TermsOfPayment) ) {
            $checkTermsOfPayment = $validateFunc->__checkValidTermsOfPayment($TermsOfPayment);
            if( !$checkTermsOfPayment )
            {
                $termsOfPayment[] = $row - 1;
                $errorCounter++;
            }
        } else {
            $termsOfPayment[] = $row - 1;
            $errorCounter++;
        }

    } // end validation loop

    if (empty($applicationRowsAppl)) {
        echo "<script>
                alert('Cannot proceed. Please check file.');
                window.location.href='index.php?token=$token';
            </script>";
        unlink($excelDetails['inputFile']);
        die();
    }

    if ($errorCounter > 0 || $errorCounter1 > 0) {

        /* ERROR MESSAGES */
        if(!empty($consignee)){ $errorLists[] = array("ErrMsg" => "Exceeds the max characters allowed (70)", "Column" => "Consignee", "Rows" => implode(", ", $consignee)); }
        if(!empty($consigneeMatch)){ $errorLists[] = array("ErrMsg" => "Only accept letters, numbers and few special characters (-_.,:;#$%()*/) - Required", "Column" => "Consignee", "Rows" => implode(", ", $consigneeMatch)); }
        if(!empty($checkConsignee)){ $errorLists[] = array("ErrMsg" => "Invalid Consignee/Buyer, please check Buyer Lookup", "Column" => "Consignee", "Rows" => implode(", ", $checkConsignee)); }

        if(!empty($conAdr1Len)){ $errorLists[] = array("ErrMsg" => "Exceeds the max characters allowed (35)", "Column" => "Con Address 1", "Rows" => implode(", ", $conAdr1Len)); }
        if(!empty($conAdr2Len)){ $errorLists[] = array("ErrMsg" => "Exceeds the max characters allowed (35)", "Column" => "Con Address 2", "Rows" => implode(", ", $conAdr2Len)); }
        if(!empty($conAdr3Len)){ $errorLists[] = array("ErrMsg" => "Exceeds the max characters allowed (35)", "Column" => "Con Address 3", "Rows" => implode(", ", $conAdr3Len)); }

        if(!empty($checkOfficeOfClearance)){ $errorLists[] = array("ErrMsg" => "Invalid Port", "Column" => "Office of Clearance", "Rows" => implode(", ", $checkOfficeOfClearance)); }
        if(!empty($checkPurposeOfExportation)){ $errorLists[] = array("ErrMsg" => "Invalid Purpose of Exportation", "Column" => "Purpose Of Exportation", "Rows" => implode(", ", $checkPurposeOfExportation)); }
        if(!empty($manifestNoMatch)){ $errorLists[] = array("ErrMsg" => "Invalid Manifest Number! Format: NNNMMMM-YY", "Column" => "Manifest Number", "Rows" => implode(", ", $manifestNoMatch)); }

        if(!empty($billOfLadingRequired)){ $errorLists[] = array("ErrMsg" => "Bill of Lading/Airbill is required", "Column" => "Bill of Lading", "Rows" => implode(", ", $billOfLadingRequired)); }
        if(!empty($billOfLading)){ $errorLists[] = array("ErrMsg" => "Exceeds the max characters allowed (26)", "Column" => "Bill of Lading", "Rows" => implode(", ", $billOfLading)); }
        if(!empty($billOfLadingMatch)){ $errorLists[] = array("ErrMsg" => "Only accept letters and numbers - Required", "Column" => "Bill of Lading", "Rows" => implode(", ", $billOfLadingMatch)); }
        if(!empty($billOfLadingInvalid)){ $errorLists[] = array("ErrMsg" => "Please enter a valid bill of lading/airway bill.", "Column" => "Bill of Lading", "Rows" => implode(", ", $billOfLadingInvalid)); }

        if(!empty($vesselAircraftRequired)){ $errorLists[] = array("ErrMsg" => "Vessel/Aircraft is required", "Column" => "Vessel / Aircraft", "Rows" => implode(", ", $vesselAircraftRequired)); }
        if(!empty($vesselAircraft)){ $errorLists[] = array("ErrMsg" => "Exceeds the max characters allowed (27)", "Column" => "Vessel / Aircraft", "Rows" => implode(", ", $vesselAircraft)); }
        if(!empty($vesselAircraftMatch)){ $errorLists[] = array("ErrMsg" => "Only letters, numbers, and spaces are allowed - no special characters", "Column" => "Vessel / Aircraft", "Rows" => implode(", ", $vesselAircraftMatch)); }

        if(!empty($checkLocationOfGoods)){ $errorLists[] = array("ErrMsg" => "Invalid Location of Goods", "Column" => "Location of Goods", "Rows" => implode(", ", $checkLocationOfGoods)); }
        if(!empty($checkProvinceOfOrigin)){ $errorLists[] = array("ErrMsg" => "Invalid Province of Origin", "Column" => "Province of Origin", "Rows" => implode(", ", $checkProvinceOfOrigin)); }
        if(!empty($checkCountryOfDestination)){ $errorLists[] = array("ErrMsg" => "Invalid Country of Destination", "Column" => "Country of Destination", "Rows" => implode(", ", $checkCountryOfDestination)); }
        if(!empty($checkPortOfLoading)){ $errorLists[] = array("ErrMsg" => "Invalid Port of Loading", "Column" => "Port of Loading", "Rows" => implode(", ", $checkPortOfLoading)); }
        if(!empty($checkPortOfDeparture)){ $errorLists[] = array("ErrMsg" => "Invalid Port of Departure", "Column" => "Port of Departure", "Rows" => implode(", ", $checkPortOfDeparture)); }

        if(!empty($containerNumberRequired)){ $errorLists[] = array("ErrMsg" => "Container Number is required", "Column" => "Container Number", "Rows" => implode(", ", $containerNumberRequired)); }
        if(!empty($containerNumber)){ $errorLists[] = array("ErrMsg" => "Exceeds the max characters allowed (100)", "Column" => "Container Number", "Rows" => implode(", ", $containerNumber)); }
        if(!empty($containerNumberMatch)){ $errorLists[] = array("ErrMsg" => "Only letters and numbers are allowed - no special characters", "Column" => "Container Number", "Rows" => implode(", ", $containerNumberMatch)); }
        if(!empty($sealNumberRequired)){ $errorLists[] = array("ErrMsg" => "Seal Number is required", "Column" => "Seal Number", "Rows" => implode(", ", $sealNumberRequired)); }
        if(!empty($sealNumber)){ $errorLists[] = array("ErrMsg" => "Exceeds the max characters allowed (100)", "Column" => "Seal Number", "Rows" => implode(", ", $sealNumber)); }
        if(!empty($sealNumberMatch)){ $errorLists[] = array("ErrMsg" => "Only letters and numbers are allowed - no special characters", "Column" => "Seal Number", "Rows" => implode(", ", $sealNumberMatch)); }
        if(!empty($containerSizeRequired)){ $errorLists[] = array("ErrMsg" => "Container Size is required", "Column" => "Container Size", "Rows" => implode(", ", $containerSizeRequired)); }
        if(!empty($checkContainerSize)){ $errorLists[] = array("ErrMsg" => "Invalid Container Size", "Column" => "Container Size", "Rows" => implode(", ", $checkContainerSize)); }
        if(!empty($containerDetailsNotAllowed)){ $errorLists[] = array("ErrMsg" => "Container details are not allowed for this mode of transport.", "Column" => "Container Number, Seal Number, Container Size", "Rows" => implode(", ", $containerDetailsNotAllowed)); }

        if(!empty($checkItemCode)){ $errorLists[] = array("ErrMsg" => "Invalid EX ITEM ID, please check Exportables Lookup", "Column" => "EX ITEM ID", "Rows" => implode(", ", $checkItemCode)); }

        if(!empty($marksAndNumberMatch)){ $errorLists[] = array("ErrMsg" => "Marks and Numbers 1 is required", "Column" => "Marks and Numbers 1", "Rows" => implode(", ", $marksAndNumberMatch)); }
        if(!empty($marksAndNumber)){ $errorLists[] = array("ErrMsg" => "Exceeds the max characters allowed (35)", "Column" => "Marks and Numbers 1", "Rows" => implode(", ", $marksAndNumber)); }
        if(!empty($marksAndNumberSpecialChar)){ $errorLists[] = array("ErrMsg" => "Only letters, numbers, and spaces are allowed - no special characters", "Column" => "Marks and Numbers 1", "Rows" => implode(", ", $marksAndNumberSpecialChar)); }
        if(!empty($marks2Len)){ $errorLists[] = array("ErrMsg" => "Exceeds the max characters allowed (35)", "Column" => "Marks and Numbers 2", "Rows" => implode(", ", $marks2Len)); }
        if(!empty($marks2Match)){ $errorLists[] = array("ErrMsg" => "Only letters, numbers, and spaces are allowed - no special characters", "Column" => "Marks and Numbers 2", "Rows" => implode(", ", $marks2Match)); }

        if(!empty($numberOfPackage)){ $errorLists[] = array("ErrMsg" => "Exceeds the max characters allowed (10)", "Column" => "Number of Package", "Rows" => implode(", ", $numberOfPackage)); }
        if(!empty($numberOfPackageMatch)){ $errorLists[] = array("ErrMsg" => "Number of Package is required and must contain numbers only", "Column" => "Number of Package", "Rows" => implode(", ", $numberOfPackageMatch)); }
        if(!empty($numberOfPackageZero)){ $errorLists[] = array("ErrMsg" => "Number of Package must be greater than 0", "Column" => "Number of Package", "Rows" => implode(", ", $numberOfPackageZero)); }
        if(!empty($numberOfPackageTooLarge)){ $errorLists[] = array("ErrMsg" => "Number of Package exceeds the maximum allowed value (2,147,483,647)", "Column" => "Number of Package", "Rows" => implode(", ", $numberOfPackageTooLarge)); }

        if(!empty($checkPackCode)){ $errorLists[] = array("ErrMsg" => "Invalid Package Code", "Column" => "Package Code", "Rows" => implode(", ", $checkPackCode)); }

        if(!empty($invoiceNumberRequired)){ $errorLists[] = array("ErrMsg" => "Invoice Number is required", "Column" => "Invoice Number", "Rows" => implode(", ", $invoiceNumberRequired)); }
        if(!empty($invoiceNumber)){ $errorLists[] = array("ErrMsg" => "Exceeds the max characters allowed (300)", "Column" => "Invoice Number", "Rows" => implode(", ", $invoiceNumber)); }
        if(!empty($invoiceNumberMatch)){ $errorLists[] = array("ErrMsg" => "Only letters, numbers, and spaces are allowed - no special characters", "Column" => "Invoice Number", "Rows" => implode(", ", $invoiceNumberMatch)); }

        if(!empty($suplementaryValueLength)){ $errorLists[] = array("ErrMsg" => "Exceeds the max characters allowed (15)", "Column" => "Supplementary Value", "Rows" => implode(", ", $suplementaryValueLength)); }
        if(!empty($suplementaryValueMatch)){ $errorLists[] = array("ErrMsg" => "Invalid format. Only whole numbers are allowed - decimals and scientific notation (e.g. 1E+10) are not accepted", "Column" => "Supplementary Value", "Rows" => implode(", ", $suplementaryValueMatch)); }
        if(!empty($checkSuplementaryValue)){ $errorLists[] = array("ErrMsg" => "Supplementary Value required for the following item", "Column" => "Supplementary Value", "Rows" => implode(", ", $checkSuplementaryValue)); }
        if(!empty($checkSuplementaryValue1)){ $errorLists[] = array("ErrMsg" => "Supplementary Value is not allowed for the following item", "Column" => "Supplementary Value", "Rows" => implode(", ", $checkSuplementaryValue1)); }

        if(!empty($checkProcedureCode)){ $errorLists[] = array("ErrMsg" => "Invalid or missing Procedure Code", "Column" => "Procedure Code", "Rows" => implode(", ", $checkProcedureCode)); }
        if(!empty($checkExtCode)){ $errorLists[] = array("ErrMsg" => "Invalid or missing Extended Code", "Column" => "Extended Code", "Rows" => implode(", ", $checkExtCode)); }

        if(!empty($itemGrossWeightRequired)){ $errorLists[] = array("ErrMsg" => "Item Gross Weight is required", "Column" => "Item Gross Weight", "Rows" => implode(", ", $itemGrossWeightRequired)); }
        if(!empty($itemNetWeightRequired)){ $errorLists[] = array("ErrMsg" => "Item Net Weight is required", "Column" => "Item Net Weight", "Rows" => implode(", ", $itemNetWeightRequired)); }
        if(!empty($itemGrossWeightMatch)){ $errorLists[] = array("ErrMsg" => "Invalid entry (e.g. 0.00) - Required", "Column" => "Item Gross Weight", "Rows" => implode(", ", $itemGrossWeightMatch)); }
        if(!empty($itemGrossWeightLength)){ $errorLists[] = array("ErrMsg" => "Exceeds the max characters allowed (10)", "Column" => "Item Gross Weight", "Rows" => implode(", ", $itemGrossWeightLength)); }
        if(!empty($itemNetWeightMatch)){ $errorLists[] = array("ErrMsg" => "Invalid entry (e.g. 0.00) - Required", "Column" => "Item Net Weight", "Rows" => implode(", ", $itemNetWeightMatch)); }
        if(!empty($itemNetWeightLength)){ $errorLists[] = array("ErrMsg" => "Exceeds the max characters allowed (10)", "Column" => "Item Net Weight", "Rows" => implode(", ", $itemNetWeightLength)); }
        if(!empty($itemNetExceedsGross)){ $errorLists[] = array("ErrMsg" => "Item Net Weight must not exceed Item Gross Weight", "Column" => "Item Net Weight", "Rows" => implode(", ", $itemNetExceedsGross)); }

        if(!empty($itemInvoiceValueRequired)){ $errorLists[] = array("ErrMsg" => "Item Invoice Value is required", "Column" => "Item Invoice Value", "Rows" => implode(", ", $itemInvoiceValueRequired)); }
        if(!empty($itemInvoiceValueLength)){ $errorLists[] = array("ErrMsg" => "Exceeds the max characters allowed (11)", "Column" => "Item Invoice Value", "Rows" => implode(", ", $itemInvoiceValueLength)); }
        if(!empty($itemInvoiceValueMatch)){ $errorLists[] = array("ErrMsg" => "Invalid entry (e.g. 1000.00) - Required", "Column" => "Item Invoice Value", "Rows" => implode(", ", $itemInvoiceValueMatch)); }

        if(!empty($termsOfDelivery)){ $errorLists[] = array("ErrMsg" => "Invalid or missing Terms of Delivery code", "Column" => "Terms Of Delivery", "Rows" => implode(", ", $termsOfDelivery)); }
        if(!empty($termsOfPayment)){ $errorLists[] = array("ErrMsg" => "Invalid or missing Terms of Payment code", "Column" => "Terms Of Payment", "Rows" => implode(", ", $termsOfPayment)); }

        $_SESSION['errormsg'] = $errorLists;
        $_SESSION['required'] = $errorCounter;
        $_SESSION['proceed']  = $errorCounter1;

        echo "<script>
                window.location.href='index.php?msg=error&token=$token';
            </script>";
        die();

    } else {

        /* ============================ PROCESS / INSERT ============================ */

        $objReader      = PHPExcel_IOFactory::createReader($excelDetails["type"]);
        $objReader->setReadDataOnly(true);
        $objPHPExcel    = $objReader->load($excelDetails["inputFile"]);
        $objWorksheetAppl = $objPHPExcel->getSheetByName('Appl');

        if (!$objWorksheetAppl) {
            echo "<script>
                    alert('Cannot proceed. Please check file.');
                    window.location.href='index.php?token=$token';
            </script>";
            die();
        }

        $generatedApplNos = array();

        foreach ($applicationRowsAppl as $row) {

            $dataRow = $objWorksheetAppl->rangeToArray('A'.$row.':AE'.$row, null, true, true, true);

            $Consignee            = strtoupper($validateFunc->trim_val($dataRow[$row]['A']));
            $ConAdr1              = $validateFunc->trim_val($dataRow[$row]['B']);
            $ConAdr2              = $validateFunc->trim_val($dataRow[$row]['C']);
            $ConAdr3              = $validateFunc->trim_val($dataRow[$row]['D']);
            $Port                 = strtoupper(trim($validateFunc->trim_val($dataRow[$row]['E'])));
            $PurposeOfExportation = strtoupper(trim($validateFunc->trim_val($dataRow[$row]['F'])));
            $ManifestNo           = strtoupper($validateFunc->trim_val($dataRow[$row]['G']));
            $BillOfLading         = $validateFunc->trim_val($dataRow[$row]['H']);
            $VesselAircraft       = $validateFunc->trim_val2($dataRow[$row]['I']);
            $LocationOfGoods      = strtoupper(trim($validateFunc->trim_val2($dataRow[$row]['J'])));
            $ProvinceOfOrigin     = $validateFunc->trim_val2($dataRow[$row]['K']);
            $CountryOfDestination = strtoupper(trim($validateFunc->trim_val2($dataRow[$row]['L'])));
            $PortOfLoading        = strtoupper(trim($validateFunc->trim_val($dataRow[$row]['M'])));
            $PortOfDeparture      = strtoupper(trim($validateFunc->trim_val($dataRow[$row]['N'])));
            $ContainerNumber      = strtoupper(trim($validateFunc->trim_val($dataRow[$row]['O'])));
            $SealNumber           = strtoupper(trim($validateFunc->trim_val($dataRow[$row]['P'])));
            $ContainerSize        = strtoupper(trim($validateFunc->trim_val($dataRow[$row]['Q'])));
            $ExItemID             = trim($validateFunc->trim_val($dataRow[$row]['R']));
            $Marks1                = strtoupper($validateFunc->trim_val($dataRow[$row]['S']));
            $Marks2                = strtoupper($validateFunc->trim_val($dataRow[$row]['T']));
            $NumberOfPackage      = strtoupper($validateFunc->trim_val($dataRow[$row]['U']));
            $PackageCode          = strtoupper($validateFunc->trim_val($dataRow[$row]['V']));
            $InvoiceNumber        = strtoupper($validateFunc->trim_val($dataRow[$row]['W']));
            $SuplementaryValue    = strtoupper($validateFunc->trim_val($dataRow[$row]['X']));
            $ProcedureCode        = strtoupper($validateFunc->trim_val($dataRow[$row]['Y']));
            $ExtendedCode         = strtoupper($validateFunc->trim_val($dataRow[$row]['Z']));
            $ItemGrossWeight      = $validateFunc->trim_val($dataRow[$row]['AA']);
            $ItemNetWeight        = $validateFunc->trim_val($dataRow[$row]['AB']);
            $ItemInvoiceValue     = $validateFunc->trim_val($dataRow[$row]['AC']);
            $TermsOfDelivery      = strtoupper($validateFunc->trim_val($dataRow[$row]['AD']));
            $TermsOfPayment       = strtoupper($validateFunc->trim_val($dataRow[$row]['AE']));

            $applNo = $validateFunc->generateApplNo($conn, $csncod);
            $generatedApplNos[] = $applNo;

            // ------------ tblEXPAPL_Master ------------ //

            $forwarder = $lookupData->getForwarders($conn, $lstexporter);
            $importer  = $lookupData->getImporters($conn, $loccod);
            $exchRate  = $lookupData->getExchangeRate($conn, 'USD');
            $modeofTransport = $lookupData->getModeofTransport($conn, $Port);

            $insert_master = "INSERT INTO tblEXPAPL_Master (Applno, ConName, ConAdr1, ConAdr2, ConAdr3, OffClear, Manifest, Waybill, DECTIN, DECname, DecAdr1, DecAdr2, DecAdr3, Cexp, Cdest, Vessel, ExpCode, ExpName, ExpAdr1, ExpAdr2, RegOfc, mdec, mdec2, Exhrate, PortofLoad, PortofDept, ProvofOrig, CreationDate, Stat, ConTIN, IAN, LGoods, Purpose, cltcode, SenderID, modeOfTransport, isExcelFileAppl) 
                            VALUES (:applno, :conname, :conadr1, :conadr2, :conadr3, :offclear, :manifest, :waybill, :dectin, :decname, :decadr1, :decadr2, :decadr3, :cexp, :cdest, :vessel, :expcode, :expname, :expadr1, :expadr2, :regofc, :mdec, :mdec2, :exhrate, :portofload, :portofdept, :provoforig, :creationdate, :stat, :contin, :ian, :lgoods, :purpose, :cltcode, :senderid, :modeoftransport, :isExcelFileAppl)";

            try {
                $stmt3 = $conn->connectIPPEZA()->prepare($insert_master);
                $stmt3->execute([
                    ':applno'           => $applNo,
                    ':conname'          => $Consignee,
                    ':conadr1'          => $ConAdr1,
                    ':conadr2'          => $ConAdr2,
                    ':conadr3'          => $ConAdr3,
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
                    ':stat'             => 'C',
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

            // ------------ tblEXPAPL_ContPEZA (only if by sea) ------------ //

            $isByAir = ($modeofTransport['offClrMode'] === "BY AIR");

            if (!$isByAir) {
                try {
                    $stmt = $conn->connectIPPEZA()->prepare("INSERT INTO tblEXPAPL_ContPEZA (Applno, Container, Seal, ContainerSize, ModeOfShipment) VALUES (:applno, :container, :seal, :containerSize, :modeOfShipment)");
                    $stmt->execute([
                        ':applno'        => $applNo,
                        ':container'     => $ContainerNumber,
                        ':seal'          => $SealNumber,
                        ':containerSize' => $ContainerSize,
                        ':modeOfShipment'=> "FCL"
                    ]);
                } catch (PDOException $e) {
                    echo "ERROR: " . $e->getMessage();
                    die();
                }
            }

            // ------------ TBLEXPAPL_DETAIL (single item, resolved by PTOPS_ROWID) ------------ //

            $checkItemCodeExists = __lookupItemByPTOPSRowID($conn, $ExItemID, $allaccids);

            $isRegulated = "";
            if ($checkItemCodeExists['status'] == "M") {
                $isRegulated = "True";
            } else if ($checkItemCodeExists['status'] == "A") {
                $isRegulated = "False";
            }

            $Regulated    = $isRegulated;
            $goodsdesc1   = $checkItemCodeExists['commodityDesc'];
            $HSCode       = $checkItemCodeExists['HsCode'];
            $HSCode_Tar   = $checkItemCodeExists['HsCode_Tar'];
            $PTOPS_ROWID  = $checkItemCodeExists['PTOPS_ROWID'];
            $ecai_no_list = $checkItemCodeExists['ecai_no'];
            $ItemCodeText = $checkItemCodeExists['commodityCode'];

            $quo_cod       = 'NNNNN';
            $quo_dsc       = 'NOT RELATED, NO RSTRCTN/CNDTN/RYLTS/ARRNGMNTS';
            $ValMethodNum  = '1';
            $ValMethodDesc = 'TRANSACTION VALUE';
            $Ocharges      = '0';
            $IFreight      = '0';
            $InvCurr       = 'USD';
            $Pref          = 'NONE';
            $ProcDesc      = $ProcedureCode;
            $CoCode        = 'PH';

            $ItemGrossWeight  = ($ItemGrossWeight === '') ? '' : number_format(round((float)str_replace(',', '', $ItemGrossWeight), 2), 2, '.', '');
            $ItemNetWeight    = ($ItemNetWeight === '') ? '' : number_format(round((float)str_replace(',', '', $ItemNetWeight), 2), 2, '.', '');
            $ItemInvoiceValue = number_format(round((float)str_replace(',', '', $ItemInvoiceValue), 2), 2, '.', '');

            $itemNo = 1; // single-item lodgement: always item #1 of its own new application

            $insert_sql1 = "INSERT INTO TBLEXPAPL_DETAIL (ApplNo, ItemNo, itemcode, Marks1, Marks2, NoPack, PackCode, InvNo, SupVal1, [Procedure], ExtCode, ItemGWeight, ItemNWeight, InvValue, quo_cod, quo_dsc, ValMethodNum, ValMethodDesc, Ocharges, IFreight, InvCurr, Pref, ProcDesc, CoCode, Regulated, goodsdesc1, HSCode, HSCode_Tar, PTOPS_ROWID, ecai_no_list)
                            VALUES (:applno, :itemNo, :itemcode, :marks1, :marks2, :nopack, :packcode, :invno, :supval1, :procedure, :extcode, :itemgrossweight, :itemnetweight, :iteminvoicevalue, :quo_cod, :quo_dsc, :valMethodNum, :valMethodDesc, :Ocharges, :IFreight, :InvCurr, :Pref, :ProcDesc, :CoCode, :Regulated, :goodsdesc1, :HSCode, :HSCode_Tar, :PTOPS_ROWID, :ecai_no_list)";

            try {
                $stmt1 = $conn->connectIPPEZA()->prepare($insert_sql1);
                $stmt1->execute([
                    ':applno'          => $applNo,
                    ':itemNo'          => $itemNo,
                    ':itemcode'        => $ItemCodeText,
                    ':marks1'          => $Marks1,
                    ':marks2'          => $Marks2,
                    ':nopack'          => $NumberOfPackage,
                    ':packcode'        => $PackageCode,
                    ':invno'           => $InvoiceNumber,
                    ':supval1'         => $SuplementaryValue,
                    ':procedure'       => $ProcedureCode,
                    ':extcode'         => $ExtendedCode,
                    ':itemgrossweight' => ($ItemGrossWeight === '') ? null : $ItemGrossWeight,
                    ':itemnetweight'   => ($ItemNetWeight === '') ? null : $ItemNetWeight,
                    ':iteminvoicevalue'=> $ItemInvoiceValue,
                    ':quo_cod'         => $quo_cod,
                    ':quo_dsc'         => $quo_dsc,
                    ':valMethodNum'    => $ValMethodNum,
                    ':valMethodDesc'   => $ValMethodDesc,
                    ':Ocharges'        => $Ocharges,
                    ':IFreight'        => $IFreight,
                    ':InvCurr'         => $InvCurr,
                    ':Pref'            => $Pref,
                    ':ProcDesc'        => $ProcDesc,
                    ':CoCode'          => $CoCode,
                    ':Regulated'       => $Regulated,
                    ':goodsdesc1'      => $goodsdesc1,
                    ':HSCode'          => $HSCode,
                    ':HSCode_Tar'      => $HSCode_Tar,
                    ':PTOPS_ROWID'     => $PTOPS_ROWID,
                    ':ecai_no_list'    => $ecai_no_list,
                ]);
            } catch (PDOException $e1) {
                echo "ERROR : " . $e1->getMessage();
                die();
            }

            // ------------ tblEXPAPL_FIN ------------ //

            $BankCode     = "998";
            $BranchCode   = "N.A.";
            $CustomVal    = "300.00";
            $CustCurr     = "USD";
            $WharCurr     = "PHP";
            $ArrasCurr    = "PHP";
            $WOBankCharge = "0";
            $Forex        = "0";
            $BRN          = "000000000-0000000";

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

            // ------------ Update totals for THIS application ------------ //

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

            $updateQuery = "UPDATE TBLEXPAPL_MASTER SET ItemCon = '$totalItems', Items = '$totalItems', Packs = '$totalPacks' WHERE ApplNo = '$applNo'";

            try {
                $stmtUpdate = $conn->connectIPPEZA()->prepare($updateQuery);
                $stmtUpdate->execute();
            } catch (PDOException $a) {
                echo "ERROR : " . $a->getMessage();
                die();
            }

        } // end foreach application row

        // ------------ REMOVE UPLOADED EXCEL FILE / SESSION FLOW ------------ //

        $getFilename = $processFunc->__getPHPExcelDetails($_FILES['file']['name']);
        unlink($getFilename['inputFile']);

        unset($_SESSION['flows'][$token]);

        $applNoList = implode(',', $generatedApplNos);

        echo "<script>
                window.location.href='index.php?redirection=$redirection&msg=success&applno=$applNoList&count=" . count($generatedApplNos) . "&mode=bulk';
            </script>";

    }

    } else {

        // ================================================================
        // GENERAL MODE
        // ================================================================


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
                        $ManifestNo             =   strtoupper($validateFunc->trim_val($dataRow[$row]['E']));
                        $BillOfLading           =   $validateFunc->trim_val($dataRow[$row]['F']);
                        $VesselAircraft         =   $validateFunc->trim_val2($dataRow[$row]['G']);
                        $LocationOfGoods        =   strtoupper($validateFunc->trim_val2($dataRow[$row]['H']));
                        $ProvinceOfOrigin       =   $validateFunc->trim_val2($dataRow[$row]['I']);
                        $CountryOfDestination   =   strtoupper($validateFunc->trim_val2($dataRow[$row]['J']));
                        $PortOfLoading          =   strtoupper($validateFunc->trim_val($dataRow[$row]['K']));
                        $PortOfDeparture        =   strtoupper($validateFunc->trim_val($dataRow[$row]['L']));
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
                    if( empty($BillOfLading) )
                    {
                        $billOfLadingRequired[] = $row - 1;
                        $errorCounter++;
                    }
                    else if( $validateFunc->max_length($BillOfLading, 26) ) 
                    { 
                        $billOfLading[] = $row - 1; 
                        $errorCounter++; 
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
                    if( empty($VesselAircraft) )
                    {
                        $vesselAircraftRequired[] = $row - 1;
                        $errorCounter++;
                    }
                    else if( $validateFunc->max_length($VesselAircraft, 27) ) 
                    { 
                        $vesselAircraft[] = $row - 1; 
                        $errorCounter++; 
                    }
                    if( !empty($VesselAircraft) && !preg_match('/^[A-Za-z0-9 ]+$/', $VesselAircraft) )
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
                        if( empty($ContainerNumber) )
                        {
                            $containerNumberRequired[] = $row - 1;
                            $errorCounter++;
                        }
                        else if( $validateFunc->max_length($ContainerNumber, 100) ) 
                        { 
                            $containerNumber[] = $row - 1; 
                            $errorCounter1++; 
                        }
                        else if( ($validateFunc->match_alphanum($ContainerNumber)) == 0 ) 
                        { 
                            $containerNumberMatch[] = $row - 1; 
                            $errorCounter++; 
                        }

                        //SealNumber
                        if( empty($SealNumber) )
                        {
                            $sealNumberRequired[] = $row - 1;
                            $errorCounter++;
                        }
                        else if( $validateFunc->max_length($SealNumber, 100) ) 
                        { 
                            $sealNumber[] = $row - 1; $errorCounter1++; 
                        }
                        else if( ($validateFunc->match_alphanum($SealNumber)) == 0 )
                        {
                            $sealNumberMatch[] = $row - 1; 
                            $errorCounter++; 
                        }

                        //ContainerSize
                        if( empty($ContainerSize) )
                        {
                            $containerSizeRequired[] = $row - 1;
                            $errorCounter++;
                        }
                        else if( ($validateFunc->match_alphanum($ContainerSize)) == 0 )
                        {
                            $containerSizeMatch[] = $row - 1;
                            $errorCounter++;
                        }
                        else 
                        { 
                            $checkContainerSizeExists = $validateFunc->__checkValidContainerSize($ContainerSize);

                            if( !$checkContainerSizeExists ) 
                            {
                                $checkContainerSize[] = $row - 1; 
                                $errorCounter++; 
                            }
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
                    if( empty($MarksAndNumber) )
                    {
                        $marksAndNumberMatch[] = $row - 1; 
                        $errorCounter++; 
                    }
                    else if( $validateFunc->max_length($MarksAndNumber, 35) ) 
                    { 
                        $marksAndNumber[] = $row - 1; 
                        $errorCounter1++; 
                    }
                    if( !empty($MarksAndNumber) && !preg_match('/^[A-Za-z0-9 ]+$/', $MarksAndNumber) )
                    { 
                        $marksAndNumberSpecialChar[] = $row - 1; 
                        $errorCounter++; 
                    }

                    //NumberOfPackage
                    if( $validateFunc->max_length($NumberOfPackage, 10) ) 
                    { 
                        $numberOfPackage[] = $row - 1; $errorCounter1++; 
                    }

                    if ( $NumberOfPackage === '' )
                    {
                        $numberOfPackageMatch[] = $row - 1;
                        $errorCounter++;
                    }
                    else if ( ($validateFunc->match_numbers($NumberOfPackage)) == 0 )
                    {
                        $numberOfPackageMatch[] = $row - 1;
                        $errorCounter++;
                    }
                    else if ( $r == 0 && (int)$NumberOfPackage <= 0 )
                    {
                        $numberOfPackageZero[] = $row - 1;
                        $errorCounter++;
                    }
                    else if ( (float)$NumberOfPackage > 2147483647 )
                    {
                        $numberOfPackageTooLarge[] = $row - 1;
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
                    if( empty($InvoiceNumber) )
                    {
                        $invoiceNumberRequired[] = $row - 1;
                        $errorCounter++;
                    }
                    else
                    {
                        if( $validateFunc->max_length($InvoiceNumber, 300) ) 
                        { 
                            $invoiceNumber[] = $row - 1; $errorCounter1++; 
                        }
                        if( ($validateFunc->match_char($InvoiceNumber)) == 0 )
                        {
                            $invoiceNumberMatch[] = $row - 1; 
                            $errorCounter++; 
                        }
                    }

                    //SuplementaryValue
                    if( !empty($SuplementaryValue) )
                    {
                        if( $validateFunc->max_length($SuplementaryValue, 15) )
                        {
                            $suplementaryValueLength[] = $row - 1;
                            $errorCounter++;
                        }
                        else if( ($validateFunc->match_numbers($SuplementaryValue)) == 0 )
                        {
                            $suplementaryValueMatch[] = $row - 1; 
                            $errorCounter++; 
                        }
                    }

                    //ProcedureCode
                    if( !empty($ProcedureCode) ) {

                        //CHECK ProcedureCode IF EXISTS
                        $checkProcedureCodeExists = $validateFunc->__checkValidNatlCode($ProcedureCode);

                        if( !$checkProcedureCodeExists )
                        {
                            $checkProcedureCode[] = $row - 1;
                            $errorCounter++;
                        }

                    } else {

                        $checkProcedureCode[] = $row - 1;
                        $errorCounter++;
                    }

                    //ExtendedCode
                    if( !empty($ExtendedCode) ) {

                        //CHECK ExtendedCode IF EXISTS
                        $checkExtendedCodeExists = $validateFunc->__checkValidExtCode($ExtendedCode);

                        if( !$checkExtendedCodeExists )
                        {
                            $checkExtCode[] = $row - 1;
                            $errorCounter++;
                        }

                    } else {

                        $checkExtCode[] = $row - 1;
                        $errorCounter++;
                    }

                    //ItemGrossWeight
                    if( empty($ItemGrossWeight) )
                    {
                        $itemGrossWeightRequired[] = $row - 1;
                        $errorCounter++;
                    }
                    else
                    {
                        if( $validateFunc->max_length($ItemGrossWeight, 10) )
                        {
                            $itemGrossWeightLength[] = $row - 1;
                            $errorCounter++;
                        }
                        else if( !$validateFunc->match_weightFormat($ItemGrossWeight) || (float)$ItemGrossWeight <= 0 )
                        {
                            $itemGrossWeightMatch[] = $row - 1; 
                            $errorCounter++; 
                        }
                    }

                    //ItemNetWeight
                    if( empty($ItemNetWeight) )
                    {
                        $itemNetWeightRequired[] = $row - 1;
                        $errorCounter++;
                    }
                    else
                    {
                        if( $validateFunc->max_length($ItemNetWeight, 10) )
                        {
                            $itemNetWeightLength[] = $row - 1;
                            $errorCounter++;
                        }
                        else if( !$validateFunc->match_weightFormat($ItemNetWeight) || (float)$ItemNetWeight <= 0 )
                        {
                            $itemNetWeightMatch[] = $row - 1; 
                            $errorCounter++; 
                        }
                        else if( !empty($ItemGrossWeight) && $validateFunc->match_weightFormat($ItemGrossWeight) && (float)$ItemNetWeight > (float)$ItemGrossWeight )
                        {
                            $itemNetExceedsGross[] = $row - 1;
                            $errorCounter++;
                        }
                    }

                    //ItemInvoiceValue
                    if( empty($ItemInvoiceValue) )
                    {
                        $itemInvoiceValueRequired[] = $row - 1;
                        $errorCounter++;
                    }
                    else if( $validateFunc->max_length($ItemInvoiceValue, 11) )
                    {
                        $itemInvoiceValueLength[] = $row - 1;
                        $errorCounter++;
                    }
                    else if( !$validateFunc->match_weightFormat($ItemInvoiceValue) || (float)$ItemInvoiceValue <= 0 )
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
                        $TermsOfDelivery =   strtoupper($validateFunc->trim_val($dataRow[$row]['A']));
                        $TermsOfPayment  =   strtoupper($validateFunc->trim_val($dataRow[$row]['B']));
                    }

                /* VALIDATE FIELD VALUES */

                    //TermsOfDelivery
                    if( !empty($TermsOfDelivery) ) {

                        //CHECK TermsOfDelivery IF EXISTS
                        $checkTermsOfDelivery = $validateFunc->__checkValidTermsOfDelivery($TermsOfDelivery); //MODIFY

                        if( !$checkTermsOfDelivery )
                        {
                            $termsOfDelivery[] = $row - 1; 
                            $errorCounter++; 
                        }

                    }else{

                        $termsOfDelivery[] = $row - 1; 
                        $errorCounter++; 
                    }

                    //TermsOfPayment
                    if( !empty($TermsOfPayment) ) {

                        //CHECK TermsOfPayment IF EXISTS
                        $checkTermsOfPayment = $validateFunc->__checkValidTermsOfPayment($TermsOfPayment); //MODIFY

                        if( !$checkTermsOfPayment )
                        {
                            $termsOfPayment[] = $row - 1; 
                            $errorCounter++; 
                        }

                    }else{

                        $termsOfPayment[] = $row - 1; 
                        $errorCounter++; 
                    }
            }
        }
    }

    if ($errorCounter > 0 || $errorCounter1 > 0) {

        /* ERROR MESSAGES */
        // General Sheet Validation
            //Consignee
            if(!empty($marksAndNumber)){
                $errorLists[] = array(
                                    "ErrMsg" => "Exceeds the max characters allowed (35)",
                                    "Column" => "Marks and Number",
                                    "Rows" => implode(", " ,$marksAndNumber)
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
            if(!empty($billOfLadingRequired)){
                $errorLists[] = array(
                                    "ErrMsg" => "Bill of Lading/Airbill is required",
                                    "Column" => "Bill of Lading",
                                    "Rows" => implode(", " ,$billOfLadingRequired)
                                );
            }
            if(!empty($billOfLading)){
                $errorLists[] = array(
                                    "ErrMsg" => "Exceeds the max characters allowed (26)",
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
            if(!empty($vesselAircraftRequired)){
                $errorLists[] = array(
                                    "ErrMsg" => "Vessel/Aircraft is required",
                                    "Column" => "Vessel / Aircraft",
                                    "Rows" => implode(", " ,$vesselAircraftRequired)
                                );
            }
            if(!empty($vesselAircraft)){
                $errorLists[] = array(
                                    "ErrMsg" => "Exceeds the max characters allowed (27)",
                                    "Column" => "Vessel / Aircraft",
                                    "Rows" => implode(", " ,$vesselAircraft)
                                );
            }
            if(!empty($vesselAircraftMatch)){
                $errorLists[] = array(
                                    "ErrMsg" => "Only letters, numbers, and spaces are allowed - no special characters",
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
                if(!empty($containerNumberRequired)){
                    $errorLists[] = array("ErrMsg" => "Container Number is required", "Column" => "Container Number", "Rows" => implode(", ", $containerNumberRequired));
                }
                if(!empty($containerNumber)){
                    $errorLists[] = array(
                                        "ErrMsg" => "Exceeds the max characters allowed (100)",
                                        "Column" => "Container Number",
                                        "Rows" => implode(", " ,$containerNumber)
                                    );
                }
                if(!empty($containerNumberMatch)){
                    $errorLists[] = array(
                                        "ErrMsg" => "Only letters and numbers are allowed - no special characters",
                                        "Column" => "Container Number",
                                        "Rows" => implode(", " ,$containerNumberMatch)
                                    );
                }

                if(!empty($sealNumberRequired)){
                    $errorLists[] = array("ErrMsg" => "Seal Number is required", "Column" => "Seal Number", "Rows" => implode(", ", $sealNumberRequired));
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
                                        "ErrMsg" => "Only letters and numbers are allowed - no special characters",
                                        "Column" => "Seal Number",
                                        "Rows" => implode(", " ,$sealNumberMatch)
                                    );
                }

                if(!empty($containerSizeRequired)){
                    $errorLists[] = array("ErrMsg" => "Container Size is required", "Column" => "Container Size", "Rows" => implode(", ", $containerSizeRequired));
                }
                if(!empty($containerSizeMatch)){
                    $errorLists[] = array("ErrMsg" => "Only letters and numbers are allowed - no special characters", "Column" => "Container Size", "Rows" => implode(", ", $containerSizeMatch));
                }
                if(!empty($checkContainerSize)){
                    $errorLists[] = array("ErrMsg" => "Invalid Container Size", "Column" => "Container Size", "Rows" => implode(", ", $checkContainerSize));
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
            if(!empty($marksAndNumberMatch)){
                $errorLists[] = array(
                                    "ErrMsg" => "Marks and Number is required",
                                    "Column" => "Marks and Number",
                                    "Rows" => implode(", " ,$marksAndNumberMatch)
                                );
            }
            if(!empty($marksAndNumber)){
                $errorLists[] = array(
                                    "ErrMsg" => "Exceeds the max characters allowed (70)",
                                    "Column" => "Marks and Number",
                                    "Rows" => implode(", " ,$marksAndNumber)
                                );
            }
            if(!empty($marksAndNumberSpecialChar)){
                $errorLists[] = array(
                                    "ErrMsg" => "Only letters, numbers, and spaces are allowed - no special characters",
                                    "Column" => "Marks and Number",
                                    "Rows" => implode(", " ,$marksAndNumberSpecialChar)
                                );
            }
            
            // NumberOfPackage
            if(!empty($numberOfPackage)){
                $errorLists[] = array(
                                    "ErrMsg" => "Exceeds the max characters allowed (10)",
                                    "Column" => "Number of Package",
                                    "Rows" => implode(", " ,$numberOfPackage)
                                );
            }
            if(!empty($numberOfPackageMatch)){
                $errorLists[] = array(
                                    "ErrMsg" => "Number of Package is required and must contain numbers only",
                                    "Column" => "Number of Package",
                                    "Rows" => implode(", " ,$numberOfPackageMatch)
                                );
            }

            if(!empty($numberOfPackageZero)){
                $errorLists[] = array(
                                    "ErrMsg" => "Number of Package must be greater than 0 for the item 1",
                                    "Column" => "Number of Package",
                                    "Rows" => implode(", " ,$numberOfPackageZero)
                                );
            }
            if(!empty($numberOfPackageTooLarge)){
                $errorLists[] = array(
                                    "ErrMsg" => "Number of Package exceeds the maximum allowed value (2,147,483,647)",
                                    "Column" => "Number of Package",
                                    "Rows" => implode(", " ,$numberOfPackageTooLarge)
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
            if(!empty($invoiceNumberRequired)){
                $errorLists[] = array(
                                    "ErrMsg" => "Invoice Number is required",
                                    "Column" => "Invoice Number",
                                    "Rows" => implode(", " ,$invoiceNumberRequired)
                                );
            }

            if(!empty($invoiceNumber)){
                $errorLists[] = array(
                                    "ErrMsg" => "Exceeds the max characters allowed (300)",
                                    "Column" => "Invoice Number",
                                    "Rows" => implode(", " ,$invoiceNumber)
                                );
            }
            if(!empty($invoiceNumberMatch)){
                $errorLists[] = array(
                                    "ErrMsg" => "Only letters, numbers, and spaces are allowed - no special characters",
                                    "Column" => "Invoice Number",
                                    "Rows" => implode(", " ,$invoiceNumberMatch)
                                );
            }
            
            // SuplementaryValue
            if(!empty($suplementaryValueLength)){
                $errorLists[] = array(
                                    "ErrMsg" => "Exceeds the max characters allowed (15)",
                                    "Column" => "Supplementary Value",
                                    "Rows" => implode(", " ,$suplementaryValueLength)
                                );
            }
            if(!empty($suplementaryValueMatch)){
                $errorLists[] = array(
                                    "ErrMsg" => "Invalid format. Only whole numbers are allowed - decimals and scientific notation (e.g. 1E+10) are not accepted",
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
            
            // ProcedureCode
            if(!empty($checkProcedureCode)){
                $errorLists[] = array(
                                    "ErrMsg" => "Invalid or missing Procedure Code",
                                    "Column" => "Procedure Code",
                                    "Rows" => implode(", " ,$checkProcedureCode)
                                );
            }

            // ExtendedCode
            if(!empty($checkExtCode)){
                $errorLists[] = array(
                                    "ErrMsg" => "Invalid or missing Extended Code",
                                    "Column" => "Extended Code",
                                    "Rows" => implode(", " ,$checkExtCode)
                                );
            }

            //ItemGrossWeight
            if(!empty($itemGrossWeightMatch)){ 
                $errorLists[] = array(
                                    "ErrMsg" => "Invalid entry (e.g. 0.00) - Required",
                                    "Column" => "Item Gross Weight",
                                    "Rows" => implode(", " ,$itemGrossWeightMatch)
                                );
            }

            if(!empty($itemGrossWeightLength)){
                $errorLists[] = array("ErrMsg" => "Exceeds the max characters allowed (10)", "Column" => "Item Gross Weight", "Rows" => implode(", ", $itemGrossWeightLength));
            }

            if(!empty($itemGrossWeightRequired)){
                $errorLists[] = array("ErrMsg" => "Item Gross Weight is required", "Column" => "Item Gross Weight", "Rows" => implode(", ", $itemGrossWeightRequired));
            }

            if(!empty($itemNetWeightRequired)){
                $errorLists[] = array("ErrMsg" => "Item Net Weight is required", "Column" => "Item Net Weight", "Rows" => implode(", ", $itemNetWeightRequired));
            }

            //ItemNetWeight
            if(!empty($itemNetWeightMatch)){ 
                $errorLists[] = array(
                                    "ErrMsg" => "Invalid entry (e.g. 0.00) - Required",
                                    "Column" => "Item Net Weight",
                                    "Rows" => implode(", " ,$itemNetWeightMatch)
                                );
            }

            if(!empty($itemNetWeightLength)){
                $errorLists[] = array("ErrMsg" => "Exceeds the max characters allowed (10)", "Column" => "Item Net Weight", "Rows" => implode(", ", $itemNetWeightLength));
            }

            if(!empty($itemNetExceedsGross)){
                $errorLists[] = array("ErrMsg" => "Item Net Weight must not exceed Item Gross Weight", "Column" => "Item Net Weight", "Rows" => implode(", ", $itemNetExceedsGross));
            }

            //AIRBILLNO
            if(!empty($airbillNo)){
                $errorLists[] = array(
                                    "ErrMsg" => "Exceeds the max characters allowed (26)",
                                    "Column" => "Airbill/BL Number",
                                    "Rows" => implode(", " ,$airbillNo)
                                );
            }
            if(!empty($airbillNoMatch)){
                $errorLists[] = array(
                                    "ErrMsg" => "Only accept letters and numbers - Required",
                                    "Column" => "Airbill/BL Number",
                                    "Rows" => implode(", " ,$airbillNoMatch)
                                );
            }

            //ItemInvoiceValue
            if(!empty($itemInvoiceValueRequired)){
                $errorLists[] = array("ErrMsg" => "Item Invoice Value is required", "Column" => "Item Invoice Value", "Rows" => implode(", ", $itemInvoiceValueRequired));
            }
            if(!empty($itemInvoiceValueLength)){
                $errorLists[] = array("ErrMsg" => "Exceeds the max characters allowed (11)", "Column" => "Item Invoice Value", "Rows" => implode(", ", $itemInvoiceValueLength));
            }
            if(!empty($itemInvoiceValueMatch)){ 
                $errorLists[] = array(
                                    "ErrMsg" => "Invalid entry (e.g. 1000.00) - Required",
                                    "Column" => "Item Invoice Value",
                                    "Rows" => implode(", " ,$itemInvoiceValueMatch)
                                );
            }

        // Financial Sheet Validation 

            //TermsOfDelivery
            if(!empty($termsOfDelivery)){ 
                $errorLists[] = array(
                                    "ErrMsg" => "Invalid or missing Terms of Delivery code",
                                    "Column" => "Terms Of Delivery",
                                    "Rows" => implode(", " ,$termsOfDelivery)
                                );
            }  

            //TermsOfPayment
            if(!empty($termsOfPayment)){ 
                $errorLists[] = array(
                                    "ErrMsg" => "Invalid or missing Terms of Payment code",
                                    "Column" => "Terms Of Payment",
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
                
                $Port                 = strtoupper(trim($Port));
                $LocationOfGoods      = strtoupper(trim($LocationOfGoods));
                $CountryOfDestination = strtoupper(trim($CountryOfDestination));
                $PortOfLoading        = strtoupper(trim($PortOfLoading));
                $PurposeOfExportation = strtoupper(trim($PurposeOfExportation));
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

                    $goodsdesc1             = $checkItemCodeExists['commodityDesc'];
                    $HSCode                 = $checkItemCodeExists['HsCode'];
                    
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
                $ProcDesc               = $ProcedureCode;  
                $CoCode                 = 'PH';

                $ItemGrossWeight  = ($ItemGrossWeight === '') ? '' : number_format(round((float)str_replace(',', '', $ItemGrossWeight), 2), 2, '.', '');
                $ItemNetWeight    = ($ItemNetWeight === '') ? '' : number_format(round((float)str_replace(',', '', $ItemNetWeight), 2), 2, '.', '');
                $ItemInvoiceValue = number_format(round((float)str_replace(',', '', $ItemInvoiceValue), 2), 2, '.', '');
                
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
                        ':itemgrossweight' => ($ItemGrossWeight === '') ? null : $ItemGrossWeight,
                        ':itemnetweight'   => ($ItemNetWeight === '') ? null : $ItemNetWeight,
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

    } // end $isBulkClient else-branch