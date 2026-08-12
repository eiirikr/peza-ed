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

$processFunc    = new ProcessFile();
$validateFunc   = new ValidateFields();
$conn           = new DBConnection();

// print_r($conn->connect());
// die();

$applNo         = $_POST["applno"];
$stats          = $_POST["status"];
$version2       = $_POST["version2"];

//GET FILE
$excelDetails = $processFunc->__getUploadedFile();
// die(var_dump($excelDetails));
        
    // START PROCESSING  //
    $xml = new DOMDocument('1.0', "utf-8"); 
    $xml->formatOutput = true;
    $xml->preserveWhiteSpace = true;
    
    $xmlRoot1   = $xml->createElement('IMPORT');
    $xmlRoot1   = $xml->appendChild($xmlRoot1);

    $xmlRoot2   = $xmlRoot1->appendChild($xml->createElement('ITEMS')); //items
    $xmlRoot4   = $xmlRoot1->appendChild($xml->createElement('ADDITIONALCTN')); //additionalctn

    $objReader      = PHPExcel_IOFactory::createReader($excelDetails["type"]);
    $objReader->setReadDataOnly(true);
    $objPHPExcel    = $objReader->load($excelDetails["inputFile"]);

    $objWorksheet   = $objPHPExcel->getSheetByName('Items');
    $objWorksheet2  = $objPHPExcel->getSheetByName('AdditionalCTN');

    if(!$objWorksheet || !$objWorksheet2){
        echo "<script>
                alert('Cannot proceed. Please check file.');
                window.location.href='index.php?applno=$applNo&status=$stats&version2=$version2';
        </script>";
        die();
    }

    // ------------ START Items WORKSHEET (MAPPING and DB INSERTION) ------------ //
    if($objWorksheet){

        $highestRow     = $objWorksheet->getHighestRow();
        $highestColumn  = $objWorksheet->getHighestColumn();
        $headingsArray  = $objWorksheet->rangeToArray('A1:'.$highestColumn.'1',null, true, true, true);
        $headingsArray  = $headingsArray[1];

        $r = -1;
        $namedDataArray = array();

        for ($row = 2; $row <= $highestRow; ++$row) {

            $dataRow = $objWorksheet->rangeToArray('A'.$row.':'.$highestColumn.$row,null, true, true, true);
            //$dataRow = $validateFunc->trim_val($dataRow);

            if ((isset($dataRow[$row]['A'])) && ($dataRow[$row]['A'] > '')) {
                ++$r;
                
                foreach($headingsArray as $columnKey => $columnHeading) {
                    
                    $Marks1             =   $validateFunc->trim_val($dataRow[$row]['A']);
                    $Marks2             =   $validateFunc->trim_val($dataRow[$row]['B']);
                    $PackageCode        =   $validateFunc->trim_val($dataRow[$row]['C']);
                    $NoOfPackage        =   $dataRow[$row]['D'];
                    $InvoiceValue       =   $dataRow[$row]['E'];
                    $InvoiceCurr        =   $validateFunc->trim_val($dataRow[$row]['F']);
                    $ContainerNo1       =   $validateFunc->trim_val2($dataRow[$row]['G']);
                    $ContainerNo2       =   $validateFunc->trim_val2($dataRow[$row]['H']);
                    $ContainerNo3       =   $validateFunc->trim_val2($dataRow[$row]['I']);
                    $ContainerNo4       =   $validateFunc->trim_val2($dataRow[$row]['J']);
                    $GoodsDescription   =   $validateFunc->trim_val($dataRow[$row]['K']);
                    $OtherChargesFlag   =   $dataRow[$row]['L'];
                    $InsuranceFlag      =   $dataRow[$row]['M'];
                    $InvoiceNo          =   $validateFunc->trim_val($dataRow[$row]['N']);
                    $SupValue           =   $dataRow[$row]['O'];
                    $TariffHeading      =   $dataRow[$row]['P'];
                    $TariffExtension    =   $dataRow[$row]['Q'];
                    $TarSpecAICode      =   $dataRow[$row]['R'];
                    $COCode             =   $validateFunc->trim_val($dataRow[$row]['S']);
                    $ItemGrossWeight    =   $dataRow[$row]['T'];
                    $ItemNetWeight      =   $dataRow[$row]['U'];
                    $Preferential       =   $validateFunc->trim_val($dataRow[$row]['V']);
                    $NationalCode       =   $dataRow[$row]['W'];
                    $ExtensionCode      =   $dataRow[$row]['X'];
                    $ValuationCode      =   $validateFunc->trim_val($dataRow[$row]['Y']);
                    $AirbillBLNo        =   $validateFunc->trim_val($dataRow[$row]['Z']);
                    $PrevDoc            =   $validateFunc->trim_val($dataRow[$row]['AA']);
                    $SpecCode           =   $dataRow[$row]['AB'];
                    $Atrig              =   $dataRow[$row]['AC'];
                    $AtrigDate          =   $dataRow[$row]['AD'];
                    $MSP                =   $dataRow[$row]['AE'];

                }

                $getValuation = $validateFunc->__checkValidValCode($ValuationCode);
                $valuationDesc = '';
                if( $getValuation['status'] )
                {
                    $valuationDesc = $getValuation['quo_dsc'];
                }
                
                /*DEFAULT VALUES*/
                //$valuationCode          = 'NNNNN';
                //$valuationDesc          = 'NOT RELATED, NO RSTRCTN/CNDTN/RYLTS/ARRNGMNTS';
                $valuationMethod        = 'NV';
                $valuationMethodDesc    = 'No Valuation';
                $adjustment             = '1.00';
                $fines                  = '0.00';
                //$InvoiceCurr                = 'USD';
                $XML_flag               = 1;

                //$applNo = $_POST["applno"];

                $ItemGrossWeight    = number_format((float)str_replace(',', '', $ItemGrossWeight), 2, '.', '');

                $ItemNetWeight      = number_format((float)str_replace(',', '', $ItemNetWeight), 2, '.', '');

                $InvoiceValue       = number_format((float)str_replace(',', '', $InvoiceValue), 2, '.', '');

                if(!empty($MSP)){
                    $MSP = number_format((float)str_replace(',', '', $MSP), 2, '.', '');
                }

                if(empty($Preferential)){
                    $Preferential = 'NONE';
                }
                
                //ATRIGDATE
                if(!is_null($AtrigDate)){
                    if (strlen($AtrigDate) == 10) {
                        $UNIX_DATE = $AtrigDate;
                        $AtrigDate = $UNIX_DATE;
                    }elseif(strlen($AtrigDate) == 9){
                        $UNIX_DATE = $AtrigDate;
                        $AtrigDate = $UNIX_DATE;
                    }else{
                        $UNIX_DATE = ($AtrigDate - 25569) * 86400;
                        $AtrigDate =  gmdate("Y-m-d", $UNIX_DATE);
                    }
                }
                
                $Marks1             = strtoupper($Marks1);
                $Marks2             = strtoupper($Marks2);
                $PackageCode        = strtoupper($PackageCode);
                $GoodsDescription   = strtoupper($GoodsDescription);
                $OtherChargesFlag   = strtoupper($OtherChargesFlag);
                $InsuranceFlag      = strtoupper($InsuranceFlag);
                $InvoiceNo          = strtoupper($InvoiceNo);
                $COCode             = strtoupper($COCode);
                $Preferential       = strtoupper($Preferential);
                $ExtensionCode      = strtoupper($ExtensionCode);
                $AirbillBLNo        = strtoupper($AirbillBLNo);
                $Atrig              = strtoupper($Atrig);

                $ContainerNo1       = strtoupper($ContainerNo1);
                $ContainerNo2       = strtoupper($ContainerNo2);
                $ContainerNo3       = strtoupper($ContainerNo3);
                $ContainerNo4       = strtoupper($ContainerNo4);
                $GoodsDescription   = str_replace('&amp;','and', $GoodsDescription);

                //MAX LENGTH
                if (strlen($Marks1) > 35) { $Marks1 = substr($Marks1, 0, 35); } 
                if (strlen($Marks2) > 35) { $Marks2 = substr($Marks2, 0, 35); } 
                if (strlen($GoodsDescription) > 255) { $GoodsDescription = substr($GoodsDescription, 0, 255); } 

                if($OtherChargesFlag != 1){
                    $OtherChargesFlag = 'FALSE';
                }else{
                    $OtherChargesFlag = 'TRUE';
                }

                if($InsuranceFlag != 1){
                    $InsuranceFlag = 'FALSE';
                }else{
                    $InsuranceFlag = 'TRUE';
                }
                
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

                //die(print_r($getItemNo));
                
                // ----------------- END GET ITEMNO ----------------- //

                $insert_sql1 = "";
                $insert_sql2 = "";

                if(!empty($SupValue)) {

                    $hs6_cod = substr($TariffHeading, 0, 6);
                    $tar_pr1 = substr($TariffHeading, 6, 7);

                    //TEST
                    $uom_cod1 = '';

                    $getUom = $validateFunc->__getUomValue($TariffExtension, $hs6_cod, $tar_pr1);

                    if( !empty($getUom) ) 
                    {
                        $uom_cod1  = $getUom;
                    }
                    else
                    {
                        //$uom_cod1 = '';
                        $SupValueRemove = 1;
                    }

                    //INSERT STATEMENT
                    $insert_sql1 = "INSERT INTO tblimpapl_detail (ApplNo, ItemNo, Marks1, Marks2, PackCode, NoPack, Cont1, Cont2, Cont3, Cont4, GoodsDesc, Ocharges, Ifreight, InvNo, HSCode, HSCODE_TAR, TARSPEC, COCode, ItemGWeight, ItemNweight, Pref, ProcDsc, ExtCode, AirBill, InvValue, SupUnit1, SupUnit2, ATRIG, ATRIGDATE, quo_cod, quo_dsc, ValMethodNum, ValMethodDesc, Adjustment, SupUnit3, InvCurr, Item_uploading_flag, PrevDoc, SupVal1, SupVal2, MSP) 
                    VALUES ('".$applNo."', '".$itemNo."', '".$Marks1."', '".$Marks2."', '".$PackageCode."', '".$NoOfPackage."', '".$ContainerNo1."', '".$ContainerNo2."', '".$ContainerNo3."', '".$ContainerNo4."', '".$GoodsDescription."', '".$OtherChargesFlag."', '".$InsuranceFlag."', '".$InvoiceNo."', '".$TariffHeading."', '".$TariffExtension."', '".$TarSpecAICode."', '".$COCode."', '".$ItemGrossWeight."', '".$ItemNetWeight."', '".$Preferential."', '".$NationalCode."', '".$ExtensionCode."', '".$AirbillBLNo."', '".$InvoiceValue."', '".$uom_cod1."', '".$SpecCode."', '".$Atrig."', '".$AtrigDate."', '".$ValuationCode."', '".$valuationDesc."', '".$valuationMethod."', '".$valuationMethodDesc."', '".$adjustment."', '".$fines."', '".$InvoiceCurr."', '".$XML_flag."', '".$PrevDoc."'";

                    $insert_sql2 = "INSERT INTO tblimpapl_cons (ApplNo, ItemNo, Marks1, Marks2, PackCode, NoPack, Cont1, Cont2, Cont3, Cont4, GoodsDesc, Ocharges, Ifreight, InvNo, HSCode, HSCODE_TAR, TARSPEC, COCode, ItemGWeight, ItemNweight, Pref, ProcDsc, ExtCode, AirBill, InvValue, SupUnit1, SupUnit2, ATRIG, ATRIGDATE, quo_cod, quo_dsc, ValMethodNum, ValMethodDesc, Adjustment, SupUnit3, InvCurr, Item_uploading_flag, PrevDoc, SupVal1, SupVal2, MSP) 
                    VALUES ('".$applNo."', '".$itemNo."', '".$Marks1."', '".$Marks2."', '".$PackageCode."', '".$NoOfPackage."', '".$ContainerNo1."', '".$ContainerNo2."', '".$ContainerNo3."', '".$ContainerNo4."', '".$GoodsDescription."', '".$OtherChargesFlag."', '".$InsuranceFlag."', '".$InvoiceNo."', '".$TariffHeading."', '".$TariffExtension."', '".$TarSpecAICode."', '".$COCode."', '".$ItemGrossWeight."', '".$ItemNetWeight."', '".$Preferential."', '".$NationalCode."', '".$ExtensionCode."', '".$AirbillBLNo."', '".$InvoiceValue."', '".$uom_cod1."', '".$SpecCode."', '".$Atrig."', '".$AtrigDate."', '".$ValuationCode."', '".$valuationDesc."', '".$valuationMethod."', '".$valuationMethodDesc."', '".$adjustment."', '".$fines."', '".$InvoiceCurr."', '".$XML_flag."', '".$PrevDoc."'";
                    
                    if(!empty($SupValueRemove)){
                        $insert_sql1.=", NULL, NULL";
                        $insert_sql2.=", NULL, NULL";
                        
                    }else {
                        $insert_sql1.=", '".$SupValue."', '".$SupValue."'";
                        $insert_sql2.=", '".$SupValue."', '".$SupValue."'";
                    }
                    
                    if(isset($MSP) || !empty($MSP) || $MSP != ''){
                        $insert_sql1.=", '".$MSP."')";
                        $insert_sql2.=", '".$MSP."')";
                    }else {
                        $insert_sql1.=", NULL)";
                        $insert_sql2.=", NULL)";
                    }

                    //UNSET VALUES
                    unset($SupValueRemove);

                }else{

                    //INSERT STATEMENT
                    $insert_sql1 = "INSERT INTO TBLIMPAPL_DETAIL (ApplNo, ItemNo, Marks1, Marks2, PackCode, NoPack, Cont1, Cont2, Cont3, Cont4, GoodsDesc, Ocharges, Ifreight, InvNo, HSCode, HSCODE_TAR, TARSPEC, COCode, ItemGWeight, ItemNweight, Pref, ProcDsc, ExtCode, AirBill, InvValue, SupVal1, SupUnit1, SupVal2, SupUnit2, ATRIG, ATRIGDATE, quo_cod, quo_dsc, ValMethodNum, ValMethodDesc, Adjustment, SupUnit3, InvCurr, Item_uploading_flag, PrevDoc, MSP)
                    VALUES ('".$applNo."', '".$itemNo."', '".$Marks1."', '".$Marks2."', '".$PackageCode."', '".$NoOfPackage."', '".$ContainerNo1."', '".$ContainerNo2."', '".$ContainerNo3."', '".$ContainerNo4."', '".$GoodsDescription."', '".$OtherChargesFlag."', '".$InsuranceFlag."', '".$InvoiceNo."', '".$TariffHeading."', '".$TariffExtension."', '".$TarSpecAICode."', '".$COCode."', '".$ItemGrossWeight."', '".$ItemNetWeight."', '".$Preferential."', '".$NationalCode."', '".$ExtensionCode."', '".$AirbillBLNo."', '".$InvoiceValue."', NULL, NULL, NULL, '".$SpecCode."', '".$Atrig."', '".$AtrigDate."', '".$ValuationCode."', '".$valuationDesc."', '".$valuationMethod."', '".$valuationMethodDesc."', '".$adjustment."', '".$fines."', '".$InvoiceCurr."', '".$XML_flag."', '".$PrevDoc."'";

                    $insert_sql2 = "INSERT INTO TBLIMPAPL_CONS (ApplNo, ItemNo, Marks1, Marks2, PackCode, NoPack, Cont1, Cont2, Cont3, Cont4, GoodsDesc, Ocharges, Ifreight, InvNo, HSCode, HSCODE_TAR, TARSPEC, COCode, ItemGWeight, ItemNweight, Pref, ProcDsc, ExtCode, AirBill, InvValue, SupVal1, SupUnit1, SupVal2, SupUnit2, ATRIG, ATRIGDATE, quo_cod, quo_dsc, ValMethodNum, ValMethodDesc, Adjustment, SupUnit3, InvCurr, Item_uploading_flag, PrevDoc, MSP) 
                    VALUES ('".$applNo."', '".$itemNo."', '".$Marks1."', '".$Marks2."', '".$PackageCode."', '".$NoOfPackage."', '".$ContainerNo1."', '".$ContainerNo2."', '".$ContainerNo3."', '".$ContainerNo4."', '".$GoodsDescription."', '".$OtherChargesFlag."', '".$InsuranceFlag."', '".$InvoiceNo."', '".$TariffHeading."', '".$TariffExtension."', '".$TarSpecAICode."', '".$COCode."', '".$ItemGrossWeight."', '".$ItemNetWeight."', '".$Preferential."', '".$NationalCode."', '".$ExtensionCode."', '".$AirbillBLNo."', '".$InvoiceValue."', NULL, NULL, NULL, '".$SpecCode."', '".$Atrig."', '".$AtrigDate."', '".$ValuationCode."', '".$valuationDesc."', '".$valuationMethod."', '".$valuationMethodDesc."', '".$adjustment."', '".$fines."', '".$InvoiceCurr."', '".$XML_flag."', '".$PrevDoc."'";

                    if(isset($MSP) || !empty($MSP) || $MSP != ''){
                        $insert_sql1.=", '".$MSP."')";
                        $insert_sql2.=", '".$MSP."')";
                    }else {
                        $insert_sql1.=", NULL)";
                        $insert_sql2.=", NULL)";
                    }
                    
                }

                //EXECUTE QUERY TO INSERT
                try{

                    $sqlExecute1  = $insert_sql1;
                    $stmt1 = $conn->connect()->prepare($sqlExecute1);
                    $stmt1->execute();
                                    
                } catch (PDOException $e1) {
                    echo "ERROR : " . $e1->getMessage();
                    die();
                }

                try{

                    $sqlExecute2  = $insert_sql2;
                    $stmt2 = $conn->connect()->prepare($sqlExecute2);
                    $stmt2->execute();
                                    
                } catch (PDOException $e2) {
                    echo "ERROR : " . $e2->getMessage();
                    die();
                }

                    //XML CONTENT 
                    $xmlRoot3 = $xmlRoot2->appendChild($xml->createElement('Item'));
                    $element1 = $xmlRoot3->appendChild($xml->createElement('MarksandNumbers1',      htmlspecialchars($Marks1)));
                    $element1 = $xmlRoot3->appendChild($xml->createElement('MarksandNumbers2',      htmlspecialchars($Marks2)));
                    $element1 = $xmlRoot3->appendChild($xml->createElement('PackageCode',           $PackageCode));
                    $element1 = $xmlRoot3->appendChild($xml->createElement('NoOfPackage',           $NoOfPackage));
                    $element1 = $xmlRoot3->appendChild($xml->createElement('InvoiceValue',          $InvoiceValue));
                    $element1 = $xmlRoot3->appendChild($xml->createElement('InvoiceCurr',           $InvoiceCurr));
                    $element1 = $xmlRoot3->appendChild($xml->createElement('ContainerNo1',          $ContainerNo1));
                    $element1 = $xmlRoot3->appendChild($xml->createElement('ContainerNo2',          $ContainerNo2));
                    $element1 = $xmlRoot3->appendChild($xml->createElement('ContainerNo3',          $ContainerNo3));
                    $element1 = $xmlRoot3->appendChild($xml->createElement('ContainerNo4',          $ContainerNo4));
                    $element1 = $xmlRoot3->appendChild($xml->createElement('GoodsDescription',      htmlspecialchars($GoodsDescription)));
                    $element1 = $xmlRoot3->appendChild($xml->createElement('OtherChargesFlag',      $OtherChargesFlag));
                    $element1 = $xmlRoot3->appendChild($xml->createElement('InsuranceFlag',         $InsuranceFlag));
                    $element1 = $xmlRoot3->appendChild($xml->createElement('InvoiceNumber',         htmlspecialchars($InvoiceNo)));
                    $element1 = $xmlRoot3->appendChild($xml->createElement('SupplementaryValue',    $SupValue));
                    $element1 = $xmlRoot3->appendChild($xml->createElement('TariffHeading',         $TariffHeading));
                    $element1 = $xmlRoot3->appendChild($xml->createElement('TariffExtension',       $TariffExtension));
                    $element1 = $xmlRoot3->appendChild($xml->createElement('TARSPEC_AICODE',        $TarSpecAICode));
                    $element1 = $xmlRoot3->appendChild($xml->createElement('CountryofOriginCode',   $COCode));
                    $element1 = $xmlRoot3->appendChild($xml->createElement('ItemGrossWeight',       $ItemGrossWeight));
                    $element1 = $xmlRoot3->appendChild($xml->createElement('ItemNetweight',         $ItemNetWeight));
                    $element1 = $xmlRoot3->appendChild($xml->createElement('Preferential',          $Preferential));
                    $element1 = $xmlRoot3->appendChild($xml->createElement('NationalCode',          $NationalCode));
                    $element1 = $xmlRoot3->appendChild($xml->createElement('ExtensionCode',         $ExtensionCode));
                    $element1 = $xmlRoot3->appendChild($xml->createElement('ValuationCode',         $ValuationCode));
                    $element1 = $xmlRoot3->appendChild($xml->createElement('Airbill_BLNumber',      $AirbillBLNo));
                    $element1 = $xmlRoot3->appendChild($xml->createElement('PrevDoc',               htmlspecialchars($PrevDoc)));
                    $element1 = $xmlRoot3->appendChild($xml->createElement('SPECode',               $SpecCode));
                    $element1 = $xmlRoot3->appendChild($xml->createElement('ATRIG',                 $Atrig));
                    $element1 = $xmlRoot3->appendChild($xml->createElement('ATRIGDATE',             $AtrigDate));
                    $element1 = $xmlRoot3->appendChild($xml->createElement('MSP',                   $MSP));

                    //UNSET VALUES
                    unset($ValuationCode);
                    unset($valuationDesc);
                            
            }
        
        }
    
    }else{

        echo "<script>
                alert('Cannot proceed, could not find Items/AdditionalCTN worksheet. Please check file.');
                window.location.href='index.php?applno=$applNo&status=$stats&version2=$version2';
            </script>";
        die();
    }
    // ------------ END Items WORKSHEET ------------ //

    // ------------ START AdditionalCTN WORKSHEET (MAPPING and DB INSERTION) ------------ //
    if($objWorksheet2){

        $highestRow2    = $objWorksheet2->getHighestRow();
        $highestColumn2 = $objWorksheet2->getHighestColumn();
        $headingsArray2 = $objWorksheet2->rangeToArray('A1:'.$highestColumn2.'1',null, true, true, true);
        $headingsArray2 = $headingsArray2[1];

        $r2 = -1;
       
        $namedDataArray2 = array();

        for ($row2 = 2; $row2 <= $highestRow2; ++$row2) {

            $dataRow2 = $objWorksheet2->rangeToArray('A'.$row2.':'.$highestColumn2.$row2, null, true, true, true);

            if ((isset($dataRow2[$row2]['A'])) && ($dataRow2[$row2]['A'] > '')) {
                ++$r2;

                foreach($headingsArray2 as $columnKey2 => $columnHeading2) {

                    $con = $row2 + 3;
                    $additionalctn_val  =   $dataRow2[$row2]['A'];
                    $additionalctn_val  = strtoupper($additionalctn_val);
                    
                }

                //EXECUTE QUERY TO INSERT
                $insert_sql3 = "INSERT INTO tblIMPAPL_Container (ApplNo, Container, Item_uploading_flag) VALUES ('".$applNo."', '".$additionalctn_val."', '".$XML_flag."')";

                try{

                    $sqlExecute3  = $insert_sql3;
                    $stmt3 = $conn->connect()->prepare($sqlExecute3);
                    $stmt3->execute();
                                    
                } catch (PDOException $e3) {
                    echo "ERROR : " . $e3->getMessage();
                    die();
                }

                    //XML CONTENT
                    $element2 = $xmlRoot4->appendChild($xml->createElement('ContainerNo'.$con, $additionalctn_val));
               
            }
        }

    }else {

        echo "<script>
                alert('Cannot proceed, could not find Items/AdditionalCTN worksheet. Please check file.');
                window.location.href='index.php?applno=$applNo&status=$stats&version2=$version2';
             </script>";
        die();

    }

    // ------------ END AdditionalCTN WORKSHEET ------------ //

    // ------------ DB UPDATE - TBLIMPAPL_MASTER ------------ //

    $totalItems = 0;
    $totalPacks = 0;
    //$applNo = $_POST["applno"];

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

    $updateQuery = "UPDATE TBLIMPAPL_MASTER SET ItemCon = '$totalItems', Items = '$totalItems', Packs = '$totalPacks' WHERE ApplNo = '$applNo'";

    try{

        $sqlUpdate  = $updateQuery;
        $stmtUpdate = $conn->connect()->prepare($sqlUpdate);
        $stmtUpdate->execute();
                        
    } catch (PDOException $a) {
        echo "ERROR : " . $a->getMessage();
        die();
    }

    // ------------ END DB UPDATE - TBLIMPAPL_MASTER ------------ //

    // ------------ REMOVE UPLOADED EXCEL FILE ------------ //

    $getFilename = $processFunc->__getUploadedFile();
    unlink($getFilename['inputFile']);

    // ------------ END REMOVE UPLOADED EXCEL FILE ------------ //

    echo "<script>
            window.location.href='index.php?applno=$applNo&status=$stats&version2=$version2&msg=success';
        </script>";