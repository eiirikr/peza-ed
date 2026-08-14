<?php

include_once('DBConnection.php');

class ValidateFields extends DBConnection{

    // public function __construct(){

    //     parent::__construct();
    // }

    public function generateApplNo($conn, $csncod)
    {
        $cntr = 1;
        $today = getdate();

        do {
            $newApplNumber = $csncod . "ED" .
                str_pad(substr($today['year'], -2), 2, "0", STR_PAD_LEFT) .
                str_pad($today['mon'], 2, "0", STR_PAD_LEFT) .
                str_pad($today['mday'], 2, "0", STR_PAD_LEFT) .
                str_pad($cntr, 3, "0", STR_PAD_LEFT);

            $stmtCheck = $conn->connectIPPEZA()->prepare(
                "SELECT ApplNo FROM tblEXPApl_Master WHERE ApplNo = :applNo"
            );

            $stmtCheck->execute([':applNo' => $newApplNumber]);

            if ($stmtCheck->rowCount() == 0) {
                return $newApplNumber;
            }

            $cntr++;

        } while (true);
    }

    public function isEmptyRow($row) {
        foreach($row as $cell)
        {
            if(NULL !== $cell){
                return false;
            }
        }
        return true;
    }

    public function trim_val($value){
        $value = preg_replace('/\s+/', ' ', $value);
        $value = preg_replace('/[\x00-\x1F\x7F-\xFF]/', '', $value);

        return $value;
    }

    public function trim_val2($value){
        $value = preg_replace('/\s+/', '', $value);
        $value = preg_replace('/[\x00-\x1F\x7F-\xFF]/', '', $value);

        return $value;
    }

    public function match_char($itemCol){
        if(preg_match("/^[A-Za-z0-9 _#$%()*+,\-.:&\/;_]+$/i", $itemCol)){
            return true;
        }
    }

    public function match_letters($itemCol){
        if(preg_match("/^[A-Za-z]+$/i", $itemCol)){
            return true;
        }
    }

    public function match_numbers($itemCol){
        if(preg_match("/^[0-9]+$/i", $itemCol)){
            return true;
        }
    }

    public function match_alphanum($itemCol){
        if(preg_match("/^[A-Za-z0-9]+$/i", $itemCol)){
            return true;
        }
    }

    public function match_amount($itemCol){
        if(preg_match("/^[0-9.]+$/i", $itemCol)){
            return true;
        }
    }

    public function match_weightFormat($itemCol){
        if(preg_match("/^\d+(\.\d+)?$/", $itemCol)){
            return true;
        }
        return false;
    }

    public function truncateDecimal($value, $decimals = 2){
        $value = (string) $value;
        if (strpos($value, '.') === false) {
            return $value;
        }
        list($intPart, $decPart) = explode('.', $value, 2);
        $decPart = substr($decPart, 0, $decimals);
        return $decPart === '' ? $intPart : $intPart . '.' . $decPart;
    }

    public function max_length($itemCol, $maxlength){
        if(strlen($itemCol) > $maxlength) { 
            return true;
        }
    }

    public function match_manifestFormat($manifest){
        if(preg_match('/^[A-Z]{3}[0-9]{4}-[0-9]{2}$/', $manifest)) { 
            return true;
        }
        return false;
    }

    public function match_containerFormat($cont){
        if(preg_match('/^[A-Z]{4}[0-9]{7}+$/i', $cont)) { 
            return true;
        }
    }

    public function match_dateFormat($value){
        if(preg_match('/^[0-9-]+$/D', $value)) { 
            return true;
        }
    }

    public function match_hscodeFormat($value){
        if(preg_match('/^[0-9]{6,8}+$/i', $value)) {
            return true;
        }
    }

    public function match_natlcodeFormat($value){
        if(preg_match('/^[0-9]{4}+$/i', $value)) {
            return true;
        }
    }

    public function match_extcodeFormat($value){
        if(preg_match('/^[A-Z0-9]{3}+$/i', $value)) {
            return true;
        }
    }

    public function match_prefFormat($value){
        if(preg_match('/^[A-Z]{0,17}+$/i', $value)) {
            return true;
        }
    }

    public function match_packcodeFormat($value){
        if(preg_match('/^[A-Z0-9]{0,17}+$/i', $value)) {
            return true;
        }
    }

    public function match_cocodeFormat($value){
        if(preg_match('/^[A-Z]{2,3}+$/i', $value)) {
            return true;
        }
    }

    public function match_speccodeFormat($value){
        if(preg_match('/^[0-9]{6}+$/i', $value)) {
            return true;
        }
    }

    public function match_valcodeFormat($value){
        if(preg_match('/^[A-Z]{5}+$/i', $value)) {
            return true;
        }
    }
    
    public function match_invCurrFormat($value){
        if(preg_match('/^[A-Z]{3}+$/i', $value)) {
            return true;
        }
    }

    public function __checkValidConsignee($conn, $Consignee, $cltcode)
    {
        $sql = "SELECT ExpCode, ExpName, Expadr1, Expadr2, Expadr3, Expadr4, expCoCode 
                FROM BUYER 
                WHERE cltcode = :cltcode 
                AND UPPER(LTRIM(RTRIM(Expname))) = UPPER(LTRIM(RTRIM(:consignee)))";

        try {
            $stmt = $conn->connectIPPEZA()->prepare($sql);
            $stmt->execute([
                ':cltcode'   => $cltcode,
                ':consignee' => $Consignee
            ]);

            $result = $stmt->fetch(PDO::FETCH_ASSOC);

            return $result ? $result : false;

        } catch (PDOException $e) {
            return false;
        }
    }

    public function __checkValidAddress($Address, $buyerRecord)
    {
        if (empty($buyerRecord)) {
            return false;
        }

        $masterAddressParts = array_filter([
            isset($buyerRecord['Expadr1']) ? $buyerRecord['Expadr1'] : '',
            isset($buyerRecord['Expadr2']) ? $buyerRecord['Expadr2'] : '',
            isset($buyerRecord['Expadr3']) ? $buyerRecord['Expadr3'] : '',
            isset($buyerRecord['Expadr4']) ? $buyerRecord['Expadr4'] : ''
        ]);

        $masterAddress = strtoupper(trim(implode(' ', $masterAddressParts)));
        $excelAddress  = strtoupper(trim(preg_replace('/\s+/', ' ', $Address)));

        return ($masterAddress === $excelAddress);
    }

    public function __checkValidLocationOfGoods($val){

        if ($val !== strtoupper($val)) {
            return false;
        }

        $sql    = "SELECT TOP 1 SHD_COD FROM dbo.GBSHDTAB_D WHERE SHD_COD = '$val'";
        $stmt   = $this->connect()->query($sql);
        $result = $stmt->fetchAll(PDO::FETCH_OBJ);

        return (count($result) > 0 ? true : false);
    }

    public function __checkValidProvinceOfOrigin($val){

        $sql    = "SELECT TOP 1 prov_cod from GBPRVORG WHERE prov_cod = '$val'";
        $stmt   = $this->connect()->query($sql);
        $result = $stmt->fetchAll(PDO::FETCH_OBJ);

        return (count($result) > 0 ? true : false);
    }

    public function __checkValidCountryOfDestination($val){

        if ($val !== strtoupper($val)) {
            return false;
        }

        $sql    = "SELECT TOP 1 cityCode FROM DmCityOrigin WHERE cityCode = '$val'";
        $stmt   = $this->connect()->query($sql);
        $result = $stmt->fetchAll(PDO::FETCH_OBJ);

        return (count($result) > 0 ? true : false);
    }

    public function __checkValidPortOfLoading($val){

        if ($val !== strtoupper($val)) {
            return false;
        }

        $sql    = "SELECT TOP 1 loc_cod FROM GBLOCTAB where loc_cod = '$val'";
        $stmt   = $this->connect()->query($sql);
        $result = $stmt->fetchAll(PDO::FETCH_OBJ);

        return (count($result) > 0 ? true : false);
    }

    public function __checkValidPortOfDeparture($val){

        if ($val !== strtoupper($val)) {
            return null;
        }

        $sql    = "SELECT TOP 1 offClrCod, offClrMode FROM DmOffClr where offClrCod = '$val'";
        $stmt   = $this->connect()->query($sql);
        $result = $stmt->fetch(PDO::FETCH_ASSOC);

        if ($result) {
            return $result;
        } else {
            return null;
        }
    }

    public function __checkValidPurposeOfExportation($val){

        $result = [
            "OTHERS",
            "FOR REPAIR",
            "RETURN TO SOURCE",
            "SALE"
        ];

        return (count(array_filter($result, function($item) use ($val){
            return strtoupper(trim($item)) == strtoupper(trim($val));
        })) > 0 ? true : false);
    }

    public function __checkValidContainerSize($val){

        $result = [
            "20",
            "40",
            "45"
        ];

        return (count(array_filter($result, function($item) use ($val){
            return strtoupper(trim($item)) == strtoupper(trim($val));
        })) > 0 ? true : false);
    }

    public function __checkValidHscode($TariffExtension, $hs6_cod, $tar_pr1){

        $sql    = "SELECT TOP 1 hs6_cod, uom_cod1, rul_cod FROM GBTARTAB WHERE hs6_cod = '$hs6_cod' AND tar_pr1 = '$tar_pr1' AND tar_pr2 = '$TariffExtension'";
        $stmt   = $this->connect()->query($sql);
        $result = $stmt->fetchAll(PDO::FETCH_OBJ);

        return (count($result) > 0) ? true : false;
    }

    public function __checkValidSpeccode($TariffExtension, $hs6_cod, $tar_pr1, $SpecCode){

        $sql    = "SELECT TOP 1 spc_cod FROM GBSPECTAB WHERE hs6_cod = '$hs6_cod' AND tar_pr1 = '$tar_pr1' AND tar_pr2 = '$TariffExtension'  AND spc_cod = '$SpecCode'";
        $stmt   = $this->connect()->query($sql);
        $result = $stmt->fetchAll(PDO::FETCH_OBJ);

        return (count($result) > 0) ? true : false;
    }

    public function __checkItemCode($itemCode, $allaccids, $accountType){
       
        // staging
        $ptopsLOELookup = "http://192.168.5.26:90/api/get-ex-item-lookup";
        // prodlike
        // $ptopsLOELookup = "http://192.168.5.26:88/api/get-ex-item-lookup";
        // production
        // $ptopsLOELookup = "http://192.168.1.67:81/api/get-ex-item-lookup";

        $data = [
            'allAccIDs'     => $allaccids,
            'commodityCode' => $itemCode,
            'accountType'   => $accountType
        ];

        // Initialize cURL session
        $ch = curl_init($ptopsLOELookup);

        // Set options for the POST request
        curl_setopt($ch, CURLOPT_RETURNTRANSFER, true);
        curl_setopt($ch, CURLOPT_POST, true);
        curl_setopt($ch, CURLOPT_POSTFIELDS, http_build_query($data)); // Use json_encode($data) for JSON APIs

        // Execute request and fetch response
        $response = curl_exec($ch);

        $result = json_decode($response, true);

        if (isset($result['data']['data'][0])) {
            return $result['data']['data'][0];
        } else {
            return null;
        }

    }

    public function __getHscodeData($TariffExtension, $hs6_cod, $tar_pr1){

        $data = array();

        $sql    = "SELECT TOP 1 hs6_cod, uom_cod1, rul_cod FROM GBTARTAB WHERE hs6_cod = '$hs6_cod' AND tar_pr1 = '$tar_pr1' AND tar_pr2 = '$TariffExtension'";
        $stmt   = $this->connect()->query($sql);
        $result = $stmt->fetchAll(PDO::FETCH_OBJ);

        if (count($result) > 0) {

           $data = array(
            "uom_cod1" =>   $result[0]->uom_cod1,
            "rul_cod" =>    $result[0]->rul_cod
           );

        }

        return $data;
    }

    public function __checkValidNatlCode($val){

        $sql    = "SELECT TOP 1 cp4_cod FROM GBCP4CP3 WHERE cp4_cod = '$val'";
        $stmt   = $this->connect()->query($sql);
        $result = $stmt->fetchAll(PDO::FETCH_OBJ);

        return (count($result) > 0 ? true : false);
    }

    public function __checkValidExtCode($val){

        $sql    = "SELECT TOP 1 cp3_cod FROM GBCP3TAB WHERE cp3_cod = '$val'";
        $stmt   = $this->connect()->query($sql);
        $result = $stmt->fetchAll(PDO::FETCH_OBJ);

        return (count($result) > 0 ? true : false);
    }

    public function __checkValidPref($val){

        $sql    = "SELECT TOP 1 prf_cod FROM GBPRFTAB WHERE prf_cod = '$val'";
        $stmt   = $this->connect()->query($sql);
        $result = $stmt->fetchAll(PDO::FETCH_OBJ);

        return (count($result) > 0 ? true : false);
    }

    public function __checkValidPackCode($val){

        $sql    = "SELECT TOP 1 pkg_cod FROM GBPKGTAB WHERE pkg_cod = '$val'";
        $stmt   = $this->connect()->query($sql);
        $result = $stmt->fetchAll(PDO::FETCH_OBJ);

        return (count($result) > 0 ? true : false);
    }

    public function __checkValidCOCode($val){

        $sql    = "SELECT TOP 1 cityCode FROM DmCityOrigin WHERE cityCode = '$val'";
        $stmt   = $this->connect()->query($sql);
        $result = $stmt->fetchAll(PDO::FETCH_OBJ);

        return (count($result) > 0 ? true : false);
    }

    public function __checkSpecReq($TariffExtension, $hs6_cod, $tar_pr1){

        $sql    = "SELECT TOP 1 spc_cod FROM GBSPECTAB WHERE hs6_cod = '$hs6_cod' AND tar_pr1 = '$tar_pr1' AND tar_pr2 = '$TariffExtension'";
        $stmt   = $this->connect()->query($sql);
        $result = $stmt->fetchAll(PDO::FETCH_OBJ);

        return (count($result) > 0) ? true : false;
    }

    public function __checkInvCurrReq($val){

        $sql    = "SELECT TOP 1 CUR_COD FROM GBRATTAB WHERE CUR_COD = '$val' ORDER BY EEA_DOV DESC";
        $stmt   = $this->connect()->query($sql);
        $result = $stmt->fetchAll(PDO::FETCH_OBJ);

        return (count($result) > 0) ? true : false;
    }

    public function __checkValidTermsOfDelivery($val){

        if ($val !== strtoupper($val)) {
            return false;
        }

        $sql    = "SELECT Distinct (tod_dsc), tod_cod FROM GBTODTAB WHERE tod_cod = '$val' ORDER BY tod_dsc ASC";
        $stmt   = $this->connect()->query($sql);
        $result = $stmt->fetchAll(PDO::FETCH_OBJ);

        return (count($result) > 0) ? true : false;
    }

    public function __checkValidTermsOfPayment($val){

        if ($val !== strtoupper($val)) {
            return false;
        }

        $sql    = "SELECT top_cod FROM GBTOPTAB WHERE top_cod = '$val' ORDER BY top_cod ASC";
        $stmt   = $this->connect()->query($sql);
        $result = $stmt->fetchAll(PDO::FETCH_OBJ);

        return (count($result) > 0) ? true : false;
    }

    public function __checkValidValCode($val){

        $sql    = "SELECT * FROM GBQUOTAB WHERE quo_cod = '$val'";
        $stmt   = $this->connect()->query($sql);
        $result = $stmt->fetchAll(PDO::FETCH_OBJ);

        if (count($result) > 0) {
            $data = array(
             "quo_cod"  =>  $result[0]->quo_cod,
             "quo_dsc"  =>  $result[0]->quo_dsc,
             "status"   =>  true
            );
 
        }else{
            $data = array(
                "status" => false
            );
        }

        return $data;
    }

    public function __checkPrevDocReq($val){

        $sql    = "SELECT TOP 1 Manifest FROM TBLIMPAPL_MASTER WHERE Manifest LIKE 'TS%' AND ApplNo = '$val'";
        $stmt   = $this->connect()->query($sql);
        $result = $stmt->fetchAll(PDO::FETCH_OBJ);

        $data = false;
        if (count($result) > 0) {
            if( strlen($result[0]->Manifest) == 12 )
            {
                $data = true;
            }
        }

        return $data;
    }

    public function __getMdec($val){

        $data = array();

        $sql    = "SELECT TOP 1 MDec, MDec2 FROM TBLIMPAPL_MASTER WHERE ApplNo = '$val'";
        $stmt   = $this->connect()->query($sql);
        $result = $stmt->fetchAll(PDO::FETCH_OBJ);

        if (count($result) > 0) {
           $data = array(
            "MDec" =>   $result[0]->MDec,
            "MDec2" =>   $result[0]->MDec2
           );
        }

        return $data;
    }

    public function __getItemNo($val){

        $sql    = "SELECT TOP 1 ItemNo FROM TBLEXPAPL_DETAIL WHERE ApplNo = '$val' ORDER BY ItemNo DESC";
        $stmt   = $this->connectIPPEZA()->query($sql);
        $result = $stmt->fetchAll(PDO::FETCH_OBJ);

        return (count($result) > 0 ? $result[0]->ItemNo : NULL);

    }
    
    public function __getUomValue($TariffExtension, $hs6_cod, $tar_pr1){
        
        $sql    = " SELECT TOP 1 uom_cod1 FROM GBTARTAB WHERE hs6_cod = '$hs6_cod' AND tar_pr1 = '$tar_pr1' AND tar_pr2 = '$TariffExtension'";
        $stmt   = $this->connect()->query($sql);
        $result = $stmt->fetchAll(PDO::FETCH_OBJ);

        return (count($result) > 0 ? $result[0]->uom_cod1 : NULL);

    }

    public function __getTotalItems($applno){

        $data = array();

        $sql    = "SELECT COUNT(*) AS totalItems, SUM(CAST(NoPack AS INT)) AS totalPacks FROM TBLEXPAPL_DETAIL WHERE ApplNo = '$applno'";
        $stmt   = $this->connectIPPEZA()->query($sql);
        $result = $stmt->fetchAll(PDO::FETCH_OBJ);

        if (count($result) > 0) {

           $data = array(
            "totalItems" =>   $result[0]->totalItems,
            "totalPacks" =>   $result[0]->totalPacks
           );

        }

        return $data;
    }
 
}