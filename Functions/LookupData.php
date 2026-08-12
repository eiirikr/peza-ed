<?php

include_once('DBConnection.php');

class LookupData extends DBConnection{

    public function getForwarders($conn, $lstexporter) {

        $stmtCheck = $conn->connectPEZAexpPTOPS()->prepare(
            "SELECT For_adr1, For_adr2, For_adr3 
            FROM tblForwarders 
            WHERE For_Name = :lstexporter"
        );

        $stmtCheck->execute([':lstexporter' => $lstexporter]);

        if ($stmtCheck->rowCount() == 0) {
            return null;
        }

        return $stmtCheck->fetch(PDO::FETCH_ASSOC);
    }

    public function getImporters($conn, $loccod) {

        $stmtCheck = $conn->connectPEZA()->prepare(
            "SELECT address1, address2, zonecode 
            FROM tblImporters 
            WHERE PezaImpCode = :loccod"
        );

        $stmtCheck->execute([':loccod' => $loccod]);

        if ($stmtCheck->rowCount() == 0) {
            return null;
        }

        return $stmtCheck->fetch(PDO::FETCH_ASSOC);
    }

    public function getExchangeRate($conn, $cur_cod) {

        $stmtCheck = $conn->connect()->prepare(
            "SELECT rat_exc from GBRATTAB where cur_cod=:cur_cod order by eea_dov desc"
        );
        
        $stmtCheck->execute([':cur_cod' => $cur_cod]);

        if ($stmtCheck->rowCount() == 0) {
            return null;
        }

        return $stmtCheck->fetch(PDO::FETCH_ASSOC);
    }

    public function getModeofTransport($conn, $offclr_cod) {

        $stmtCheck = $conn->connect()->prepare(
            "SELECT offClrcod, offClrName, offClrMode FROM dbo.DmOffClr WHERE offClrcod=:offclrcod"
        );
        
        $stmtCheck->execute([':offclrcod' => $offclr_cod]);

        if ($stmtCheck->rowCount() == 0) {
            return null;
        }

        return $stmtCheck->fetch(PDO::FETCH_ASSOC);
    }

    

}