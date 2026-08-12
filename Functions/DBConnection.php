<?php

class DBConnection{

    public function connect() {
        
        try
        {
            
            $options = array(
                            PDO::ATTR_ERRMODE => PDO::ERRMODE_EXCEPTION, 
                            PDO::ATTR_DEFAULT_FETCH_MODE => PDO::FETCH_ASSOC
                        );

            $pdo = new PDO( 'sqlsrv:server=192.168.5.70,1477;Database=INSCUSTSTDB','sa', 'df0rc3', $options );

            return $pdo;
        }
        catch(PDOException $ex)
        {
            return "ERROR : " . $ex->getMessage();
        }

    }

    public function connectIPPEZA() {
        
        try
        {
            
            $options = array(
                            PDO::ATTR_ERRMODE => PDO::ERRMODE_EXCEPTION, 
                            PDO::ATTR_DEFAULT_FETCH_MODE => PDO::FETCH_ASSOC
                        );

            $pdo = new PDO( 'sqlsrv:server=192.168.5.70,1477;Database=INSIPPEZA','sa', 'df0rc3', $options );

            return $pdo;
        }
        catch(PDOException $ex)
        {
            return "ERROR : " . $ex->getMessage();
        }

    }

    public function connectPEZA() {
        
        try
        {
            
            $options = array(
                            PDO::ATTR_ERRMODE => PDO::ERRMODE_EXCEPTION, 
                            PDO::ATTR_DEFAULT_FETCH_MODE => PDO::FETCH_ASSOC
                        );

            $pdo = new PDO( 'sqlsrv:server=192.168.5.70,1477;Database=PEZA','sa', 'df0rc3', $options );

            return $pdo;
        }
        catch(PDOException $ex)
        {
            return "ERROR : " . $ex->getMessage();
        }

    }
    
    public function connectPEZAexpPTOPS() {
        
        try
        {
            
            $options = array(
                            PDO::ATTR_ERRMODE => PDO::ERRMODE_EXCEPTION, 
                            PDO::ATTR_DEFAULT_FETCH_MODE => PDO::FETCH_ASSOC
                        );

            $pdo = new PDO( 'sqlsrv:server=192.168.5.70,1477;Database=PEZAexpPTOPS','sa', 'df0rc3', $options );

            return $pdo;
        }
        catch(PDOException $ex)
        {
            return "ERROR : " . $ex->getMessage();
        }

    }

}

?>