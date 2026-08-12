<?php
 
class ProcessFile{

    public function __getUploadedFile(){

        $result = array();
        
        //GET FILE EXTENSION
        $dir = scandir("uploads/");
        $ext = "";
        
        foreach($dir as $files) {
        
            if($files !== '.' && $files !== '..') { 
        
                $ext = pathinfo($files, PATHINFO_EXTENSION);        
            }
        }

        $filename = "items.".$ext;
        $filepath = dirname(__DIR__).'/uploads/'.$filename;

        $fileExt = pathinfo($filepath); 
        $fileExt = strtolower($fileExt["extension"]);

        if ($fileExt=='xlsx') { 

            $result = array(
                            "inputFile" => dirname(__DIR__)."/uploads/{$filename}",
                            "type"      => "Excel2007"
                    );
            
        }else{

            $result = array(
                            "inputFile" => dirname(__DIR__)."/uploads/{$filename}",
                            "type"      => "Excel5"
                    );
        }
       
        return $result;

    }

    public function __checkFileFormat($file){

        $allowedFileType = array(
                                    'application/vnd.ms-excel',
                                    'text/xls',
                                    'text/xlsx',
                                    'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
                            );
        
        if(!in_array($file, $allowedFileType)){
            
            return false;
        }
        else{
            
            return true;
        }
        
    }

    public function __getFileExtension($file){

        $ext = pathinfo($file, PATHINFO_EXTENSION);
        return $ext;
    }

    public function __uploadFileInDirectory($uploadedfile){

        //GET FILE EXTENSION
        $fileExt = $this->__getFileExtension($uploadedfile);

        $filename   = "items.".$fileExt;
        $filepath   = dirname(__DIR__).'/uploads/'.$filename;

        //UNLINK FILE IF ALREADY EXISTS
        if (file_exists($filepath)) {
            unlink($filepath);
        }
        
        //EXCEL FILE SUCCESSFULLY UPLOADED
        if(move_uploaded_file($_FILES["file"]["tmp_name"], "uploads/{$filename}")){
            return true;
        }

    }

    public function __getPHPExcelDetails($file){

        $result = array();
        
        //GET FILE EXTENSION
        $fileExt = $this->__getFileExtension($file);
        $filename = "items.".$fileExt;
        $filepath = dirname(__DIR__).'/uploads/'.$filename;

        $fileExt = pathinfo($filepath); 
        $fileExt = strtolower($fileExt["extension"]);

        if ($fileExt=='xlsx') { 

            $result = array(
                            "inputFile" => dirname(__DIR__)."/uploads/{$filename}",
                            "type"      => "Excel2007"
                    );
            
        }else{

            $result = array(
                            "inputFile" => dirname(__DIR__)."/uploads/{$filename}",
                            "type"      => "Excel5"
                    );
        }
       
        return $result;
    }
 
}