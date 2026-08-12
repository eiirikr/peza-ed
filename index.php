<?php
session_start();

$_POST = array_change_key_case($_POST, CASE_LOWER);

$token = '';

if (isset($_POST['flow_token'])) {
    $token = $_POST['flow_token'];
} elseif (isset($_GET['token'])) {
    $token = $_GET['token'];
} elseif (isset($_SESSION['current_flow_token'])) {
    $token = $_SESSION['current_flow_token'];
}

/* SAVE TOKEN */
$_SESSION['current_flow_token'] = $token;

/* init container */
if (!isset($_SESSION['flows'])) {
    $_SESSION['flows'] = array();
}

/* store per token */
if (
    isset($_POST['csncod']) &&
    !isset($_SESSION['flows'][$token])
) {

    $_SESSION['flows'][$token] = array(
        'csncod'         => isset($_POST['csncod']) ? $_POST['csncod'] : '',
        'loctin'         => isset($_POST['loctin']) ? $_POST['loctin'] : '',
        'zonecode'       => isset($_POST['zonecode']) ? $_POST['zonecode'] : '',
        'ptopstin'       => isset($_POST['ptops_tin']) ? $_POST['ptops_tin'] : '',
        'enterprisetype' => isset($_POST['enterprisetype']) ? $_POST['enterprisetype'] : '',
        'compnam'        => isset($_POST['compnam']) ? $_POST['compnam'] : '',
        'userid'         => isset($_POST['userid']) ? $_POST['userid'] : '',
        'lstexporter'    => isset($_POST['lstexporter']) ? $_POST['lstexporter'] : '',
        'locbroktin'     => isset($_POST['locbroktin']) ? $_POST['locbroktin'] : '',
        'loccod'         => isset($_POST['loccod']) ? $_POST['loccod'] : '',
        'allaccids'      => isset($_POST['allaccids']) ? $_POST['allaccids'] : '',
        'mod_cod'        => isset($_POST['mod_cod']) ? $_POST['mod_cod'] : '',
        'mod_cod2'       => isset($_POST['mod_cod2']) ? $_POST['mod_cod2'] : '',
        'cltcode'        => isset($_POST['cltcode']) ? $_POST['cltcode'] : '',
        'redirection'    => isset($_POST['redirection']) ? $_POST['redirection'] : ''
    );

    header("Location: index.php?token=" . $token);
    exit;
}

if ( isset($_GET['msg']) && $_GET['msg'] == 'error' )
{
    echo '<script>alert("ERROR OCCURS.")</script>'; 
}

if ( isset($_GET['msg']) && $_GET['msg'] == 'success' )
{
    $redirection = $_GET['redirection'];

    echo "<script>
            alert('File uploaded successfully');
            window.location.href='../../WebCWS/".$redirection.".asp?ApplNo=". $_GET['applno']  . "&Status=I';
        </script>";
}

?>

<!DOCTYPE html>
<html lang="en">

<head>
     <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <title>Create ED from Excel File </title>
        <!-- Bootstrap core CSS -->
        <link rel="stylesheet" href="https://maxcdn.bootstrapcdn.com/bootstrap/4.0.0/css/bootstrap.min.css" integrity="sha384-Gn5384xqQ1aoWXA+058RXPxPg6fy4IWvTNh0E263XmFcJlSAwiGgFAW/dAiS6JXm" crossorigin="anonymous">
        <!-- Datatables -->
        <link rel="stylesheet" href="https://cdn.datatables.net/1.10.20/css/dataTables.bootstrap4.min.css">
        <!-- <link href="css/header.css" rel="stylesheet"/> -->
</head>

<body>
       
    <div class="container" style="margin-top:50px;">

        <div class="row" align="center">
            <div class="col-md-12" align="center">
            <img src="img/header.png" border="0" height="auto" width="auto">
        </div>
        </div>
        <div class="row">
            <div class="col-md-12" align="center">
             
            </br></br>
                
                <h1>
                    UPLOAD EXCEL FILE
                </h1>
                <div class="form-group col-md-4">
                    <label for="UploadExcel">Choose Excel file to Process</label>
                </div>

                <form method="POST" action="process.php" enctype="multipart/form-data">
                    <div class="form-group col-md-4">
                        <input type="hidden" name="flow_token" value="<?php echo $token ?>">
                        <input type="file" name="file" id="fileSelect" accept=".xls, .xlsx, application/vnd.openxmlformats-officedocument.spreadsheetml.sheet, application/vnd.ms-excel" class="form-control" required/>
                    </div>
                    
                    <div class="form-group col-md-4">
                        <button type="submit" name="btn" value="Submit" id="submitBtn" class="btn btn-primary btn-lg btn-block">PROCESS</button>
                    </div>
                </form>
            </div>
        </div>
        
        <div class="row">
            <div class="col-md-12" align="center">
                <div class="form-group">
                   <a href='template/Bulk Upload Template for AEDS.xlsx' download>[CLICK HERE TO DOWNLOAD TEMPLATE]</a>
                </div>
            </div>
        </div>

        <div class="row">
            <div class="col-md-12" align="center">
                <div class="form-group">
                   <button type="button" id="submitBtn" class="btn btn-light border" onclick="openLookupPopup()"> Exportables Lookup </button>
                </div>
            </div>
        </div>

        <div class="row">
            <div class="col-md-12" align="center">
                <div class="form-group col-md-2">
                    <?php 
						if ( isset($_GET['msg']) && $_GET['msg'] == 'error' ) 
						{
							echo "<a href='http://testweb.intercommerce.com.ph/BULK-ITEMUPLOAD/IMPORT/index.php?applno={$applNo}&status={$stats}' id='back' class='btn btn-default btn-sm btn-block'><< Back to page </a>";
						}
						else 
						{
							echo "<a href='http://testweb.intercommerce.com.ph/WebCWS/cws_impdec.asp?applno={$applNo}&status={$stats}' id='back' class='btn btn-default btn-sm btn-block'><< Back to page </a>";
						}
					?>
                </div>
            </div>
        </div>

       <div class="row"><br><br></div>
        <div class="row">
            <div class="col-md-12">
                <?php if ( isset($_GET['msg']) && $_GET['msg'] == 'error' ) { 
                        $errorLists     = $_SESSION['errormsg'];  
                        $requiredFields = $_SESSION['required'];
                        $proceedFields  = $_SESSION['proceed'];
                ?>
                <table id="myTable" class="table table-hover table-bordered" cellspacing="0"> 
                    <thead> 
                        <tr align="center">
                                <th colspan="3" class="table-dark"> ERROR SUMMARY </th>
                            </tr>
                            <tr>
                                <th class="table-danger">ITEM FIELD</th>
                                <th class="table-danger">ERROR MESSAGE</th>
                                <th class="text-center table-danger">ITEM NO.</th>
                            </tr>
                    </thead> 
                    <tbody> 
                    <tr> 
                    <?php foreach($errorLists as $lists) { ?>
                        <td class="col-md-4"><?php echo $lists['Column']; ?></td>
                        <td><?php echo $lists['ErrMsg']; ?></td>
                        <td align="center"><?php echo $lists['Rows']; ?></td>
                    </tr>
                    <?php } ?>
                    </tbody>    
                </table>
                <br><br>
            </div>
            <div class="col-md-12" align="center">
                 <?php if (!empty($proceedFields) && $requiredFields == 0) { ?>
                <form method="POST" action="proceed.php" enctype="multipart/form-data"  id="proceedForm" name="proceedForm">
                    <div class="form-group col-md-3">
                        <button type="button" name="proceedBtn" value="Proceed" id="proceedBtn" data-toggle="modal" class="btn btn-success btn-sm btn-block">PROCEED WITH THE UPLOADING</button>
                    </div>
                </form>
                <?php } ?>
            </div>
            <?php } ?>
        </div>
        
    </div>

    <!-- Footer --> 
    <footer class="text-lg-start text-white fixed-bottom" style="background-color: #3e4551">
    <div class="text-center p-3" style="background-color: rgba(0, 0, 0, 0.2)">
        Copyright © <?php echo date("Y"); ?> | <a class="text-white" href="https://www.intercommerce.com.ph/login.asp?home=home">www.intercommerce.com.ph</a>
    </div>
    </footer>
    <!-- END Footer --> 

    <!-- START PROCEED MODAL -->
    <div class="modal fade" id="confirm-submit" tabindex="-1" role="dialog" aria-labelledby="myModalLabel" aria-hidden="true">
        <div class="modal-dialog">
            <div class="modal-content">
                <div class="modal-header">
                    <b>KINDLY CONFIRM BEFORE PROCEEDING!</b> 
                </div>
                <div class="modal-body">
                    <p>Characters of the Item field that exceeds beyond the maximum length allowed will be excluded or eliminated.</p>
                    <p>Do you still wish to proceed? </p>
                </div>
                <div class="modal-footer">
                    <button type="button" class="btn btn-default" data-dismiss="modal">Cancel</button>
                    <a href="#" id="submit" class="btn btn-success success">Yes, Proceed</a>
                </div>
            </div>
        </div>
    </div>
    <!-- END PROCEED MODAL -->

        <!-- Core JS -->                  
        <script src="https://code.jquery.com/jquery-3.3.1.js"></script>
        <!-- Bootstrap JS -->
        <script src="https://maxcdn.bootstrapcdn.com/bootstrap/4.0.0/js/bootstrap.min.js" integrity="sha384-JZR6Spejh4U02d8jOt6vLEHfe/JQGiRRSQQxSfFWpi1MquVdAyjUar5+76PVCmYl" crossorigin="anonymous"></script>
        <!-- Datatables -->
        <script src="https://cdn.datatables.net/1.10.20/js/jquery.dataTables.min.js"></script>
        <script src="https://cdn.datatables.net/1.10.20/js/dataTables.bootstrap4.min.js"></script>

</body>

</html>

<script>


$(document).ready(function(){ 

    $('#myTable').DataTable({
            "ordering": false
    });

    $('#proceedBtn').click(function() {
        $('#confirm-submit').modal('show');
        e.preventDefault();
    });

    $('#submit').click(function(){
        $('#proceedForm').submit();
    });

});


function openLookupPopup() {

    var allaccids = "<?= urlencode($_SESSION['flows'][$token]['allaccids']) ?>";
    

    var popupUrl =
        "http://testweb.intercommerce.com.ph/webcws/ptops-lookup-importables-pezaexploc.asp?allaccids=" + allaccids;

    window.open(
        popupUrl,
        "lookupPopup",
        "width=1000,height=700,resizable=yes,scrollbars=yes"
    );
}

</script>