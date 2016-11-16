<?php
//include_once('../dss.php');
include('../library/global_methods.php');
if($_SERVER['REQUEST_METHOD'] == "POST"){
	$url = urlToApi('/montessori/library/admin.php');
	// Get data
  if($_POST['action'] == 'register_admin'){
	$obj = jsonFromRequest($_POST, $url);
	if($obj["response"] == 1){
        $response = $obj["message"];
		$delay = 1500;
        $redirecturl = '../index.php';
	}else{
		$response = $obj["message"];
		$delay = 1500;
        $redirecturl = '../admin.php';
	}
  }
}

redirectToURL($response, $redirecturl, $delay);

 ?>
