<?php

//include_once('dss.php');
include('../library/global_methods.php');
if($_SERVER['REQUEST_METHOD'] == "POST"){
	$url = urlToApi('/montessori/library/login.php');
    if($_POST['action'] == 'login'){

        $obj = jsonFromRequest($_POST, $url);
        if($obj["response"] == 1){
            setcookie('usrn', $_POST["usrn"], time() + 86400 * 5, "/");
            setcookie('pssw', $_POST["pssw"], time() + 86400 * 5, "/");
            //setcookie('email', $email, time() + 86400 * 5, "/");

            $response = $obj["message"];
        }else{
            $response = $obj["message"];
        }
    }else if($_POST['action'] == 'logout'){
        $action = $_POST['action'];
        $usrn = $_COOKIE['usrn'];
        $fields = array(
           'action' => $action,
           'usrn' => $usrn
        );

        $obj = jsonFromRequest($fields, $url);

        if($obj["response"] == 1){
            setcookie('usrn', '', time() - 7200, "/");
            setcookie('token', '', time() - 7200, "/");
            setcookie('first_name', '', time() - 7200, "/");
            setcookie('last_name', '', time() - 7200, "/");
        }
        $response = $obj["message"];
    }
}else{
	$url = urlToApi('../api/login.php');
	if($_GET['action'] == 'logout'){
        $action = $_GET['action'];
        $usrn = $_COOKIE['usrn'];
        $fields = array(
           'action' => $action,
           'usrn' => $usrn
        );

        $obj = jsonFromRequest($fields, $url);

        if($obj["response"] == 1){
            setcookie('usrn', '', time() - 7200, "/");
            setcookie('token', '', time() - 7200, "/");
            setcookie('first_name', '', time() - 7200, "/");
            setcookie('last_name', '', time() - 7200, "/");
        }
        $response = $obj["message"];
    }
}

$delay = 1500;
$redirecturl = '../index.php';

redirectToURL($response, $redirecturl, $delay);
?>
