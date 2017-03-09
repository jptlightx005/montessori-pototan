<?php
/* global functions */

function generateSalt(){
    $algorithm = "2a";
    $length = "12";

    $salt = "$" . $algorithm . "$" . $length . "$";

    $salt .= substr( str_replace( "+", ".", base64_encode( mcrypt_create_iv( 128, MCRYPT_DEV_URANDOM ) ) ), 0, 22 );

    return $salt;
}

function redirectToURL($msg, $rurl, $timeout){
    echo "<h1>$msg</h1>";

    if($rurl != null)
        echo "<script>setTimeout(\"location.href = '$rurl';\",$timeout);</script>";
}

function expiredSession(){
    $usrn = $_COOKIE['usrn'];
    $query = "UPDATE `tbl_admin` SET expiration=CURRENT_TIMESTAMP, token='' WHERE usrn='$usrn'";

    if(mysql_query($query)){
        return true;
    }else{
        return false;
    }
}

function resetCookies(){
    setcookie('usrn', '', time() - 7200, "/");
    setcookie('token', '', time() - 7200, "/");
}
function jsonFromRequest($fields, $url){
    // build the urlencoded data
    $postvars = http_build_query($fields);

    // open connection
    $ch = curl_init();

    // set the url, number of POST vars, POST data
    curl_setopt($ch, CURLOPT_RETURNTRANSFER, 1);
    curl_setopt($ch, CURLOPT_URL, $url);
    curl_setopt($ch, CURLOPT_POST, count($fields));
    curl_setopt($ch, CURLOPT_POSTFIELDS, $postvars);

    // execute post
    $result = curl_exec($ch);

    // close connection
    curl_close($ch);

    return json_decode($result, true);
}

function jsonFromRequestForTest($fields, $url){
    // build the urlencoded data
    $postvars = http_build_query($fields);

    // open connection
    $ch = curl_init();

    // set the url, number of POST vars, POST data
    //curl_setopt($ch, CURLOPT_RETURNTRANSFER, 1);
    curl_setopt($ch, CURLOPT_URL, $url);
    curl_setopt($ch, CURLOPT_POST, count($fields));
    curl_setopt($ch, CURLOPT_POSTFIELDS, $postvars);

    // execute post
    $result = curl_exec($ch);

    // close connection
    curl_close($ch);

    return json_decode($result, true);
}

function urlToApi($relative){
	$domain = $_SERVER['HTTP_HOST'];
    $ishttps = (isset( $_SERVER["HTTPS"]) && strtolower($_SERVER["HTTPS"]) == "on");
    $prefix = $ishttps ? 'https://' : 'http://';

	return $prefix.$domain.$relative;
}

function generateSelect($name, $placeholder, $options, $default) {
    $html = "<select name='$name'>";
    $html .= "<option disabled selected value>$placeholder</option>";
    foreach ($options as $option => $value) {
        if ($option == $default) {
            $html .= "<option value='$value' selected='selected'>$option</option>";
        } else {
            $html .= "<option value='$value'>$option</option>";
        }
    }

    $html .= '</select>';
    return $html;
}
function generateSelectOptions($options, $default) {
    $html = "";
    foreach ($options as $option => $value) {
        if ($option == $default) {
            $html .= "<option value='$value' selected='selected'>$option</option>";
        } else {
            $html .= "<option value='$value'>$option</option>";
        }
    }
    return $html;
}
?>
