 <?php
 
include_once('db.php');

if($_SERVER['REQUEST_METHOD'] == "POST"){
	// Get data
	$usrn = isset($_POST['usrn']) ? mysql_real_escape_string($_POST['usrn']) : "";
	$pssw = isset($_POST['pssw']) ? mysql_real_escape_string($_POST['pssw']) : "";
	$role = isset($_POST['role']) ? mysql_real_escape_string($_POST['role']) : "";
	if(!empty($usrn) && !empty($pssw)){
		// check account
		$query = "SELECT * FROM `montessori_admin` WHERE `usrn` = '$usrn' AND `pssw` = '$pssw' AND `role` = '$role'";
		$result = mysql_query($query);
		$num = mysql_num_rows($result);

		if($num > 0){
			$json = array("response" => $num, "message" => "Successfully logged in!");
			$query = "UPDATE `montessori_admin` SET login_count = login_count + 1 WHERE `usrn` = '$usrn' AND `pssw` = '$pssw'";
			mysql_query($query);
		}
		else{
			$json = array("response" => $num, "message" => "Failed to log in!");
		}
	}else{
		$json = array("response" => -1, "message" => "Please enter username and password!");
	}
}
	 
 @mysql_close($conn);
 
 /* Output header */
 header('Content-type: application/json');
 
 echo json_encode($json);
	 
?>