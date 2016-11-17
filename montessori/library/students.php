 <?php

include_once('db.php');
if($_SERVER['REQUEST_METHOD'] == "POST"){
	// Get data
	$usrn = isset($_POST['usrn']) ? mysql_real_escape_string($_POST['usrn']) : "";
	$pssw = isset($_POST['pssw']) ? mysql_real_escape_string($_POST['pssw']) : "";
	$role = isset($_POST['role']) ? mysql_real_escape_string($_POST['role']) : "";
	$action = isset($_POST['action']) ? mysql_real_escape_string($_POST['action']) : "";

	if(!empty($usrn) && !empty($pssw)){
		// check authorization
		$query = "SELECT * FROM `montessori_admin` WHERE `usrn` = '$usrn' AND `pssw` = '$pssw' AND (`role` = '$role' OR `role` = 'master')";
		$result = mysql_query($query);
		$num = mysql_num_rows($result);

		if($num > 0){
			if($action == "search_student"){
				$filter_key = isset($_POST['filter_key']) ? mysql_real_escape_string($_POST['filter_key']) : "";
				$query = "SELECT * FROM `montessori_records` AS r JOIN montessori_queue AS a ON r.ID = a.Student_ID WHERE status = 'enrolled'";
				if($filter_key != ""){
					$filter_value = isset($_POST['filter_value']) ? mysql_real_escape_string($_POST['filter_value']) : "";
					$query .= " AND `$filter_key` LIKE '%$filter_value%'";
				}
				$result = mysql_query($query);
				$studentcount = mysql_num_rows($result);
				if($studentcount > 0){
					while($row = mysql_fetch_assoc($result)){
						 $rows[] = $row;
					}

					$json = array("response" => 1, "message" => $rows);
				}else{
					$json = array("response" => 0, "message" => "No results");
				}
			}
		}else{
			$json = array("response" => -1, "message" => "Invalid Request");
		}
	}else{
		$json = array("response" => -1, "message" => "Invalid Request");
	}
}else{
	$json = array("response" => -1, "message" => "Unknown Method");
}

 @mysql_close($conn);

 /* Output header */
 header('Content-type: application/json');
 //echo $json;
 echo json_encode($json);

?>
