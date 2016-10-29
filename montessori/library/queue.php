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
		$query = "SELECT * FROM `montessori_admin` WHERE `usrn` = '$usrn' AND `pssw` = '$pssw' AND `role` = '$role'";
		$result = mysql_query($query);
		$num = mysql_num_rows($result);

		if($num > 0){
			if($action == "queue_list"){
				$query = "SELECT * FROM `montessori_queue` WHERE `status` = 'onprocess'";
				$result = mysql_query($query);
				$onprocesscount = mysql_num_rows($result);

				$query = "SELECT Queue_ID, student_info FROM `montessori_queue` WHERE `status` = 'onqueue'";
				$result = mysql_query($query);
				$onqueuecount = mysql_num_rows($result);

				if($onqueuecount > 0){
					while($row = mysql_fetch_assoc($result)){
						 $rows[] = $row;
					}

					$message = array("onqueue" => $onqueuecount, "onprocess" => $onprocesscount, "list" => $rows);
					$json = array("response" => 1, "message" => $message);
				}else{
					$message = array("onqueue" => $onqueuecount, "onprocess" => $onprocesscount, "list" => []);
					$json = array("response" => 1, "message" => $message);
				}
			}else if($action == "register_student"){
				$student_info = isset($_POST['student_info']) ? mysql_real_escape_string($_POST['student_info']) : "";
				$registered_ip = isset($_POST['registered_ip']) ? mysql_real_escape_string($_POST['registered_ip']) : "";
				$query = "SELECT * FROM `montessori_queue` WHERE `student_info` = '$student_info'";
				$result = mysql_query($query);
				$sameinfocount =  mysql_num_rows($result);

				if($sameinfocount == 0){
					$query = "INSERT INTO `montessori_queue` VALUES (NULL, '$usrn', '$registered_ip', '$student_info', 'onqueue', CURRENT_TIMESTAMP)";
					if(mysql_query($query)){
						$query = "SELECT Queue_ID FROM montessori_queue WHERE `student_info` = '$student_info'";
						$result = mysql_query($query);
						if($result){
							$record = mysql_fetch_assoc($result);
							if($record)
								$json = array("response" => 1, "message" => $record['Queue_ID']);
							else
								$json = array("response" => 0, "message" => "Student not found!");
						}else{
							$json = array("response" => 0, "message" => "An error has occured while fetching!");
						}

					}else{
						$json = array("response" => 0, "message" => $query);
					}

				}else{
					$json = array("response" => 0, "message" => "The student already exists!");
				}
			}else if($action == "drop_student"){
				$queue_id = isset($_POST['queue_id']) ? mysql_real_escape_string($_POST['queue_id']) : "";

				$query = "UPDATE `montessori_queue` SET `status` = 'dropped' WHERE `Queue_ID` = '$queue_id'";

				if(mysql_query($query)){
					$json = array("response" => 1, "message" => "The student has been dropped!");
				}
			}
		}else{
			$json = array("response" => -1, "message" => "Invalid Request");
		}

	}else{
		$json = array("response" => -1, "message" => "Invalid Request");
	}
}

 @mysql_close($conn);

 /* Output header */
 header('Content-type: application/json');
 //echo $json;
 echo json_encode($json);

?>
