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
			if($action == "queue_list"){
				$query = "SELECT * FROM `montessori_queue` WHERE `status` = 'onprocess'";
				$result = mysql_query($query);
				$onprocesscount = mysql_num_rows($result);

				// $query = "SELECT Student_ID, first_name, middle_name, last_name, current_grade, status FROM `montessori_queue` INNER JOIN montessori_records ON montessori_records.ID = montessori_queue.Student_ID WHERE `status` = 'onqueue'";
                $query = "SELECT * FROM montessori_records AS r JOIN montessori_queue AS q ON r.ID = q.Student_ID  WHERE `status` = 'onqueue'";
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
                $is_new = isset($_POST['is_new']) ? mysql_real_escape_string($_POST['is_new']) : "";
                $registered_ip = isset($_POST['registered_ip']) ? mysql_real_escape_string($_POST['registered_ip']) : "";
                $school_year = isset($_POST['school_year']) ? mysql_real_escape_string($_POST['school_year']) : "";
                $fields = "(";
				$values = "(";

				foreach($_POST as $key => $value){

					if($key != "usrn" &&
						$key != "pssw" &&
						$key != "role" &&
						$key != "action" &&
                        $key != "is_new" &&
                        $key != "school_year" &&
                        $key != "registered_ip"){
							$newValue = addslashes($value);
							$fields .= "$key, ";
							$values .= "'$newValue', ";
					}
				}

				$fields = substr($fields, 0, strlen($fields) - 2) . ")";
				$values = substr($values, 0, strlen($values) - 2) . ")";
                $theQueries = "";
				$query = "INSERT INTO `montessori_records` $fields VALUES $values";

				$result = mysql_query($query);
				if($result){
					$query = "INSERT INTO `montessori_queue` VALUES ((SELECT LAST_INSERT_ID()), '$usrn', '$registered_ip', '$is_new', '$school_year', 'onqueue', CURRENT_TIMESTAMP)";
					if(mysql_query($query)){
						$json = array("response" => 1, "message" => "Successfully registered!");
					}else{
						$json = array("response" => 0, "message" => "An error has occured while saving!", "query2" => $query);
					}
				}else{
					$json = array("response" => 0, "message" => "An error has occured while saving!", "query" => $query);
				}
			}else if($action == "drop_student"){
				$student_id = isset($_POST['student_id']) ? mysql_real_escape_string($_POST['student_id']) : "";

				$query = "UPDATE `montessori_queue` SET `status` = 'dropped' WHERE `Student_ID` = '$student_id'";

				if(mysql_query($query)){
					$json = array("response" => 1, "message" => "The student has been dropped!");
				}
			}else if($action == "update_student"){
                $is_new = isset($_POST['is_new']) ? mysql_real_escape_string($_POST['is_new']) : "";
                $school_year = isset($_POST['school_year']) ? mysql_real_escape_string($_POST['school_year']) : "";
                $registered_ip = isset($_POST['registered_ip']) ? mysql_real_escape_string($_POST['registered_ip']) : "";
                $student_id = isset($_POST['student_id']) ? mysql_real_escape_string($_POST['student_id']) : "";
                $query = "UPDATE `montessori_queue` SET usrn = '$usrn', rf_ip = '$registered_ip', is_new = '$is_new', temp_school_year = '$school_year', status = 'onqueue', date_registered = CURRENT_TIMESTAMP WHERE Student_ID = $student_id";
                if(mysql_query($query)){
                    $json = array("response" => 1, "message" => $student_id);
                }else{
                    $json = array("response" => 0, "message" => "An error has occured while saving!", "query2" => $query);
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
