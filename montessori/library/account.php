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
			if($action == "register_student"){
				$fields = "(";
				$values = "(";
				foreach($_POST as $key => $value){
					if($key != "usrn" &&
						$key != "pssw" &&
						$key != "role" &
						$key != "action"){
							$newValue = addslashes($value);
							$fields .= "$key, ";
							$values .= "'$newValue', ";
					}
				}

				$fields = substr($fields, 0, strlen($fields) - 2) . ")";
				$values = substr($values, 0, strlen($values) - 2) . ")";

				$query = "INSERT INTO `montessori_records` $fields VALUES $values";
				$result = mysql_query($query);
				if($result){
					$queue_id = $_POST['Queue_ID'];
					$query = "UPDATE `montessori_queue` SET `status` = 'onprocess' WHERE `Queue_ID` = '$queue_id'";

					if(mysql_query($query)){
						$query = "SELECT Student_ID FROM montessori_records WHERE `Queue_ID` = '$queue_id'";
						$result = mysql_query($query);
						if($result){
							$record = mysql_fetch_assoc($result);
							if($record){
								$student_id = addslashes($record['Student_ID']);
								$first_name = addslashes($_POST['first_name']);
								$middle_name = addslashes($_POST['middle_name']);
								$last_name = addslashes($_POST['last_name']);
								$home_address = addslashes($_POST['home_address']);
                                $month_now = date('n');
                                $year_now = date('Y');
                                if($month_now > 6){
                                    $y = $year_now + 1;
                                    $school_year =  "$year_now-$y";
                                }else{
                                    $y = $year_now - 1;
                                    $school_year =  "$y-$year_now";
                                }
								$total_matriculation = addslashes("25000");
								$current_grade = addslashes($_POST['current_grade']);
								$query = "INSERT INTO `montessori_accounts` (Student_ID, Queue_ID, first_name, middle_name, last_name, home_address, school_year, current_grade, total_matriculation, total_payment) VALUES ('$student_id', '$queue_id', '$first_name', '$middle_name', '$last_name', '$home_address', '$school_year', '$current_grade', $total_matriculation, 0)";
								$result = mysql_query($query);
								if($result){
									$json = array("response" => 1, "message" => $student_id);
								}else{
									$json = array("response" => 0, "message" => "An error has occured while saving!");
								}
							}else{
								$json = array("response" => 0, "message" => "Student not found!");
							}
						}else{
							$json = array("response" => 0, "message" => "An error has occured while fetching!");
						}
					}else{
						$json = array("response" => 0, "message" => "An error has occured while saving!");
					}
				}else{
					$json = array("response" => 0, "message" => "An error has occured while saving!", "query" => $query);
				}
			}else if($action == "search_student"){
				$student_id = isset($_POST['student_id']) ? mysql_real_escape_string($_POST['student_id']) : "";
				$query = "SELECT * FROM montessori_accounts WHERE Student_ID = '$student_id'";
				$result = mysql_query($query);
				if($result){
					$record = mysql_fetch_assoc($result);
					if($record)
						$json = array("response" => 1, "message" => $record);
					else
						$json = array("response" => 0, "message" => "Student not found!");
				}else{
					$json = array("response" => 0, "message" => "An error has occured while fetching!");
				}
			}else if($action == "enroll_student"){
				$queue_id = isset($_POST['queue_id']) ? mysql_real_escape_string($_POST['queue_id']) : "";
				$query = "SELECT `status` FROM `montessori_queue` WHERE `Queue_ID` = '$queue_id'";
				$status = mysql_fetch_assoc(mysql_query($query));
				if($status['status'] == "onprocess"){
					$query = "UPDATE `montessori_queue` SET `status` = 'enrolled' WHERE `Queue_ID` = '$queue_id'";
					if(mysql_query($query))
						$json = array("response" => 1, "message" => "Successfully enrolled!");
					else
						$json = array("response" => 0, "message" => "An error has occured while saving!");
				}else if($status['status'] == "enrolled"){
					$json = array("response" => 0, "message" => "The student is already enrolled!");
				}else{
					$json = array("response" => 0, "message" => $query);
				}
			}else if($action == "student_payment"){
				$student_id = isset($_POST['student_id']) ? mysql_real_escape_string($_POST['student_id']) : "";
				$balance_paid = isset($_POST['balance_paid']) ? mysql_real_escape_string($_POST['balance_paid']) : "";
				$query = "UPDATE `montessori_accounts` SET `total_payment` = $balance_paid, `date_of_payment` = CURRENT_TIMESTAMP WHERE `Student_ID` = '$student_id'";

				if(mysql_query($query))
					$json = array("response" => 1, "message" => "Balance successfully updated!", "query" => $query);
				else
					$json = array("response" => 0, "message" => "An error has occured while updating!");
			}else if ($action == "update_student"){
				$student_id = isset($_POST['student_id']) ? mysql_real_escape_string($_POST['student_id']) : "";

				$setFieldValue = "";

				foreach($_POST as $key => $value){
					if($key != "usrn" &&
						$key != "pssw" &&
						$key != "role" &&
						$key != "action" &&
						$key != "student_id"){
							$newValue = addslashes($value);
							$setFieldValue .= "`$key` = '$newValue', ";
					}
				}
				$setFieldValue = substr($setFieldValue, 0, strlen($setFieldValue) - 2);

				$query = "UPDATE `montessori_records` SET $setFieldValue WHERE `Student_ID` = '$student_id'";

				if(mysql_query($query)){
					$current_grade = $_POST['current_grade'];
					$first_name = $_POST['first_name'];
					$middle_name = $_POST['middle_name'];
					$last_name = $_POST['last_name'];
					$home_address = $_POST['home_address'];

					$query = "UPDATE `montessori_accounts` SET first_name = '$first_name', middle_name = '$middle_name', last_name = '$last_name', home_address = '$home_address', current_grade = '$current_grade' WHERE `Student_ID` = '$student_id'";
					if(mysql_query($query))
						$json = array("response" => 1, "message" => "Student Information successfully updated!", "query" => $query);
					else
						$json = array("response" => 0, "message" => "An error has occured while updating!");
				}else{
					$json = array("response" => 0, "message" => "An error has occured while updating!");
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
