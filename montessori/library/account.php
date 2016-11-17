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
			if($action == "register_student"){
                $year_now = date("Y");
                $student_id = $_POST["Student_ID"];
                $studentID = "$year_now-$student_id";
				$setFieldValue = "StudentID = '$studentID', ";

                foreach($_POST as $key => $value){
					if($key != "usrn" &&
						$key != "pssw" &&
						$key != "role" &&
						$key != "action" &&
                        $key != "Student_ID" &&
                        $key != "is_new" &&
                        $key != "school_year" &&
						$key != "total_matriculation"){
							$newValue = addslashes($value);
							$setFieldValue .= "`$key` = '$newValue', ";
					}
				}
				$setFieldValue = substr($setFieldValue, 0, strlen($setFieldValue) - 2);

				$query = "UPDATE `montessori_records` SET $setFieldValue WHERE `ID` = '$student_id'";
				$result = mysql_query($query);
				if($result){
					$query = "UPDATE `montessori_queue` SET `status` = 'onprocess' WHERE `Student_ID` = '$student_id'";

					if(mysql_query($query)){

            $school_year =  addslashes($_POST['school_year']);
  					$total_matriculation = addslashes($_POST['total_matriculation']);

                        $is_new = isset($_POST['is_new']) ? mysql_real_escape_string($_POST['is_new']) : "";
                        if($is_new == 1){
                            $query = "INSERT INTO `montessori_accounts` (Student_ID, school_year, total_matriculation, total_payment) VALUES ('$student_id', '$school_year', $total_matriculation, 0)";
                        }else{
                            $query = "UPDATE `montessori_accounts` SET school_year = '$school_year', total_matriculation = $total_matriculation, total_payment = 0 WHERE Student_ID = '$student_id'";
                        }
						$result = mysql_query($query);
						if($result){
							$json = array("response" => 1, "message" => $studentID);
						}else{
							$json = array("response" => 0, "message" => "An error has occured while saving!", "query" => $query);
						}
					}else{
						$json = array("response" => 0, "message" => "An error has occured while saving!", "query1" => $query);
					}
				}else{
					$json = array("response" => 0, "message" => "An error has occured while saving!", "query2" => $query);
				}
			}else if($action == "search_student"){
				$student_id = isset($_POST['student_id']) ? mysql_real_escape_string($_POST['student_id']) : "";
				// $query = "SELECT Student_ID, Queue_ID, first_name, middle_name, last_name, home_address, school_year, current_grade, total_payment, total_matriculation, latest_payment FROM montessori_records INNER JOIN montessori_accounts ON montessori_records.ID = montessori_accounts.Student_ID WHERE Student_ID = '$student_id'";
                $query = "SELECT * FROM montessori_records AS r JOIN montessori_accounts AS a ON r.ID = a.Student_ID  WHERE StudentID = '$student_id'";
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
				$student_id = isset($_POST['student_id']) ? mysql_real_escape_string($_POST['student_id']) : "";
				$query = "SELECT `status` FROM `montessori_queue` WHERE `Student_ID` = '$student_id'";
				$status = mysql_fetch_assoc(mysql_query($query));
				if($status['status'] == "onprocess"){
					$query = "UPDATE `montessori_queue` SET `status` = 'enrolled' WHERE `Student_ID` = '$student_id'";
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
				$query = "UPDATE `montessori_accounts` SET total_payment = total_payment + $balance_paid, `latest_payment` = CURRENT_TIMESTAMP WHERE `Student_ID` = '$student_id'";

				if(mysql_query($query)){
                    $query = "INSERT INTO `montessori_transactions` VALUES (NULL, $student_id, $balance_paid, CURRENT_TIMESTAMP)";
                    if(mysql_query($query)){
                        $json = array("response" => 1, "message" => "Balance successfully updated!");
                    }else {
                        $json = array("response" => 0, "message" => "An error has occured while updating!");
                    }
				}else{
					$json = array("response" => 0, "message" => "An error has occured while updating!");
                }
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

				$query = "UPDATE `montessori_records` SET $setFieldValue WHERE `ID` = '$student_id'";

                if(mysql_query($query))
                    $json = array("response" => 1, "message" => "Student Information successfully updated!");
                else
                    $json = array("response" => 0, "message" => "An error has occured while updating!", "query" => $query);

			}else if ($action == "transaction_list"){

                $filter_date =  $_POST['filter_date'];
                $query = "SELECT montessori_transactions.ID, first_name, last_name, current_grade, payment, date_of_payment FROM montessori_records INNER JOIN montessori_transactions ON montessori_records.ID = montessori_transactions.Student_ID WHERE CAST(date_of_payment AS Date) = '$filter_date'";

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
