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
							$fields .= "$key, ";
							$values .= "'$value', ";
					}
				}
				$fields = substr($fields, 0, strlen($fields) - 2) . ")";
				$values = substr($values, 0, strlen($values) - 2) . ")";
				
				$query = "INSERT INTO `montessori_records` $fields VALUES $values";
				$result = mysql_query($query);
				if($result){
					$queue_id = $_POST['Queue_ID'];
					$query = "UPDATE `montessori_queue` SET `status` = 'onprocess' WHERE `Queue_ID` = '$queue_id'";
					
					if(mysql_query($query))
						$json = array("response" => 1, "message" => "Successfully registered!");
					else
						$json = array("response" => 0, "message" => "An error has occured while saving!");
				}else{
					$json = array("response" => 0, "message" => "An error has occured while saving!");
				}
			}else if($action == "search_student"){
				$student_id = isset($_POST['student_id']) ? mysql_real_escape_string($_POST['student_id']) : "";
				$query = "SELECT * FROM montessori_records WHERE Student_ID = '$student_id'";
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
				$query = "UPDATE `montessori_queue` SET `status` = 'enrolled' WHERE `Queue_ID` = '$queue_id'";
				
				if(mysql_query($query))
					$json = array("response" => 1, "message" => "Successfully enrolled!");
				else
					$json = array("response" => 0, "message" => "An error has occured while saving!");
			}else if($action == "student_payment"){
				$student_id = isset($_POST['student_id']) ? mysql_real_escape_string($_POST['student_id']) : "";
				$balance_paid = isset($_POST['balance_paid']) ? mysql_real_escape_string($_POST['balance_paid']) : "";
				$query = "UPDATE `montessori_records` SET `balance_paid` = $balance_paid, `date_of_payment` = CURRENT_TIMESTAMP WHERE `Student_ID` = '$student_id'";
				
				if(mysql_query($query))
					$json = array("response" => 1, "message" => "Balance successfully updated!");
				else
					$json = array("response" => 0, "message" => "An error has occured while updating!");
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