<?php

include_once('db.php');

if($_SERVER['REQUEST_METHOD'] == "POST"){
   // Get data
   $usrn = isset($_POST['usrn']) ? mysql_real_escape_string($_POST['usrn']) : "";
   $pssw = isset($_POST['pssw']) ? mysql_real_escape_string($_POST['pssw']) : "";
   $role = isset($_POST['role']) ? mysql_real_escape_string($_POST['role']) : "";
   $action = isset($_POST['action']) ? mysql_real_escape_string($_POST['action']) : "";
   if(!empty($usrn) && !empty($pssw)){
       // check account
       $query = "SELECT * FROM `montessori_admin` WHERE `usrn` = '$usrn' AND `pssw` = '$pssw' AND `role` = '$role'";
       $result = mysql_query($query);
       $num = mysql_num_rows($result);

       if($num > 0){
           if($action == "register_admin"){
               $username = isset($_POST['username']) ? mysql_real_escape_string($_POST['username']) : "";
               $password = isset($_POST['password']) ? mysql_real_escape_string($_POST['password']) : "";
               $admin_role = isset($_POST['admin_role']) ? mysql_real_escape_string($_POST['admin_role']) : "";
               $full_name = isset($_POST['full_name']) ? mysql_real_escape_string($_POST['full_name']) : "";
               $email = isset($_POST['email']) ? mysql_real_escape_string($_POST['email']) : "";

               $query = "INSERT INTO `montessori_admin` VALUES(NULL, '$username', '$password', '$admin_role', '$full_name', '$email', 0)";
               $result = mysql_query($query);
               if($result){
                   $json = array("response" => 1, "message" => "Successfully registered admin!");
               }else{
                   $json = array("response" => 0, "message" => "An error has occured while saving!", "query" => $query);
               }
           }else if($action == "get_admin_list"){
               $query = "SELECT ID, usrn, role, admin_name, login_count FROM `montessori_admin`";
               $result = mysql_query($query);

               if($result){
                   while($row = mysql_fetch_assoc($result)){
                        $rows[] = $row;
                   }
                   $json = array("response" => 1, "message" => $rows);
               }else{
                   $json = array("response" => 1, "message" => "Failed to fetch admin list.");
               }
           }
       }
       else{
           $json = array("response" => $num, "message" => "Invalid!");
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
