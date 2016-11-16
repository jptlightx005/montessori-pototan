<?php
	// include_once('dss.php');
    session_start();
    include_once('library/global_methods.php');
    $site_logo = "assets/logo.png";
    $site_title = "Exel Montessori de Pototan";
    $_SESSION['isLoggedIn'] = isset($_COOKIE['usrn']) && isset($_COOKIE['pssw']);
    if($_SESSION['isLoggedIn']){
        $usrn = $_COOKIE['usrn'];
        $pssw = $_COOKIE['pssw'];

        $url = urlToApi('/montessori/library/admin.php');
        $fields = array("usrn" => $usrn,
                        "pssw" => $pssw,
                        "role" => "master",
                        "action" => "get_admin_list");
        $obj = jsonFromRequest($fields, $url);
        if($obj["response"] == 1){
            $list = $obj["message"];
        }else{
            echo $obj["message"];
        }
    }
?>

<html>
	<head>
		<title><?php echo $site_title; ?></title>
		<link href="css/general.css" type="text/css" rel="stylesheet">
		<link rel="icon" type="image/png" href="<?php echo $site_logo; ?>" />
		<style>
			table #header{
				font-weight: bold;
				text-align: center;
			}
		</style>
	</head>
	<body>
		<header>
			<img src='<?php echo $site_logo; ?>' alt='insert logo here' />
			<h1><?php echo $site_title; ?></h1>
			<nav>
				<a href="index.php"><span>Home</span></a>
<?php if($_SESSION['isLoggedIn']){ ?>
				<a href="web/login.php?action=logout"><span>Log out</span></a>
<?php } ?>
			</nav>
		</header>
		<div id="content">
<?php if(!$_SESSION['isLoggedIn']){ ?>
				<form id="login_form" action="web/login.php" method="post">
                    <input type='hidden' name="role" value="master" />
					<label>Username:</label><br/>
					<input type="text" name="usrn"/><br/>
					<label>Password:</label><br/>
					<input type="password" name="pssw"/><br/><br/>
					<button type="submit" name="action" value="login">Log-in</button>
				</form>
<?php }else{ ?>
					<h1>Welcome, <?php echo $usrn; ?>!</h1>
					<table border="1">
						<tr id="header">
							<td width="100px">Admin ID</td>
                            <td width="300px">Admin Name</td>
							<td width="300px">Username</td>
							<td width="100px">Role</td>
							<td width="300px">Log-in Count</td>
						</tr>
    <?php foreach($list as $dict){?>
                        <tr id="admin_row">
                            <td width="100px"><?php echo $dict["ID"]; ?></td>
                            <td width="300px"><?php echo $dict["admin_name"]; ?></td>
							<td width="300px"><?php echo $dict["usrn"]; ?></td>
							<td width="100px"><?php echo $dict["role"]; ?></td>
							<td width="300px"><?php echo $dict["login_count"]; ?></td>
                        </tr>
    <?php } ?>
					</table>
					<a href="admin.php">Add Account</a>
<?php } ?>
		</div>
	</body>
</html>
