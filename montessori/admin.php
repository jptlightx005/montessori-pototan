<?php
	// include_once('dss.php');
    include_once('library/global_methods.php');
    session_start();
    $site_logo = "assets/logo.png";
    $site_title = "Exel Montessori de Pototan";
    $_SESSION['isLoggedIn'] = isset($_COOKIE['usrn']) && isset($_COOKIE['pssw']);
    if($_SESSION['isLoggedIn']){
        $usrn = $_COOKIE['usrn'];
        $pssw = $_COOKIE['pssw'];
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
        <script>
		function validateForm() {
			var login = document.forms["login_form"];
            var usrn = login["username"].value;
			var pssw = login["password"].value;
			var conf_pssw = login["conf_pssw"].value;
            var email = login["email"];
            var full_name = login["full_name"];
            var admin_role = login["admin_role"];
            if (usrn.length == 0){
                alert("Please enter admin's username!");
				return false;
            }
            if (pssw.length == 0){
                alert("Please enter admin's password!");
				return false;
            }
			if (pssw != conf_pssw) {
				alert("Passwords didn't match!");
				return false;
			}
            if (email.length == 0){
                alert("Please enter admin's email!");
				return false;
            }
            if (full_name.length == 0){
                alert("Please enter admin's Full Name!");
				return false;
            }
            if (admin_role.length == 0){
                alert("Please select a Role!");
				return false;
            }
		}
		</script>
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
<?php if($_SESSION['isLoggedIn']){ ?>
    <form name="login_form" action="web/register.php" method="post" onsubmit="return validateForm()">
        <label>Username:</label><br/>
        <input type='hidden' name="usrn" value="<?php echo $_COOKIE['usrn']; ?>" />
        <input type='hidden' name="pssw" value="<?php echo $_COOKIE['pssw']; ?>" />
        <input type='hidden' name="role" value="master" />
        <input type="text" name="username"/><br/>
        <label>Password:</label><br/>
        <input type="password" name="password"/><br/>
        <label>Confirm Password:</label><br/>
        <input type="password" name="conf_pssw"/><br/>
        <label>Email:</label><br/>
        <input type="email" name="email"/><br/>
        <label>Full Name:</label><br/>
        <input type="text" name="full_name"/><br/>
        <label>Gender:</label><br>
        <select name="admin_role">
            <option disabled selected value>-- Select a Role --</option>
<?php
      echo generateSelectOptions(array('Admin' => 'admin', 'Registrar' => 'registrar', 'Accountant' => 'accountant'), '');
?>
          </select><br>
        <br/>
        <button type="submit" name="action" value="register_admin">Register</button>
    </form>
<?php } ?>
		</div>
	</body>
</html>
