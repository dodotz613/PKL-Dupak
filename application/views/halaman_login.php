<!DOCTYPE html>
<html>
<head>
	<title>Login</title>
</head>
<link rel="stylesheet" href="https://stackpath.bootstrapcdn.com/bootstrap/4.5.2/css/bootstrap.min.css" integrity="sha384-JcKb8q3iqJ61gNV9KGb8thSsNjpSL0n8PARn9HuZOnIxN0hoP+VmmDGMN5t9UJ0Z" crossorigin="anonymous">
<link rel="stylesheet" type="text/css" href="<?php echo base_url()?>css/css_login.css">
<body>
<img src="<?php echo base_url(); ?>assets/unila.png" width="350">
<div class = "container">
	<div class="wrapper">
		<form action="<?php echo site_url('login/aksi_login')?>" method="post" name="Login_Form" class="form-signin">
		    <h3 class="form-signin-heading">Hello Selamat Datang silahkan login</h3>
			  <hr class="colorgraph"><br>

			  <?php
				$info = $this->session->flashdata('info');
				if (!empty($info)){
					echo $info;
				}
			  ?>

			  <input type="text" class="form-control" name="username" placeholder="Username" required="" autofocus="" />
			  <input type="password" class="form-control" name="password" placeholder="Password" required=""/>

			  <button class="btn btn-lg btn-primary btn-block"  name="Submit" value="Login" type="Submit">Login</button>
		</form>
	</div>
</div>
</body>
</html>
