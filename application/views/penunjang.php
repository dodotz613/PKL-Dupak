<?php
defined('BASEPATH') OR exit('No direct script access allowed');
?>

<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <meta http-equiv="x-ua-compatible" content="ie=edge">
  <title>Penunjang</title>
  <link rel="stylesheet" href="<?php echo base_url(); ?>assets/AdminLTE-3.0.1/plugins/fontawesome-free/css/all.min.css">
  <link rel="stylesheet" href="<?php echo base_url(); ?>assets/AdminLTE-3.0.1/dist/css/adminlte.min.css">
  <link rel="stylesheet" href="http://code.ionicframework.com/ionicons/2.0.1/css/ionicons.min.css">
  <link href="https://fonts.googleapis.com/css?family=Source+Sans+Pro:300,400,400i,700" rel="stylesheet">
</head>

<body class="hold-transition sidebar-mini">
  <!-- Site wrapper -->
  <div class="wrapper">
  <!-- Navbar -->
  <nav class="main-header navbar navbar-expand navbar-white navbar-light">
    <!-- Left navbar links -->
    <ul class="navbar-nav">
      <li class="nav-item">
        <a class="nav-link" data-widget="pushmenu" href="#"><i class="fas fa-bars"></i></a>
      </li>
    </ul>


    <!-- Right navbar links -->
    <ul class="navbar-nav ml-auto">
      <!-- Messages Dropdown Menu -->
      <li class="nav-item dropdown">
        <a class="nav-link" data-toggle="dropdown" href="#">
          <i class="fas fa-power-off"> <?php echo $this->session->userdata("nama");?> </i>
         </a>
        <div class="dropdown-menu dropdown-menu-lg dropdown-menu-right">
             <a href="<?php echo site_url('login/logout'); ?>" class="dropdown-item dropdown-footer">Keluar</a>
        </div>
       </ul>
  </nav>

  <!-- Main Sidebar Container -->
  <aside class="main-sidebar sidebar-dark-primary elevation-4">
    <!-- Brand Logo -->
    <a href="#" class="brand-link">
      <img src="<?php echo base_url(); ?>assets/AdminLTE-3.0.1/dist/img/himakom.jpg" alt="AdminLTE Logo" class="brand-image img-circle elevation-3"
           style="opacity: .8">
      <span class="brand-text font-weight-light">Ilkom Unila</span>
    </a>

    <!-- Sidebar -->
    <div class="sidebar">
      <!-- Sidebar user panel (optional) -->
      <div class="user-panel mt-3 pb-3 mb-3 d-flex">
        <div class="image">
          <img src="<?php echo base_url(); ?>assets/AdminLTE-3.0.1/dist/img/default.jpg" class="img-circle elevation-2" alt="User Image">
        </div>
        <div class="info">
          <a href="#" class="d-block"><?php echo $this->session->userdata("nama");?></a>
        </div>
      </div>

      <!-- Sidebar Menu -->
      <nav class="mt-2">
        <ul class="nav nav-pills nav-sidebar flex-column" data-widget="treeview" role="menu" data-accordion="false">
          <!-- Add icons to the links using the .nav-icon class
               with font-awesome or any other icon font library -->
               <li class="nav-item">
              <a href="<?php echo site_url('dashboard')?>" class="nav-link">
                <i class="nav-icon fas fa-tachometer-alt"></i>
                 <p>Dashboard</p>
              </a>
                </li>

              <li class="nav-item">
                <a href="<?php echo site_url('pendidikan')?>" class="nav-link">
                <i class="nav-icon fas fa-graduation-cap"></i>
                  <p>Pendidikan</p>
                </a>
              </li>

              <li class="nav-item">
                <a href="<?php echo site_url('penelitian')?>" class="nav-link">
                <i class="nav-icon fas fa-book"></i>
                  <p>Penelitian</p>
                </a>
              </li>

                <li class="nav-item">
                <a href="<?php echo site_url('pengabdian')?>" class="nav-link">
                <i class="nav-icon fas fa-briefcase"></i>
                    <p>Pengabdian</p>
                  </a>
              </li>

              <li class="nav-item">
                <a href="<?php echo site_url('penunjang')?>" class="nav-link">
                <i class="nav-icon fas fa-book"></i>
                  <p>Penunjang</p>
                </a>
              </li>

              <li class="nav-item">
                <a href="<?php echo site_url('report')?>" class="nav-link">
                <i class="nav-icon fas fa-book"></i>
                  <p>Report</p>
                </a>
              </li>
        </ul>
      </nav>
    </div>
  </aside>

  <!-- Content Wrapper. Contains page content -->
  <div class="content-wrapper">
    <!-- Content Header (Page header) -->
    <div class="content-header">
      <div class="container-fluid">
        <div class="row mb-2">
          <div class="col-sm-6">
            <h1 class="m-0 text-dark">Penunjang</h1>
          </div>
        </div>
      </div>
    </div>

    <!-- Main content -->
    <div class="content">
      <div class="container-fluid">
        <div class="row">
          <div class="col-lg-6">
            <div class="card">
              <div class="card-header border-0">
                <div class="d-flex justify-content-between">
                  <h3 class="card-title">Halaman Penunjang</h3>
                </div>
              </div>
              <div class="card-body">
                <div class="d-flex">
                  <p class="d-flex flex-column">
                    <span class="text-bold text-lg">Jika ingin mendownload file data penunjang dosen silahkan klik tombol di bawah ini</span>
                  </p>
                </div>

                            <div class="timeline-footer">
                            <a href="<?=base_url()?>index.php/penunjang/export" class="btn btn-primary">Download</a>
                            </div>

              </div>
            </div>
          </div>
          </div>
        </div>
      </div>
    </div>
  </div>

  <!-- Control Sidebar -->
  <aside class="control-sidebar control-sidebar-dark">
    <!-- Control sidebar content goes here -->
  </aside>
  <!-- Main Footer -->
  <footer class="main-footer">
    <strong>Copyright &copy; 2020 <a>Hengki - Sulung - Udin</a></strong>
    <div class="float-right d-none d-sm-inline-block">
    </div>
  </footer>

<!-- REQUIRED SCRIPTS -->

<!-- jQuery -->
<script src="<?php echo base_url(); ?>assets/AdminLTE-3.0.1/plugins/jquery/jquery.min.js"></script>
<!-- Bootstrap -->
<script src="<?php echo base_url(); ?>assets/AdminLTE-3.0.1/plugins/bootstrap/js/bootstrap.bundle.min.js"></script>
<!-- AdminLTE -->
<script src="<?php echo base_url(); ?>assets/AdminLTE-3.0.1/dist/js/adminlte.js"></script>

<!-- OPTIONAL SCRIPTS -->
<script src="<?php echo base_url(); ?>assets/AdminLTE-3.0.1/plugins/chart.js/Chart.min.js"></script>
<script src="<?php echo base_url(); ?>assets/AdminLTE-3.0.1/dist/js/demo.js"></script>
<script src="<?php echo base_url(); ?>assets/AdminLTE-3.0.1/dist/js/pages/dashboard3.js"></script>
</body>
</html>
