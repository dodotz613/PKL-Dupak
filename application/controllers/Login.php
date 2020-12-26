<?php
defined('BASEPATH') OR exit('No direct script access allowed');

class Login extends CI_Controller {

	function __construct(){
        parent::__construct();
        $this->load->model('model_login');

    }


	function index(){
        $this->load->view('halaman_login');
    }

    function aksi_login(){
        $username = $this->input->post('username');
        $password = $this->input->post('password');
        $where = array(
            'username' => $username,
            'password' => md5($password)
            );
        $cek = $this->model_login->cek_login("login",$where)->num_rows();
        if($cek > 0){

            $data_session = array(
                'nama' => $username,
                'status' => "login"
                );

            $this->session->set_userdata($data_session);

            redirect('user');

        }else{
            $this->session->set_flashdata('info','Username dan Password Salah !!!');
			redirect('login');
        }
    }

    function logout(){
        $this->session->sess_destroy();
        redirect('login');
    }
}
