<?php
defined('BASEPATH') OR exit('No direct script access allowed');

class Dashboard extends CI_Controller {

	 public function __construct(){
			parent::__construct();

			$this->load->model('Dosen');
			$this->load->model('Lektor');

			if($this->session->userdata('status') != "login"){
				redirect('login');
			}
		}


		public function index() {

		$this->load->view('dashboard');
  }

}
