<?php
defined('BASEPATH') OR exit('No direct script access allowed');

class Model_login extends CI_Model {

	function cek_login($table,$where){
        return $this->db->get_where($table,$where);
	}

	public function keamanan() {

		$username = $this->session->sess_destroy('username');
		if (!empty($username)){
			$this->session->sess_destroy();
			redirect('login');
		}
	}

}
