<?php
if ( ! defined('BASEPATH')) exit('No direct script access allowed');

class Lektor extends CI_Model {

  public function view(){

    return $this->db->get('dosen_penunjang')->result();
  }

}
