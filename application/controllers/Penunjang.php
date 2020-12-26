<?php
defined('BASEPATH') OR exit('No direct script access allowed');

class Penunjang extends CI_Controller {

	public function __construct(){
		parent::__construct();

		$this->load->model('Dosen');
		$this->load->model('Lektor');

		if($this->session->userdata('status') != "login"){
            redirect('login');
        }
	  }

	public function index(){

		$data['data_dosen'] = $this->Dosen->view();
		$data['dosen_penunjang'] = $this->Lektor->view();
		$this->load->view('penunjang', $data);
		}

	public function export(){
			// Load plugin PHPExcel nya
		include APPPATH.'third_party/PHPExcel/PHPExcel.php';

			// Panggil class PHPExcel nya
		$excel = new PHPExcel();

		// Settingan awal fil excel
		$excel->getProperties()->setCreator('Sulung')
							   ->setLastModifiedBy('Sulung')
							   ->setTitle("Penunjang")
							   ->setSubject("Dupak")
							   ->setDescription("Laporan")
							   ->setKeywords("Penunjang");

		$style_standar = array(

	  	// Set font nya jadi bold
			'borders' => array(
			'top' => array('style'  => PHPExcel_Style_Border::BORDER_THIN), // Set border top dengan garis tipis
			'right' => array('style'  => PHPExcel_Style_Border::BORDER_THIN),  // Set border right dengan garis tipis
			'bottom' => array('style'  => PHPExcel_Style_Border::BORDER_THIN), // Set border bottom dengan garis tipis
			'left' => array('style'  => PHPExcel_Style_Border::BORDER_THIN) // Set border left dengan garis tipis

						 )
			);

		// Buat sebuah variabel untuk menampung pengaturan style dari header tabel
		$style_col = array(

			// Set font nya jadi bold
			'alignment' => array(
			'horizontal' => PHPExcel_Style_Alignment::HORIZONTAL_CENTER, // Set text jadi ditengah secara horizontal (center)
			'vertical' => PHPExcel_Style_Alignment::VERTICAL_CENTER // Set text jadi di tengah secara vertical (middle)
				),

				'borders' => array(
				'top' => array('style'  => PHPExcel_Style_Border::BORDER_THIN), // Set border top dengan garis tipis
				'right' => array('style'  => PHPExcel_Style_Border::BORDER_THIN),  // Set border right dengan garis tipis
				'bottom' => array('style'  => PHPExcel_Style_Border::BORDER_THIN), // Set border bottom dengan garis tipis
				'left' => array('style'  => PHPExcel_Style_Border::BORDER_THIN) // Set border left dengan garis tipis
				)
			);

		// Buat sebuah variabel untuk menampung pengaturan style dari isi tabel
		$style_row = array(
			'alignment' => array(
			'vertical' => PHPExcel_Style_Alignment::VERTICAL_CENTER // Set text jadi di tengah secara vertical (middle)
				),

				'borders' => array(
				'top' => array('style'  => PHPExcel_Style_Border::BORDER_THIN), // Set border top dengan garis tipis
				'right' => array('style'  => PHPExcel_Style_Border::BORDER_THIN),  // Set border right dengan garis tipis
				'bottom' => array('style'  => PHPExcel_Style_Border::BORDER_THIN), // Set border bottom dengan garis tipis
				'left' => array('style'  => PHPExcel_Style_Border::BORDER_THIN) // Set border left dengan garis tipis
				)
			);

		$excel->getDefaultStyle()
			->getFont()
			->setName('Arial')
			->setSize(11);


		$excel->setActiveSheetIndex(0)->setCellValue('A1', "SURAT PERNYATAAN");
		$excel->setActiveSheetIndex(0)->setCellValue('A2', "MELAKSANAKAN PENUNJANG TUGAS DOSEN");
		$excel->getActiveSheet()->mergeCells('A1:R1');
		$excel->getActiveSheet()->mergeCells('A2:R2');
		$excel->getActiveSheet()->getStyle('A1:A2')->getFont()->setSize(11);
		$excel->getActiveSheet()->getStyle('A1:A2')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);

		$excel->setActiveSheetIndex(0)->setCellValue('B4', "Nama ");
		$excel->setActiveSheetIndex(0)->setCellValue('J4', ":");
		$excel->setActiveSheetIndex(0)->setCellValue('B5', "NIP");
		$excel->setActiveSheetIndex(0)->setCellValue('J5', ":");
		$excel->setActiveSheetIndex(0)->setCellValue('B6', "Pangkat");
		$excel->setActiveSheetIndex(0)->setCellValue('J6', ":");
		$excel->setActiveSheetIndex(0)->setCellValue('B7', "Golongan");
		$excel->setActiveSheetIndex(0)->setCellValue('J7', ":");
		$excel->setActiveSheetIndex(0)->setCellValue('B8', "Jabatan");
		$excel->setActiveSheetIndex(0)->setCellValue('J8', ":");
		$excel->setActiveSheetIndex(0)->setCellValue('B9', "Unit Kerja");
		$excel->setActiveSheetIndex(0)->setCellValue('J9', ":");
		$excel->setActiveSheetIndex(0)->setCellValue('B10', "Menyatakan ");
		$excel->setActiveSheetIndex(0)->setCellValue('B11', "Nama");
		$excel->setActiveSheetIndex(0)->setCellValue('J11', ":");
		$excel->setActiveSheetIndex(0)->setCellValue('B12', "NIP");
		$excel->setActiveSheetIndex(0)->setCellValue('J12', ":");
		$excel->setActiveSheetIndex(0)->setCellValue('B13', "Pangkat");
		$excel->setActiveSheetIndex(0)->setCellValue('J13', ":");
		$excel->setActiveSheetIndex(0)->setCellValue('B14', "Golongan");
		$excel->setActiveSheetIndex(0)->setCellValue('J14', ":");
		$excel->setActiveSheetIndex(0)->setCellValue('B15', "Jabatan");
		$excel->setActiveSheetIndex(0)->setCellValue('J15', ":");
		$excel->setActiveSheetIndex(0)->setCellValue('B16', "Unit Kerja");
		$excel->setActiveSheetIndex(0)->setCellValue('J16', ":");
		$excel->setActiveSheetIndex(0)->setCellValue('B18', "Telah menyatakan pengabdian masyarakat sebagai berikut : ");

			$data_dosen = $this->Dosen->view();

			foreach($data_dosen as $data) {

				$excel->setActiveSheetIndex(0)->setCellValue('K4', $data->nama);
				$excel->setActiveSheetIndex(0)->setCellValue('K5', $data->nip);
				$excel->setActiveSheetIndex(0)->setCellValue('K6', $data->pangkat);
				$excel->setActiveSheetIndex(0)->setCellValue('K7', $data->golongan);
				$excel->setActiveSheetIndex(0)->setCellValue('K8', $data->jabatan);
				$excel->setActiveSheetIndex(0)->setCellValue('K9', $data->unit_kerja);
			}

			$dosen_penunjang = $this->Lektor->view();

			foreach($dosen_penunjang as $data) {

				$excel->setActiveSheetIndex(0)->setCellValue('K11', $data->nama);
				$excel->setActiveSheetIndex(0)->setCellValue('K12', $data->nip);
				$excel->setActiveSheetIndex(0)->setCellValue('K13', $data->pangkat);
				$excel->setActiveSheetIndex(0)->setCellValue('K14', $data->golongan);
				$excel->setActiveSheetIndex(0)->setCellValue('K15', $data->jabatan);
				$excel->setActiveSheetIndex(0)->setCellValue('K16', $data->unit_kerja);
			}



		$excel->setActiveSheetIndex(0)->setCellValue('A20', "No");
		$excel->getActiveSheet()->getStyle('A20')->getAlignment()->setWrapText(TRUE);
		$excel->setActiveSheetIndex(0)->setCellValue('B20', "Uraian Kegiatan");
		$excel->getActiveSheet()->mergeCells('B20:K20');
		$excel->getActiveSheet()->getStyle('B20')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
		$excel->getActiveSheet()->getStyle('B20')->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
		$excel->setActiveSheetIndex(0)->setCellValue('L20', "Tanggal");
		$excel->setActiveSheetIndex(0)->setCellValue('M20', "Satuan Hasil");
		$excel->getActiveSheet()->getStyle('M20')->getAlignment()->setWrapText(TRUE);
		$excel->setActiveSheetIndex(0)->setCellValue('N20', "Jumlah Volume Kegiatan");
		$excel->getActiveSheet()->getStyle('N20')->getAlignment()->setWrapText(TRUE);
		$excel->setActiveSheetIndex(0)->setCellValue('O20', "Angka Kredit");
		$excel->getActiveSheet()->getStyle('O20')->getAlignment()->setWrapText(TRUE);
		$excel->setActiveSheetIndex(0)->setCellValue('P20', "Jumlah Angka Kredit");
		$excel->getActiveSheet()->getStyle('P20')->getAlignment()->setWrapText(TRUE);
		$excel->setActiveSheetIndex(0)->setCellValue('Q20', "Keterangan/Bukti Fisik");
		$excel->getActiveSheet()->mergeCells('Q20:R20');
		$excel->getActiveSheet()->getStyle('Q20')->getAlignment()->setWrapText(TRUE);

		$excel->getActiveSheet()->getStyle('A20:A203')->applyFromArray($style_col);
		$excel->getActiveSheet()->getStyle('L20:L203')->applyFromArray($style_col);
		$excel->getActiveSheet()->getStyle('M20:M203')->applyFromArray($style_col);
		$excel->getActiveSheet()->getStyle('N20:N203')->applyFromArray($style_col);
		$excel->getActiveSheet()->getStyle('O20:O203')->applyFromArray($style_col);
		$excel->getActiveSheet()->getStyle('P20:P203')->applyFromArray($style_col);
		$excel->getActiveSheet()->getStyle('R21:R203')->applyFromArray($style_col);
		$excel->getActiveSheet()->getStyle('A20')->applyFromArray($style_standar);
		$excel->getActiveSheet()->getStyle('B20')->applyFromArray($style_standar);
		$excel->getActiveSheet()->getStyle('C20')->applyFromArray($style_standar);
		$excel->getActiveSheet()->getStyle('D20')->applyFromArray($style_standar);
		$excel->getActiveSheet()->getStyle('E20')->applyFromArray($style_standar);
		$excel->getActiveSheet()->getStyle('F20')->applyFromArray($style_standar);
		$excel->getActiveSheet()->getStyle('G20')->applyFromArray($style_standar);
		$excel->getActiveSheet()->getStyle('H20')->applyFromArray($style_standar);
		$excel->getActiveSheet()->getStyle('I20')->applyFromArray($style_standar);
		$excel->getActiveSheet()->getStyle('J20')->applyFromArray($style_standar);
		$excel->getActiveSheet()->getStyle('K20')->applyFromArray($style_standar);
		$excel->getActiveSheet()->getStyle('L20:Q20')->applyFromArray($style_standar);
		$excel->getActiveSheet()->getStyle('R20')->applyFromArray($style_standar);

		$excel->setActiveSheetIndex(0)->setCellValue('A21', "(1)");
		$excel->setActiveSheetIndex(0)->setCellValue('B21', "(2)");
		$excel->getActiveSheet()->getStyle('B21')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
		$excel->getActiveSheet()->mergeCells('B21:K21');
		$excel->getActiveSheet()->getStyle('A21:Q21')->applyFromArray($style_standar);
		$excel->setActiveSheetIndex(0)->setCellValue('L21', "(3)");
		$excel->setActiveSheetIndex(0)->setCellValue('M21', "(4)");
		$excel->setActiveSheetIndex(0)->setCellValue('N21', "(5)");
		$excel->setActiveSheetIndex(0)->setCellValue('O21', "(6)");
		$excel->setActiveSheetIndex(0)->setCellValue('P21', "(7)");
		$excel->setActiveSheetIndex(0)->setCellValue('R21', "(8)");
		$excel->getActiveSheet()->getStyle('R21')->getAlignment()->setWrapText(TRUE);
		$excel->getActiveSheet()->getStyle('R21')->applyFromArray($style_standar);

		$excel->setActiveSheetIndex(0)->setCellValue('A22', "IV.");
		$excel->getActiveSheet()->getStyle('A22')->getFont()->setBold(TRUE);
		$excel->getActiveSheet()->getStyle('A22')->getFont()->setSize(11);
		$excel->setActiveSheetIndex(0)->setCellValue('B22', "PENUNJANG TUGAS");
		$excel->getActiveSheet()->mergeCells('B22:K22');
		$excel->getActiveSheet()->getStyle('B22')->getFont()->setBold(TRUE);
		$excel->getActiveSheet()->getStyle('B22')->getFont()->setSize(11);
		$excel->getActiveSheet()->getStyle('A22:R22')->applyFromArray($style_standar);

		$excel->setActiveSheetIndex(0)->setCellValue('A23', "A");
		$excel->setActiveSheetIndex(0)->setCellValue('B23', "Menjadi anggota dalam suatu Panitia/Badan");
		$excel->getActiveSheet()->mergeCells('B23:K23');
		$excel->setActiveSheetIndex(0)->setCellValue('B24', "2.");
		$excel->setActiveSheetIndex(0)->setCellValue('C24', "Sebagai anggota");
		$excel->setActiveSheetIndex(0)->setCellValue('Q24', "VI.A.2");
		$excel->setActiveSheetIndex(0)->setCellValue('C25', "1)");
		$excel->setActiveSheetIndex(0)->setCellValue('D25', "Anggota Panitia Tim Penjamin Mutu Program Sarjana (S1) FMIPA Unila ");
		$excel->setActiveSheetIndex(0)->setCellValue('L25', "Sem Ganjil ");
		$excel->setActiveSheetIndex(0)->setCellValue('M25', "1 Semester");
		$excel->setActiveSheetIndex(0)->setCellValue('N25', "1,0");
		$excel->setActiveSheetIndex(0)->setCellValue('O25', "1,0");
		$excel->setActiveSheetIndex(0)->setCellValue('P25', "1,0");
		$excel->setActiveSheetIndex(0)->setCellValue('Q25', "1");
		$excel->setActiveSheetIndex(0)->setCellValue('R25', "SK Dekan FMIPA Unila");
		$excel->setActiveSheetIndex(0)->setCellValue('L26', "2013/2014");
		$excel->setActiveSheetIndex(0)->setCellValue('R26', "No 473/UN26/7/DT/2013");
		$excel->setActiveSheetIndex(0)->setCellValue('R27', "Tanggal 4 Januari 2013");

		$excel->setActiveSheetIndex(0)->setCellValue('C29', "2)");
		$excel->setActiveSheetIndex(0)->setCellValue('D29', "Anggota panitia Seminar dan Rapat Tahunan Bidang Ilmu MIPA (Semirata BKS PTN Barat)");
		$excel->setActiveSheetIndex(0)->setCellValue('L29', "Sem Ganjil ");
		$excel->setActiveSheetIndex(0)->setCellValue('M29', "1 Semester");
		$excel->setActiveSheetIndex(0)->setCellValue('N29', "1,0");
		$excel->setActiveSheetIndex(0)->setCellValue('O29', "1,0");
		$excel->setActiveSheetIndex(0)->setCellValue('P29', "1,0");
		$excel->setActiveSheetIndex(0)->setCellValue('Q29', "1");
		$excel->setActiveSheetIndex(0)->setCellValue('R29', "SK Dekan FMIPA Unila");
		$excel->setActiveSheetIndex(0)->setCellValue('L30', "2013/2014");
		$excel->setActiveSheetIndex(0)->setCellValue('R30', "4517/UN26/7/DT/2013");
		$excel->setActiveSheetIndex(0)->setCellValue('R31', "Tanggal 20 Maret 2013");

		$excel->setActiveSheetIndex(0)->setCellValue('C33', "3)");
		$excel->setActiveSheetIndex(0)->setCellValue('D33', "Anggota Panitia Tim Penjamin Mutu Program Sarjana (S1) FMIPA Unila ");
		$excel->setActiveSheetIndex(0)->setCellValue('L33', "Sem Ganjil ");
		$excel->setActiveSheetIndex(0)->setCellValue('M33', "1 Semester");
		$excel->setActiveSheetIndex(0)->setCellValue('N33', "1,0");
		$excel->setActiveSheetIndex(0)->setCellValue('O33', "1,0");
		$excel->setActiveSheetIndex(0)->setCellValue('P33', "1,0");
		$excel->setActiveSheetIndex(0)->setCellValue('Q33', "1");
		$excel->setActiveSheetIndex(0)->setCellValue('R33', "SK Dekan FMIPA Unila");
		$excel->setActiveSheetIndex(0)->setCellValue('L34', "2013/2014");
		$excel->setActiveSheetIndex(0)->setCellValue('R34', "No 1462a/UN26/7/DT/2013");
		$excel->setActiveSheetIndex(0)->setCellValue('R35', "Tanggal 1 Oktober 2013");

		$excel->setActiveSheetIndex(0)->setCellValue('C37', "4)");
		$excel->setActiveSheetIndex(0)->setCellValue('D37', "Anggota Tim Verifikasi Rekayasa Perangkat Lunak Tugas Akhir D3 MI Semester Ganjil ");
		$excel->setActiveSheetIndex(0)->setCellValue('L37', "Sem Ganjil ");
		$excel->setActiveSheetIndex(0)->setCellValue('M37', "1 Semester");
		$excel->setActiveSheetIndex(0)->setCellValue('N37', "1,0");
		$excel->setActiveSheetIndex(0)->setCellValue('O37', "1,0");
		$excel->setActiveSheetIndex(0)->setCellValue('P37', "1,0");
		$excel->setActiveSheetIndex(0)->setCellValue('Q37', "1");
		$excel->setActiveSheetIndex(0)->setCellValue('R37', "SK Dekan FMIPA Unila");
		$excel->setActiveSheetIndex(0)->setCellValue('L38', "2013/2014");
		$excel->setActiveSheetIndex(0)->setCellValue('R38', "No 2604/UN26/7/DT/2013");
		$excel->setActiveSheetIndex(0)->setCellValue('R39', "Tanggal 7 Oktober 2013");

		$excel->setActiveSheetIndex(0)->setCellValue('C41', "5)");
		$excel->setActiveSheetIndex(0)->setCellValue('D41', "Tim Penilai Sertifikasi Dosen FMIPA Unila");
		$excel->setActiveSheetIndex(0)->setCellValue('L41', "Sem Ganjil ");
		$excel->setActiveSheetIndex(0)->setCellValue('M41', "1 Semester");
		$excel->setActiveSheetIndex(0)->setCellValue('N41', "1,0");
		$excel->setActiveSheetIndex(0)->setCellValue('O41', "1,0");
		$excel->setActiveSheetIndex(0)->setCellValue('P41', "1,0");
		$excel->setActiveSheetIndex(0)->setCellValue('Q41', "1");
		$excel->setActiveSheetIndex(0)->setCellValue('R41', "SK Dekan FMIPA Unila");
		$excel->setActiveSheetIndex(0)->setCellValue('L42', "2013/2014");
		$excel->setActiveSheetIndex(0)->setCellValue('R42', "No 2679/UN26/7/DT/2013");
		$excel->setActiveSheetIndex(0)->setCellValue('R43', "Tanggal 18 Oktober 2013");

		$excel->setActiveSheetIndex(0)->setCellValue('C45', "6)");
		$excel->setActiveSheetIndex(0)->setCellValue('D45', "Anggota Panitia Seminar Nasional Sain dan Teknologi (SATEK) V ");
		$excel->setActiveSheetIndex(0)->setCellValue('L45', "Sem Ganjil ");
		$excel->setActiveSheetIndex(0)->setCellValue('M45', "1 Semester");
		$excel->setActiveSheetIndex(0)->setCellValue('N45', "1,0");
		$excel->setActiveSheetIndex(0)->setCellValue('O45', "1,0");
		$excel->setActiveSheetIndex(0)->setCellValue('P45', "1,0");
		$excel->setActiveSheetIndex(0)->setCellValue('Q45', "1");
		$excel->setActiveSheetIndex(0)->setCellValue('R45', "SK Rektor Unila");
		$excel->setActiveSheetIndex(0)->setCellValue('L46', "2013/2014");
		$excel->setActiveSheetIndex(0)->setCellValue('R46', "No 767/UN26/LP/2013");
		$excel->setActiveSheetIndex(0)->setCellValue('R47', "Tanggal November 2013");

		$excel->setActiveSheetIndex(0)->setCellValue('C49', "7)");
		$excel->setActiveSheetIndex(0)->setCellValue('D49', "Anggota Tim Verifikasi Rekayasa Perangkat Lunak Tugas Akhir D3 MI Semester Genap");
		$excel->setActiveSheetIndex(0)->setCellValue('L49', "Sem Genap");
		$excel->setActiveSheetIndex(0)->setCellValue('M49', "1 Semester");
		$excel->setActiveSheetIndex(0)->setCellValue('N49', "1,0");
		$excel->setActiveSheetIndex(0)->setCellValue('O49', "1,0");
		$excel->setActiveSheetIndex(0)->setCellValue('P49', "1,0");
		$excel->setActiveSheetIndex(0)->setCellValue('Q49', "1");
		$excel->setActiveSheetIndex(0)->setCellValue('R49', "SK Dekan FMIPA Unila");
		$excel->setActiveSheetIndex(0)->setCellValue('L50', "2013/2014");
		$excel->setActiveSheetIndex(0)->setCellValue('R50', "No 135a/UN26/7/DT/2014");
		$excel->setActiveSheetIndex(0)->setCellValue('R51', "Tanggal 14 Januari 2014");

		$excel->setActiveSheetIndex(0)->setCellValue('C54', "8)");
		$excel->setActiveSheetIndex(0)->setCellValue('D54', "Tim Audit Internal ISO 9001:2008 Jurusan Ilmu Komputer FMIPA Unila");
		$excel->setActiveSheetIndex(0)->setCellValue('L54', "Sem Genap");
		$excel->setActiveSheetIndex(0)->setCellValue('M54', "1 Semester");
		$excel->setActiveSheetIndex(0)->setCellValue('N54', "1,0");
		$excel->setActiveSheetIndex(0)->setCellValue('O54', "1,0");
		$excel->setActiveSheetIndex(0)->setCellValue('P54', "1,0");
		$excel->setActiveSheetIndex(0)->setCellValue('Q54', "1");
		$excel->setActiveSheetIndex(0)->setCellValue('R54', "SK Dekan FMIPA UNILA");
		$excel->setActiveSheetIndex(0)->setCellValue('L55', "2014/2015");
		$excel->setActiveSheetIndex(0)->setCellValue('R55', "No. 476a/UN26/7/KM/2014");
		$excel->setActiveSheetIndex(0)->setCellValue('R56', "Tanggal 3 Maret 2014");

		$excel->setActiveSheetIndex(0)->setCellValue('C58', "9)");
		$excel->setActiveSheetIndex(0)->setCellValue('D58', "Tim Jaminan Mutu ISO 9001:2008 Jurusan Ilmu Komputer FMIPA Unila");
		$excel->setActiveSheetIndex(0)->setCellValue('L58', "Sem Genap");
		$excel->setActiveSheetIndex(0)->setCellValue('M58', "1 Semester");
		$excel->setActiveSheetIndex(0)->setCellValue('N58', "1,0");
		$excel->setActiveSheetIndex(0)->setCellValue('O58', "1,0");
		$excel->setActiveSheetIndex(0)->setCellValue('P58', "1,0");
		$excel->setActiveSheetIndex(0)->setCellValue('Q58', "1");
		$excel->setActiveSheetIndex(0)->setCellValue('R58', "SK Dekan FMIPA UNILA");
		$excel->setActiveSheetIndex(0)->setCellValue('L59', "2014/2015");
		$excel->setActiveSheetIndex(0)->setCellValue('R59', "No. 483a/UN26/7/KM/2014");
		$excel->setActiveSheetIndex(0)->setCellValue('R60', "Tanggal 4 Maret 2014");

		$excel->setActiveSheetIndex(0)->setCellValue('C62', "10)");
		$excel->setActiveSheetIndex(0)->setCellValue('D62', "Anggota Tim Verifikasi Rekayasa Perangkat Lunak Tugas Akhir D3 MI Semester Ganjil");
		$excel->setActiveSheetIndex(0)->setCellValue('L62', "Sem Ganjil");
		$excel->setActiveSheetIndex(0)->setCellValue('M62', "1 Semester");
		$excel->setActiveSheetIndex(0)->setCellValue('N62', "1,0");
		$excel->setActiveSheetIndex(0)->setCellValue('O62', "1,0");
		$excel->setActiveSheetIndex(0)->setCellValue('P62', "1,0");
		$excel->setActiveSheetIndex(0)->setCellValue('Q62', "1");
		$excel->setActiveSheetIndex(0)->setCellValue('R62', "SK Dekan FMIPA UNILA");
		$excel->setActiveSheetIndex(0)->setCellValue('L63', "2014/2015");
		$excel->setActiveSheetIndex(0)->setCellValue('R63', "No 2012/UN26/7/DT/2014");
		$excel->setActiveSheetIndex(0)->setCellValue('R64', "Tanggal 9 Oktober 2014");

		$excel->setActiveSheetIndex(0)->setCellValue('C67', "11)");
		$excel->setActiveSheetIndex(0)->setCellValue('D67', "Anggota tim audit pembelajaran program sarjana Jurusan Ilmu Komputer FMIPA Unila");
		$excel->setActiveSheetIndex(0)->setCellValue('L67', "Smstr Genap");
		$excel->setActiveSheetIndex(0)->setCellValue('M67', "1 Semester");
		$excel->setActiveSheetIndex(0)->setCellValue('N67', "1,0");
		$excel->setActiveSheetIndex(0)->setCellValue('O67', "1,0");
		$excel->setActiveSheetIndex(0)->setCellValue('P67', "1,0");
		$excel->setActiveSheetIndex(0)->setCellValue('Q67', "1");
		$excel->setActiveSheetIndex(0)->setCellValue('R67', "SK Dekan FMIPA UNILA");
		$excel->setActiveSheetIndex(0)->setCellValue('L68', "2014/2015");
		$excel->setActiveSheetIndex(0)->setCellValue('R68', "No. 16a/UN26/7/DT/2015");
		$excel->setActiveSheetIndex(0)->setCellValue('R69', "Tanggal 7 Januari 2015");

		$excel->setActiveSheetIndex(0)->setCellValue('C71', "12)");
		$excel->setActiveSheetIndex(0)->setCellValue('D71', "Pengurus Badan Pelaksana Kuliah Kerja Nyata (BP-KKN) Unila");
		$excel->setActiveSheetIndex(0)->setCellValue('L71', "Smstr Genap");
		$excel->setActiveSheetIndex(0)->setCellValue('M71', "1 Semester");
		$excel->setActiveSheetIndex(0)->setCellValue('N71', "1,0");
		$excel->setActiveSheetIndex(0)->setCellValue('O71', "1,0");
		$excel->setActiveSheetIndex(0)->setCellValue('P71', "1,0");
		$excel->setActiveSheetIndex(0)->setCellValue('Q71', "1");
		$excel->setActiveSheetIndex(0)->setCellValue('R71', "SK Rektor Unila");
		$excel->setActiveSheetIndex(0)->setCellValue('L72', "2014/2015");
		$excel->setActiveSheetIndex(0)->setCellValue('R72', "No 140/UN26/KP/2015");
		$excel->setActiveSheetIndex(0)->setCellValue('R73', "Tanggal 24 Maret 2015");

		$excel->setActiveSheetIndex(0)->setCellValue('C76', "13)");
		$excel->setActiveSheetIndex(0)->setCellValue('D76', "Tim Auditor ISO 9001:2008 Jurusan Ilmu Komputer FMIPA Unila");
		$excel->setActiveSheetIndex(0)->setCellValue('L76', "Smstr Genap");
		$excel->setActiveSheetIndex(0)->setCellValue('M76', "1 Semester");
		$excel->setActiveSheetIndex(0)->setCellValue('N76', "1,0");
		$excel->setActiveSheetIndex(0)->setCellValue('O76', "1,0");
		$excel->setActiveSheetIndex(0)->setCellValue('P76', "1,0");
		$excel->setActiveSheetIndex(0)->setCellValue('Q76', "1");
		$excel->setActiveSheetIndex(0)->setCellValue('R76', "SK Dekan FMIPA UNILA");
		$excel->setActiveSheetIndex(0)->setCellValue('L77', "2014/2015");
		$excel->setActiveSheetIndex(0)->setCellValue('R77', "No. 1637/UN26/7/KM/2014");
		$excel->setActiveSheetIndex(0)->setCellValue('R78', "Tanggal 4 Mei 2015");

		$excel->setActiveSheetIndex(0)->setCellValue('C80', "14)");
		$excel->setActiveSheetIndex(0)->setCellValue('D80', "Anggota Panitia Tim Penjamin Mutu Program Sarjana (S1) FMIPA Unila ");
		$excel->setActiveSheetIndex(0)->setCellValue('L80', "Smstr Genap");
		$excel->setActiveSheetIndex(0)->setCellValue('M80', "1 Semester");
		$excel->setActiveSheetIndex(0)->setCellValue('N80', "1,0");
		$excel->setActiveSheetIndex(0)->setCellValue('O80', "1,0");
		$excel->setActiveSheetIndex(0)->setCellValue('P80', "1,0");
		$excel->setActiveSheetIndex(0)->setCellValue('Q80', "1");
		$excel->setActiveSheetIndex(0)->setCellValue('R80', "SK Dekan FMIPA UNILA");
		$excel->setActiveSheetIndex(0)->setCellValue('L81', "2014/2015");
		$excel->setActiveSheetIndex(0)->setCellValue('R81', "No 1642a/UN26/7/DT/2015");
		$excel->setActiveSheetIndex(0)->setCellValue('R82', "Tanggal 5 Mei 2015");

		$excel->setActiveSheetIndex(0)->setCellValue('C84', "15)");
		$excel->setActiveSheetIndex(0)->setCellValue('D84', "Anggota tim penyusun akreditasi program sarjana Jurusan Ilmu Komputer FMIPA Unila");
		$excel->setActiveSheetIndex(0)->setCellValue('L84', "Smstr Genap");
		$excel->setActiveSheetIndex(0)->setCellValue('M84', "1 Semester");
		$excel->setActiveSheetIndex(0)->setCellValue('N84', "1,0");
		$excel->setActiveSheetIndex(0)->setCellValue('O84', "1,0");
		$excel->setActiveSheetIndex(0)->setCellValue('P84', "1,0");
		$excel->setActiveSheetIndex(0)->setCellValue('Q84', "1");
		$excel->setActiveSheetIndex(0)->setCellValue('R84', "SK Dekan FMIPA UNILA");
		$excel->setActiveSheetIndex(0)->setCellValue('L85', "2014/2015");
		$excel->setActiveSheetIndex(0)->setCellValue('R85', "No. 1941/UN26/7/DT/2015");
		$excel->setActiveSheetIndex(0)->setCellValue('R86', "Tanggal 15 Juni 2015");

		$excel->setActiveSheetIndex(0)->setCellValue('C88', "16)");
		$excel->setActiveSheetIndex(0)->setCellValue('D88', "Juri Mahasiswa Berprestasi FMIPA Unila");
		$excel->setActiveSheetIndex(0)->setCellValue('L88', "Smstr Genap");
		$excel->setActiveSheetIndex(0)->setCellValue('M88', "1 Semester");
		$excel->setActiveSheetIndex(0)->setCellValue('N88', "1,0");
		$excel->setActiveSheetIndex(0)->setCellValue('O88', "1,0");
		$excel->setActiveSheetIndex(0)->setCellValue('P88', "1,0");
		$excel->setActiveSheetIndex(0)->setCellValue('Q88', "1");
		$excel->setActiveSheetIndex(0)->setCellValue('R88', "SK Dekan FMIPA UNILA");
		$excel->setActiveSheetIndex(0)->setCellValue('L89', "2015/2016");
		$excel->setActiveSheetIndex(0)->setCellValue('R89', "No. 636/UN26/7/KM/2016");
		$excel->setActiveSheetIndex(0)->setCellValue('R90', "Tanggal 21 Maret 2016");

		$excel->setActiveSheetIndex(0)->setCellValue('C92', "17)");
		$excel->setActiveSheetIndex(0)->setCellValue('D92', "Anggota Tim Verifikasi Rekayasa Perangkat Lunak Tugas Akhir D3 MI Semester Genap");
		$excel->setActiveSheetIndex(0)->setCellValue('L92', "Smstr Genap");
		$excel->setActiveSheetIndex(0)->setCellValue('M92', "1 Semester");
		$excel->setActiveSheetIndex(0)->setCellValue('N92', "1,0");
		$excel->setActiveSheetIndex(0)->setCellValue('O92', "1,0");
		$excel->setActiveSheetIndex(0)->setCellValue('P92', "1,0");
		$excel->setActiveSheetIndex(0)->setCellValue('Q92', "1");
		$excel->setActiveSheetIndex(0)->setCellValue('R92', "SK Dekan FMIPA UNILA");
		$excel->setActiveSheetIndex(0)->setCellValue('L93', "2015/2016");
		$excel->setActiveSheetIndex(0)->setCellValue('R93', "No 771/UN26/7/DT/2014");
		$excel->setActiveSheetIndex(0)->setCellValue('R94', "Tanggal 7 April 2016");

		$excel->setActiveSheetIndex(0)->setCellValue('C96', "18)");
		$excel->setActiveSheetIndex(0)->setCellValue('D96', "Anggota tim penyusun akreditasi program sarjana Jurusan Ilmu Komputer FMIPA Unila");
		$excel->setActiveSheetIndex(0)->setCellValue('L96', "Smstr Genap");
		$excel->setActiveSheetIndex(0)->setCellValue('M96', "1 Semester");
		$excel->setActiveSheetIndex(0)->setCellValue('N96', "1,0");
		$excel->setActiveSheetIndex(0)->setCellValue('O96', "1,0");
		$excel->setActiveSheetIndex(0)->setCellValue('P96', "1,0");
		$excel->setActiveSheetIndex(0)->setCellValue('Q96', "1");
		$excel->setActiveSheetIndex(0)->setCellValue('R96', "SK Dekan FMIPA UNILA");
		$excel->setActiveSheetIndex(0)->setCellValue('L97', "2015/2016");
		$excel->setActiveSheetIndex(0)->setCellValue('R97', "No. 814/UN26/7/DT/2016");
		$excel->setActiveSheetIndex(0)->setCellValue('R98', "Tanggal 13 April 2016");

		$excel->setActiveSheetIndex(0)->setCellValue('C100', "19)");
		$excel->setActiveSheetIndex(0)->setCellValue('D100', "Anggota Tim Pengelola Lokakarya Revisi Kurikulum PS S1 Ilmu Komputer FMIPA Unila");
		$excel->setActiveSheetIndex(0)->setCellValue('L100', "Smstr Genap");
		$excel->setActiveSheetIndex(0)->setCellValue('M100', "1 Semester");
		$excel->setActiveSheetIndex(0)->setCellValue('N100', "1,0");
		$excel->setActiveSheetIndex(0)->setCellValue('O100', "1,0");
		$excel->setActiveSheetIndex(0)->setCellValue('P100', "1,0");
		$excel->setActiveSheetIndex(0)->setCellValue('Q100', "1");
		$excel->setActiveSheetIndex(0)->setCellValue('R100', "SK Dekan FMIPA UNILA");
		$excel->setActiveSheetIndex(0)->setCellValue('L101', "2015/2016");
		$excel->setActiveSheetIndex(0)->setCellValue('R101', "No 961/UN26/7/DT/2016");
		$excel->setActiveSheetIndex(0)->setCellValue('R102', "Tanggal 26 April 2016");

		$excel->setActiveSheetIndex(0)->setCellValue('C104', "20)");
		$excel->setActiveSheetIndex(0)->setCellValue('D104', "Anggota Tim Pengelola Lokakarya Revisi Kurikulum PS D3 Manajemen Informatika FMIPA Unila");
		$excel->setActiveSheetIndex(0)->setCellValue('L104', "Smstr Genap");
		$excel->setActiveSheetIndex(0)->setCellValue('M104', "1 Semester");
		$excel->setActiveSheetIndex(0)->setCellValue('N104', "1,0");
		$excel->setActiveSheetIndex(0)->setCellValue('O104', "1,0");
		$excel->setActiveSheetIndex(0)->setCellValue('P104', "1,0");
		$excel->setActiveSheetIndex(0)->setCellValue('Q104', "1");
		$excel->setActiveSheetIndex(0)->setCellValue('R104', "SK Dekan FMIPA UNILA");
		$excel->setActiveSheetIndex(0)->setCellValue('L105', "2015/2016");
		$excel->setActiveSheetIndex(0)->setCellValue('R105', "No 963/UN26/7/DT/2016");
		$excel->setActiveSheetIndex(0)->setCellValue('R106', "Tanggal 28 April 2016");

		$excel->setActiveSheetIndex(0)->setCellValue('C108', "21)");
		$excel->setActiveSheetIndex(0)->setCellValue('D108', "Tim Audit Internal ISO 9001:2008 Jurusan Ilmu Komputer FMIPA Unila");
		$excel->setActiveSheetIndex(0)->setCellValue('L108', "Smstr Ganjil");
		$excel->setActiveSheetIndex(0)->setCellValue('M108', "1 Semester");
		$excel->setActiveSheetIndex(0)->setCellValue('N108', "1,0");
		$excel->setActiveSheetIndex(0)->setCellValue('O108', "1,0");
		$excel->setActiveSheetIndex(0)->setCellValue('P108', "1,0");
		$excel->setActiveSheetIndex(0)->setCellValue('Q108', "1");
		$excel->setActiveSheetIndex(0)->setCellValue('R108', "SK Dekan FMIPA UNILA");
		$excel->setActiveSheetIndex(0)->setCellValue('L109', "2016/2017");
		$excel->setActiveSheetIndex(0)->setCellValue('R109', "No. 2597/UN26/7/DT/2016");
		$excel->setActiveSheetIndex(0)->setCellValue('R110', "Tanggal 24 Oktober 2016");

		$excel->setActiveSheetIndex(0)->setCellValue('C112', "22)");
		$excel->setActiveSheetIndex(0)->setCellValue('D112', "Pengurus Badan Pelaksana Kuliah Kerja Nyata (BP-KKN) Unila");
		$excel->setActiveSheetIndex(0)->setCellValue('L112', "Smstr Genap");
		$excel->setActiveSheetIndex(0)->setCellValue('M112', "1 Semester");
		$excel->setActiveSheetIndex(0)->setCellValue('N112', "1,0");
		$excel->setActiveSheetIndex(0)->setCellValue('O112', "1,0");
		$excel->setActiveSheetIndex(0)->setCellValue('P112', "1,0");
		$excel->setActiveSheetIndex(0)->setCellValue('Q112', "1");
		$excel->setActiveSheetIndex(0)->setCellValue('R112', "SK Rektor Unila");
		$excel->setActiveSheetIndex(0)->setCellValue('L113', "2016/2017");
		$excel->setActiveSheetIndex(0)->setCellValue('R113', "No. 129/UN26/KP/2017");
		$excel->setActiveSheetIndex(0)->setCellValue('R114', "Tanggal 01 Februari 2017");

		$excel->setActiveSheetIndex(0)->setCellValue('C117', "23)");
		$excel->setActiveSheetIndex(0)->setCellValue('D117', "Anggota Panitia Seleksi Dosen Kontrak FMIPA Unila");
		$excel->setActiveSheetIndex(0)->setCellValue('L117', "Smstr Ganjil");
		$excel->setActiveSheetIndex(0)->setCellValue('M117', "1 Semester");
		$excel->setActiveSheetIndex(0)->setCellValue('N117', "1,0");
		$excel->setActiveSheetIndex(0)->setCellValue('O117', "1,0");
		$excel->setActiveSheetIndex(0)->setCellValue('P117', "1,0");
		$excel->setActiveSheetIndex(0)->setCellValue('Q117', "1");
		$excel->setActiveSheetIndex(0)->setCellValue('R117', "SK Dekan FMIPA Unila");
		$excel->setActiveSheetIndex(0)->setCellValue('L118', "2017/2018");
		$excel->setActiveSheetIndex(0)->setCellValue('R118', "No 2723/UN26/7/KP/2017");
		$excel->setActiveSheetIndex(0)->setCellValue('R119', "Tanggal 15 Juni 2017");

		$excel->setActiveSheetIndex(0)->setCellValue('C121', "24)");
		$excel->setActiveSheetIndex(0)->setCellValue('D121', "Anggota Tim Pelaksana Sie Koreksi Tugas Mahasiswa Kegiatan Program Pengenalan Kehidupan Kampus bagi Mahasiswa baru Unila");
		$excel->setActiveSheetIndex(0)->setCellValue('L121', "Smstr Genap");
		$excel->setActiveSheetIndex(0)->setCellValue('M121', "1 Semester");
		$excel->setActiveSheetIndex(0)->setCellValue('N121', "1,0");
		$excel->setActiveSheetIndex(0)->setCellValue('O121', "1,0");
		$excel->setActiveSheetIndex(0)->setCellValue('P121', "1,0");
		$excel->setActiveSheetIndex(0)->setCellValue('Q121', "1");
		$excel->setActiveSheetIndex(0)->setCellValue('R121', "SK Rektor Unila");
		$excel->setActiveSheetIndex(0)->setCellValue('L122', "2017/2018");
		$excel->setActiveSheetIndex(0)->setCellValue('R122', "No 952/UN26/DT/2017");
		$excel->setActiveSheetIndex(0)->setCellValue('R123', "Tanggal 07 Agustus 2017");

		$excel->setActiveSheetIndex(0)->setCellValue('C125', "25)");
		$excel->setActiveSheetIndex(0)->setCellValue('D125', "Anggota Tim Verifikasi Rekayasa Perangkat Lunak Tugas Akhir D3 MI Semester Ganjil");
		$excel->setActiveSheetIndex(0)->setCellValue('L125', "Smstr Ganjil");
		$excel->setActiveSheetIndex(0)->setCellValue('M125', "1 Semester");
		$excel->setActiveSheetIndex(0)->setCellValue('N125', "1,0");
		$excel->setActiveSheetIndex(0)->setCellValue('O125', "1,0");
		$excel->setActiveSheetIndex(0)->setCellValue('P125', "1,0");
		$excel->setActiveSheetIndex(0)->setCellValue('Q125', "1");
		$excel->setActiveSheetIndex(0)->setCellValue('R125', "SK Dekan FMIPA Unila");
		$excel->setActiveSheetIndex(0)->setCellValue('L126', "2017/2018");
		$excel->setActiveSheetIndex(0)->setCellValue('R126', "No 4216/UN26/7/DT/2017");
		$excel->setActiveSheetIndex(0)->setCellValue('R127', "Tanggal 27 Oktober 2017");

		$excel->setActiveSheetIndex(0)->setCellValue('C130', "26)");
		$excel->setActiveSheetIndex(0)->setCellValue('D130', "Anggota IT Kuliah Kerja Nyata Kebangsaan ");
		$excel->setActiveSheetIndex(0)->setCellValue('L130', "Smstr Genap");
		$excel->setActiveSheetIndex(0)->setCellValue('M130', "1 Semester");
		$excel->setActiveSheetIndex(0)->setCellValue('N130', "1,0");
		$excel->setActiveSheetIndex(0)->setCellValue('O130', "1,0");
		$excel->setActiveSheetIndex(0)->setCellValue('P130', "1,0");
		$excel->setActiveSheetIndex(0)->setCellValue('Q130', "1");
		$excel->setActiveSheetIndex(0)->setCellValue('R130', "SK Rektor Unila");
		$excel->setActiveSheetIndex(0)->setCellValue('L131', "2017/2018");
		$excel->setActiveSheetIndex(0)->setCellValue('R131', "No 1247/UN26/PM.03/2018");
		$excel->setActiveSheetIndex(0)->setCellValue('R132', "Tanggal 08 Juni 2018");

		$excel->setActiveSheetIndex(0)->setCellValue('C134', "27)");
		$excel->setActiveSheetIndex(0)->setCellValue('D134', "Kepala Divisi pada UPT Bahasa Unila");
		$excel->setActiveSheetIndex(0)->setCellValue('L134', "Smstr Ganjil");
		$excel->setActiveSheetIndex(0)->setCellValue('M134', "1 Semester");
		$excel->setActiveSheetIndex(0)->setCellValue('N134', "1,0");
		$excel->setActiveSheetIndex(0)->setCellValue('O134', "1,0");
		$excel->setActiveSheetIndex(0)->setCellValue('P134', "1,0");
		$excel->setActiveSheetIndex(0)->setCellValue('Q134', "1");
		$excel->setActiveSheetIndex(0)->setCellValue('R134', "SK Rektor Unila");
		$excel->setActiveSheetIndex(0)->setCellValue('L135', "2018/2019");
		$excel->setActiveSheetIndex(0)->setCellValue('R135', "No 1903/UN26/KP/2018");
		$excel->setActiveSheetIndex(0)->setCellValue('R136', "Tanggal 26 Oktober 2018");

		$excel->setActiveSheetIndex(0)->setCellValue('O138', "Jumlah");
		$excel->setActiveSheetIndex(0)->setCellValue('P138', "27,00");
		$excel->getActiveSheet()->getStyle('O138:P138')->getFont()->setBold(TRUE);
		$excel->getActiveSheet()->getStyle('O138:P138')->getFont()->setSize(11);
		$excel->getActiveSheet()->getStyle('A138:R138')->applyFromArray($style_standar);

		$excel->setActiveSheetIndex(0)->setCellValue('A139', "B");
		$excel->setActiveSheetIndex(0)->setCellValue('B139', "Menjadi anggota panitia/badan pada lembaga pemerintah");
		$excel->setActiveSheetIndex(0)->setCellValue('B140', "1.");
		$excel->setActiveSheetIndex(0)->setCellValue('C140', "Panitia pusat");
		$excel->setActiveSheetIndex(0)->setCellValue('C141', "a.");
		$excel->setActiveSheetIndex(0)->setCellValue('D141', "Ketua/Wakil Ketua");
		$excel->setActiveSheetIndex(0)->setCellValue('C142', "b.");
		$excel->setActiveSheetIndex(0)->setCellValue('D142', "Anggota");
		$excel->setActiveSheetIndex(0)->setCellValue('B143', "2.");
		$excel->setActiveSheetIndex(0)->setCellValue('C143', "Panitia daerah");
		$excel->setActiveSheetIndex(0)->setCellValue('C144', "a.");
		$excel->setActiveSheetIndex(0)->setCellValue('D144', "Ketua/Wakil Ketua");
		$excel->setActiveSheetIndex(0)->setCellValue('C145', "b.");
		$excel->setActiveSheetIndex(0)->setCellValue('D145', "Anggota");
		$excel->getActiveSheet()->getStyle('A145:R145')->applyFromArray($style_standar);

		$excel->setActiveSheetIndex(0)->setCellValue('A146', "C");
		$excel->setActiveSheetIndex(0)->setCellValue('B146', "Menjadi anggota organisasi profesi");
		$excel->setActiveSheetIndex(0)->setCellValue('B147', "1.");
		$excel->setActiveSheetIndex(0)->setCellValue('C147', "Tingkat Internasional");
		$excel->getActiveSheet()->mergeCells('C147:F147');
		$excel->setActiveSheetIndex(0)->setCellValue('C148', "a");
		$excel->setActiveSheetIndex(0)->setCellValue('D148', "Pengurus");
		$excel->setActiveSheetIndex(0)->setCellValue('C149', "b");
		$excel->setActiveSheetIndex(0)->setCellValue('D149', "Anggota atas permintaan");
		$excel->setActiveSheetIndex(0)->setCellValue('C150', "c");
		$excel->setActiveSheetIndex(0)->setCellValue('D150', "Anggota");
		$excel->setActiveSheetIndex(0)->setCellValue('B151', "2.");
		$excel->setActiveSheetIndex(0)->setCellValue('C151', "Tingkat Nasional");
		$excel->getActiveSheet()->mergeCells('C151:F151');
		$excel->setActiveSheetIndex(0)->setCellValue('C152', "a");
		$excel->setActiveSheetIndex(0)->setCellValue('D152', "Pengurus");
		$excel->setActiveSheetIndex(0)->setCellValue('C153', "b");
		$excel->setActiveSheetIndex(0)->setCellValue('D153', "Anggota atas permintaan");
		$excel->setActiveSheetIndex(0)->setCellValue('C154', "c");
		$excel->setActiveSheetIndex(0)->setCellValue('D154', "Anggota");
		$excel->setActiveSheetIndex(0)->setCellValue('Q154', "VI.C2.c");
		$excel->setActiveSheetIndex(0)->setCellValue('D155', "Anggota Ikatan Ahli Informatika");
		$excel->setActiveSheetIndex(0)->setCellValue('L555', "Periode");
		$excel->setActiveSheetIndex(0)->setCellValue('M155', "Setiap");
		$excel->setActiveSheetIndex(0)->setCellValue('N155', "1");
		$excel->setActiveSheetIndex(0)->setCellValue('O155', "0,50");
		$excel->setActiveSheetIndex(0)->setCellValue('P155', "0,50");
		$excel->setActiveSheetIndex(0)->setCellValue('R155', "Kartu Anggota IAII");
		$excel->setActiveSheetIndex(0)->setCellValue('D156', "Indonesia");
		$excel->setActiveSheetIndex(0)->setCellValue('L156', "2016-2018");
		$excel->setActiveSheetIndex(0)->setCellValue('M156', "Periode");
		$excel->setActiveSheetIndex(0)->setCellValue('R156', "No. 16.10.10002");
		$excel->setActiveSheetIndex(0)->setCellValue('O157', "Jumlah");
		$excel->setActiveSheetIndex(0)->setCellValue('P157', "0,50");
		$excel->getActiveSheet()->getStyle('O157:P157')->getFont()->setBold(TRUE);
		$excel->getActiveSheet()->getStyle('O157:P157')->getFont()->setSize(11);
		$excel->getActiveSheet()->getStyle('A157:R157')->applyFromArray($style_standar);

		$excel->setActiveSheetIndex(0)->setCellValue('A158', "D");
		$excel->setActiveSheetIndex(0)->setCellValue('B158', "Mewakili perguruan tinggi/lembaga pemerintah");
		$excel->getActiveSheet()->mergeCells('B158:K158');
		$excel->setActiveSheetIndex(0)->setCellValue('C159', "Mewakili perguruan tinggi/lembaga pemerintah duduk dalam panitia antar lembaga");
		$excel->getActiveSheet()->mergeCells('C159:K159');
		$excel->getActiveSheet()->getStyle('A159:R159')->applyFromArray($style_standar);

		$excel->setActiveSheetIndex(0)->setCellValue('A160', "E");
		$excel->setActiveSheetIndex(0)->setCellValue('B160', "Menjadi anggota delegasi nasional ke pertemuan internasional");
		$excel->setActiveSheetIndex(0)->setCellValue('B161', "1.");
		$excel->setActiveSheetIndex(0)->setCellValue('C161', "Sebagai ketua delegasi");
		$excel->getActiveSheet()->mergeCells('C161:K161');
		$excel->setActiveSheetIndex(0)->setCellValue('B162', "2.");
		$excel->setActiveSheetIndex(0)->setCellValue('C162', "Sebagai anggota delegasi");
		$excel->getActiveSheet()->mergeCells('C162:K162');
		$excel->getActiveSheet()->getStyle('A162:R162')->applyFromArray($style_standar);

		$excel->setActiveSheetIndex(0)->setCellValue('A163', "F");
		$excel->setActiveSheetIndex(0)->setCellValue('B163', "Berperan serta aktif dalam pertemuan ilmiah");
		$excel->getActiveSheet()->mergeCells('B163:K163');
		$excel->setActiveSheetIndex(0)->setCellValue('B164', "1.");
		$excel->setActiveSheetIndex(0)->setCellValue('C164', "Tingkat internasional/nasional/regional sebagai :");
		$excel->getActiveSheet()->mergeCells('C164:K164');
		$excel->setActiveSheetIndex(0)->setCellValue('C165', "a.");
		$excel->setActiveSheetIndex(0)->setCellValue('D165', "Ketua");
		$excel->getActiveSheet()->mergeCells('D165:F165');
		$excel->getActiveSheet()->getStyle('A166:R166')->applyFromArray($style_standar);
		$excel->setActiveSheetIndex(0)->setCellValue('C167', "b.");
		$excel->setActiveSheetIndex(0)->setCellValue('D167', "Anggota");
		$excel->getActiveSheet()->mergeCells('D167:F167');
		$excel->setActiveSheetIndex(0)->setCellValue('Q167', "VI.F1.b");
		$excel->getActiveSheet()->getStyle('A167:R167')->applyFromArray($style_standar);
		$excel->getActiveSheet()->getStyle('A168:R168')->applyFromArray($style_standar);
		$excel->setActiveSheetIndex(0)->setCellValue('B169', "2.");
		$excel->setActiveSheetIndex(0)->setCellValue('C169', "Di lingkungan perguruan tinggi sebagai :");
		$excel->setActiveSheetIndex(0)->setCellValue('C170', "a.");
		$excel->setActiveSheetIndex(0)->setCellValue('D170', "Ketua");
		$excel->getActiveSheet()->mergeCells('D170:F170');
		$excel->setActiveSheetIndex(0)->setCellValue('O171', "Jumlah");
		$excel->setActiveSheetIndex(0)->setCellValue('P171', "0,00");
		$excel->getActiveSheet()->getStyle('O171:P171')->getFont()->setBold(TRUE);
		$excel->getActiveSheet()->getStyle('O171:P171')->getFont()->setSize(11);
		$excel->getActiveSheet()->getStyle('A171:R171')->applyFromArray($style_standar);
		$excel->setActiveSheetIndex(0)->setCellValue('C172', "b.");
		$excel->setActiveSheetIndex(0)->setCellValue('D172', "Anggota");
		$excel->getActiveSheet()->mergeCells('D172:F172');
		$excel->setActiveSheetIndex(0)->setCellValue('O174', "Jumlah");
		$excel->setActiveSheetIndex(0)->setCellValue('P174', "0,00");
		$excel->getActiveSheet()->getStyle('O174:P174')->getFont()->setBold(TRUE);
		$excel->getActiveSheet()->getStyle('O174:P174')->getFont()->setSize(11);
		$excel->getActiveSheet()->getStyle('A174:R174')->applyFromArray($style_standar);
		$excel->setActiveSheetIndex(0)->setCellValue('A175', "G");
		$excel->setActiveSheetIndex(0)->setCellValue('B175', "Mendapat penghargaan/ tanda jasa");
		$excel->getActiveSheet()->mergeCells('B175:K175');
		$excel->setActiveSheetIndex(0)->setCellValue('B176', "1.");
		$excel->setActiveSheetIndex(0)->setCellValue('C176', "Penghargaan/tanda jasa Satya Lancana Karya Satya");
		$excel->getActiveSheet()->mergeCells('C176:K176');
		$excel->setActiveSheetIndex(0)->setCellValue('C177', "a");
		$excel->setActiveSheetIndex(0)->setCellValue('D177', "30 (tiga puluh) tahun");
		$excel->getActiveSheet()->mergeCells('D177:K177');
		$excel->getActiveSheet()->getStyle('A178:R178')->applyFromArray($style_standar);
		$excel->setActiveSheetIndex(0)->setCellValue('C179', "b");
		$excel->setActiveSheetIndex(0)->setCellValue('D179', "20 (dua puluh) tahun");
		$excel->getActiveSheet()->mergeCells('D179:K179');
		$excel->getActiveSheet()->getStyle('A179:R179')->applyFromArray($style_standar);
		$excel->setActiveSheetIndex(0)->setCellValue('O180', "Jumlah");
		$excel->setActiveSheetIndex(0)->setCellValue('P180', "0,00");
		$excel->getActiveSheet()->getStyle('O180:P180')->getFont()->setBold(TRUE);
		$excel->getActiveSheet()->getStyle('O180:P180')->getFont()->setSize(11);
		$excel->getActiveSheet()->getStyle('A180:R180')->applyFromArray($style_standar);
		$excel->setActiveSheetIndex(0)->setCellValue('C181', "c");
		$excel->setActiveSheetIndex(0)->setCellValue('D181', "10 (sepuluh) tahun");
		$excel->getActiveSheet()->mergeCells('D181:K181');
		$excel->setActiveSheetIndex(0)->setCellValue('B182', "2.");
		$excel->setActiveSheetIndex(0)->setCellValue('C182', "Memperoleh penghargaan lainnya");
		$excel->getActiveSheet()->mergeCells('C182:K182');
		$excel->setActiveSheetIndex(0)->setCellValue('C183', "a");
		$excel->setActiveSheetIndex(0)->setCellValue('D183', "Tingkat internasional");
		$excel->setActiveSheetIndex(0)->setCellValue('C184', "b");
		$excel->setActiveSheetIndex(0)->setCellValue('D184', "Tingkat nasional");
		$excel->setActiveSheetIndex(0)->setCellValue('C185', "c");
		$excel->setActiveSheetIndex(0)->setCellValue('D185', "Tingkat provinsi");
		$excel->setActiveSheetIndex(0)->setCellValue('O187', "Jumlah");
		$excel->setActiveSheetIndex(0)->setCellValue('P187', "0,00");
		$excel->getActiveSheet()->getStyle('O187:P187')->getFont()->setBold(TRUE);
		$excel->getActiveSheet()->getStyle('O187:P187')->getFont()->setSize(11);
		$excel->getActiveSheet()->getStyle('A187:R187')->applyFromArray($style_standar);
		$excel->setActiveSheetIndex(0)->setCellValue('A188', "H");
		$excel->setActiveSheetIndex(0)->setCellValue('B188', "Menulis buku pelajaran SLTA ke bawah yang diterbitkan dan diedarkan secara nasional");
		$excel->getActiveSheet()->mergeCells('B188:K188');
		$excel->setActiveSheetIndex(0)->setCellValue('B189', "1");
		$excel->setActiveSheetIndex(0)->setCellValue('C189', "Buku SLTA atau setingkat");
		$excel->getActiveSheet()->mergeCells('C189:F189');
		$excel->setActiveSheetIndex(0)->setCellValue('B190', "2");
		$excel->setActiveSheetIndex(0)->setCellValue('C190', "Buku SLTP atau setingkat");
		$excel->getActiveSheet()->mergeCells('C190:F190');
		$excel->setActiveSheetIndex(0)->setCellValue('B191', "3");
		$excel->setActiveSheetIndex(0)->setCellValue('C191', "Buku SD atau setingkat");
		$excel->getActiveSheet()->mergeCells('C191:F191');
		$excel->getActiveSheet()->getStyle('A191:R191')->applyFromArray($style_standar);
		$excel->setActiveSheetIndex(0)->setCellValue('A192', "I");
		$excel->setActiveSheetIndex(0)->setCellValue('B192', "Mempunyai prestasi di bidang olahraga/-humaniora");
		$excel->getActiveSheet()->mergeCells('B192:K192');
		$excel->setActiveSheetIndex(0)->setCellValue('B193', "1.");
		$excel->setActiveSheetIndex(0)->setCellValue('C193', "Tingkat internasional");
		$excel->setActiveSheetIndex(0)->setCellValue('B194', "2.");
		$excel->setActiveSheetIndex(0)->setCellValue('C194', "Tingkat nasional");
		$excel->setActiveSheetIndex(0)->setCellValue('B195', "3.");
		$excel->setActiveSheetIndex(0)->setCellValue('C195', "Tingkat daerah/lokal");
		$excel->getActiveSheet()->getStyle('A195:R195')->applyFromArray($style_standar);
		$excel->setActiveSheetIndex(0)->setCellValue('A196', "J");
		$excel->setActiveSheetIndex(0)->setCellValue('B196', "Keanggotaan dalam tim penilaian ");
		$excel->getActiveSheet()->mergeCells('B196:K196');
		$excel->setActiveSheetIndex(0)->setCellValue('C197', "Menjadi anggota tim penilaian  jabatan Akademik Dosen");
		$excel->getActiveSheet()->mergeCells('C197:K197');
		$excel->getActiveSheet()->getStyle('A197:R197')->applyFromArray($style_standar);
		$excel->setActiveSheetIndex(0)->setCellValue('O202', "Jumlah");
		$excel->setActiveSheetIndex(0)->setCellValue('P202', "0,00");
		$excel->getActiveSheet()->getStyle('O202:P202')->getFont()->setBold(TRUE);
		$excel->getActiveSheet()->getStyle('O202:P202')->getFont()->setSize(11);
		$excel->getActiveSheet()->getStyle('A202:R202')->applyFromArray($style_standar);
		$excel->setActiveSheetIndex(0)->setCellValue('L203', "Total Penunjang");
		$excel->setActiveSheetIndex(0)->setCellValue('P203', "27,50");
		$excel->getActiveSheet()->getStyle('P203')->getFont()->setBold(TRUE);
		$excel->getActiveSheet()->getStyle('P203')->getFont()->setSize(11);
		$excel->getActiveSheet()->getStyle('A203:R203')->applyFromArray($style_standar);

		$excel->setActiveSheetIndex(0)->setCellValue('O204', "Bandar Lampung,  31 Juli 2019");
		$excel->setActiveSheetIndex(0)->setCellValue('O205', "Ketua Jurusan Ilmu Komputer");
		$excel->setActiveSheetIndex(0)->setCellValue('O209', "Dr.Ir. Kurnia Muludi, M.S.Sc");
		$excel->setActiveSheetIndex(0)->setCellValue('O210', "NIP. 19640616 198902 1 001");




		// Set height semua kolom menjadi auto (mengikuti height isi dari kolommnya, jadi otomatis)
		$excel->getActiveSheet()->getDefaultRowDimension()->setRowHeight(-1);
		// Set orientasi kertas jadi LANDSCAPE
		$excel->getActiveSheet()->getPageSetup()->setOrientation(PHPExcel_Worksheet_PageSetup::ORIENTATION_LANDSCAPE);

		// Set judul file excel nya
		$excel->getActiveSheet(0)->setTitle("Penunjang");
		$excel->setActiveSheetIndex(0);

		// Proses file excel
		header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
		header('Content-Disposition: attachment; filename="Penunjang.xlsx"'); // Set nama file excel nya
		header('Cache-Control: max-age=0');

		$write = PHPExcel_IOFactory::createWriter($excel, 'Excel2007');
		$write->save('php://output');

	}

}
