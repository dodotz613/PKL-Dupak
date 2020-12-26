<?php
defined('BASEPATH') OR exit('No direct script access allowed');

class Pendidikan extends CI_Controller {

	public function __construct(){
		parent::__construct();

		$this->load->model('Dosen');
		$this->load->model('Lektor');

		if($this->session->userdata('status') != "login"){
            redirect('login');
		}

	  }

	public function index()
	{

		$data['data_dosen'] = $this->Dosen->view();
		$data['dosen_penunjang'] = $this->Lektor->view();
		$this->load->view('pendidikan', $data);
	}

	public function export(){
		// Load plugin PHPExcel nya
		include APPPATH.'third_party/PHPExcel/PHPExcel.php';

		// Panggil class PHPExcel nya
		$excel = new PHPExcel();

		// Settingan awal fil excel
		$excel->getProperties()->setCreator('ASUS')
							   ->setLastModifiedBy('ASUS')
							   ->setTitle("Pendidikan")
							   ->setSubject("Dupak")
							   ->setDescription("Laporan")
							   ->setKeywords("Pendidikan");

		$style_standar = array(

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

		$excel->setActiveSheetIndex(0)->setCellValue('A1', "SURAT PERNYATAAN");
		$excel->setActiveSheetIndex(0)->setCellValue('A2', "MELAKSANAKAN PENDIDIKAN");


		$excel->getActiveSheet()->mergeCells('A1:L1');
		$excel->getActiveSheet()->mergeCells('A2:L2');
		$excel->getActiveSheet()->getStyle('A1:A2')->getFont()->setBold(TRUE); // Set bold kolom A1
		$excel->getActiveSheet()->getStyle('A1:A2')->getFont()->setSize(15); // Set font size 15 untuk kolom A1
		$excel->getActiveSheet()->getStyle('A1:A2')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER); // Set text center untuk kolom A1


		// Buat header tabel nya pada baris ke 3

		$excel->setActiveSheetIndex(0)->setCellValue('B3', "Yang bertanda tangan di bawah ini : ");
		$excel->setActiveSheetIndex(0)->setCellValue('B5', "Nama ");
		$excel->setActiveSheetIndex(0)->setCellValue('F5', ":");
		$excel->setActiveSheetIndex(0)->setCellValue('B6', "NIP");
		$excel->setActiveSheetIndex(0)->setCellValue('F6', ":");
		$excel->setActiveSheetIndex(0)->setCellValue('B7', "Pangkat");
		$excel->setActiveSheetIndex(0)->setCellValue('F7', ":");
		$excel->setActiveSheetIndex(0)->setCellValue('B8', "Golongan");
		$excel->setActiveSheetIndex(0)->setCellValue('F8', ":");
		$excel->setActiveSheetIndex(0)->setCellValue('B9', "Jabatan");
		$excel->setActiveSheetIndex(0)->setCellValue('F9', ":");
		$excel->setActiveSheetIndex(0)->setCellValue('B10', "Unit Kerja");
		$excel->setActiveSheetIndex(0)->setCellValue('F10', ":");
		$excel->setActiveSheetIndex(0)->setCellValue('B11', "Menyatakan ");
		$excel->setActiveSheetIndex(0)->setCellValue('B12', "Nama");
		$excel->setActiveSheetIndex(0)->setCellValue('F12', ":");
		$excel->setActiveSheetIndex(0)->setCellValue('B13', "NIP");
		$excel->setActiveSheetIndex(0)->setCellValue('F13', ":");
		$excel->setActiveSheetIndex(0)->setCellValue('B14', "Pangkat");
		$excel->setActiveSheetIndex(0)->setCellValue('F14', ":");
		$excel->setActiveSheetIndex(0)->setCellValue('B15', "Golongan");
		$excel->setActiveSheetIndex(0)->setCellValue('F15', ":");
		$excel->setActiveSheetIndex(0)->setCellValue('B16', "Jabatan");
		$excel->setActiveSheetIndex(0)->setCellValue('F16', ":");
		$excel->setActiveSheetIndex(0)->setCellValue('B17', "Unit Kerja");
		$excel->setActiveSheetIndex(0)->setCellValue('F17', ":");
		$excel->setActiveSheetIndex(0)->setCellValue('B18', "Telah menyatakan pengabdian masyarakat sebagai berikut : ");

		$data_dosen = $this->Dosen->view();

		foreach($data_dosen as $data) {

			$excel->setActiveSheetIndex(0)->setCellValue('G5', $data->nama);
			$excel->setActiveSheetIndex(0)->setCellValue('G6', $data->nip);
			$excel->setActiveSheetIndex(0)->setCellValue('G7', $data->pangkat);
			$excel->setActiveSheetIndex(0)->setCellValue('G8', $data->golongan);
			$excel->setActiveSheetIndex(0)->setCellValue('G9', $data->jabatan);
			$excel->setActiveSheetIndex(0)->setCellValue('G10', $data->unit_kerja);
		}

		$dosen_penunjang = $this->Lektor->view();

		foreach($dosen_penunjang as $data) {

			$excel->setActiveSheetIndex(0)->setCellValue('G12', $data->nama);
			$excel->setActiveSheetIndex(0)->setCellValue('G13', $data->nip);
			$excel->setActiveSheetIndex(0)->setCellValue('G14', $data->pangkat);
			$excel->setActiveSheetIndex(0)->setCellValue('G15', $data->golongan);
			$excel->setActiveSheetIndex(0)->setCellValue('G16', $data->jabatan);
			$excel->setActiveSheetIndex(0)->setCellValue('G17', $data->unit_kerja);
		}

		// Set width kolom
		$excel->getActiveSheet()->getColumnDimension('A')->setWidth(5); // Set width kolom A
		$excel->getActiveSheet()->getColumnDimension('B')->setWidth(13); // Set width kolom B
		$excel->getActiveSheet()->getColumnDimension('C')->setWidth(25); // Set width kolom C
		$excel->getActiveSheet()->getColumnDimension('D')->setWidth(18); // Set width kolom D
		$excel->getActiveSheet()->getColumnDimension('E')->setWidth(15); // Set width kolom E
		$excel->getActiveSheet()->getColumnDimension('F')->setWidth(15);
		$excel->getActiveSheet()->getColumnDimension('G')->setWidth(15);
		$excel->getActiveSheet()->getColumnDimension('H')->setWidth(15);
		$excel->getActiveSheet()->getColumnDimension('I')->setWidth(15);


		$excel->setActiveSheetIndex(0)->setCellValue('A19', "No");
		$excel->getActiveSheet()->mergeCells('B19:D19');
		$excel->setActiveSheetIndex(0)->setCellValue('B19', "Uraian Kegiatan");
		$excel->setActiveSheetIndex(0)->setCellValue('E19', "Tanggal");
		$excel->getActiveSheet()->getStyle('E19')->getAlignment()->setWrapText(TRUE);
		$excel->setActiveSheetIndex(0)->setCellValue('F19', "Satuan Hasil");
		$excel->getActiveSheet()->getStyle('F19')->getAlignment()->setWrapText(TRUE);
		$excel->setActiveSheetIndex(0)->setCellValue('G19', "Jumlah Volume Kegiatan");
		$excel->getActiveSheet()->getStyle('G19')->getAlignment()->setWrapText(TRUE);
		$excel->setActiveSheetIndex(0)->setCellValue('H19', "Angka Kredit");
		$excel->getActiveSheet()->getStyle('H19')->getAlignment()->setWrapText(TRUE);
		$excel->setActiveSheetIndex(0)->setCellValue('I19', "Jumlah Angka Kredit");
		$excel->getActiveSheet()->getStyle('I19')->getAlignment()->setWrapText(TRUE);
		$excel->getActiveSheet()->mergeCells('J19:L19');
		$excel->setActiveSheetIndex(0)->setCellValue('J19', "Keterangan/Bukti Fisik");
		$excel->getActiveSheet()->getStyle('A19')->applyFromArray($style_col);
		$excel->getActiveSheet()->getStyle('B19:D19')->applyFromArray($style_col);
		$excel->getActiveSheet()->getStyle('E19')->applyFromArray($style_col);
		$excel->getActiveSheet()->getStyle('F19')->applyFromArray($style_col);
		$excel->getActiveSheet()->getStyle('G19')->applyFromArray($style_col);
		$excel->getActiveSheet()->getStyle('H19')->applyFromArray($style_col);
		$excel->getActiveSheet()->getStyle('I19')->applyFromArray($style_col);
		$excel->getActiveSheet()->getStyle('J19:L19')->applyFromArray($style_col);//1

		$excel->setActiveSheetIndex(0)->setCellValue('A20', "(1)");
		$excel->getActiveSheet()->mergeCells('B20:D20');
		$excel->setActiveSheetIndex(0)->setCellValue('B20', "(2)");
		$excel->setActiveSheetIndex(0)->setCellValue('E20', "(3)");
		$excel->setActiveSheetIndex(0)->setCellValue('F20', "(4)");
		$excel->setActiveSheetIndex(0)->setCellValue('G20', "(5)");
		$excel->setActiveSheetIndex(0)->setCellValue('H20', "(6)");
		$excel->setActiveSheetIndex(0)->setCellValue('I20', "(7)");
		$excel->getActiveSheet()->mergeCells('J20:L20');
		$excel->setActiveSheetIndex(0)->setCellValue('J20', "(8)");
		$excel->getActiveSheet()->getStyle('A20')->applyFromArray($style_col);
		$excel->getActiveSheet()->getStyle('B20:D20')->applyFromArray($style_col);
		$excel->getActiveSheet()->getStyle('E20')->applyFromArray($style_col);
		$excel->getActiveSheet()->getStyle('F20')->applyFromArray($style_col);
		$excel->getActiveSheet()->getStyle('G20')->applyFromArray($style_col);
		$excel->getActiveSheet()->getStyle('H20')->applyFromArray($style_col);
		$excel->getActiveSheet()->getStyle('I20')->applyFromArray($style_col);
		$excel->getActiveSheet()->getStyle('J20:L20')->applyFromArray($style_col);//2

		$excel->setActiveSheetIndex(0)->setCellValue('A21', "I.");
		$excel->getActiveSheet()->getStyle('A21')->getFont()->setBold(TRUE); // Set bold kolom
		$excel->getActiveSheet()->mergeCells('B21:D21');
		$excel->setActiveSheetIndex(0)->setCellValue('B21', "Unsur Pendidikan");
		$excel->getActiveSheet()->getStyle('B21')->getFont()->setBold(TRUE); // Set bold kolom
		$excel->getActiveSheet()->mergeCells('J21:L21');
		$excel->getActiveSheet()->getStyle('A21')->applyFromArray($style_col);
		$excel->getActiveSheet()->getStyle('B21:D21')->applyFromArray($style_standar);
		$excel->getActiveSheet()->getStyle('E21')->applyFromArray($style_col);
		$excel->getActiveSheet()->getStyle('F21')->applyFromArray($style_col);
		$excel->getActiveSheet()->getStyle('G21')->applyFromArray($style_col);
		$excel->getActiveSheet()->getStyle('H21')->applyFromArray($style_col);
		$excel->getActiveSheet()->getStyle('I21')->applyFromArray($style_col);
		$excel->getActiveSheet()->getStyle('J21:L21')->applyFromArray($style_col);//3

		$excel->setActiveSheetIndex(0)->setCellValue('A22', "A");
		$excel->getActiveSheet()->mergeCells('B22:D22');
		$excel->setActiveSheetIndex(0)->setCellValue('B22', "Pendidikan");
		$excel->getActiveSheet()->getStyle('B22')->getFont()->setBold(TRUE); // Set bold kolom
		$excel->getActiveSheet()->mergeCells('J22:L22');
		$excel->getActiveSheet()->getStyle('A22:A24')->applyFromArray($style_col);
		$excel->getActiveSheet()->getStyle('B22:D22')->applyFromArray($style_standar);
		$excel->getActiveSheet()->getStyle('E22')->applyFromArray($style_col);
		$excel->getActiveSheet()->getStyle('F22')->applyFromArray($style_col);
		$excel->getActiveSheet()->getStyle('G22')->applyFromArray($style_col);
		$excel->getActiveSheet()->getStyle('H22')->applyFromArray($style_col);
		$excel->getActiveSheet()->getStyle('I22')->applyFromArray($style_col);
		$excel->getActiveSheet()->getStyle('J22:L22')->applyFromArray($style_col);//4

		$excel->setActiveSheetIndex(0)->setCellValue('B23', "'1. Mengikuti pendidikan formal dan memperoleh gelar/sebutan/ijazah:");
		$excel->getActiveSheet()->getStyle('B23')->getFont()->setBold(TRUE); // Set bold kolom
		$excel->getActiveSheet()->getStyle('B23:L23')->applyFromArray($style_standar);//5



		$excel->getActiveSheet()->mergeCells('B24:D24');
		$excel->setActiveSheetIndex(0)->setCellValue('J24', "I.A.1.a");

		$excel->getActiveSheet()->mergeCells('J24:L24');
		$excel->getActiveSheet()->getStyle('B24:D24')->applyFromArray($style_standar);
		$excel->getActiveSheet()->getStyle('E24')->applyFromArray($style_col);
		$excel->getActiveSheet()->getStyle('F24')->applyFromArray($style_col);
		$excel->getActiveSheet()->getStyle('G24')->applyFromArray($style_col);
		$excel->getActiveSheet()->getStyle('H24')->applyFromArray($style_col);
		$excel->getActiveSheet()->getStyle('I24')->applyFromArray($style_col);
		$excel->getActiveSheet()->getStyle('J24:L24')->applyFromArray($style_standar);//6

		$excel->setActiveSheetIndex(0)->setCellValue('H25', "Jumlah");
		$excel->getActiveSheet()->getStyle('H25')->getFont()->setBold(TRUE); // Set bold kolom
		$excel->getActiveSheet()->mergeCells('B25:D25');
		$excel->setActiveSheetIndex(0)->setCellValue('I25', "0.00");
		$excel->getActiveSheet()->getStyle('I25')->getFont()->setBold(TRUE); // Set bold kolom
		$excel->getActiveSheet()->mergeCells('J25:L25');
		$excel->getActiveSheet()->getStyle('A25')->applyFromArray($style_col);
		$excel->getActiveSheet()->getStyle('B25:D25')->applyFromArray($style_col);
		$excel->getActiveSheet()->getStyle('E25')->applyFromArray($style_col);
		$excel->getActiveSheet()->getStyle('F25')->applyFromArray($style_col);
		$excel->getActiveSheet()->getStyle('G25')->applyFromArray($style_col);
		$excel->getActiveSheet()->getStyle('H25')->applyFromArray($style_standar);
		$excel->getActiveSheet()->getStyle('I25')->applyFromArray($style_standar);
		$excel->getActiveSheet()->getStyle('J25:L25')->applyFromArray($style_col);//7

		$excel->setActiveSheetIndex(0)->setCellValue('B26', "'2. Mengikuti pendidikan dan pelatihan, dan prajabatan");
		$excel->getActiveSheet()->getStyle('B26')->getFont()->setBold(TRUE); // Set bold kolom
		$excel->getActiveSheet()->getStyle('B26')->applyFromArray($style_standar);//8

		$excel->getActiveSheet()->mergeCells('B26:D26');
		$excel->setActiveSheetIndex(0)->setCellValue('J26', "I.A.2.");

		$excel->getActiveSheet()->mergeCells('J26:L26');
		$excel->getActiveSheet()->getStyle('B26:D26')->applyFromArray($style_standar);
		$excel->getActiveSheet()->getStyle('E26')->applyFromArray($style_col);
		$excel->getActiveSheet()->getStyle('F26')->applyFromArray($style_col);
		$excel->getActiveSheet()->getStyle('G26')->applyFromArray($style_col);
		$excel->getActiveSheet()->getStyle('H26')->applyFromArray($style_col);
		$excel->getActiveSheet()->getStyle('I26')->applyFromArray($style_col);
		$excel->getActiveSheet()->getStyle('J26:L26')->applyFromArray($style_standar);//9

		$excel->setActiveSheetIndex(0)->setCellValue('H27', "Jumlah");
		$excel->getActiveSheet()->getStyle('H27')->getFont()->setBold(TRUE); // Set bold kolom
		$excel->getActiveSheet()->mergeCells('B27:D27');
		$excel->setActiveSheetIndex(0)->setCellValue('I27', "0.00");
		$excel->getActiveSheet()->getStyle('I27')->getFont()->setBold(TRUE); // Set bold kolom
		$excel->getActiveSheet()->mergeCells('J27:L27');
		$excel->getActiveSheet()->getStyle('A27')->applyFromArray($style_col);
		$excel->getActiveSheet()->getStyle('B27:D27')->applyFromArray($style_col);
		$excel->getActiveSheet()->getStyle('E27')->applyFromArray($style_col);
		$excel->getActiveSheet()->getStyle('F27')->applyFromArray($style_col);
		$excel->getActiveSheet()->getStyle('G27')->applyFromArray($style_col);
		$excel->getActiveSheet()->getStyle('H27')->applyFromArray($style_standar);
		$excel->getActiveSheet()->getStyle('I27')->applyFromArray($style_standar);
		$excel->getActiveSheet()->getStyle('J27:L27')->applyFromArray($style_col);//10

		$excel->getActiveSheet()->getStyle('B28:L253')->applyFromArray($style_standar);

		$excel->setActiveSheetIndex(0)->setCellValue('B28', "A");
		$excel->setActiveSheetIndex(0)->setCellValue('C28', "Pendidikan Formal");
		$excel->setActiveSheetIndex(0)->setCellValue('C29', "1. Doktor (S3)");
		$excel->setActiveSheetIndex(0)->setCellValue('C30', "2. Magister (S2)");
		$excel->setActiveSheetIndex(0)->setCellValue('B31', "B");
		$excel->setActiveSheetIndex(0)->setCellValue('C31', "Pendidikan dan pelatihan Prajabatan Golongan  III");//11

		$excel->setActiveSheetIndex(0)->setCellValue('A33', "II.");
		$excel->getActiveSheet()->getStyle('A33')->getFont()->setBold(TRUE); // Set bold kolom
		$excel->getActiveSheet()->mergeCells('B33:D33');
		$excel->setActiveSheetIndex(0)->setCellValue('B33', "UNSUR PELAKSANAAN PENDIDIKAN");
		$excel->getActiveSheet()->getStyle('B33')->getFont()->setBold(TRUE); // Set bold kolom
		$excel->getActiveSheet()->mergeCells('J33:L33');
		$excel->getActiveSheet()->getStyle('A33')->applyFromArray($style_col);
		$excel->getActiveSheet()->getStyle('B33:D33')->applyFromArray($style_standar);
		$excel->getActiveSheet()->getStyle('E33')->applyFromArray($style_col);
		$excel->getActiveSheet()->getStyle('F33')->applyFromArray($style_col);
		$excel->getActiveSheet()->getStyle('G33')->applyFromArray($style_col);
		$excel->getActiveSheet()->getStyle('H33')->applyFromArray($style_col);
		$excel->getActiveSheet()->getStyle('I33')->applyFromArray($style_col);
		$excel->getActiveSheet()->getStyle('J33:L33')->applyFromArray($style_col);//12

		$excel->setActiveSheetIndex(0)->setCellValue('A34', "A");
		$excel->setActiveSheetIndex(0)->setCellValue('B34', "Melaksanakan perkulihan/tutorial dan membimbing, menguji serta  menyelenggarakan pendidikan  di laboratorium, praktik keguruan bengkel/studio/Kebun percobaan/");
		$excel->getActiveSheet()->getStyle('B34')->getFont()->setBold(TRUE); // Set bold kolom
		$excel->getActiveSheet()->getStyle('A34')->applyFromArray($style_col);
		$excel->getActiveSheet()->getStyle('B34:L34')->applyFromArray($style_standar);//13

		$excel->setActiveSheetIndex(0)->setCellValue('A35', "1");
		$excel->getActiveSheet()->mergeCells('B35:D35');
		$excel->setActiveSheetIndex(0)->setCellValue('B35', "Lektor");
		$excel->getActiveSheet()->getStyle('B35')->getFont()->setBold(TRUE); // Set bold kolom
		$excel->getActiveSheet()->mergeCells('J35:L35');
		//$excel->getActiveSheet()->getStyle('A25:A14')->applyFromArray($style_col);
		$excel->getActiveSheet()->getStyle('B35:D35')->applyFromArray($style_standar);
		$excel->getActiveSheet()->getStyle('E35')->applyFromArray($style_col);
		$excel->getActiveSheet()->getStyle('F35')->applyFromArray($style_col);
		$excel->getActiveSheet()->getStyle('G35')->applyFromArray($style_col);
		$excel->getActiveSheet()->getStyle('H35')->applyFromArray($style_col);
		$excel->getActiveSheet()->getStyle('I35')->applyFromArray($style_col);
		$excel->getActiveSheet()->getStyle('J35:L35')->applyFromArray($style_col);//14

		$excel->getActiveSheet()->getStyle('E36:I171')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);


		$excel->setActiveSheetIndex(0)->setCellValue('A36', "a.");
		$excel->setActiveSheetIndex(0)->setCellValue('B36', "Beban mengajar 10 sks pertama");
		$excel->setActiveSheetIndex(0)->setCellValue('J36', "II.A.1.a");
		$excel->getActiveSheet()->getStyle('J36')->getFont()->setBold(TRUE); // Set bold kolom //15

		$excel->getActiveSheet()->getStyle('E36:E177')->applyFromArray($style_standar);
		$excel->getActiveSheet()->getStyle('F36:F177')->applyFromArray($style_standar);
		$excel->getActiveSheet()->getStyle('G36:G177')->applyFromArray($style_standar);
		$excel->getActiveSheet()->getStyle('H36:H177')->applyFromArray($style_standar);
		$excel->getActiveSheet()->getStyle('I36:I177')->applyFromArray($style_standar);
		$excel->getActiveSheet()->getStyle('J36:L177')->applyFromArray($style_standar);

		$excel->setActiveSheetIndex(0)->setCellValue('B37', "Mengajar Mata Kuliah/Praktikum/Responsi:");

		$excel->setActiveSheetIndex(0)->setCellValue('A38', "(1)");
		$excel->setActiveSheetIndex(0)->setCellValue('B38', "1) Pemrograman Desktop (T) 2.0 sks (M)");
		$excel->setActiveSheetIndex(0)->setCellValue('E38', "Smstr Ganjil");
		$excel->setActiveSheetIndex(0)->setCellValue('F38', "10 sks");
		$excel->setActiveSheetIndex(0)->setCellValue('G38', "9.00");
		$excel->setActiveSheetIndex(0)->setCellValue('H38', "1.00");
		$excel->setActiveSheetIndex(0)->setCellValue('I38', "9.00");
		$excel->setActiveSheetIndex(0)->setCellValue('J38', "1. SK Dekan FMIPA Unila No. ");//17


		$excel->setActiveSheetIndex(0)->setCellValue('B39', "2) Pemrograman Desktop (P) 1.0 sks (M)");
		$excel->setActiveSheetIndex(0)->setCellValue('E39', "2013/2014");
		$excel->setActiveSheetIndex(0)->setCellValue('F39', "Pertama");
		$excel->setActiveSheetIndex(0)->setCellValue('J39', "1504a/UN26/7/DT/2013 ");//18

		$excel->setActiveSheetIndex(0)->setCellValue('B40', "3) Pengujian Perangkat Lunak(T) 2.0 sks (M)");
		$excel->setActiveSheetIndex(0)->setCellValue('J40', "tanggal 06 September2013 ");//19

		$excel->setActiveSheetIndex(0)->setCellValue('B41', "4) Pengujian Perangkat Lunak(P) 1.0 sks (M)");//20
		$excel->setActiveSheetIndex(0)->setCellValue('B42', "5) Rekayasa Perangkat Lunak (T) 2.0 sks (M)");//21
		$excel->setActiveSheetIndex(0)->setCellValue('B43', "6) Rekayasa Perangkat Lunak (P) 1.0 sks (M)");//22
		$excel->setActiveSheetIndex(0)->setCellValue('B44', "Total :    9.0 sks  : 1  =      9.0 sks");//23

		$excel->setActiveSheetIndex(0)->setCellValue('A46', "b.");
		$excel->setActiveSheetIndex(0)->setCellValue('B46', "Beban mengajar 2 sks berikutnya");
		$excel->setActiveSheetIndex(0)->setCellValue('J46', "II.A.1.b");//24

		$excel->setActiveSheetIndex(0)->setCellValue('A47', "(2)");
		$excel->setActiveSheetIndex(0)->setCellValue('B47', "1)                   0.0 sks (M)");//25

		$excel->setActiveSheetIndex(0)->setCellValue('B48', "Total: 0.0 sks : 1 = 0.0 sks");//26
		$excel->setActiveSheetIndex(0)->setCellValue('A50', "a.");
		$excel->setActiveSheetIndex(0)->setCellValue('B50', "Beban mengajar 10 sks pertama");
		$excel->setActiveSheetIndex(0)->setCellValue('J50', "II.A.1.a");//27

		$excel->setActiveSheetIndex(0)->setCellValue('B51', "Mengajar Mata Kuliah/Praktikum/Responsi:");//28
		$excel->setActiveSheetIndex(0)->setCellValue('A52', "(3)");
		$excel->setActiveSheetIndex(0)->setCellValue('B52', "1) Interaksi Manusia dan Komputer (T) 2.0 sks (M)");
		$excel->setActiveSheetIndex(0)->setCellValue('E52', "Smstr Ganjil");
		$excel->setActiveSheetIndex(0)->setCellValue('F52', "10 sks");
		$excel->setActiveSheetIndex(0)->setCellValue('G52', "10.00");
		$excel->setActiveSheetIndex(0)->setCellValue('H52', "1.00");
		$excel->setActiveSheetIndex(0)->setCellValue('I52', "10.00");
		$excel->setActiveSheetIndex(0)->setCellValue('J52', "1. SK Dekan FMIPA Unila No. ");//29

		$excel->setActiveSheetIndex(0)->setCellValue('B53', "2) Interaksi Manusia dan Komputer (P) 1.0 sks (M)");
		$excel->setActiveSheetIndex(0)->setCellValue('E53', "2013/2014");
		$excel->setActiveSheetIndex(0)->setCellValue('F53', "Pertama");
		$excel->setActiveSheetIndex(0)->setCellValue('J53', "141/UN26/7/DT/2014 ");//30

		$excel->setActiveSheetIndex(0)->setCellValue('B54', "3) Rekayasa Perangkat Lunak(T) 2.0 sks (M)");
		$excel->setActiveSheetIndex(0)->setCellValue('J54', "tanggal 06 Januari 2014 ");//31

		$excel->setActiveSheetIndex(0)->setCellValue('B55', "4) Rekayasa Perangkat Lunak(P) 1.0 sks (M)");
		$excel->setActiveSheetIndex(0)->setCellValue('J55', "2. SK Dekan FMIPA Unila No. ");//32

		$excel->setActiveSheetIndex(0)->setCellValue('B56', "5) Interaksi Manusia dan Komputer (T) 2.0 (M)");
		$excel->setActiveSheetIndex(0)->setCellValue('J56', "748/UN26/7/DT/2014");//33

		$excel->setActiveSheetIndex(0)->setCellValue('B57', "6) Interaksi Manusia dan Komputer (T) 2.0 (M)");
		$excel->setActiveSheetIndex(0)->setCellValue('J57', "tanggal 12 Maret 2014");//34

		$excel->setActiveSheetIndex(0)->setCellValue('B58', "Sub Total (M)  10.0  sks : 1  =  10.0 sks");
		$excel->setActiveSheetIndex(0)->setCellValue('B59', "Total :                       =  10.0 sks");//35

		$excel->setActiveSheetIndex(0)->setCellValue('A61', "b.");
		$excel->setActiveSheetIndex(0)->setCellValue('B61', "Beban mengajar 2 sks berikutnya");
		$excel->setActiveSheetIndex(0)->setCellValue('J61', "II.A.1.b");//36

		$excel->setActiveSheetIndex(0)->setCellValue('A62', "(4)");
		$excel->setActiveSheetIndex(0)->setCellValue('B62', "1) Interaksi Manusia dan Komputer (P) 1.0 sks (M)");
		$excel->setActiveSheetIndex(0)->setCellValue('E62', "Smstr Genap");
		$excel->setActiveSheetIndex(0)->setCellValue('F62', "2 sks");
		$excel->setActiveSheetIndex(0)->setCellValue('G62', "1.00");
		$excel->setActiveSheetIndex(0)->setCellValue('H62', "0.50");
		$excel->setActiveSheetIndex(0)->setCellValue('I62', "0.50");
		$excel->setActiveSheetIndex(0)->setCellValue('J62', "1. SK Dekan FMIPA Unila No. ");//37

		$excel->setActiveSheetIndex(0)->setCellValue('B63', "Total (M)  1.0  sks : 1  =  1.0 sks");
		$excel->setActiveSheetIndex(0)->setCellValue('E63', "2013/2014");
		$excel->setActiveSheetIndex(0)->setCellValue('F63', "Berikutnya");
		$excel->setActiveSheetIndex(0)->setCellValue('J63', "748/UN26/7/DT/2014");//38
		$excel->setActiveSheetIndex(0)->setCellValue('J64', "tanggal 12 Maret 2014");//39

		$excel->setActiveSheetIndex(0)->setCellValue('A66', "a.");
		$excel->setActiveSheetIndex(0)->setCellValue('B66', "Beban mengajar 10 sks pertama");
		$excel->setActiveSheetIndex(0)->setCellValue('J66', "II.A.1.a");//40

		$excel->setActiveSheetIndex(0)->setCellValue('A67', "(5)");
		$excel->setActiveSheetIndex(0)->setCellValue('B67', "1) Basis Data (T) 2.0 sks (M)");
		$excel->setActiveSheetIndex(0)->setCellValue('E67', "Smstr Ganjil");
		$excel->setActiveSheetIndex(0)->setCellValue('F67', "10 sks");
		$excel->setActiveSheetIndex(0)->setCellValue('G67', "10.00");
		$excel->setActiveSheetIndex(0)->setCellValue('H67', "1.00");
		$excel->setActiveSheetIndex(0)->setCellValue('I67', "10.00");
		$excel->setActiveSheetIndex(0)->setCellValue('J67', "1. SK Dekan FMIPA Unila No. ");//41

		$excel->setActiveSheetIndex(0)->setCellValue('B68', "2) Basis Data (P) 1.0 sks (M)");
		$excel->setActiveSheetIndex(0)->setCellValue('E68', "2014/2015");
		$excel->setActiveSheetIndex(0)->setCellValue('F68', "Pertama");
		$excel->setActiveSheetIndex(0)->setCellValue('J68', "2013/UN26/7/DT/2014");//42

		$excel->setActiveSheetIndex(0)->setCellValue('B69', "3) Pemrograman Desktop (T) 2.0 sks (M)");
		$excel->setActiveSheetIndex(0)->setCellValue('J69', "tanggal 09 Oktober 2014");//43

		$excel->setActiveSheetIndex(0)->setCellValue('B70', "4) Rekayasa Perangkat Lunak II (T) 2.0 sks (M)");
		$excel->setActiveSheetIndex(0)->setCellValue('J70', "2. SK Dekan FMIPA Unila No. ");//44

		$excel->setActiveSheetIndex(0)->setCellValue('B71', "5) Pengujian Perangkat Lunak 3.0 (M)");
		$excel->setActiveSheetIndex(0)->setCellValue('J71', "2363/UN26/7/DT/2014");//45
		$excel->setActiveSheetIndex(0)->setCellValue('B72', "Total :                       =  10.0 sks");
		$excel->setActiveSheetIndex(0)->setCellValue('J72', "tanggal 28 November 2014");//46

		$excel->setActiveSheetIndex(0)->setCellValue('A74', "b.");
		$excel->setActiveSheetIndex(0)->setCellValue('B74', "Beban mengajar 2 sks berikutnya");
		$excel->setActiveSheetIndex(0)->setCellValue('J74', "II.A.1.b");//47

		$excel->setActiveSheetIndex(0)->setCellValue('A75', "(6)");
		$excel->setActiveSheetIndex(0)->setCellValue('B75', "1) Pemrograman Desktop (P) 1.0 sks (M)");
		$excel->setActiveSheetIndex(0)->setCellValue('E75', "Smstr Genap");
		$excel->setActiveSheetIndex(0)->setCellValue('F75', "2 sks");
		$excel->setActiveSheetIndex(0)->setCellValue('G75', "2.00");
		$excel->setActiveSheetIndex(0)->setCellValue('H75', "0.50");
		$excel->setActiveSheetIndex(0)->setCellValue('I75', "1.00");
		$excel->setActiveSheetIndex(0)->setCellValue('J75', "1. SK Dekan FMIPA Unila No. ");//48

		$excel->setActiveSheetIndex(0)->setCellValue('B76', "2) Pemrosesan Bahasa Alami 1.5 sks (T/2)");
		$excel->setActiveSheetIndex(0)->setCellValue('E76', "2014/2015");
		$excel->setActiveSheetIndex(0)->setCellValue('F76', "Berikutnya");
		$excel->setActiveSheetIndex(0)->setCellValue('J76', "2363/UN26/7/DT/2014");//49
		$excel->setActiveSheetIndex(0)->setCellValue('B77', "Total :       2.0 sks : 1     =  2.0 sks");
		$excel->setActiveSheetIndex(0)->setCellValue('J77', "tanggal 28 November 2014");//50

		$excel->setActiveSheetIndex(0)->setCellValue('A79', "a.");
		$excel->setActiveSheetIndex(0)->setCellValue('B79', "Beban mengajar 10 sks pertama");
		$excel->setActiveSheetIndex(0)->setCellValue('J79', "II.A.1.a");//51

		$excel->setActiveSheetIndex(0)->setCellValue('B80', "Mengajar Mata Kuliah/Praktikum/Responsi:");//52
		$excel->setActiveSheetIndex(0)->setCellValue('A81', "(7)");
		$excel->setActiveSheetIndex(0)->setCellValue('B81', "1) Interaksi Manusia dan Komputer (T) 2.0 sks (M)");
		$excel->setActiveSheetIndex(0)->setCellValue('E81', "Smstr Ganjil");
		$excel->setActiveSheetIndex(0)->setCellValue('F81', "10 sks");
		$excel->setActiveSheetIndex(0)->setCellValue('G81', "10.00");
		$excel->setActiveSheetIndex(0)->setCellValue('H81', "1.00");
		$excel->setActiveSheetIndex(0)->setCellValue('I81', "10.00");
		$excel->setActiveSheetIndex(0)->setCellValue('J81', "1. SK Dekan FMIPA Unila No. ");//53

		$excel->setActiveSheetIndex(0)->setCellValue('B82', "2) Interaksi Manusia dan Komputer (P) 1.0 sks (M)");
		$excel->setActiveSheetIndex(0)->setCellValue('E82', "2014/2015");
		$excel->setActiveSheetIndex(0)->setCellValue('F82', "Pertama");
		$excel->setActiveSheetIndex(0)->setCellValue('J82', "969/UN26/7/DT/2015 ");//54

		$excel->setActiveSheetIndex(0)->setCellValue('B83', "3) Pemrograman Client Server (T) 2.0 sks (M)");
		$excel->setActiveSheetIndex(0)->setCellValue('J83', "tanggal  23 Maret 2015 ");//55

		$excel->setActiveSheetIndex(0)->setCellValue('B84', "4) Pemrograman Client Server (P) 1.0 sks (M)");
		$excel->setActiveSheetIndex(0)->setCellValue('J84', "2. SK Dekan FMIPA Unila No. ");//56

		$excel->setActiveSheetIndex(0)->setCellValue('B85', "5) Interaksi Manusia dan Komputer A 2.0 sks (M)");
		$excel->setActiveSheetIndex(0)->setCellValue('J85', "1726/UN26/7/DT/2015");//57

		$excel->setActiveSheetIndex(0)->setCellValue('B86', "6) Interaksi Manusia dan Komputer B 2.0 sks (M)");
		$excel->setActiveSheetIndex(0)->setCellValue('J86', "tanggal  15 Mei 2015");//58

		$excel->setActiveSheetIndex(0)->setCellValue('B87', "Sub Total (M)  10.0  sks : 1  =  10.0 sks");//59
		$excel->setActiveSheetIndex(0)->setCellValue('B88', "Total :                       =  10.0 sks");//60

		$excel->setActiveSheetIndex(0)->setCellValue('A90', "b.");
		$excel->setActiveSheetIndex(0)->setCellValue('B90', "Beban mengajar 2 sks berikutnya");
		$excel->setActiveSheetIndex(0)->setCellValue('J90', "II.A.1.b");//61

		$excel->setActiveSheetIndex(0)->setCellValue('A91', "(8)");
		$excel->setActiveSheetIndex(0)->setCellValue('B91', "1) Pengujian Perangkat Lunak (T) 2.0 sks (M)");
		$excel->setActiveSheetIndex(0)->setCellValue('E91', "Smstr Genap");
		$excel->setActiveSheetIndex(0)->setCellValue('F91', "2 sks");
		$excel->setActiveSheetIndex(0)->setCellValue('G91', "2.00");
		$excel->setActiveSheetIndex(0)->setCellValue('H91', "0.50");
		$excel->setActiveSheetIndex(0)->setCellValue('I91', "1.00");
		$excel->setActiveSheetIndex(0)->setCellValue('J91', "1. SK Dekan FMIPA Unila No. ");//62

		$excel->setActiveSheetIndex(0)->setCellValue('B92', "2) Pengujian Perangkat Lunak (P) 1.0 sks (M)");
		$excel->setActiveSheetIndex(0)->setCellValue('E92', "2014/2015");
		$excel->setActiveSheetIndex(0)->setCellValue('F92', "Berikutnya");
		$excel->setActiveSheetIndex(0)->setCellValue('J92', "1726/UN26/7/DT/2015");//63
		$excel->setActiveSheetIndex(0)->setCellValue('B93', "Total :       2.0 sks : 1     =  2.0 sks");
		$excel->setActiveSheetIndex(0)->setCellValue('J93', "tanggal  15 Mei 2015");//64

		$excel->setActiveSheetIndex(0)->setCellValue('A95', "a.");
		$excel->setActiveSheetIndex(0)->setCellValue('B95', "Beban mengajar 10 sks pertama");
		$excel->setActiveSheetIndex(0)->setCellValue('J95', "II.A.1.a");//65

		$excel->setActiveSheetIndex(0)->setCellValue('A96', "(9)");
		$excel->setActiveSheetIndex(0)->setCellValue('B96', "1) Basis Data A (T) 2.0 sks (M)");
		$excel->setActiveSheetIndex(0)->setCellValue('E96', "Smstr Ganjil");
		$excel->setActiveSheetIndex(0)->setCellValue('F96', "10 sks");
		$excel->setActiveSheetIndex(0)->setCellValue('G96', "10.00");
		$excel->setActiveSheetIndex(0)->setCellValue('H96', "1.00");
		$excel->setActiveSheetIndex(0)->setCellValue('I96', "10.00");
		$excel->setActiveSheetIndex(0)->setCellValue('J96', "1. SK Dekan FMIPA Unila No. ");//66

		$excel->setActiveSheetIndex(0)->setCellValue('B97', "2) Basis Data B (T) 2.0 sks (M)");
		$excel->setActiveSheetIndex(0)->setCellValue('E97', "2015/2016");
		$excel->setActiveSheetIndex(0)->setCellValue('F97', "Pertama");
		$excel->setActiveSheetIndex(0)->setCellValue('J97', "2926/UN26/7/DT/2015");//67

		$excel->setActiveSheetIndex(0)->setCellValue('B98', "3) Pemrograman Desktop A (T) 2.0 sks (M)");
		$excel->setActiveSheetIndex(0)->setCellValue('J98', "tanggal 12 Oktober 2015");//68

		$excel->setActiveSheetIndex(0)->setCellValue('B99', "4) Pemrograman Desktop B (T) 2.0 sks (M)");//69

		$excel->setActiveSheetIndex(0)->setCellValue('B100', "5) Manajemen Proyek Sistem Informasi(T) 2.0 sks (M)");//70
		$excel->setActiveSheetIndex(0)->setCellValue('B101', "Total :                       =  10.0 sks");//71


		$excel->setActiveSheetIndex(0)->setCellValue('A103', "b.");
		$excel->setActiveSheetIndex(0)->setCellValue('B103', "Beban mengajar 2 sks berikutnya");
		$excel->setActiveSheetIndex(0)->setCellValue('J103', "II.A.1.b");//72

		$excel->setActiveSheetIndex(0)->setCellValue('A104', "(10)");
		$excel->setActiveSheetIndex(0)->setCellValue('B104', "1) Basis Data A (T) 2.0 sks (M)");
		$excel->setActiveSheetIndex(0)->setCellValue('E104', "Smstr Ganjil");
		$excel->setActiveSheetIndex(0)->setCellValue('F104', "2 sks");
		$excel->setActiveSheetIndex(0)->setCellValue('G104', "2.00");
		$excel->setActiveSheetIndex(0)->setCellValue('H104', "0.5");
		$excel->setActiveSheetIndex(0)->setCellValue('I104', "1.00");
		$excel->setActiveSheetIndex(0)->setCellValue('J104', "1. SK Dekan FMIPA Unila No. ");//73

		$excel->setActiveSheetIndex(0)->setCellValue('B105', "2) Basis Data B (T) 2.0 sks (M)");
		$excel->setActiveSheetIndex(0)->setCellValue('E105', "2015/2016");
		$excel->setActiveSheetIndex(0)->setCellValue('F105', "Berikutnya");
		$excel->setActiveSheetIndex(0)->setCellValue('J105', "2911/UN26/7/DT/2015");//74

		$excel->setActiveSheetIndex(0)->setCellValue('B106', "3) Pemrograman Desktop A (T) 1.0 sks (M)");
		$excel->setActiveSheetIndex(0)->setCellValue('J106', "tanggal 12 Oktober 2015");//75

		$excel->setActiveSheetIndex(0)->setCellValue('B107', "4) Pemrograman Desktop B (T) 1.0 sks (M)");//76

		$excel->setActiveSheetIndex(0)->setCellValue('B108', "5) Manajemen Proyek Sistem Informasi(T) 1.0 sks (M)");//77
		$excel->setActiveSheetIndex(0)->setCellValue('B109', "6) Pemrosesan Bahasa Alami (T) 1.0 sks (T/2)");//78
		$excel->setActiveSheetIndex(0)->setCellValue('B110', "7) Memrosesan Bahasa Alami (P) 0.5 sks (T/2)");//79
		$excel->setActiveSheetIndex(0)->setCellValue('B111', "Total :   2.0  sks : 1    =  2.0 sks");//80

		$excel->setActiveSheetIndex(0)->setCellValue('A113', "a.");
		$excel->setActiveSheetIndex(0)->setCellValue('B113', "Beban mengajar 10 sks pertama");
		$excel->setActiveSheetIndex(0)->setCellValue('J113', "II.A.1.a");//81

		$excel->setActiveSheetIndex(0)->setCellValue('A114', "(11)");
		$excel->setActiveSheetIndex(0)->setCellValue('B114', "1) Interaksi Manusia dan Komputer (T) 2.0 sks (M)");
		$excel->setActiveSheetIndex(0)->setCellValue('E114', "Smstr Genap");
		$excel->setActiveSheetIndex(0)->setCellValue('F114', "10 sks");
		$excel->setActiveSheetIndex(0)->setCellValue('G114', "10.00");
		$excel->setActiveSheetIndex(0)->setCellValue('H114', "1.00");
		$excel->setActiveSheetIndex(0)->setCellValue('I114', "10.00");
		$excel->setActiveSheetIndex(0)->setCellValue('J114', "1. SK Dekan FMIPA Unila No. ");//82

		$excel->setActiveSheetIndex(0)->setCellValue('B115', "2) Interaksi Manusia dan Komputer (P) 1.0 sks (M)");
		$excel->setActiveSheetIndex(0)->setCellValue('E115', "2015/2016");
		$excel->setActiveSheetIndex(0)->setCellValue('F115', "Pertama");
		$excel->setActiveSheetIndex(0)->setCellValue('J115', "770/UN26/7/DT/2016");//83

		$excel->setActiveSheetIndex(0)->setCellValue('B116', "3) Proyek Khusus A (T) 1.0 sks (M)");
		$excel->setActiveSheetIndex(0)->setCellValue('J116', "tanggal  6 April 2016");//84

		$excel->setActiveSheetIndex(0)->setCellValue('B117', "4) Proyek Khusus A (R) 1.0 sks (M)");//85

		$excel->setActiveSheetIndex(0)->setCellValue('B118', "5) Proyek Khusus B (T)) 1.0 sks (M)");//86
		$excel->setActiveSheetIndex(0)->setCellValue('B119', "6) Proyek Khusus B (R)) 1.0 sks (M)");//87
		$excel->setActiveSheetIndex(0)->setCellValue('B120', "7) Pengujian Perangkat Lunak A (T) 1.0 sks (T/2)");//88
		$excel->setActiveSheetIndex(0)->setCellValue('B121', "8) Pengujian Perangkat Lunak A (P) 0.5 sks (T/2)");//89
		$excel->setActiveSheetIndex(0)->setCellValue('B122', "9) Pengujian Perangkat Lunak B (T) 1.0 sks (T/2)");//90
		$excel->setActiveSheetIndex(0)->setCellValue('B123', "10) Pengujian Perangkat Lunak B (P) 0.5 sks (T/2)");//91
		$excel->setActiveSheetIndex(0)->setCellValue('B124', "Sub Total (M) :   10.00  sks : 1    =  10.00 sks");//92
		$excel->setActiveSheetIndex(0)->setCellValue('B125', "Total:                                 10.00 sks");//93

		$excel->setActiveSheetIndex(0)->setCellValue('A127', "b.");
		$excel->setActiveSheetIndex(0)->setCellValue('B127', "Beban mengajar 2 sks berikutnya");
		$excel->setActiveSheetIndex(0)->setCellValue('J127', "II.A.1.b");//94

		$excel->setActiveSheetIndex(0)->setCellValue('A128', "(12)");
		$excel->setActiveSheetIndex(0)->setCellValue('B128', "1) Temu Kembali Informasi A (T) 1.0 sks (T/2)");
		$excel->setActiveSheetIndex(0)->setCellValue('E128', "Smstr Genap");
		$excel->setActiveSheetIndex(0)->setCellValue('F128', "2 sks");
		$excel->setActiveSheetIndex(0)->setCellValue('G128', "2.0");
		$excel->setActiveSheetIndex(0)->setCellValue('H128', "0.5");
		$excel->setActiveSheetIndex(0)->setCellValue('I128', "1.0");
		$excel->setActiveSheetIndex(0)->setCellValue('J128', "1. SK Dekan FMIPA Unila No. ");//95

		$excel->setActiveSheetIndex(0)->setCellValue('B129', "2) Temu Kembali Informasi A (P) 0.5 sks (T/2)");
		$excel->setActiveSheetIndex(0)->setCellValue('E129', "2015/2016");
		$excel->setActiveSheetIndex(0)->setCellValue('F129', "Berikutnya");
		$excel->setActiveSheetIndex(0)->setCellValue('J129', "769/UN26/7/DT/2016");//96

		$excel->setActiveSheetIndex(0)->setCellValue('B130', "3) Temu Kembali Informasi B (T) 1.0 sks (T/2)");
		$excel->setActiveSheetIndex(0)->setCellValue('J130', "tanggal  6 April 2016");//97

		$excel->setActiveSheetIndex(0)->setCellValue('B131', "4) Temu Kembali Informasi B (P) 0.5 sks (T/2)");//98

		$excel->setActiveSheetIndex(0)->setCellValue('B131', "5) Pemrograman Client Server A (T) 2.0 sks (M)");//99
		$excel->setActiveSheetIndex(0)->setCellValue('J131', "2. SK Dekan FMIPA Unila No. ");
		$excel->setActiveSheetIndex(0)->setCellValue('B132', "6) Pemrograman Client Server A (P) 1.0 sks (M)");//100
		$excel->setActiveSheetIndex(0)->setCellValue('J132', "770/UN26/7/DT/2016");
		$excel->setActiveSheetIndex(0)->setCellValue('J133', "tanggal  6 April 2016");
		$excel->setActiveSheetIndex(0)->setCellValue('B133', "7) Pemrograman Client Server B (T) 2.0 sks (M)");//101
		$excel->setActiveSheetIndex(0)->setCellValue('B134', "8) Pemrograman Client Server B (P) 1.0 sks (M)");//102
		$excel->setActiveSheetIndex(0)->setCellValue('B135', "9) Interaksi Manusia dan Komputer A (T) 2.0 sks (M)");//103
		$excel->setActiveSheetIndex(0)->setCellValue('B136', "10) Interaksi Manusia dan Komputer A (P) 1.0 sks (M)");//104
		$excel->setActiveSheetIndex(0)->setCellValue('B137', "11) Interaksi Manusia dan Komputer B (T) 2.0 sks (M)");//105
		$excel->setActiveSheetIndex(0)->setCellValue('B138', "12) Interaksi Manusia dan Komputer B (P) 1.0 sks (M)");//106
		$excel->setActiveSheetIndex(0)->setCellValue('B139', "Total :   2.0  sks : 1    =  2.0 sks");//107

		$excel->setActiveSheetIndex(0)->setCellValue('A141', "a.");
		$excel->setActiveSheetIndex(0)->setCellValue('B141', "Beban mengajar 10 sks pertama");
		$excel->setActiveSheetIndex(0)->setCellValue('J141', "II.A.1.a");//108

		$excel->setActiveSheetIndex(0)->setCellValue('A142', "(13)");
		$excel->setActiveSheetIndex(0)->setCellValue('B142', "1) Basis Data A (T) 2.0 sks (M)");
		$excel->setActiveSheetIndex(0)->setCellValue('E142', "Smstr Ganjil");
		$excel->setActiveSheetIndex(0)->setCellValue('F142', "10 sks");
		$excel->setActiveSheetIndex(0)->setCellValue('G142', "10.00");
		$excel->setActiveSheetIndex(0)->setCellValue('H142', "1.0");
		$excel->setActiveSheetIndex(0)->setCellValue('I142', "10.00");
		$excel->setActiveSheetIndex(0)->setCellValue('J142', "1. SK Dekan FMIPA Unila No. ");//109

		$excel->setActiveSheetIndex(0)->setCellValue('B143', "2) Basis Data B (T) 2.0 sks (M)");
		$excel->setActiveSheetIndex(0)->setCellValue('E143', "2016/2017");
		$excel->setActiveSheetIndex(0)->setCellValue('F143', "Pertama");
		$excel->setActiveSheetIndex(0)->setCellValue('J143', "2678/UN26/7/DT/2016");//110

		$excel->setActiveSheetIndex(0)->setCellValue('B144', "3) Pemrograman Desktop A (T) 2.0 sks (M)");
		$excel->setActiveSheetIndex(0)->setCellValue('J144', "tanggal 31 Oktober 2016");//11

		$excel->setActiveSheetIndex(0)->setCellValue('B145', "4) Pemrograman Desktop B (T) 2.0 sks (M)");//112

		$excel->setActiveSheetIndex(0)->setCellValue('B146', "5) Basis Data A (P) 1.0 sks (M)");//113
		$excel->setActiveSheetIndex(0)->setCellValue('B147', "6) Basis Data B (P) 1.0 sks (M)");//114
		$excel->setActiveSheetIndex(0)->setCellValue('B148', "Total :   10.0  sks : 1    =  10.0 sks");//115

		$excel->setActiveSheetIndex(0)->setCellValue('A150', "b.");
		$excel->setActiveSheetIndex(0)->setCellValue('B150', "Beban mengajar 2 sks berikutnya");
		$excel->setActiveSheetIndex(0)->setCellValue('J150', "II.A.1.b");//116

		$excel->setActiveSheetIndex(0)->setCellValue('A151', "(14)");
		$excel->setActiveSheetIndex(0)->setCellValue('B151', "1) Pemrograman Desktop A (P) 1.0 sks (M)");
		$excel->setActiveSheetIndex(0)->setCellValue('E151', "Smstr Ganjil");
		$excel->setActiveSheetIndex(0)->setCellValue('F151', "2 sks");
		$excel->setActiveSheetIndex(0)->setCellValue('G151', "2.0");
		$excel->setActiveSheetIndex(0)->setCellValue('H151', "0.5");
		$excel->setActiveSheetIndex(0)->setCellValue('I151', "1.0");
		$excel->setActiveSheetIndex(0)->setCellValue('J151', "1. SK Dekan FMIPA Unila No. ");//117

		$excel->setActiveSheetIndex(0)->setCellValue('B152', "2) Pemrograman Desktop A (P) 1.0 sks (M)");
		$excel->setActiveSheetIndex(0)->setCellValue('E152', "2016/2017");
		$excel->setActiveSheetIndex(0)->setCellValue('F152', "Berikutnya");
		$excel->setActiveSheetIndex(0)->setCellValue('J152', "2678/UN26/7/DT/2016");//118

		$excel->setActiveSheetIndex(0)->setCellValue('B153', "3) Pemrosesan Bahasa Alami A 3.0 sks (M)");
		$excel->setActiveSheetIndex(0)->setCellValue('J153', "tanggal 31 Oktober 2016");//119

		$excel->setActiveSheetIndex(0)->setCellValue('B154', "4) Pemrosesan Bahasa Alami A 3.0 sks (M)");//120
		$excel->setActiveSheetIndex(0)->setCellValue('J154', "2. SK Dekan FMIPA Unila No. ");
		$excel->setActiveSheetIndex(0)->setCellValue('B155', "5) Pemrosesan Bahasa Alami A 3.0 sks (M)");//121
		$excel->setActiveSheetIndex(0)->setCellValue('J155', "2677/UN26/7/DT/2016");
		$excel->setActiveSheetIndex(0)->setCellValue('B156', "6) Sistem Pakar 3.0 sks (M)");//122
		$excel->setActiveSheetIndex(0)->setCellValue('J156', "tanggal 31 Oktober 2016");
		$excel->setActiveSheetIndex(0)->setCellValue('B157', "Total :   2.0  sks : 1    =  2.0 sks");//123

		$excel->setActiveSheetIndex(0)->setCellValue('A160', "a.");
		$excel->setActiveSheetIndex(0)->setCellValue('B160', "Beban mengajar 10 sks pertama");
		$excel->setActiveSheetIndex(0)->setCellValue('J160', "II.A.1.a");//124

		$excel->setActiveSheetIndex(0)->setCellValue('B161', "Mengajar Mata Kuliah/Praktikum/Responsi:");//125
		$excel->setActiveSheetIndex(0)->setCellValue('A162', "(1)");
		$excel->setActiveSheetIndex(0)->setCellValue('B162', "1) Manajemen Resiko A (T) 1.5 sks (T/2)");
		$excel->setActiveSheetIndex(0)->setCellValue('E162', "Smstr Ganjil");
		$excel->setActiveSheetIndex(0)->setCellValue('F162', "10 sks");
		$excel->setActiveSheetIndex(0)->setCellValue('G162', "10.00");
		$excel->setActiveSheetIndex(0)->setCellValue('H162', "1.00");
		$excel->setActiveSheetIndex(0)->setCellValue('I162', "10.00");
		$excel->setActiveSheetIndex(0)->setCellValue('J162', "1. SK Dekan FMIPA Unila No. ");//126

		$excel->setActiveSheetIndex(0)->setCellValue('B163', "2) Manajemen Resiko A (T) 1.5 sks (T/2)");
		$excel->setActiveSheetIndex(0)->setCellValue('E163', "2017/2018");
		$excel->setActiveSheetIndex(0)->setCellValue('F163', "Pertama");
		$excel->setActiveSheetIndex(0)->setCellValue('J163', "4507/UN26/7/DT/2017");//127

		$excel->setActiveSheetIndex(0)->setCellValue('B164', "3) Sistem Pakar A 3.0 sks (M)");
		$excel->setActiveSheetIndex(0)->setCellValue('J164', "tanggal 14 November 2017");//128

		$excel->setActiveSheetIndex(0)->setCellValue('B165', "4) Sistem Pakar B 3.0 sks (M)");//129
		$excel->setActiveSheetIndex(0)->setCellValue('B166', "5) Kapita Selekta Sistem Informasi (P) 1.0 sks (M)");//130
		$excel->setActiveSheetIndex(0)->setCellValue('B167', "Total:            10.0  sks : 1  =  10.0 sks");//131

		$excel->setActiveSheetIndex(0)->setCellValue('A169', "b.");
		$excel->setActiveSheetIndex(0)->setCellValue('B169', "Beban mengajar 2 sks berikutnya");
		$excel->setActiveSheetIndex(0)->setCellValue('J169', "II.A.1.b");//132

		$excel->setActiveSheetIndex(0)->setCellValue('A170', "(10)");
		$excel->setActiveSheetIndex(0)->setCellValue('B170', "1) Kapita Selekta Sistem Informasi (T) 2.0 sks (M)");
		$excel->setActiveSheetIndex(0)->setCellValue('E170', "Smstr Ganjil");
		$excel->setActiveSheetIndex(0)->setCellValue('F170', "2 sks");
		$excel->setActiveSheetIndex(0)->setCellValue('G170', "2.0");
		$excel->setActiveSheetIndex(0)->setCellValue('H170', "0.5");
		$excel->setActiveSheetIndex(0)->setCellValue('I170', "00");
		$excel->setActiveSheetIndex(0)->setCellValue('J170', "1. SK Dekan FMIPA Unila No. ");//133

		$excel->setActiveSheetIndex(0)->setCellValue('B171', "2) Rekayasa Perangkat Lunak 3 sks (M)");
		$excel->setActiveSheetIndex(0)->setCellValue('E171', "2017/2018");
		$excel->setActiveSheetIndex(0)->setCellValue('F171', "berikutnya");
		$excel->setActiveSheetIndex(0)->setCellValue('J171', "4508/UN26/7/DT/2017");//134

		$excel->setActiveSheetIndex(0)->setCellValue('B171', "Total:            2.0  sks : 1  =  2.0 sks");
		$excel->setActiveSheetIndex(0)->setCellValue('J171', "tanggal 14 November 2017");//135

		$excel->setActiveSheetIndex(0)->setCellValue('A174', "2");
		$excel->setActiveSheetIndex(0)->setCellValue('B174', "Lektor/Lektor Kepala/Profesor untuk:");
		$excel->getActiveSheet()->getStyle('B174')->getFont()->setBold(TRUE);
		$excel->setActiveSheetIndex(0)->setCellValue('J174', "II.A.2");//136

		$excel->setActiveSheetIndex(0)->setCellValue('A175', "a.");
		$excel->setActiveSheetIndex(0)->setCellValue('B175', "Beban Mengajar 10 sks pertama");
		$excel->setActiveSheetIndex(0)->setCellValue('J175', "II.A.2.a");//137

		$excel->setActiveSheetIndex(0)->setCellValue('A176', "b.");
		$excel->setActiveSheetIndex(0)->setCellValue('B176', "Beban Mengajar 2 sks berikutnya");//138

		$excel->setActiveSheetIndex(0)->setCellValue('H178', "Jumlah");
		$excel->setActiveSheetIndex(0)->setCellValue('I178', "84.50");
		$excel->getActiveSheet()->getStyle('H178:I178')->getFont()->setBold(TRUE);//139
		$excel->getActiveSheet()->getStyle('A178:L178')->applyFromArray($style_standar);

		$excel->setActiveSheetIndex(0)->setCellValue('A179', "B.");
		$excel->setActiveSheetIndex(0)->setCellValue('B179', "Membimbing Seminar Mahasiswa (Setiap Mahasiswa)");
		$excel->getActiveSheet()->getStyle('A179:B179')->getFont()->setBold(TRUE);//140
		$excel->getActiveSheet()->getStyle('A179')->applyFromArray($style_standar);
		$excel->getActiveSheet()->getStyle('B179:L179')->applyFromArray($style_standar);
		$excel->getActiveSheet()->getStyle('E179:E216')->applyFromArray($style_standar);
		$excel->getActiveSheet()->getStyle('F179:F216')->applyFromArray($style_standar);
		$excel->getActiveSheet()->getStyle('G179:G216')->applyFromArray($style_standar);
		$excel->getActiveSheet()->getStyle('H179:H216')->applyFromArray($style_standar);
		$excel->getActiveSheet()->getStyle('I179:I216')->applyFromArray($style_standar);
		$excel->getActiveSheet()->getStyle('J179:L216')->applyFromArray($style_standar);

		$excel->setActiveSheetIndex(0)->setCellValue('B180', "Membimbing Seminar");//141
		$excel->setActiveSheetIndex(0)->setCellValue('J181', "II.b");//142

		$excel->getActiveSheet()->getStyle('E182:I208')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);

		$excel->setActiveSheetIndex(0)->setCellValue('B182', "1) Membimbing mahasiswa Seminar Usul Proposal");
		$excel->setActiveSheetIndex(0)->setCellValue('E182', "Smstr Genap");
		$excel->setActiveSheetIndex(0)->setCellValue('F182', "Tiap");
		$excel->setActiveSheetIndex(0)->setCellValue('G182', "1.0");
		$excel->setActiveSheetIndex(0)->setCellValue('H182', "1.0");
		$excel->setActiveSheetIndex(0)->setCellValue('I182', "1.0");
		$excel->setActiveSheetIndex(0)->setCellValue('J182', "1. SK Dekan FMIPA Unila No. ");//143

		$excel->setActiveSheetIndex(0)->setCellValue('B183', "a.n Annisa Nur Fadhilah NPM 1317051010");
		$excel->setActiveSheetIndex(0)->setCellValue('E183', "2016/2017");
		$excel->setActiveSheetIndex(0)->setCellValue('F183', "Semester");
		$excel->setActiveSheetIndex(0)->setCellValue('J183', "3164/UN26/7/DT/2017");//144
		$excel->setActiveSheetIndex(0)->setCellValue('J184', "tanggal 7 Juni 2017");//145

		$excel->setActiveSheetIndex(0)->setCellValue('B185', "2) Membimbing Mahasiswa Seminar Hasil");
		$excel->setActiveSheetIndex(0)->setCellValue('E185', "Smstr Genap");
		$excel->setActiveSheetIndex(0)->setCellValue('F185', "Tiap");
		$excel->setActiveSheetIndex(0)->setCellValue('G185', "1.0");
		$excel->setActiveSheetIndex(0)->setCellValue('H185', "1.0");
		$excel->setActiveSheetIndex(0)->setCellValue('I185', "1.0");
		$excel->setActiveSheetIndex(0)->setCellValue('J185', "1. SK Dekan FMIPA Unila No. ");//146

		$excel->setActiveSheetIndex(0)->setCellValue('B186', "a.n Irfani Maharani NPM 1317051033");
		$excel->setActiveSheetIndex(0)->setCellValue('E186', "2016/2017");
		$excel->setActiveSheetIndex(0)->setCellValue('F186', "Semester");
		$excel->setActiveSheetIndex(0)->setCellValue('J186', "3164/UN26/7/DT/2017");//147
		$excel->setActiveSheetIndex(0)->setCellValue('B187', "a.n Faiq Sulthon Dani NPM 1347051060");
		$excel->setActiveSheetIndex(0)->setCellValue('J187', "tanggal 3 Agustus 2017");//148
		$excel->setActiveSheetIndex(0)->setCellValue('B188', "a.n Rizka Esa Basri NPM 1317051058");

		$excel->setActiveSheetIndex(0)->setCellValue('B189', "a.n Navia Yufitasari NPM 1347051010");
		$excel->setActiveSheetIndex(0)->setCellValue('J189', "2. SK Dekan FMIPA Unila No. ");//149
		$excel->setActiveSheetIndex(0)->setCellValue('B190', "a.n Eria Ayu Ningtias NPM 1317051020");
		$excel->setActiveSheetIndex(0)->setCellValue('J190', "3580/UN26/7/DT/2017");//150
		$excel->setActiveSheetIndex(0)->setCellValue('B191', "a.n Jonhar Lucky Adrianus NPM 1117032034");
		$excel->setActiveSheetIndex(0)->setCellValue('J191', "tanggal 7 September 2017");//151
		$excel->setActiveSheetIndex(0)->setCellValue('B192', "a.n Galih Imam W NPM 1117032028");//152
		$excel->setActiveSheetIndex(0)->setCellValue('B193', "a.n Nurmayanti NPM 1217032049");//153

		$excel->setActiveSheetIndex(0)->setCellValue('B204', "3) Membimbing Mahasiswa Seminar Hasil");
		$excel->setActiveSheetIndex(0)->setCellValue('E204', "Smstr Ganjil");
		$excel->setActiveSheetIndex(0)->setCellValue('F204', "Tiap");
		$excel->setActiveSheetIndex(0)->setCellValue('G204', "1.0");
		$excel->setActiveSheetIndex(0)->setCellValue('H204', "1.0");
		$excel->setActiveSheetIndex(0)->setCellValue('I204', "1.0");
		$excel->setActiveSheetIndex(0)->setCellValue('J204', "1. SK Dekan FMIPA Unila No. ");//154

		$excel->setActiveSheetIndex(0)->setCellValue('B205', "a.n Mita Fuljana NPM 1317051007");
		$excel->setActiveSheetIndex(0)->setCellValue('E205', "2017/2018");
		$excel->setActiveSheetIndex(0)->setCellValue('F205', "Semester");
		$excel->setActiveSheetIndex(0)->setCellValue('J205', "4407/UN26/7/DT/2017");//155
		$excel->setActiveSheetIndex(0)->setCellValue('J206', "tanggal 7 November 2017");//156

		$excel->setActiveSheetIndex(0)->setCellValue('B207', "4) Membimbing Mahasiswa Usul Proposal");
		$excel->setActiveSheetIndex(0)->setCellValue('E207', "Smstr Ganjil");
		$excel->setActiveSheetIndex(0)->setCellValue('F207', "Tiap");
		$excel->setActiveSheetIndex(0)->setCellValue('G207', "1.0");
		$excel->setActiveSheetIndex(0)->setCellValue('H207', "1.0");
		$excel->setActiveSheetIndex(0)->setCellValue('I207', "1.0");
		$excel->setActiveSheetIndex(0)->setCellValue('J207', "1. SK Dekan FMIPA Unila No. ");//157

		$excel->setActiveSheetIndex(0)->setCellValue('B208', "a.n Ichwan Almaza NPM 1417051066");
		$excel->setActiveSheetIndex(0)->setCellValue('E208', "2017/2018");
		$excel->setActiveSheetIndex(0)->setCellValue('F208', "Semester");
		$excel->setActiveSheetIndex(0)->setCellValue('J208', "4406/UN26/7/DT/2017");//158
		$excel->setActiveSheetIndex(0)->setCellValue('B209', "a.n Wisnu Lukito NPM 1417051052");
		$excel->setActiveSheetIndex(0)->setCellValue('J209', "tanggal 7 November 2017");//159

		$excel->setActiveSheetIndex(0)->setCellValue('B210', "a.n firmansyah NPM 1417051056");//160
		$excel->setActiveSheetIndex(0)->setCellValue('J211', "2. SK Dekan FMIPA Unila No. ");//161
		$excel->setActiveSheetIndex(0)->setCellValue('B211', "a.n Deddy Pratama NPM 1417051033");
		$excel->setActiveSheetIndex(0)->setCellValue('J212', "635/UN26/7/DT/2018");//162
		$excel->setActiveSheetIndex(0)->setCellValue('B212', "a.n Yudistira Fazri NPM 1417051055");
		$excel->setActiveSheetIndex(0)->setCellValue('J213', "tanggal 5 Februari 2018");//163
		$excel->setActiveSheetIndex(0)->setCellValue('B213', "a.n Annisa Nur Fadhilah NPM 1317051010");
		$excel->setActiveSheetIndex(0)->setCellValue('B214', "a.n Ade Pamungkas NPM 1017051010");//164

		$excel->setActiveSheetIndex(0)->setCellValue('H217', "Jumlah");
		$excel->setActiveSheetIndex(0)->setCellValue('I217', "4.00");
		$excel->getActiveSheet()->getStyle('H217:I217')->getFont()->setBold(TRUE);
		$excel->getActiveSheet()->getStyle('A217:L217')->applyFromArray($style_standar);//165

		$excel->setActiveSheetIndex(0)->setCellValue('A218', "C");
		$excel->setActiveSheetIndex(0)->setCellValue('B218', "Membimbing kuliah kerja nyata, pratek kerja nyata, praktek kerja lapangan");
		$excel->getActiveSheet()->getStyle('A218:B218')->getFont()->setBold(TRUE);
		$excel->getActiveSheet()->getStyle('A218')->applyFromArray($style_standar);
		$excel->getActiveSheet()->getStyle('B218:L218')->applyFromArray($style_standar);//166
		$excel->setActiveSheetIndex(0)->setCellValue('B219', "Membimbing mahasiswa kuliah kerja nyata, pratek kerja nyata, praktek kerja lapangan");//167
		$excel->setActiveSheetIndex(0)->setCellValue('B220', "1) Membimbing PKL, 5 orang mahasiswa D3");
		$excel->setActiveSheetIndex(0)->setCellValue('E220', "Smstr Ganjil");
		$excel->setActiveSheetIndex(0)->setCellValue('F220', "Tiap");
		$excel->setActiveSheetIndex(0)->setCellValue('G220', "1.0");
		$excel->setActiveSheetIndex(0)->setCellValue('H220', "1.0");
		$excel->setActiveSheetIndex(0)->setCellValue('I220', "1.0");
		$excel->setActiveSheetIndex(0)->setCellValue('J220', "II.C");//167

		$excel->getActiveSheet()->getStyle('E220:I250')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);

		$excel->getActiveSheet()->getStyle('B220:D253')->applyFromArray($style_standar);
		$excel->getActiveSheet()->getStyle('E220:E253')->applyFromArray($style_standar);
		$excel->getActiveSheet()->getStyle('F220:F253')->applyFromArray($style_standar);
		$excel->getActiveSheet()->getStyle('G220:G253')->applyFromArray($style_standar);
		$excel->getActiveSheet()->getStyle('H220:H253')->applyFromArray($style_standar);
		$excel->getActiveSheet()->getStyle('I220:I253')->applyFromArray($style_standar);
		$excel->getActiveSheet()->getStyle('J220:L253')->applyFromArray($style_standar);

		$excel->setActiveSheetIndex(0)->setCellValue('B221', "a.n Dwi Mahlawi S NPM 1107051014");
		$excel->setActiveSheetIndex(0)->setCellValue('E221', "2013/2014");
		$excel->setActiveSheetIndex(0)->setCellValue('F221', "Semester");
		$excel->setActiveSheetIndex(0)->setCellValue('J221', "1. SK Dekan FMIPA Unila No. ");//168
		$excel->setActiveSheetIndex(0)->setCellValue('B222', "a.n M. Santri Maulana NPM 1107051029");
		$excel->setActiveSheetIndex(0)->setCellValue('J222', "2605/UN26/7/DT/2013");//169
		$excel->setActiveSheetIndex(0)->setCellValue('B223', "a.n Dicky Pratama NPM 1107032011");
		$excel->setActiveSheetIndex(0)->setCellValue('J223', "tanggal  07 Oktober 2013");//70
		$excel->setActiveSheetIndex(0)->setCellValue('B224', "a.n Zefri Rahman NPM 1107032051");//171
		$excel->setActiveSheetIndex(0)->setCellValue('B224', "a.n Roby Irwansyah NPM 1107032036");

		$excel->setActiveSheetIndex(0)->setCellValue('B226', "2) Membimbing PKL, 5 orang mahasiswa D3");
		$excel->setActiveSheetIndex(0)->setCellValue('E226', "Smstr Genap");
		$excel->setActiveSheetIndex(0)->setCellValue('F226', "Tiap");
		$excel->setActiveSheetIndex(0)->setCellValue('G226', "1.0");
		$excel->setActiveSheetIndex(0)->setCellValue('H226', "1.0");
		$excel->setActiveSheetIndex(0)->setCellValue('I226', "1.0");
		$excel->setActiveSheetIndex(0)->setCellValue('J226', "II.C");//172

		$excel->setActiveSheetIndex(0)->setCellValue('B227', "a.n M. Danu Ristanto NPM 1207051037");
		$excel->setActiveSheetIndex(0)->setCellValue('E227', "2014/2015");
		$excel->setActiveSheetIndex(0)->setCellValue('F227', "Semester");
		$excel->setActiveSheetIndex(0)->setCellValue('J227', "1. SK Dekan FMIPA Unila No. ");//173
		$excel->setActiveSheetIndex(0)->setCellValue('B228', "a.n Panji Anom Wijaya NPM 1207051030");
		$excel->setActiveSheetIndex(0)->setCellValue('J228', "891/UN26/7/DT/2015");//174
		$excel->setActiveSheetIndex(0)->setCellValue('B229', "a.n Robin NPM 1207032063");
		$excel->setActiveSheetIndex(0)->setCellValue('J229', "tanggal 20 Maret 2015");//175
		$excel->setActiveSheetIndex(0)->setCellValue('B230', "a.n Cahya Baihaqi NPM 1207051014");//176
		$excel->setActiveSheetIndex(0)->setCellValue('B231', "a.n Bayu Bagus Saputra NPM 1207051011");//177

		$excel->setActiveSheetIndex(0)->setCellValue('B233', "  Membimbing PKL, 4 orang mahasiswa D3");
		$excel->setActiveSheetIndex(0)->setCellValue('E233', "Smstr Genap");;//178

		$excel->setActiveSheetIndex(0)->setCellValue('B234', "a.n M. Septian EB NPM 1307051038");
		$excel->setActiveSheetIndex(0)->setCellValue('E234', "2015/2016");
		$excel->setActiveSheetIndex(0)->setCellValue('J234', "2. SK Dekan FMIPA Unila No. ");//179
		$excel->setActiveSheetIndex(0)->setCellValue('B235', "a.n A. Aziz Al Hakim NPM 1307051001");
		$excel->setActiveSheetIndex(0)->setCellValue('J235', "816/UN26/7/DT/2016");//180
		$excel->setActiveSheetIndex(0)->setCellValue('B236', "a.n Arifki NPM 1307051011");
		$excel->setActiveSheetIndex(0)->setCellValue('J236', "tanggal 13 April 2016");//181
		$excel->setActiveSheetIndex(0)->setCellValue('B237', "a.n Tri Hartono NPM 1307051061");//182

		$excel->setActiveSheetIndex(0)->setCellValue('B239', "3) Membimbing PKL, 8 orang mahasiswa");
		$excel->setActiveSheetIndex(0)->setCellValue('E239', "Smstr Ganjil");
		$excel->setActiveSheetIndex(0)->setCellValue('F239', "Tiap");
		$excel->setActiveSheetIndex(0)->setCellValue('G239', "1.0");
		$excel->setActiveSheetIndex(0)->setCellValue('H239', "1.0");
		$excel->setActiveSheetIndex(0)->setCellValue('I239', "1.0");
		$excel->setActiveSheetIndex(0)->setCellValue('J239', "II.C");//183

		$excel->setActiveSheetIndex(0)->setCellValue('B240', "a.n M. M. Rahman Koestanto NPM 1217051043");
		$excel->setActiveSheetIndex(0)->setCellValue('E240', "2016/2017");
		$excel->setActiveSheetIndex(0)->setCellValue('F240', "Semester");
		$excel->setActiveSheetIndex(0)->setCellValue('J240', "1. SK Dekan FMIPA Unila No. ");//184
		$excel->setActiveSheetIndex(0)->setCellValue('B241', "a.n Danzen Hangga Pratama NPM 1317051003");
		$excel->setActiveSheetIndex(0)->setCellValue('J241', "3178/UN26/7/DT/2016");//185
		$excel->setActiveSheetIndex(0)->setCellValue('B242', "a.n Agung Prasetyo NPM 1317051006");
		$excel->setActiveSheetIndex(0)->setCellValue('J242', "tanggal 14 Desember 2016");//186
		$excel->setActiveSheetIndex(0)->setCellValue('B243', "a.n Annisa Nur Fadilah NPM 1317051010");//187
		$excel->setActiveSheetIndex(0)->setCellValue('B244', "a.n Dini Khanza Al Nukhaiyah NPM 1317051019");
		$excel->setActiveSheetIndex(0)->setCellValue('B245', "a.n Tika Oktavia NPM 1317051065");//188
		$excel->setActiveSheetIndex(0)->setCellValue('B246', "a.n Vandu Riski Muwisnawangsa NPM 1317051069");//189
		$excel->setActiveSheetIndex(0)->setCellValue('B247', "a.n Yeni Nuhricha Sari NPM 1317051072");//190

		$excel->setActiveSheetIndex(0)->setCellValue('B249', "   Membimbing PKL, 4 orang mahasiswa");
		$excel->setActiveSheetIndex(0)->setCellValue('E249', "Smstr Genap");
		$excel->setActiveSheetIndex(0)->setCellValue('F249', "Tiap");
		$excel->setActiveSheetIndex(0)->setCellValue('G249', "1.0");
		$excel->setActiveSheetIndex(0)->setCellValue('H249', "1.0");
		$excel->setActiveSheetIndex(0)->setCellValue('I249', "1.0");
		$excel->setActiveSheetIndex(0)->setCellValue('J249', "II.C");//191

		$excel->setActiveSheetIndex(0)->setCellValue('B250', "a.n Andhika Rizki P NPM 1507011044");
		$excel->setActiveSheetIndex(0)->setCellValue('E250', "2017/2018");
		$excel->setActiveSheetIndex(0)->setCellValue('F250', "Semester");
		$excel->setActiveSheetIndex(0)->setCellValue('J250', "1. SK Dekan FMIPA Unila No. ");//192
		$excel->setActiveSheetIndex(0)->setCellValue('B251', "a.n Ayub Chandra N NPM 1507011002");
		$excel->setActiveSheetIndex(0)->setCellValue('J251', "695/UN26/7/DT/2018");//193
		$excel->setActiveSheetIndex(0)->setCellValue('B252', "a.n Shultan Hariza S NPM 1507011031");
		$excel->setActiveSheetIndex(0)->setCellValue('J252', "tanggal 07 Februari 2018");//194
		$excel->setActiveSheetIndex(0)->setCellValue('B253', "a.n Ega Nur Khotimah NPM 1507011016");//195

		$excel->setActiveSheetIndex(0)->setCellValue('H254', "Jumlah");
		$excel->setActiveSheetIndex(0)->setCellValue('I254', "3.00");
		$excel->getActiveSheet()->getStyle('H254:I254')->getFont()->setBold(TRUE);
		$excel->getActiveSheet()->getStyle('A254:L254')->applyFromArray($style_standar);//196
		$excel->setActiveSheetIndex(0)->setCellValue('A255', "D");
		$excel->setActiveSheetIndex(0)->setCellValue('B255', "Membimbing dan ikut membimbing dalam menghasilkan disertasi, thesis, skripsi dan laporan akhir studi");
		$excel->getActiveSheet()->getStyle('A255:B255')->getFont()->setBold(TRUE);
		$excel->getActiveSheet()->getStyle('A255')->applyFromArray($style_standar);
		$excel->getActiveSheet()->getStyle('B255:L255')->applyFromArray($style_standar);//197

		$excel->setActiveSheetIndex(0)->setCellValue('B256', "1. Pembimbing Utama");//198
		$excel->setActiveSheetIndex(0)->setCellValue('B257', "a. Disertasi");
		$excel->setActiveSheetIndex(0)->setCellValue('J257', "II.D.1.a");//199
		$excel->setActiveSheetIndex(0)->setCellValue('H258', "Jumlah");
		$excel->setActiveSheetIndex(0)->setCellValue('I258', "0.00");
		$excel->getActiveSheet()->getStyle('H258:I258')->getFont()->setBold(TRUE);
		$excel->getActiveSheet()->getStyle('A258:L258')->applyFromArray($style_standar);//200

		$excel->setActiveSheetIndex(0)->setCellValue('B259', "b. Tesis");
		$excel->setActiveSheetIndex(0)->setCellValue('J259', "II.D.1.b");//201
		$excel->setActiveSheetIndex(0)->setCellValue('H260', "Jumlah");
		$excel->setActiveSheetIndex(0)->setCellValue('I260', "0.00");
		$excel->getActiveSheet()->getStyle('H260:I260')->getFont()->setBold(TRUE);
		$excel->getActiveSheet()->getStyle('A260:L260')->applyFromArray($style_standar);//202
		$excel->setActiveSheetIndex(0)->setCellValue('B261', "c. Skripsi");//203

		$excel->getActiveSheet()->getStyle('B262:D306')->applyFromArray($style_standar);

		$excel->getActiveSheet()->getStyle('E262:E306')->applyFromArray($style_standar);
		$excel->getActiveSheet()->getStyle('F262:F306')->applyFromArray($style_standar);
		$excel->getActiveSheet()->getStyle('G262:G306')->applyFromArray($style_standar);
		$excel->getActiveSheet()->getStyle('H262:H306')->applyFromArray($style_standar);
		$excel->getActiveSheet()->getStyle('I262:I306')->applyFromArray($style_standar);
		$excel->getActiveSheet()->getStyle('J262:L306')->applyFromArray($style_standar);

		$excel->getActiveSheet()->getStyle('E263:I297')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);

		$excel->setActiveSheetIndex(0)->setCellValue('B263', "1). Pembimbing Utama Skripsi Mahasiswa");
		$excel->setActiveSheetIndex(0)->setCellValue('E263', "Smstr Ganjil");
		$excel->setActiveSheetIndex(0)->setCellValue('F263', "Lulusan");
		$excel->setActiveSheetIndex(0)->setCellValue('G263', "3.0");
		$excel->setActiveSheetIndex(0)->setCellValue('H263', "1.0");
		$excel->setActiveSheetIndex(0)->setCellValue('I263', "3.0");
		$excel->setActiveSheetIndex(0)->setCellValue('J263', "II.D.1.c");//204

		$excel->setActiveSheetIndex(0)->setCellValue('B264', "a.n Eko Dwi Wibowo NPM 0917032037");
		$excel->setActiveSheetIndex(0)->setCellValue('E264', "2013/2014");
		$excel->setActiveSheetIndex(0)->setCellValue('J264', "1. SK Dekan FMIPA Unila No. ");//205
		$excel->setActiveSheetIndex(0)->setCellValue('B265', "a.n Dwi Susanto NPM 0917032004");
		$excel->setActiveSheetIndex(0)->setCellValue('J265', "2821a/UN26/7/DT/2013");//206
		$excel->setActiveSheetIndex(0)->setCellValue('B266', "a.n Dian Andrian Ginting NPM 0917032033");
		$excel->setActiveSheetIndex(0)->setCellValue('J266', "tanggal 20 November 2013");//207

		$excel->setActiveSheetIndex(0)->setCellValue('B268', "2). Pembimbing Utama Skripsi Mahasiswa");
		$excel->setActiveSheetIndex(0)->setCellValue('E268', "Smstr Ganjil");
		$excel->setActiveSheetIndex(0)->setCellValue('F268', "Lulusan");
		$excel->setActiveSheetIndex(0)->setCellValue('G268', "1.0");
		$excel->setActiveSheetIndex(0)->setCellValue('H268', "1.0");
		$excel->setActiveSheetIndex(0)->setCellValue('I268', "1.0");
		$excel->setActiveSheetIndex(0)->setCellValue('J268', "II.D.1.c");//208

		$excel->setActiveSheetIndex(0)->setCellValue('B269', "a.n Tubagus Riki Andrian NPM 1017032045");
		$excel->setActiveSheetIndex(0)->setCellValue('E269', "2014/2015");
		$excel->setActiveSheetIndex(0)->setCellValue('J269', "1. SK Dekan FMIPA Unila No. ");//209
		$excel->setActiveSheetIndex(0)->setCellValue('J270', "64/UN26/7/DT/2015");//210
		$excel->setActiveSheetIndex(0)->setCellValue('J271', "tanggal 8 Januari 2015");//211

		$excel->setActiveSheetIndex(0)->setCellValue('B273', "3). Pembimbing Utama Skripsi Mahasiswa");
		$excel->setActiveSheetIndex(0)->setCellValue('E273', "Smstr Genap");
		$excel->setActiveSheetIndex(0)->setCellValue('F273', "Lulusan");
		$excel->setActiveSheetIndex(0)->setCellValue('G273', "2.0");
		$excel->setActiveSheetIndex(0)->setCellValue('H273', "1.0");
		$excel->setActiveSheetIndex(0)->setCellValue('I273', "2.0");
		$excel->setActiveSheetIndex(0)->setCellValue('J273', "II.D.1.c");//212

		$excel->setActiveSheetIndex(0)->setCellValue('B274', "a.n Revy Firandama NPM 1017032039");
		$excel->setActiveSheetIndex(0)->setCellValue('E274', "2014/2015");
		$excel->setActiveSheetIndex(0)->setCellValue('J274', "1. SK Dekan FMIPA Unila No. ");//213
		$excel->setActiveSheetIndex(0)->setCellValue('B275', "a.n Agatha Beny Himawan NPM 0917032023");
		$excel->setActiveSheetIndex(0)->setCellValue('J275', "2281/UN26/7/DT/2015");//214
		$excel->setActiveSheetIndex(0)->setCellValue('J276', "tanggal 8 Juli 2015");//215

		$excel->setActiveSheetIndex(0)->setCellValue('B278', "4). Pembimbing Utama Skripsi Mahasiswa");
		$excel->setActiveSheetIndex(0)->setCellValue('E278', "Smstr Genap");
		$excel->setActiveSheetIndex(0)->setCellValue('F278', "Lulusan");
		$excel->setActiveSheetIndex(0)->setCellValue('G278', "5.0");
		$excel->setActiveSheetIndex(0)->setCellValue('H278', "1.0");
		$excel->setActiveSheetIndex(0)->setCellValue('I278', "5.0");
		$excel->setActiveSheetIndex(0)->setCellValue('J278', "II.D.1.c");//216

		$excel->setActiveSheetIndex(0)->setCellValue('B279', "a.n Aidha Damayanti NPM 1117032005");
		$excel->setActiveSheetIndex(0)->setCellValue('E279', "2014/2015");
		$excel->setActiveSheetIndex(0)->setCellValue('J279', "1. SK Dekan FMIPA Unila No. ");//216
		$excel->setActiveSheetIndex(0)->setCellValue('B280', "a.n Ardye Armando Pratama NPM 1117032011");
		$excel->setActiveSheetIndex(0)->setCellValue('J280', "3533/UN26/7/DT/2015");//218
		$excel->setActiveSheetIndex(0)->setCellValue('B281', "a.n Aldona Pronika NPM 1117032006");
		$excel->setActiveSheetIndex(0)->setCellValue('J281', "tanggal 28 Desember 2015");//219
		$excel->setActiveSheetIndex(0)->setCellValue('B282', "a.n Harisa Eka Septiarani NPM 1117032030");//220
		$excel->setActiveSheetIndex(0)->setCellValue('B283', "a.n Putri Marlina Sari Ridwan NPM 1117032048");//221

		$excel->setActiveSheetIndex(0)->setCellValue('B285', "5). Pembimbing Utama Skripsi Mahasiswa");
		$excel->setActiveSheetIndex(0)->setCellValue('E285', "Smstr Genap");
		$excel->setActiveSheetIndex(0)->setCellValue('F285', "Lulusan");
		$excel->setActiveSheetIndex(0)->setCellValue('G285', "4.0");
		$excel->setActiveSheetIndex(0)->setCellValue('H285', "1.0");
		$excel->setActiveSheetIndex(0)->setCellValue('I285', "4.0");
		$excel->setActiveSheetIndex(0)->setCellValue('J285', "II.D.1.c");//222

		$excel->setActiveSheetIndex(0)->setCellValue('B286', "a.n Ardhika Praseda AP NPM 1117032009");
		$excel->setActiveSheetIndex(0)->setCellValue('E286', "2015/2016");
		$excel->setActiveSheetIndex(0)->setCellValue('J286', "1. SK Dekan FMIPA Unila No. ");//223
		$excel->setActiveSheetIndex(0)->setCellValue('B287', "a.n Rifki Wardana NPM 1117032051");
		$excel->setActiveSheetIndex(0)->setCellValue('J287', "309a/UN26/7/DT/2015");//224
		$excel->setActiveSheetIndex(0)->setCellValue('B288', "a.n Anita NPM 1217051007");
		$excel->setActiveSheetIndex(0)->setCellValue('J288', "tanggal 2 Februari 2016");//225
		$excel->setActiveSheetIndex(0)->setCellValue('B289', "a.n Puja Putri A. NPM 1217051051");//226

		$excel->setActiveSheetIndex(0)->setCellValue('B291', "6). Pembimbing Utama Skripsi Mahasiswa");
		$excel->setActiveSheetIndex(0)->setCellValue('E291', "Smstr Genap");
		$excel->setActiveSheetIndex(0)->setCellValue('F291', "Lulusan");
		$excel->setActiveSheetIndex(0)->setCellValue('G291', "7.0");
		$excel->setActiveSheetIndex(0)->setCellValue('H291', "1.0");
		$excel->setActiveSheetIndex(0)->setCellValue('I291', "7.0");
		$excel->setActiveSheetIndex(0)->setCellValue('J291', "II.D.1.c");//227

		$excel->setActiveSheetIndex(0)->setCellValue('B292', "a.n Annisa Nur Fadhilah NPM 1417051010");
		$excel->setActiveSheetIndex(0)->setCellValue('E292', "2016/2017");
		$excel->setActiveSheetIndex(0)->setCellValue('J292', "1. SK Dekan FMIPA Unila No. ");//228
		$excel->setActiveSheetIndex(0)->setCellValue('J293', "2576/UN26/7/DT/2017");//229
		$excel->setActiveSheetIndex(0)->setCellValue('J294', "tanggal 7 Juni 2018");//230

		$excel->getActiveSheet()->getStyle('B256:L375')->applyFromArray($style_standar);

		$excel->setActiveSheetIndex(0)->setCellValue('B296', "7). Pembimbing Utama Skripsi Mahasiswa");
		$excel->setActiveSheetIndex(0)->setCellValue('E296', "Smstr Ganjil");
		$excel->setActiveSheetIndex(0)->setCellValue('F296', "Lulusan");
		$excel->setActiveSheetIndex(0)->setCellValue('G296', "7.0");
		$excel->setActiveSheetIndex(0)->setCellValue('H296', "1.0");
		$excel->setActiveSheetIndex(0)->setCellValue('I296', "7.0");
		$excel->setActiveSheetIndex(0)->setCellValue('J296', "II.D.1.c");//231

		$excel->setActiveSheetIndex(0)->setCellValue('B297', "a.n Ichwan Almaza NPM 1417051066");
		$excel->setActiveSheetIndex(0)->setCellValue('E297', "2017/2018");
		$excel->setActiveSheetIndex(0)->setCellValue('J297', "1. SK Dekan FMIPA Unila No. ");//232
		$excel->setActiveSheetIndex(0)->setCellValue('B298', "a.n Wisnu Lukito NPM 1417051052");
		$excel->setActiveSheetIndex(0)->setCellValue('J298', "4409/UN26/7/DT/2018");//233
		$excel->setActiveSheetIndex(0)->setCellValue('B299', "a.n Firmansyah NPM 1417051056");
		$excel->setActiveSheetIndex(0)->setCellValue('J299', "tanggal 7 November 2018");//234
		$excel->setActiveSheetIndex(0)->setCellValue('B300', "a.n Deddy Pratama NPM 1417051033");//235
		$excel->setActiveSheetIndex(0)->setCellValue('B301', "a.n David Abror NPM 1417051032");//236
		$excel->setActiveSheetIndex(0)->setCellValue('B302', "a.n Yudistira Fazri NPM 1417051055");//237
		$excel->setActiveSheetIndex(0)->setCellValue('B303', "a.n Novianti NPM 1417051104");//238
		$excel->setActiveSheetIndex(0)->setCellValue('B304', "a.n Danis Sela Valena NPM 1417051031");//239

		$excel->setActiveSheetIndex(0)->setCellValue('H307', "Jumlah");
		$excel->setActiveSheetIndex(0)->setCellValue('I307', "46.00");
		$excel->getActiveSheet()->getStyle('H307:I307')->getFont()->setBold(TRUE);
		$excel->getActiveSheet()->getStyle('A307:L307')->applyFromArray($style_standar);//240
		$excel->setActiveSheetIndex(0)->setCellValue('B308', "d. Laporan Akhir");//241

		$excel->getActiveSheet()->getStyle('B309:D340')->applyFromArray($style_standar);

		$excel->getActiveSheet()->getStyle('E309:E340')->applyFromArray($style_standar);
		$excel->getActiveSheet()->getStyle('F309:F340')->applyFromArray($style_standar);
		$excel->getActiveSheet()->getStyle('G309:G340')->applyFromArray($style_standar);
		$excel->getActiveSheet()->getStyle('H309:H340')->applyFromArray($style_standar);
		$excel->getActiveSheet()->getStyle('I309:I340')->applyFromArray($style_standar);
		$excel->getActiveSheet()->getStyle('J309:L340')->applyFromArray($style_standar);

		$excel->getActiveSheet()->getStyle('E309:I338')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);

		$excel->setActiveSheetIndex(0)->setCellValue('B309', "1). Pembimbing Utama Tugas Akhir Mahasiswa");
		$excel->setActiveSheetIndex(0)->setCellValue('E309', "Smstr Ganjil");
		$excel->setActiveSheetIndex(0)->setCellValue('F309', "Lulusan");
		$excel->setActiveSheetIndex(0)->setCellValue('G309', "4.0");
		$excel->setActiveSheetIndex(0)->setCellValue('H309', "1.0");
		$excel->setActiveSheetIndex(0)->setCellValue('I309', "4.0");
		$excel->setActiveSheetIndex(0)->setCellValue('J309', "II.D.1.c");//242

		$excel->setActiveSheetIndex(0)->setCellValue('B310', "a.n Yoga Setiawan NPM 0807051082");
		$excel->setActiveSheetIndex(0)->setCellValue('E310', "2013/2014");
		$excel->setActiveSheetIndex(0)->setCellValue('J310', "1. SK Dekan FMIPA Unila No. ");//243
		$excel->setActiveSheetIndex(0)->setCellValue('B311', "a.n Dwi Anggraini NPM 0907051029");
		$excel->setActiveSheetIndex(0)->setCellValue('J311', "2607/UN26/7/DT/2013");//244
		$excel->setActiveSheetIndex(0)->setCellValue('B312', "a.n Freddy Sidauruk NPM 1007051024");
		$excel->setActiveSheetIndex(0)->setCellValue('J312', "tanggal 13 Oktober 2013");//245
		$excel->setActiveSheetIndex(0)->setCellValue('B313', "a.n Ardiansyah NPM 1007051010");//246
		$excel->setActiveSheetIndex(0)->setCellValue('J314', "2. SK Dekan FMIPA Unila No. ");//247
		$excel->setActiveSheetIndex(0)->setCellValue('J315', "2829/UN26/7/DT/2013 ");//248
		$excel->setActiveSheetIndex(0)->setCellValue('J316', "tanggal 21 November 2013");//249

		$excel->setActiveSheetIndex(0)->setCellValue('B318', "1). Pembimbing Utama Tugas Akhir Mahasiswa");
		$excel->setActiveSheetIndex(0)->setCellValue('E318', "Smstr Genap");
		$excel->setActiveSheetIndex(0)->setCellValue('F318', "Lulusan");
		$excel->setActiveSheetIndex(0)->setCellValue('G318', "3.0");
		$excel->setActiveSheetIndex(0)->setCellValue('H318', "1.0");
		$excel->setActiveSheetIndex(0)->setCellValue('I318', "3.0");
		$excel->setActiveSheetIndex(0)->setCellValue('J318', "II.D.1.c");//250

		$excel->setActiveSheetIndex(0)->setCellValue('B319', "a.n M. Reza P NPM 1007051038");
		$excel->setActiveSheetIndex(0)->setCellValue('E319', "2013/2014");
		$excel->setActiveSheetIndex(0)->setCellValue('J319', "1. SK Dekan FMIPA Unila No. ");//251
		$excel->setActiveSheetIndex(0)->setCellValue('B320', "a.n Senja Putri Arinda NPM 1007051057");
		$excel->setActiveSheetIndex(0)->setCellValue('J320', "880/UN26/7/DT/2014");//252
		$excel->setActiveSheetIndex(0)->setCellValue('B321', "a.n Angga Nurhadiansyah NPM 1007051008");
		$excel->setActiveSheetIndex(0)->setCellValue('J321', "tanggal 6 Mei 2014");//253

		$excel->setActiveSheetIndex(0)->setCellValue('B323', "1). Pembimbing Utama Tugas Akhir Mahasiswa");
		$excel->setActiveSheetIndex(0)->setCellValue('E323', "Smstr Genap");
		$excel->setActiveSheetIndex(0)->setCellValue('F323', "Lulusan");
		$excel->setActiveSheetIndex(0)->setCellValue('G323', "4.0");
		$excel->setActiveSheetIndex(0)->setCellValue('H323', "1.0");
		$excel->setActiveSheetIndex(0)->setCellValue('I323', "4.0");
		$excel->setActiveSheetIndex(0)->setCellValue('J323', "II.D.1.c");//254

		$excel->setActiveSheetIndex(0)->setCellValue('B324', "a.n Evi Dwi Jayanti NPM 1107051018");
		$excel->setActiveSheetIndex(0)->setCellValue('E324', "2014/2015");
		$excel->setActiveSheetIndex(0)->setCellValue('J324', "1. SK Dekan FMIPA Unila No. ");//255
		$excel->setActiveSheetIndex(0)->setCellValue('B325', "a.n Yesi Putriana NPM 1107051049");
		$excel->setActiveSheetIndex(0)->setCellValue('J325', "3326/UN26/7/DT/2015");//256
		$excel->setActiveSheetIndex(0)->setCellValue('B326', "a.n Apri Singgih P NPM 1007051009");
		$excel->setActiveSheetIndex(0)->setCellValue('J326', "tanggal 21 Agustus 2015");//267
		$excel->setActiveSheetIndex(0)->setCellValue('B327', "a.n Tri Yuliana NPM 1107051047");//258
		$excel->setActiveSheetIndex(0)->setCellValue('J328', "2. SK Dekan FMIPA Unila No. ");//259
		$excel->setActiveSheetIndex(0)->setCellValue('J329', "892/UN26/7/DT/2016");//260
		$excel->setActiveSheetIndex(0)->setCellValue('J330', "tanggal 20 Maret 2015");//261

		$excel->setActiveSheetIndex(0)->setCellValue('B332', "2). Pembimbing Utama Tugas Akhir Mahasiswa");
		$excel->setActiveSheetIndex(0)->setCellValue('E332', "Smstr Ganjil");
		$excel->setActiveSheetIndex(0)->setCellValue('F332', "Lulusan");
		$excel->setActiveSheetIndex(0)->setCellValue('G332', "2.0");
		$excel->setActiveSheetIndex(0)->setCellValue('H332', "1.0");
		$excel->setActiveSheetIndex(0)->setCellValue('I332', "2.0");
		$excel->setActiveSheetIndex(0)->setCellValue('J332', "II.D.1.c");//262

		$excel->setActiveSheetIndex(0)->setCellValue('B333', "a.n Siska Pertiwi NPM 1307051059");
		$excel->setActiveSheetIndex(0)->setCellValue('E333', "2016/2017");
		$excel->setActiveSheetIndex(0)->setCellValue('J333', "1. SK Dekan FMIPA Unila No. ");//263
		$excel->setActiveSheetIndex(0)->setCellValue('B334', "a.n Ridwan Syaifudin NPM 1207051059");
		$excel->setActiveSheetIndex(0)->setCellValue('J334', "2504/UN26/7/DT/2016");//264
		$excel->setActiveSheetIndex(0)->setCellValue('J335', "tanggal 5 Oktober 2016");//265

		$excel->setActiveSheetIndex(0)->setCellValue('B337', "3). Pembimbing Utama Tugas Akhir Mahasiswa");
		$excel->setActiveSheetIndex(0)->setCellValue('E337', "Smstr Genap");
		$excel->setActiveSheetIndex(0)->setCellValue('F337', "Lulusan");
		$excel->setActiveSheetIndex(0)->setCellValue('G337', "2.0");
		$excel->setActiveSheetIndex(0)->setCellValue('H337', "1.0");
		$excel->setActiveSheetIndex(0)->setCellValue('I337', "2.0");
		$excel->setActiveSheetIndex(0)->setCellValue('J337', "II.D.1.c");//267

		$excel->setActiveSheetIndex(0)->setCellValue('B338', "a.n Herman Jaya Saputra NPM 1407051025");
		$excel->setActiveSheetIndex(0)->setCellValue('E338', "2017/2018");
		$excel->setActiveSheetIndex(0)->setCellValue('J338', "1. SK Dekan FMIPA Unila No. ");//268
		$excel->setActiveSheetIndex(0)->setCellValue('B339', "a.n Maimunah NPM 1207051040");
		$excel->setActiveSheetIndex(0)->setCellValue('J339', "693/UN26/7/DT/2018");//269
		$excel->setActiveSheetIndex(0)->setCellValue('J340', "tanggal 7 Februari 2018");//270

		$excel->setActiveSheetIndex(0)->setCellValue('H341', "Jumlah");
		$excel->setActiveSheetIndex(0)->setCellValue('I341', "15.00");
		$excel->getActiveSheet()->getStyle('H341:I341')->getFont()->setBold(TRUE);
		$excel->getActiveSheet()->getStyle('A341:L341')->applyFromArray($style_standar);//271

		$excel->setActiveSheetIndex(0)->setCellValue('B342', "2  Pembimbing pendamping/pembantu");//272
		$excel->setActiveSheetIndex(0)->setCellValue('B343', "a. Disertasi");//273
		$excel->setActiveSheetIndex(0)->setCellValue('B344', "b. Thesis");//274
		$excel->setActiveSheetIndex(0)->setCellValue('B345', "c. Skripsi");//275

		$excel->getActiveSheet()->getStyle('B346:D365')->applyFromArray($style_standar);

		$excel->getActiveSheet()->getStyle('E346:E365')->applyFromArray($style_standar);
		$excel->getActiveSheet()->getStyle('F346:F365')->applyFromArray($style_standar);
		$excel->getActiveSheet()->getStyle('G346:G365')->applyFromArray($style_standar);
		$excel->getActiveSheet()->getStyle('H346:H365')->applyFromArray($style_standar);
		$excel->getActiveSheet()->getStyle('I346:I365')->applyFromArray($style_standar);
		$excel->getActiveSheet()->getStyle('J346:L365')->applyFromArray($style_standar);

		$excel->getActiveSheet()->getStyle('E346:I365')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);

		$excel->setActiveSheetIndex(0)->setCellValue('B346', "1). Pembimbing pendamping/pembantu Skripsi mahasiswa");
		$excel->setActiveSheetIndex(0)->setCellValue('E346', "Smstr Ganjil");
		$excel->setActiveSheetIndex(0)->setCellValue('F346', "Lulusan");
		$excel->setActiveSheetIndex(0)->setCellValue('G346', "1.0");
		$excel->setActiveSheetIndex(0)->setCellValue('H346', "1.0");
		$excel->setActiveSheetIndex(0)->setCellValue('I346', "1.0");
		$excel->setActiveSheetIndex(0)->setCellValue('J346', "II.D.1.c");//276

		$excel->setActiveSheetIndex(0)->setCellValue('B347', "a.n Valentina Ambarwati NPM 0617032105");
		$excel->setActiveSheetIndex(0)->setCellValue('E347', "2013/2014");
		$excel->setActiveSheetIndex(0)->setCellValue('J347', "1. SK Dekan FMIPA Unila No. ");//277
		$excel->setActiveSheetIndex(0)->setCellValue('J348', "2821a/UN26/7/DT/2013");//278
		$excel->setActiveSheetIndex(0)->setCellValue('J349', "tanggal 20 November 2013");//279

		$excel->setActiveSheetIndex(0)->setCellValue('B351', "1). Pembimbing pendamping/pembantu Skripsi mahasiswa");
		$excel->setActiveSheetIndex(0)->setCellValue('E351', "Smstr Genap");
		$excel->setActiveSheetIndex(0)->setCellValue('F351', "Lulusan");
		$excel->setActiveSheetIndex(0)->setCellValue('G351', "1.0");
		$excel->setActiveSheetIndex(0)->setCellValue('H351', "1.0");
		$excel->setActiveSheetIndex(0)->setCellValue('I351', "1.0");
		$excel->setActiveSheetIndex(0)->setCellValue('J351', "II.D.1.c");//280

		$excel->setActiveSheetIndex(0)->setCellValue('B352', "a.n Frank Sabas Fajar Basuki NPM 0917032020");
		$excel->setActiveSheetIndex(0)->setCellValue('E352', "2013/2014");
		$excel->setActiveSheetIndex(0)->setCellValue('J352', "1. SK Dekan FMIPA Unila No. ");//281
		$excel->setActiveSheetIndex(0)->setCellValue('J353', "879a/UN26/7/DT/2014");//282
		$excel->setActiveSheetIndex(0)->setCellValue('J354', "tanggal 7 Mei 2014");//283

		$excel->setActiveSheetIndex(0)->setCellValue('B356', "2). Pembimbing pendamping/pembantu Skripsi mahasiswa");
		$excel->setActiveSheetIndex(0)->setCellValue('E356', "Smstr Genap");
		$excel->setActiveSheetIndex(0)->setCellValue('F356', "Lulusan");
		$excel->setActiveSheetIndex(0)->setCellValue('G356', "2.0");
		$excel->setActiveSheetIndex(0)->setCellValue('H356', "1.0");
		$excel->setActiveSheetIndex(0)->setCellValue('I356', "2.0");
		$excel->setActiveSheetIndex(0)->setCellValue('J356', "II.D.1.c");//284

		$excel->setActiveSheetIndex(0)->setCellValue('B357', "a.n Aryanti Dwiastuti NPM 1117032013");
		$excel->setActiveSheetIndex(0)->setCellValue('E357', "2014/2015");
		$excel->setActiveSheetIndex(0)->setCellValue('J357', "1. SK Dekan FMIPA Unila No. ");//285
		$excel->setActiveSheetIndex(0)->setCellValue('B358', "a.n Riska Malinda NPM 1117032052");
		$excel->setActiveSheetIndex(0)->setCellValue('J358', "2281/UN26/7/DT/2015");//286
		$excel->setActiveSheetIndex(0)->setCellValue('J359', "tanggal 8 Juli 2015");//287

		$excel->setActiveSheetIndex(0)->setCellValue('B361', "3). Pembimbing pendamping/pembantu Skripsi mahasiswa");
		$excel->setActiveSheetIndex(0)->setCellValue('E361', "Smstr Ganjil");
		$excel->setActiveSheetIndex(0)->setCellValue('F361', "Lulusan");
		$excel->setActiveSheetIndex(0)->setCellValue('G361', "2.0");
		$excel->setActiveSheetIndex(0)->setCellValue('H361', "0.5");
		$excel->setActiveSheetIndex(0)->setCellValue('I361', "1.0");
		$excel->setActiveSheetIndex(0)->setCellValue('J361', "II.D.2.c");//289

		$excel->setActiveSheetIndex(0)->setCellValue('B362', "a.n Ahmad Amirudin NPM 1017032066");
		$excel->setActiveSheetIndex(0)->setCellValue('E362', "2015/2016");
		$excel->setActiveSheetIndex(0)->setCellValue('J362', "1. SK Dekan FMIPA Unila No. ");//290
		$excel->setActiveSheetIndex(0)->setCellValue('B363', "a.n M. Fathan Kurniawan NPM 1017032040");
		$excel->setActiveSheetIndex(0)->setCellValue('J363', "3533/UN26/7/DT/2015");//291
		$excel->setActiveSheetIndex(0)->setCellValue('J364', "tanggal 28 Desember 2015");//292

		$excel->setActiveSheetIndex(0)->setCellValue('H366', "Jumlah");
		$excel->setActiveSheetIndex(0)->setCellValue('I366', "5.00");
		$excel->getActiveSheet()->getStyle('H366:I366')->getFont()->setBold(TRUE);
		$excel->getActiveSheet()->getStyle('A366:L366')->applyFromArray($style_standar);//293

		$excel->getActiveSheet()->getStyle('B367:D375')->applyFromArray($style_standar);

		$excel->getActiveSheet()->getStyle('E367:E375')->applyFromArray($style_standar);
		$excel->getActiveSheet()->getStyle('F367:F375')->applyFromArray($style_standar);
		$excel->getActiveSheet()->getStyle('G367:G375')->applyFromArray($style_standar);
		$excel->getActiveSheet()->getStyle('H367:H375')->applyFromArray($style_standar);
		$excel->getActiveSheet()->getStyle('I367:I375')->applyFromArray($style_standar);
		$excel->getActiveSheet()->getStyle('J367:L375')->applyFromArray($style_standar);

		$excel->setActiveSheetIndex(0)->setCellValue('B367', "d. Laporan Akhir");
		$excel->setActiveSheetIndex(0)->setCellValue('J367', "II.D.2.d");//294

		$excel->getActiveSheet()->getStyle('E369:I375')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);

		$excel->setActiveSheetIndex(0)->setCellValue('B369', "1). Pembimbing pendamping/pembantu");
		$excel->setActiveSheetIndex(0)->setCellValue('E369', "Smstr Genap");
		$excel->setActiveSheetIndex(0)->setCellValue('F369', "Lulusan");
		$excel->setActiveSheetIndex(0)->setCellValue('G369', "3.0");
		$excel->setActiveSheetIndex(0)->setCellValue('H369', "0.5");
		$excel->setActiveSheetIndex(0)->setCellValue('I369', "1.5");
		$excel->setActiveSheetIndex(0)->setCellValue('J369', "II.D.2.c");//295

		$excel->setActiveSheetIndex(0)->setCellValue('B370', "a.n Hipzon Rosadi NPM 0917051043");
		$excel->setActiveSheetIndex(0)->setCellValue('E370', "2014/2015");
		$excel->setActiveSheetIndex(0)->setCellValue('J370', "1. SK Dekan FMIPA Unila No. ");//296
		$excel->setActiveSheetIndex(0)->setCellValue('B371', "a.n Nurfeti Ambarsari NPM 1207051048");
		$excel->setActiveSheetIndex(0)->setCellValue('J371', "3326/UN26/7/DT/2015");//297
		$excel->setActiveSheetIndex(0)->setCellValue('J372', "tanggal 21 Agustus 2015");//298

		$excel->setActiveSheetIndex(0)->setCellValue('B373', "a.n Wahyu Pangestu NPM 1107051048");
		$excel->setActiveSheetIndex(0)->setCellValue('J373', "2. SK Dekan FMIPA Unila No. ");//299
		$excel->setActiveSheetIndex(0)->setCellValue('J374', "892/UN26/7/DT/2015");//300
		$excel->setActiveSheetIndex(0)->setCellValue('J375', "tanggal 20 Maret 2015");//301

		$excel->setActiveSheetIndex(0)->setCellValue('H376', "Jumlah");
		$excel->setActiveSheetIndex(0)->setCellValue('I376', "1.50");
		$excel->getActiveSheet()->getStyle('H376:I376')->getFont()->setBold(TRUE);
		$excel->getActiveSheet()->getStyle('A376:L376')->applyFromArray($style_standar);//302

		$excel->setActiveSheetIndex(0)->setCellValue('A377', "E");
		$excel->setActiveSheetIndex(0)->setCellValue('B377', "Bertugas sebagai penguji pada ujian akhir");
		$excel->getActiveSheet()->getStyle('A377:B377')->getFont()->setBold(TRUE);
		$excel->getActiveSheet()->getStyle('A377')->applyFromArray($style_standar);
		$excel->getActiveSheet()->getStyle('B377:L377')->applyFromArray($style_standar);//303

		$excel->setActiveSheetIndex(0)->setCellValue('B378', "1. Ketua Penguji");//304
		$excel->setActiveSheetIndex(0)->setCellValue('B379', "a. Disertasi");
		$excel->setActiveSheetIndex(0)->setCellValue('J379', "II.E.1");//305
		$excel->setActiveSheetIndex(0)->setCellValue('H380', "Jumlah");
		$excel->setActiveSheetIndex(0)->setCellValue('I380', "0.0");
		$excel->getActiveSheet()->getStyle('H380:I380')->getFont()->setBold(TRUE);
		$excel->getActiveSheet()->getStyle('A380:L380')->applyFromArray($style_standar);//306
		$excel->setActiveSheetIndex(0)->setCellValue('B381', "b. Thesis");
		$excel->setActiveSheetIndex(0)->setCellValue('J381', "II.E.1");//307
		$excel->setActiveSheetIndex(0)->setCellValue('H382', "Jumlah");
		$excel->setActiveSheetIndex(0)->setCellValue('I382', "0.0");//308
		$excel->getActiveSheet()->getStyle('H382:I382')->getFont()->setBold(TRUE);
		$excel->getActiveSheet()->getStyle('A382:L382')->applyFromArray($style_standar);
		$excel->setActiveSheetIndex(0)->setCellValue('B383', "c. Skripsi");
		$excel->setActiveSheetIndex(0)->setCellValue('J383', "II.E.1");//309

		$excel->getActiveSheet()->getStyle('B383:D399')->applyFromArray($style_standar);

		$excel->getActiveSheet()->getStyle('E383:E399')->applyFromArray($style_standar);
		$excel->getActiveSheet()->getStyle('F383:F399')->applyFromArray($style_standar);
		$excel->getActiveSheet()->getStyle('G383:G399')->applyFromArray($style_standar);
		$excel->getActiveSheet()->getStyle('H383:H399')->applyFromArray($style_standar);
		$excel->getActiveSheet()->getStyle('I383:I399')->applyFromArray($style_standar);
		$excel->getActiveSheet()->getStyle('J383:L399')->applyFromArray($style_standar);

		$excel->getActiveSheet()->getStyle('E384:I399')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);

		$excel->setActiveSheetIndex(0)->setCellValue('B384', "1) Ketua Penguji Skripsi mahasiswa");
		$excel->setActiveSheetIndex(0)->setCellValue('E384', "Smstr Genap");
		$excel->setActiveSheetIndex(0)->setCellValue('F384', "Lulusan");
		$excel->setActiveSheetIndex(0)->setCellValue('G384', "4.0");
		$excel->setActiveSheetIndex(0)->setCellValue('H384', "1.0");
		$excel->setActiveSheetIndex(0)->setCellValue('I384', "4.0");
		$excel->setActiveSheetIndex(0)->setCellValue('J384', "II.E.1");//310

		$excel->setActiveSheetIndex(0)->setCellValue('B385', "a.n Irfani Maharani NPM 1317051033");
		$excel->setActiveSheetIndex(0)->setCellValue('E385', "2016/2017");
		$excel->setActiveSheetIndex(0)->setCellValue('J385', "1. SK Dekan FMIPA Unila No. ");//311
		$excel->setActiveSheetIndex(0)->setCellValue('B386', "a.n Faiq Sulthon Dani NPM 1347051060");
		$excel->setActiveSheetIndex(0)->setCellValue('J386', "3164/UN26/7/DT/2017");//312
		$excel->setActiveSheetIndex(0)->setCellValue('B387', "a.n Rizka Esa Basri NPM 1317051058");
		$excel->setActiveSheetIndex(0)->setCellValue('J387', "tanggal 3 Agustus 2017");//313
		$excel->setActiveSheetIndex(0)->setCellValue('B388', "a.n Navia Yufitasari NPM 1347051010");//314

		$excel->setActiveSheetIndex(0)->setCellValue('J389', "2. SK Dekan FMIPA Unila No. ");//415
		$excel->setActiveSheetIndex(0)->setCellValue('J390', "3579/UN26/7/DT/2017");//316
		$excel->setActiveSheetIndex(0)->setCellValue('J391', "tanggal 7 September 2017");//317

		$excel->setActiveSheetIndex(0)->setCellValue('B393', "2) Ketua Penguji Skripsi mahasiswa");
		$excel->setActiveSheetIndex(0)->setCellValue('E393', "Smstr Genap");
		$excel->setActiveSheetIndex(0)->setCellValue('F393', "Lulusan");
		$excel->setActiveSheetIndex(0)->setCellValue('G393', "4.0");
		$excel->setActiveSheetIndex(0)->setCellValue('H393', "1.0");
		$excel->setActiveSheetIndex(0)->setCellValue('I393', "4.0");
		$excel->setActiveSheetIndex(0)->setCellValue('J393', "II.E.1");//318

		$excel->setActiveSheetIndex(0)->setCellValue('B394', "a.n Irfani Maharani NPM 1317051033");
		$excel->setActiveSheetIndex(0)->setCellValue('E394', "2016/2017");
		$excel->setActiveSheetIndex(0)->setCellValue('J394', "1. SK Dekan FMIPA Unila No. ");//319
		$excel->setActiveSheetIndex(0)->setCellValue('B395', "a.n Faiq Sulthon Dani NPM 1347051060");
		$excel->setActiveSheetIndex(0)->setCellValue('J395', "3164/UN26/7/DT/2017");//320
		$excel->setActiveSheetIndex(0)->setCellValue('B396', "a.n Rizka Esa Basri NPM 1317051058");
		$excel->setActiveSheetIndex(0)->setCellValue('J397', "tanggal 3 Agustus 2017");//321
		$excel->setActiveSheetIndex(0)->setCellValue('B398', "a.n Navia Yufitasari NPM 1347051010");//322

		$excel->getActiveSheet()->getStyle('B378:L399')->applyFromArray($style_standar);

		$excel->setActiveSheetIndex(0)->setCellValue('H400', "Jumlah");
		$excel->setActiveSheetIndex(0)->setCellValue('I400', "8.0");//323
		$excel->getActiveSheet()->getStyle('H400:I400')->getFont()->setBold(TRUE);
		$excel->getActiveSheet()->getStyle('A400:L400')->applyFromArray($style_standar);//324

		$excel->setActiveSheetIndex(0)->setCellValue('B401', "2 Anggota penguji");//325
		$excel->setActiveSheetIndex(0)->setCellValue('B402', "a. Disertasi");//326
		$excel->setActiveSheetIndex(0)->setCellValue('H403', "Jumlah");
		$excel->setActiveSheetIndex(0)->setCellValue('I403', "0.0");//327
		$excel->getActiveSheet()->getStyle('H403:I403')->getFont()->setBold(TRUE);
		$excel->getActiveSheet()->getStyle('A403:L403')->applyFromArray($style_standar);
		$excel->setActiveSheetIndex(0)->setCellValue('B404', "b. Thesis");
		$excel->setActiveSheetIndex(0)->setCellValue('J404', "II.E.1");//328
		$excel->setActiveSheetIndex(0)->setCellValue('H405', "Jumlah");
		$excel->setActiveSheetIndex(0)->setCellValue('I405', "0.0");//329
		$excel->getActiveSheet()->getStyle('H405:I405')->getFont()->setBold(TRUE);
		$excel->getActiveSheet()->getStyle('A405:L405')->applyFromArray($style_standar);
		$excel->setActiveSheetIndex(0)->setCellValue('B406', "c. Skripsi");
		$excel->setActiveSheetIndex(0)->setCellValue('J406', "II.E.1");//330
		$excel->setActiveSheetIndex(0)->setCellValue('H408', "Jumlah");
		$excel->setActiveSheetIndex(0)->setCellValue('I408', "0.0");//331
		$excel->getActiveSheet()->getStyle('H408:I408')->getFont()->setBold(TRUE);
		$excel->getActiveSheet()->getStyle('A408:L408')->applyFromArray($style_standar);

		$excel->getActiveSheet()->getStyle('B401:L407')->applyFromArray($style_standar);


		$excel->setActiveSheetIndex(0)->setCellValue('A409', "F");
		$excel->setActiveSheetIndex(0)->setCellValue('B409', "Melakukan pembinaan kegiatan mahasiswa di bidang Akademik dan kemahasiswaan");
		$excel->getActiveSheet()->getStyle('A409:B409')->getFont()->setBold(TRUE);
		$excel->getActiveSheet()->getStyle('A409')->applyFromArray($style_standar);
		$excel->getActiveSheet()->getStyle('B409:L409')->applyFromArray($style_standar);//332

		$excel->getActiveSheet()->getStyle('E410:I508')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);

		$excel->setActiveSheetIndex(0)->setCellValue('B410', "1) Pembimbing Akademik Mhs S1 Ilmu Komputer");
		$excel->setActiveSheetIndex(0)->setCellValue('E410', "Smstr Ganjil ");
		$excel->setActiveSheetIndex(0)->setCellValue('F410', "Semester");
		$excel->setActiveSheetIndex(0)->setCellValue('G410', "1.0");
		$excel->setActiveSheetIndex(0)->setCellValue('H410', "2.0");
		$excel->setActiveSheetIndex(0)->setCellValue('I410', "2.0");
		$excel->setActiveSheetIndex(0)->setCellValue('J410', "II.F");//333

		$excel->setActiveSheetIndex(0)->setCellValue('B411', "33 mahasiswa");
		$excel->setActiveSheetIndex(0)->setCellValue('E411', "2013/2014");
		$excel->setActiveSheetIndex(0)->setCellValue('J411', "  SK Dekan FMIPA Unila No. ");//334
		$excel->setActiveSheetIndex(0)->setCellValue('J412', "2697a/UN26/7/DT/2013");//335
		$excel->setActiveSheetIndex(0)->setCellValue('J413', "tanggal 24 Oktober 2013");//336

		$excel->setActiveSheetIndex(0)->setCellValue('B414', "2) Pembimbing Akademik Mhs  D3 Mnjmn Informatika ");
		$excel->setActiveSheetIndex(0)->setCellValue('E414', "Smstr Ganjil ");
		$excel->setActiveSheetIndex(0)->setCellValue('F414', "Semester");
		$excel->setActiveSheetIndex(0)->setCellValue('G414', "1.0");
		$excel->setActiveSheetIndex(0)->setCellValue('H414', "2.0");
		$excel->setActiveSheetIndex(0)->setCellValue('I414', "2.0");
		$excel->setActiveSheetIndex(0)->setCellValue('J414', "II.F");//337

		$excel->setActiveSheetIndex(0)->setCellValue('B415', "21 mahasiswa");
		$excel->setActiveSheetIndex(0)->setCellValue('E415', "2013/2014");
		$excel->setActiveSheetIndex(0)->setCellValue('J415', "  SK Dekan FMIPA Unila No. ");//338
		$excel->setActiveSheetIndex(0)->setCellValue('J416', "2606/UN26/7/DT/2013");//339
		$excel->setActiveSheetIndex(0)->setCellValue('J417', "tanggal 7 Oktober 2013");//340

		$excel->setActiveSheetIndex(0)->setCellValue('B418', "3) Pembimbing Akademik Mhs S1 Ilmu Komputer");
		$excel->setActiveSheetIndex(0)->setCellValue('E418', "Smstr Genap ");
		$excel->setActiveSheetIndex(0)->setCellValue('F418', "Semester");
		$excel->setActiveSheetIndex(0)->setCellValue('G418', "1.0");
		$excel->setActiveSheetIndex(0)->setCellValue('H418', "2.0");
		$excel->setActiveSheetIndex(0)->setCellValue('I418', "2.0");
		$excel->setActiveSheetIndex(0)->setCellValue('J418', "II.F");//341

		$excel->setActiveSheetIndex(0)->setCellValue('B419', "30 mahasiswa");
		$excel->setActiveSheetIndex(0)->setCellValue('E419', "2013/2014");
		$excel->setActiveSheetIndex(0)->setCellValue('J419', "  SK Dekan FMIPA Unila No. ");//342
		$excel->setActiveSheetIndex(0)->setCellValue('J420', "853/UN26/7/DT/2014");//343
		$excel->setActiveSheetIndex(0)->setCellValue('J421', "tanggal 30 April 2014 ");//344

		$excel->setActiveSheetIndex(0)->setCellValue('B422', "4) Pembimbing Akademik Mhs  D3 Mnjmn Informatika ");
		$excel->setActiveSheetIndex(0)->setCellValue('E422', "Smstr Genap ");
		$excel->setActiveSheetIndex(0)->setCellValue('F422', "Semester");
		$excel->setActiveSheetIndex(0)->setCellValue('G422', "1.0");
		$excel->setActiveSheetIndex(0)->setCellValue('H422', "2.0");
		$excel->setActiveSheetIndex(0)->setCellValue('I422', "2.0");
		$excel->setActiveSheetIndex(0)->setCellValue('J422', "II.F");//345

		$excel->setActiveSheetIndex(0)->setCellValue('B423', "21 mahasiswa");
		$excel->setActiveSheetIndex(0)->setCellValue('E423', "2013/2014");
		$excel->setActiveSheetIndex(0)->setCellValue('J423', "  SK Dekan FMIPA Unila No. ");//346
		$excel->setActiveSheetIndex(0)->setCellValue('J424', "140/UN26/7/DT/2014");//347
		$excel->setActiveSheetIndex(0)->setCellValue('J425', "tanggal 16 Januari 2014");//348


		$excel->setActiveSheetIndex(0)->setCellValue('B427', "5) Pembimbing Akademik Mhs S1 Ilmu Komputer");
		$excel->setActiveSheetIndex(0)->setCellValue('E427', "Smstr Ganjil ");
		$excel->setActiveSheetIndex(0)->setCellValue('F427', "Semester");
		$excel->setActiveSheetIndex(0)->setCellValue('G427', "1.0");
		$excel->setActiveSheetIndex(0)->setCellValue('H427', "2.0");
		$excel->setActiveSheetIndex(0)->setCellValue('I427', "2.0");
		$excel->setActiveSheetIndex(0)->setCellValue('J427', "II.F");//349

		$excel->setActiveSheetIndex(0)->setCellValue('B428', "38 mahasiswa");
		$excel->setActiveSheetIndex(0)->setCellValue('E428', "2014/2015");
		$excel->setActiveSheetIndex(0)->setCellValue('J428', "  SK Dekan FMIPA Unila No. ");//350
		$excel->setActiveSheetIndex(0)->setCellValue('J429', "2151/UN26/7/DT/2014");//351
		$excel->setActiveSheetIndex(0)->setCellValue('J430', "tanggal 4 November 2014 ");//352

		$excel->setActiveSheetIndex(0)->setCellValue('B431', "6) Pembimbing Akademik Mhs  D3 Mnjmn Informatika ");
		$excel->setActiveSheetIndex(0)->setCellValue('E431', "Smstr Ganjil ");
		$excel->setActiveSheetIndex(0)->setCellValue('F431', "Semester");
		$excel->setActiveSheetIndex(0)->setCellValue('G431', "1.0");
		$excel->setActiveSheetIndex(0)->setCellValue('H431', "2.0");
		$excel->setActiveSheetIndex(0)->setCellValue('I431', "2.0");
		$excel->setActiveSheetIndex(0)->setCellValue('J431', "II.F");//353

		$excel->setActiveSheetIndex(0)->setCellValue('B432', "24 mahasiswa");
		$excel->setActiveSheetIndex(0)->setCellValue('E432', "2014/2015");
		$excel->setActiveSheetIndex(0)->setCellValue('J432', "  SK Dekan FMIPA Unila No. ");//354
		$excel->setActiveSheetIndex(0)->setCellValue('J433', "2017/UN26/7/DT/2014");//355
		$excel->setActiveSheetIndex(0)->setCellValue('J434', "tanggal 9 Oktober 2014");//356

		$excel->setActiveSheetIndex(0)->setCellValue('B436', "7) Pembimbing Akademik Mhs S1 Ilmu Komputer");
		$excel->setActiveSheetIndex(0)->setCellValue('E436', "Smstr Genap ");
		$excel->setActiveSheetIndex(0)->setCellValue('F436', "Semester");
		$excel->setActiveSheetIndex(0)->setCellValue('G436', "1.0");
		$excel->setActiveSheetIndex(0)->setCellValue('H436', "2.0");
		$excel->setActiveSheetIndex(0)->setCellValue('I436', "2.0");
		$excel->setActiveSheetIndex(0)->setCellValue('J436', "II.F");//357

		$excel->setActiveSheetIndex(0)->setCellValue('B437', "38 mahasiswa");
		$excel->setActiveSheetIndex(0)->setCellValue('E437', "2014/2015");
		$excel->setActiveSheetIndex(0)->setCellValue('J437', "  SK Dekan FMIPA Unila No. ");//358
		$excel->setActiveSheetIndex(0)->setCellValue('J438', "1786/UN26/7/DT/2015");//359
		$excel->setActiveSheetIndex(0)->setCellValue('J439', "Tanggal 18 Mei 2015");//360

		$excel->setActiveSheetIndex(0)->setCellValue('B440', "8) Pembimbing Akademik Mhs  D3 Mnjmn Informatika ");
		$excel->setActiveSheetIndex(0)->setCellValue('E440', "Smstr Genap ");
		$excel->setActiveSheetIndex(0)->setCellValue('F440', "Semester");
		$excel->setActiveSheetIndex(0)->setCellValue('G440', "1.0");
		$excel->setActiveSheetIndex(0)->setCellValue('H440', "2.0");
		$excel->setActiveSheetIndex(0)->setCellValue('I440', "2.0");
		$excel->setActiveSheetIndex(0)->setCellValue('J440', "II.F");//362

		$excel->setActiveSheetIndex(0)->setCellValue('B441', "20 mahasiswa");
		$excel->setActiveSheetIndex(0)->setCellValue('E441', "2014/2015");
		$excel->setActiveSheetIndex(0)->setCellValue('J441', "  SK Dekan FMIPA Unila No. ");//362
		$excel->setActiveSheetIndex(0)->setCellValue('J442', "968/UN26/7/DT/2015");//363
		$excel->setActiveSheetIndex(0)->setCellValue('J443', "Tanggal 23 Maret 2015");//364


		$excel->setActiveSheetIndex(0)->setCellValue('B445', "9) Pembimbing Akademik Mhs S1 Ilmu Komputer");
		$excel->setActiveSheetIndex(0)->setCellValue('E445', "Smstr Ganjil ");
		$excel->setActiveSheetIndex(0)->setCellValue('F445', "Semester");
		$excel->setActiveSheetIndex(0)->setCellValue('G445', "1.0");
		$excel->setActiveSheetIndex(0)->setCellValue('H445', "2.0");
		$excel->setActiveSheetIndex(0)->setCellValue('I445', "2.0");
		$excel->setActiveSheetIndex(0)->setCellValue('J445', "II.F");//365

		$excel->setActiveSheetIndex(0)->setCellValue('B446', "60 mahasiswa");
		$excel->setActiveSheetIndex(0)->setCellValue('E446', "2015/2016");
		$excel->setActiveSheetIndex(0)->setCellValue('J446', "  SK Dekan FMIPA Unila No. ");//366
		$excel->setActiveSheetIndex(0)->setCellValue('J447', "2996/UN26/7/DT/2015");//367
		$excel->setActiveSheetIndex(0)->setCellValue('J447', "tanggal 19 Oktober 2015 ");//368

		$excel->setActiveSheetIndex(0)->setCellValue('B449', "10) Pembimbing Akademik Mhs  D3 Mnjmn Informatika ");
		$excel->setActiveSheetIndex(0)->setCellValue('E449', "Smstr Ganjil ");
		$excel->setActiveSheetIndex(0)->setCellValue('F449', "Semester");
		$excel->setActiveSheetIndex(0)->setCellValue('G449', "1.0");
		$excel->setActiveSheetIndex(0)->setCellValue('H449', "2.0");
		$excel->setActiveSheetIndex(0)->setCellValue('I449', "2.0");
		$excel->setActiveSheetIndex(0)->setCellValue('J449', "II.F");//369

		$excel->setActiveSheetIndex(0)->setCellValue('B450', "22 mahasiswa");
		$excel->setActiveSheetIndex(0)->setCellValue('E450', "22015/2016");
		$excel->setActiveSheetIndex(0)->setCellValue('J450', "  SK Dekan FMIPA Unila No. ");//370
		$excel->setActiveSheetIndex(0)->setCellValue('J451', "2977/UN26/7/DT/2015");//371
		$excel->setActiveSheetIndex(0)->setCellValue('J452', "tanggal 19 Oktober 2015");//372

		$excel->setActiveSheetIndex(0)->setCellValue('B454', "11) Pembimbing Akademik Mhs S1 Ilmu Komputer");
		$excel->setActiveSheetIndex(0)->setCellValue('E454', "Smstr Genap ");
		$excel->setActiveSheetIndex(0)->setCellValue('F454', "Semester");
		$excel->setActiveSheetIndex(0)->setCellValue('G454', "1.0");
		$excel->setActiveSheetIndex(0)->setCellValue('H454', "2.0");
		$excel->setActiveSheetIndex(0)->setCellValue('I454', "2.0");
		$excel->setActiveSheetIndex(0)->setCellValue('J454', "II.F");//373

		$excel->setActiveSheetIndex(0)->setCellValue('B455', "58 mahasiswa");
		$excel->setActiveSheetIndex(0)->setCellValue('E455', "2015/2016");
		$excel->setActiveSheetIndex(0)->setCellValue('J455', "  SK Dekan FMIPA Unila No. ");//374
		$excel->setActiveSheetIndex(0)->setCellValue('J456', "766/UN26/7/DT/2016");//375
		$excel->setActiveSheetIndex(0)->setCellValue('J457', "Tanggal 6 April 2016");//376

		$excel->setActiveSheetIndex(0)->setCellValue('B458', "12) Pembimbing Akademik Mhs  D3 Mnjmn Informatika ");
		$excel->setActiveSheetIndex(0)->setCellValue('E458', "Smstr Genap ");
		$excel->setActiveSheetIndex(0)->setCellValue('F458', "Semester");
		$excel->setActiveSheetIndex(0)->setCellValue('G458', "1.0");
		$excel->setActiveSheetIndex(0)->setCellValue('H458', "2.0");
		$excel->setActiveSheetIndex(0)->setCellValue('I458', "2.0");
		$excel->setActiveSheetIndex(0)->setCellValue('J458', "II.F");//377

		$excel->setActiveSheetIndex(0)->setCellValue('B459', "19 mahasiswa");
		$excel->setActiveSheetIndex(0)->setCellValue('E459', "2015/2016");
		$excel->setActiveSheetIndex(0)->setCellValue('J459', "  SK Dekan FMIPA Unila No. ");//378
		$excel->setActiveSheetIndex(0)->setCellValue('J460', "768/UN26/7/DT/2016");//379
		$excel->setActiveSheetIndex(0)->setCellValue('J461', "Tanggal 6 April 2016");//380

		$excel->setActiveSheetIndex(0)->setCellValue('B463', "13) Pembimbing Akademik Mhs S1 Ilmu Komputer");
		$excel->setActiveSheetIndex(0)->setCellValue('E463', "Smstr Ganjil ");
		$excel->setActiveSheetIndex(0)->setCellValue('F463', "Semester");
		$excel->setActiveSheetIndex(0)->setCellValue('G463', "1.0");
		$excel->setActiveSheetIndex(0)->setCellValue('H463', "2.0");
		$excel->setActiveSheetIndex(0)->setCellValue('I463', "2.0");
		$excel->setActiveSheetIndex(0)->setCellValue('J463', "II.F");//381

		$excel->setActiveSheetIndex(0)->setCellValue('B464', "59 mahasiswa");
		$excel->setActiveSheetIndex(0)->setCellValue('E464', "2016/2017");
		$excel->setActiveSheetIndex(0)->setCellValue('J464', "  SK Dekan FMIPA Unila No. ");//382
		$excel->setActiveSheetIndex(0)->setCellValue('J465', "2502/UN26/7/DT/2016");//383
		$excel->setActiveSheetIndex(0)->setCellValue('J466', "tanggal 4 Oktober 2016");//384

		$excel->setActiveSheetIndex(0)->setCellValue('B467', "14) Pembimbing Akademik Mhs  D3 Mnjmn Informatika ");
		$excel->setActiveSheetIndex(0)->setCellValue('E467', "Smstr Ganjil ");
		$excel->setActiveSheetIndex(0)->setCellValue('F467', "Semester");
		$excel->setActiveSheetIndex(0)->setCellValue('G467', "1.0");
		$excel->setActiveSheetIndex(0)->setCellValue('H467', "2.0");
		$excel->setActiveSheetIndex(0)->setCellValue('I467', "2.0");
		$excel->setActiveSheetIndex(0)->setCellValue('J467', "II.F");//385

		$excel->setActiveSheetIndex(0)->setCellValue('B468', "21 mahasiswa");
		$excel->setActiveSheetIndex(0)->setCellValue('E468', "2016/2017");
		$excel->setActiveSheetIndex(0)->setCellValue('J468', "  SK Dekan FMIPA Unila No. ");//386
		$excel->setActiveSheetIndex(0)->setCellValue('J469', "2512/UN26/7/DT/2016");//387
		$excel->setActiveSheetIndex(0)->setCellValue('J470', "tanggal 17 Oktober 2016");//388

		$excel->setActiveSheetIndex(0)->setCellValue('B472', "15) Pembimbing Akademik Mhs S1 Ilmu Komputer");
		$excel->setActiveSheetIndex(0)->setCellValue('E472', "Smstr Genap ");
		$excel->setActiveSheetIndex(0)->setCellValue('F472', "Semester");
		$excel->setActiveSheetIndex(0)->setCellValue('G472', "1.0");
		$excel->setActiveSheetIndex(0)->setCellValue('H472', "2.0");
		$excel->setActiveSheetIndex(0)->setCellValue('I472', "2.0");
		$excel->setActiveSheetIndex(0)->setCellValue('J472', "II.F");//389

		$excel->setActiveSheetIndex(0)->setCellValue('B473', "52 mahasiswa");
		$excel->setActiveSheetIndex(0)->setCellValue('E473', "2016/2017");
		$excel->setActiveSheetIndex(0)->setCellValue('J473', "  SK Dekan FMIPA Unila No. ");//390
		$excel->setActiveSheetIndex(0)->setCellValue('J474', "1948/UN26/7/DT/2017");//391
		$excel->setActiveSheetIndex(0)->setCellValue('J475', "Tanggal 25 April 2017");//392

		$excel->setActiveSheetIndex(0)->setCellValue('B476', "16) Pembimbing Akademik Mhs  D3 Mnjmn Informatika ");
		$excel->setActiveSheetIndex(0)->setCellValue('E476', "Smstr Genap ");
		$excel->setActiveSheetIndex(0)->setCellValue('F476', "Semester");
		$excel->setActiveSheetIndex(0)->setCellValue('G476', "1.0");
		$excel->setActiveSheetIndex(0)->setCellValue('H476', "2.0");
		$excel->setActiveSheetIndex(0)->setCellValue('I476', "2.0");
		$excel->setActiveSheetIndex(0)->setCellValue('J476', "II.F");//393

		$excel->setActiveSheetIndex(0)->setCellValue('B477', "15 mahasiswa");
		$excel->setActiveSheetIndex(0)->setCellValue('E477', "2016/2017");
		$excel->setActiveSheetIndex(0)->setCellValue('J477', "  SK Dekan FMIPA Unila No. ");//394
		$excel->setActiveSheetIndex(0)->setCellValue('J478', "1957/UN26/7/DT/2017");//395
		$excel->setActiveSheetIndex(0)->setCellValue('J479', "Tanggal 25 April 2017");//396

		$excel->setActiveSheetIndex(0)->setCellValue('B481', "17) Pembimbing Akademik Mhs S1 Ilmu Komputer");
		$excel->setActiveSheetIndex(0)->setCellValue('E481', "Smstr Ganjil ");
		$excel->setActiveSheetIndex(0)->setCellValue('F481', "Semester");
		$excel->setActiveSheetIndex(0)->setCellValue('G481', "1.0");
		$excel->setActiveSheetIndex(0)->setCellValue('H481', "2.0");
		$excel->setActiveSheetIndex(0)->setCellValue('I481', "2.0");
		$excel->setActiveSheetIndex(0)->setCellValue('J481', "II.F");//397

		$excel->setActiveSheetIndex(0)->setCellValue('B482', "55 mahasiswa");
		$excel->setActiveSheetIndex(0)->setCellValue('E482', "2017/2018");
		$excel->setActiveSheetIndex(0)->setCellValue('J482', "  SK Dekan FMIPA Unila No. ");//398
		$excel->setActiveSheetIndex(0)->setCellValue('J483', "4581/UN26/7/DT/2017");//399
		$excel->setActiveSheetIndex(0)->setCellValue('J484', "Tanggal 20 November 2017");//400

		$excel->setActiveSheetIndex(0)->setCellValue('B485', "18) Pembimbing Akademik Mhs  D3 Mnjmn Informatika ");
		$excel->setActiveSheetIndex(0)->setCellValue('E485', "Smstr Ganjil ");
		$excel->setActiveSheetIndex(0)->setCellValue('F485', "Semester");
		$excel->setActiveSheetIndex(0)->setCellValue('G485', "1.0");
		$excel->setActiveSheetIndex(0)->setCellValue('H485', "2.0");
		$excel->setActiveSheetIndex(0)->setCellValue('I485', "2.0");
		$excel->setActiveSheetIndex(0)->setCellValue('J485', "II.F");//401

		$excel->setActiveSheetIndex(0)->setCellValue('B486', "16 mahasiswa");
		$excel->setActiveSheetIndex(0)->setCellValue('E486', "2017/2018");
		$excel->setActiveSheetIndex(0)->setCellValue('J486', "  SK Dekan FMIPA Unila No. ");//402
		$excel->setActiveSheetIndex(0)->setCellValue('J487', "3996/UN26/7/DT/2017");//403
		$excel->setActiveSheetIndex(0)->setCellValue('J488', "Tanggal 17 Oktober 2017");//404

		$excel->setActiveSheetIndex(0)->setCellValue('B490', "19) Pembimbing Akademik Mhs S1 Ilmu Komputer");
		$excel->setActiveSheetIndex(0)->setCellValue('E490', "Smstr Genap ");
		$excel->setActiveSheetIndex(0)->setCellValue('F490', "Semester");
		$excel->setActiveSheetIndex(0)->setCellValue('G490', "1.0");
		$excel->setActiveSheetIndex(0)->setCellValue('H490', "2.0");
		$excel->setActiveSheetIndex(0)->setCellValue('I490', "2.0");
		$excel->setActiveSheetIndex(0)->setCellValue('J490', "II.F");//405

		$excel->setActiveSheetIndex(0)->setCellValue('B491', "51 mahasiswa");
		$excel->setActiveSheetIndex(0)->setCellValue('E491', "2017/2018");
		$excel->setActiveSheetIndex(0)->setCellValue('J491', "  SK Dekan FMIPA Unila No. ");//406
		$excel->setActiveSheetIndex(0)->setCellValue('J492', "2022/UN26/7/DT/2018");//407
		$excel->setActiveSheetIndex(0)->setCellValue('J493', "Tanggal 26 Juni 2018");//408

		$excel->setActiveSheetIndex(0)->setCellValue('B494', "20) Pembimbing Akademik Mhs  D3 Mnjmn Informatika ");
		$excel->setActiveSheetIndex(0)->setCellValue('E494', "Smstr Genap ");
		$excel->setActiveSheetIndex(0)->setCellValue('F494', "Semester");
		$excel->setActiveSheetIndex(0)->setCellValue('G494', "1.0");
		$excel->setActiveSheetIndex(0)->setCellValue('H494', "2.0");
		$excel->setActiveSheetIndex(0)->setCellValue('I494', "2.0");
		$excel->setActiveSheetIndex(0)->setCellValue('J494', "II.F");//409

		$excel->setActiveSheetIndex(0)->setCellValue('B495', "11 mahasiswa");
		$excel->setActiveSheetIndex(0)->setCellValue('E495', "2017/2018");
		$excel->setActiveSheetIndex(0)->setCellValue('J495', "  SK Dekan FMIPA Unila No. ");//410
		$excel->setActiveSheetIndex(0)->setCellValue('J496', "3996/UN26/7/DT/2017");//411
		$excel->setActiveSheetIndex(0)->setCellValue('J497', "Tanggal 17 Oktober 2017");//412


		$excel->setActiveSheetIndex(0)->setCellValue('B499', "21) Pembimbing Akademik Mhs S1 Ilmu Komputer");
		$excel->setActiveSheetIndex(0)->setCellValue('E499', "Smstr Ganjil ");
		$excel->setActiveSheetIndex(0)->setCellValue('F499', "Semester");
		$excel->setActiveSheetIndex(0)->setCellValue('G499', "1.0");
		$excel->setActiveSheetIndex(0)->setCellValue('H499', "2.0");
		$excel->setActiveSheetIndex(0)->setCellValue('I499', "2.0");
		$excel->setActiveSheetIndex(0)->setCellValue('J499', "II.F");//413

		$excel->setActiveSheetIndex(0)->setCellValue('B500', "45 mahasiswa");
		$excel->setActiveSheetIndex(0)->setCellValue('E500', "2018/2019");
		$excel->setActiveSheetIndex(0)->setCellValue('J500', "  SK Dekan FMIPA Unila No. ");//414
		$excel->setActiveSheetIndex(0)->setCellValue('J501', "3690/UN26/7/DT/2018");//415
		$excel->setActiveSheetIndex(0)->setCellValue('J502', "Tanggal 31 Oktober 2018");//416

		$excel->setActiveSheetIndex(0)->setCellValue('B503', "22) Pembimbing Akademik Mhs  D3 Mnjmn Informatika ");
		$excel->setActiveSheetIndex(0)->setCellValue('E503', "Smstr Ganjil ");
		$excel->setActiveSheetIndex(0)->setCellValue('F503', "Semester");
		$excel->setActiveSheetIndex(0)->setCellValue('G503', "1.0");
		$excel->setActiveSheetIndex(0)->setCellValue('H503', "2.0");
		$excel->setActiveSheetIndex(0)->setCellValue('I503', "2.0");
		$excel->setActiveSheetIndex(0)->setCellValue('J503', "II.F");//417

		$excel->setActiveSheetIndex(0)->setCellValue('B504', "16 mahasiswa");
		$excel->setActiveSheetIndex(0)->setCellValue('E504', "2018/2019");
		$excel->setActiveSheetIndex(0)->setCellValue('J504', "  SK Dekan FMIPA Unila No. ");//418
		$excel->setActiveSheetIndex(0)->setCellValue('J505', "3736/UN26/7/DT/2018");//419
		$excel->setActiveSheetIndex(0)->setCellValue('J506', "Tanggal 02 November 2018");

		$excel->setActiveSheetIndex(0)->setCellValue('H509', "Jumlah");
		$excel->setActiveSheetIndex(0)->setCellValue('I509', "36.0");//420
		$excel->getActiveSheet()->getStyle('H509:I509')->getFont()->setBold(TRUE);
		$excel->getActiveSheet()->getStyle('A509:L509')->applyFromArray($style_standar);

		$excel->getActiveSheet()->getStyle('A410:A508')->applyFromArray($style_standar);
		$excel->getActiveSheet()->getStyle('B410:D508')->applyFromArray($style_standar);
		$excel->getActiveSheet()->getStyle('E410:E508')->applyFromArray($style_standar);
		$excel->getActiveSheet()->getStyle('F410:F508')->applyFromArray($style_standar);
		$excel->getActiveSheet()->getStyle('G410:G508')->applyFromArray($style_standar);
		$excel->getActiveSheet()->getStyle('H410:H508')->applyFromArray($style_standar);
		$excel->getActiveSheet()->getStyle('I410:I508')->applyFromArray($style_standar);
		$excel->getActiveSheet()->getStyle('J410:L508')->applyFromArray($style_standar);

		$excel->setActiveSheetIndex(0)->setCellValue('A510', "G");
		$excel->setActiveSheetIndex(0)->setCellValue('B510', "Mengembangkan program kuliah");
		$excel->getActiveSheet()->getStyle('A510:B510')->getFont()->setBold(TRUE);
		$excel->getActiveSheet()->getStyle('A510')->applyFromArray($style_standar);
		$excel->getActiveSheet()->getStyle('B510:L510')->applyFromArray($style_standar);//421

		$excel->setActiveSheetIndex(0)->setCellValue('B511', "Melakukan kegiatan pengembangan program kuliah");
		$excel->setActiveSheetIndex(0)->setCellValue('J511', "II.G");//422
		$excel->setActiveSheetIndex(0)->setCellValue('H512', "Jumlah");
		$excel->setActiveSheetIndex(0)->setCellValue('I512', "0.0");//420
		$excel->getActiveSheet()->getStyle('H512:I512')->getFont()->setBold(TRUE);
		$excel->getActiveSheet()->getStyle('A512:L512')->applyFromArray($style_standar);

		$excel->setActiveSheetIndex(0)->setCellValue('A513', "H");
		$excel->setActiveSheetIndex(0)->setCellValue('B513', "Mengembangkan bahan pengajaran");
		$excel->getActiveSheet()->getStyle('A513:B513')->getFont()->setBold(TRUE);
		$excel->getActiveSheet()->getStyle('A513')->applyFromArray($style_standar);
		$excel->getActiveSheet()->getStyle('B513:L513')->applyFromArray($style_standar);//421

		$excel->setActiveSheetIndex(0)->setCellValue('B514', "1 Buku Ajar");
		$excel->setActiveSheetIndex(0)->setCellValue('J514', "II.H.1");//422
		$excel->setActiveSheetIndex(0)->setCellValue('H515', "Jumlah");
		$excel->setActiveSheetIndex(0)->setCellValue('I515', "0.0");//423
		$excel->getActiveSheet()->getStyle('H515:I515')->getFont()->setBold(TRUE);
		$excel->getActiveSheet()->getStyle('A515:L515')->applyFromArray($style_standar);

		$excel->setActiveSheetIndex(0)->setCellValue('A516', "I");
		$excel->setActiveSheetIndex(0)->setCellValue('B516', "Menyampaikan orasi ilmiah");
		$excel->getActiveSheet()->getStyle('A516:B516')->getFont()->setBold(TRUE);
		$excel->getActiveSheet()->getStyle('A516')->applyFromArray($style_standar);
		$excel->getActiveSheet()->getStyle('B516:L516')->applyFromArray($style_standar);//424

		$excel->setActiveSheetIndex(0)->setCellValue('B517', "Melakukan kegiatan orasi ilmiah pada perguruan tinggi tiap tahun ");
		$excel->setActiveSheetIndex(0)->setCellValue('J517', "II.I");//425
		$excel->setActiveSheetIndex(0)->setCellValue('H518', "Jumlah");
		$excel->setActiveSheetIndex(0)->setCellValue('I518', "0.0");//426
		$excel->getActiveSheet()->getStyle('H518:I518')->getFont()->setBold(TRUE);
		$excel->getActiveSheet()->getStyle('A518:L518')->applyFromArray($style_standar);

		$excel->setActiveSheetIndex(0)->setCellValue('A519', "J.");
		$excel->setActiveSheetIndex(0)->setCellValue('B519', "Menduduki jabatan pimpinan perguruan tinggi");
		$excel->getActiveSheet()->getStyle('A519:B519')->getFont()->setBold(TRUE);
		$excel->getActiveSheet()->getStyle('A519')->applyFromArray($style_standar);
		$excel->getActiveSheet()->getStyle('B519:L519')->applyFromArray($style_standar);//427

		$excel->setActiveSheetIndex(0)->setCellValue('B520', "1 Rektor");
		$excel->setActiveSheetIndex(0)->setCellValue('J520', "II.J.1.");//428
		$excel->setActiveSheetIndex(0)->setCellValue('B521', "2 Pembantu rektor/dekan/direktur program pasca sarjana");
		$excel->setActiveSheetIndex(0)->setCellValue('J521', "II.J.1.");//429
		$excel->setActiveSheetIndex(0)->setCellValue('B522', "3 Ketua sekolah tinggi/pembantu dekan/asisten direktur program pasca sarjana/direktur politeknik");
		$excel->setActiveSheetIndex(0)->setCellValue('J522', "II.J.1.");//430
		$excel->setActiveSheetIndex(0)->setCellValue('B523', "4 Pembantu ketua sekolah tinggi/pembantu direktur politeknik ");
		$excel->setActiveSheetIndex(0)->setCellValue('J523', "II.J.1.");//431
		$excel->setActiveSheetIndex(0)->setCellValue('B524', "5 Direktur akademi");
		$excel->setActiveSheetIndex(0)->setCellValue('J524', "II.J.1.");//432
		$excel->setActiveSheetIndex(0)->setCellValue('B525', "6 Pembantu direktur akademi/ketua jurusan/bagian pada Universitas/institut/sekolah tinggi");
		$excel->setActiveSheetIndex(0)->setCellValue('J525', "II.J.1.");//433
		$excel->setActiveSheetIndex(0)->setCellValue('B526', "7 Ketua jurusan pada politeknik/akademi/ sekretaris jurusan/bagian pada universitas/ institut/sekolah tinggi");
		$excel->setActiveSheetIndex(0)->setCellValue('J526', "II.J.1.");//434
		$excel->setActiveSheetIndex(0)->setCellValue('B527', "8 Sekretaris jurusan pada politeknik/akademik dan kepala laboratorium universitas/institut/sekolah tinggi/politeknik/akademi");
		$excel->setActiveSheetIndex(0)->setCellValue('J527', "II.J.8");//435

		$excel->getActiveSheet()->getStyle('J520:L527')->applyFromArray($style_standar);

		$excel->getActiveSheet()->getStyle('B520:L548')->applyFromArray($style_standar);

		$excel->getActiveSheet()->getStyle('B528:D548')->applyFromArray($style_standar);

		$excel->getActiveSheet()->getStyle('E528:E548')->applyFromArray($style_standar);
		$excel->getActiveSheet()->getStyle('F528:F548')->applyFromArray($style_standar);
		$excel->getActiveSheet()->getStyle('G528:G548')->applyFromArray($style_standar);
		$excel->getActiveSheet()->getStyle('H528:H548')->applyFromArray($style_standar);
		$excel->getActiveSheet()->getStyle('I528:I548')->applyFromArray($style_standar);
		$excel->getActiveSheet()->getStyle('J528:L548')->applyFromArray($style_standar);

		$excel->getActiveSheet()->getStyle('E528:I548')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);


		$excel->setActiveSheetIndex(0)->setCellValue('B528', "    1 Menjadi Kepala Laboratorium Rekayasa Perangkat Lunak");
		$excel->setActiveSheetIndex(0)->setCellValue('E528', "Smstr Ganjil ");
		$excel->setActiveSheetIndex(0)->setCellValue('F528', "1 Semester");
		$excel->setActiveSheetIndex(0)->setCellValue('G528', "1.0");
		$excel->setActiveSheetIndex(0)->setCellValue('H528', "3.0");
		$excel->setActiveSheetIndex(0)->setCellValue('I528', "3.0");
		$excel->setActiveSheetIndex(0)->setCellValue('J528', "SK Rektor Nomor");//436

		$excel->setActiveSheetIndex(0)->setCellValue('E529', "2015/2016");
		$excel->setActiveSheetIndex(0)->setCellValue('J529', "646/UN26/KP/2016");//437
		$excel->setActiveSheetIndex(0)->setCellValue('J530', "Tanggal 30 Mei 216");//438

		$excel->setActiveSheetIndex(0)->setCellValue('B531', "    2 Menjadi Kepala Laboratorium Rekayasa Perangkat Lunak");
		$excel->setActiveSheetIndex(0)->setCellValue('E531', "Smstr Genap");
		$excel->setActiveSheetIndex(0)->setCellValue('F531', "1 Semester");
		$excel->setActiveSheetIndex(0)->setCellValue('G531', "1.0");
		$excel->setActiveSheetIndex(0)->setCellValue('H531', "3.0");
		$excel->setActiveSheetIndex(0)->setCellValue('I531', "3.0");
		$excel->setActiveSheetIndex(0)->setCellValue('J531', "SK Rektor Nomor");//439

		$excel->setActiveSheetIndex(0)->setCellValue('E532', "2015/2016");
		$excel->setActiveSheetIndex(0)->setCellValue('J532', "646/UN26/KP/2016");//440
		$excel->setActiveSheetIndex(0)->setCellValue('J533', "Tanggal 30 Mei 216");//441

		$excel->setActiveSheetIndex(0)->setCellValue('B534', "    3 Menjadi Kepala Laboratorium Rekayasa Perangkat Lunak");
		$excel->setActiveSheetIndex(0)->setCellValue('E534', "Smstr Ganjil ");
		$excel->setActiveSheetIndex(0)->setCellValue('F534', "1 Semester");
		$excel->setActiveSheetIndex(0)->setCellValue('G534', "1.0");
		$excel->setActiveSheetIndex(0)->setCellValue('H534', "3.0");
		$excel->setActiveSheetIndex(0)->setCellValue('I534', "3.0");
		$excel->setActiveSheetIndex(0)->setCellValue('J534', "SK Rektor Nomor");//442

		$excel->setActiveSheetIndex(0)->setCellValue('E535', "2016/2017");
		$excel->setActiveSheetIndex(0)->setCellValue('J535', "646/UN26/KP/2016");//443
		$excel->setActiveSheetIndex(0)->setCellValue('J536', "Tanggal 30 Mei 216");//444

		$excel->setActiveSheetIndex(0)->setCellValue('B537', "    4 Menjadi Kepala Laboratorium Rekayasa Perangkat Lunak");
		$excel->setActiveSheetIndex(0)->setCellValue('E537', "Smstr Genap");
		$excel->setActiveSheetIndex(0)->setCellValue('F537', "1 Semester");
		$excel->setActiveSheetIndex(0)->setCellValue('G537', "1.0");
		$excel->setActiveSheetIndex(0)->setCellValue('H537', "3.0");
		$excel->setActiveSheetIndex(0)->setCellValue('I537', "3.0");
		$excel->setActiveSheetIndex(0)->setCellValue('J537', "SK Rektor Nomor");//445

		$excel->setActiveSheetIndex(0)->setCellValue('E538', "2016/2017");
		$excel->setActiveSheetIndex(0)->setCellValue('J538', "646/UN26/KP/2016");//446
		$excel->setActiveSheetIndex(0)->setCellValue('J539', "Tanggal 30 Mei 216");//447

		$excel->setActiveSheetIndex(0)->setCellValue('B540', "    5 Menjadi Kepala Laboratorium Rekayasa Perangkat Lunak");
		$excel->setActiveSheetIndex(0)->setCellValue('E540', "Smstr Ganjil ");
		$excel->setActiveSheetIndex(0)->setCellValue('F540', "1 Semester");
		$excel->setActiveSheetIndex(0)->setCellValue('G540', "1.0");
		$excel->setActiveSheetIndex(0)->setCellValue('H540', "3.0");
		$excel->setActiveSheetIndex(0)->setCellValue('I540', "3.0");
		$excel->setActiveSheetIndex(0)->setCellValue('J540', "SK Rektor Nomor");//448

		$excel->setActiveSheetIndex(0)->setCellValue('E541', "2017/2018");
		$excel->setActiveSheetIndex(0)->setCellValue('J541', "646/UN26/KP/2016");//449
		$excel->setActiveSheetIndex(0)->setCellValue('J542', "Tanggal 30 Mei 216");//450

		$excel->setActiveSheetIndex(0)->setCellValue('B543', "    6 Menjadi Kepala Laboratorium Rekayasa Perangkat Lunak");
		$excel->setActiveSheetIndex(0)->setCellValue('E543', "Smstr Genap");
		$excel->setActiveSheetIndex(0)->setCellValue('F543', "1 Semester");
		$excel->setActiveSheetIndex(0)->setCellValue('G543', "1.0");
		$excel->setActiveSheetIndex(0)->setCellValue('H543', "3.0");
		$excel->setActiveSheetIndex(0)->setCellValue('I543', "3.0");
		$excel->setActiveSheetIndex(0)->setCellValue('J543', "SK Rektor Nomor");//451

		$excel->setActiveSheetIndex(0)->setCellValue('E544', "2017/2018");
		$excel->setActiveSheetIndex(0)->setCellValue('J544', "646/UN26/KP/2016");//452
		$excel->setActiveSheetIndex(0)->setCellValue('J545', "Tanggal 30 Mei 216");//453

		$excel->setActiveSheetIndex(0)->setCellValue('B546', "    7 Menjadi Kepala Laboratorium Rekayasa Perangkat Lunak");
		$excel->setActiveSheetIndex(0)->setCellValue('E546', "Smstr Genap");
		$excel->setActiveSheetIndex(0)->setCellValue('F546', "1 Semester");
		$excel->setActiveSheetIndex(0)->setCellValue('G546', "1.0");
		$excel->setActiveSheetIndex(0)->setCellValue('H546', "3.0");
		$excel->setActiveSheetIndex(0)->setCellValue('I546', "3.0");
		$excel->setActiveSheetIndex(0)->setCellValue('J546', "SK Rektor Nomor");//451

		$excel->setActiveSheetIndex(0)->setCellValue('E547', "2018/2019");
		$excel->setActiveSheetIndex(0)->setCellValue('J547', "646/UN26/KP/2016");//452
		$excel->setActiveSheetIndex(0)->setCellValue('J548', "Tanggal 30 Mei 216");

		$excel->setActiveSheetIndex(0)->setCellValue('H549', "Jumlah");
		$excel->setActiveSheetIndex(0)->setCellValue('I549', "21.0");//453
		$excel->getActiveSheet()->getStyle('H549:I549')->getFont()->setBold(TRUE);
		$excel->getActiveSheet()->getStyle('A549:L549')->applyFromArray($style_standar);

		$excel->setActiveSheetIndex(0)->setCellValue('A550', "K.");
		$excel->setActiveSheetIndex(0)->setCellValue('B550', "Membimbing Akademik Dosen yang lebih rendah jabatannya");
		$excel->getActiveSheet()->getStyle('A550:B550')->getFont()->setBold(TRUE);
		$excel->getActiveSheet()->getStyle('A550')->applyFromArray($style_standar);
		$excel->getActiveSheet()->getStyle('B550:L550')->applyFromArray($style_standar);//454
//cek lagi
		$excel->getActiveSheet()->getStyle('J551:L569')->applyFromArray($style_standar);

		$excel->setActiveSheetIndex(0)->setCellValue('B551', "1 Pembimbing pencangkokan");//455
		$excel->setActiveSheetIndex(0)->setCellValue('J552', "II.K.1");//456
		$excel->setActiveSheetIndex(0)->setCellValue('B553', "2 Reguler");//457
		$excel->setActiveSheetIndex(0)->setCellValue('J554', "II.K.2");
		$excel->getActiveSheet()->getStyle('B551:L555')->applyFromArray($style_standar);

		$excel->setActiveSheetIndex(0)->setCellValue('A556', "L.");
		$excel->setActiveSheetIndex(0)->setCellValue('B556', "Melaksanakan kegiatan Detasering dan pencangkokan Akademik Dosen");
		$excel->getActiveSheet()->getStyle('A556:B556')->getFont()->setBold(TRUE);
		$excel->getActiveSheet()->getStyle('A556')->applyFromArray($style_standar);
		$excel->getActiveSheet()->getStyle('B556:L556')->applyFromArray($style_standar);//458

		$excel->setActiveSheetIndex(0)->setCellValue('B557', "1 Detasering");//459
		$excel->setActiveSheetIndex(0)->setCellValue('H558', "Jumlah");
		$excel->setActiveSheetIndex(0)->setCellValue('I558', "0.0");
		$excel->getActiveSheet()->getStyle('H558:I558')->getFont()->setBold(TRUE);
		$excel->getActiveSheet()->getStyle('A558:L558')->applyFromArray($style_standar);
		$excel->setActiveSheetIndex(0)->setCellValue('J558', "II.L.1");//460

		$excel->setActiveSheetIndex(0)->setCellValue('B559', "2 Pencangkokan");//461
		$excel->setActiveSheetIndex(0)->setCellValue('H560', "Jumlah");
		$excel->setActiveSheetIndex(0)->setCellValue('I560', "0.0");
		$excel->getActiveSheet()->getStyle('H560:I560')->getFont()->setBold(TRUE);
		$excel->getActiveSheet()->getStyle('A560:L560')->applyFromArray($style_standar);
		$excel->setActiveSheetIndex(0)->setCellValue('J560', "II.L.2");//462
		$excel->getActiveSheet()->getStyle('B560:L560')->applyFromArray($style_standar);

		$excel->setActiveSheetIndex(0)->setCellValue('A561', "M.");
		$excel->setActiveSheetIndex(0)->setCellValue('B561', "Melakukan kegiatan pengembangan diri untuk meningkatkan kompetensi");
		$excel->getActiveSheet()->getStyle('A561:B561')->getFont()->setBold(TRUE);
		$excel->getActiveSheet()->getStyle('A561')->applyFromArray($style_standar);
		$excel->getActiveSheet()->getStyle('B561:L561')->applyFromArray($style_standar);//463

		$excel->setActiveSheetIndex(0)->setCellValue('B562', "1 Lamanya lebih dari 960 jam");
		$excel->setActiveSheetIndex(0)->setCellValue('J562', "II.M.1");//464
		$excel->setActiveSheetIndex(0)->setCellValue('B563', "2 Lamanya 641-960 jam");
		$excel->setActiveSheetIndex(0)->setCellValue('J563', "II.M.2");//465
		$excel->setActiveSheetIndex(0)->setCellValue('B564', "3 Lamanya 481-640 jam");
		$excel->setActiveSheetIndex(0)->setCellValue('J564', "II.M.3");//466
		$excel->setActiveSheetIndex(0)->setCellValue('B565', "4 Lamanya 161-480 jam");
		$excel->setActiveSheetIndex(0)->setCellValue('J565', "II.M.4");//467
		$excel->setActiveSheetIndex(0)->setCellValue('B566', "5 Lamanya 81-160 jam");
		$excel->setActiveSheetIndex(0)->setCellValue('J566', "II.M.5");//468
		$excel->setActiveSheetIndex(0)->setCellValue('B567', "6 Lamanya 31-80 jam");
		$excel->setActiveSheetIndex(0)->setCellValue('J567', "II.M.6");//469
		$excel->setActiveSheetIndex(0)->setCellValue('B568', "7 Lamanya 10-30 jam");
		$excel->setActiveSheetIndex(0)->setCellValue('J568', "II.M.7");//470
			//balik cuy ke baris 1880
		$excel->getActiveSheet()->getStyle('B562:L569')->applyFromArray($style_standar);

		$excel->setActiveSheetIndex(0)->setCellValue('H570', "Jumlah");
		$excel->setActiveSheetIndex(0)->setCellValue('I570', "0.0");
		$excel->getActiveSheet()->getStyle('H570:I570')->getFont()->setBold(TRUE);
		$excel->getActiveSheet()->getStyle('A570:L570')->applyFromArray($style_standar);//471

		$excel->setActiveSheetIndex(0)->setCellValue('H571', "Total pendidikan");
		$excel->setActiveSheetIndex(0)->setCellValue('I571', "204.0");
		$excel->getActiveSheet()->getStyle('H571:I571')->getFont()->setBold(TRUE);
		$excel->getActiveSheet()->getStyle('A571:L571')->applyFromArray($style_standar);//472

		$excel->setActiveSheetIndex(0)->setCellValue('A573', "Demikian Pernyataan ini dibuat untuk dapat dipergunakan sebagaimana mestinya.");//473
		$excel->setActiveSheetIndex(0)->setCellValue('J575', "Bandar Lampung,  31 Juli 2019");//474
		$excel->setActiveSheetIndex(0)->setCellValue('J576', "Ketua Jurusan Ilmu Komputer");//475
		$excel->setActiveSheetIndex(0)->setCellValue('J580', "Dr.Ir. Kurnia Muludi, M.S.Sc");//476
		$excel->setActiveSheetIndex(0)->setCellValue('J581', "NIP. 19640616 198902 1 001");//477








		// Set height semua kolom menjadi auto (mengikuti height isi dari kolommnya, jadi otomatis)
		$excel->getActiveSheet()->getDefaultRowDimension()->setRowHeight(-1);

		// Set orientasi kertas jadi LANDSCAPE
		$excel->getActiveSheet()->getPageSetup()->setOrientation(PHPExcel_Worksheet_PageSetup::ORIENTATION_LANDSCAPE);

		// Set judul file excel nya
		$excel->getActiveSheet(0)->setTitle("Pendidikan");
		$excel->setActiveSheetIndex(0);

		// Proses file excel
		header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
		header('Content-Disposition: attachment; filename="Pendidikan.xlsx"'); // Set nama file excel nya
		header('Cache-Control: max-age=0');

		$write = PHPExcel_IOFactory::createWriter($excel, 'Excel2007');
		$write->save('php://output');
	}
}
