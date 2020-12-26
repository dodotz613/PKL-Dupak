<?php
defined('BASEPATH') OR exit('No direct script access allowed');

class Pengabdian extends CI_Controller {

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
		$this->load->view('pengabdian', $data);

		}


	public function export(){
			// Load plugin PHPExcel nya
		include APPPATH.'third_party/PHPExcel/PHPExcel.php';

			// Panggil class PHPExcel nya
		$excel = new PHPExcel();

		// Settingan awal fil excel
		$excel->getProperties()->setCreator('Sulung')
							   ->setLastModifiedBy('Sulung')
							   ->setTitle("Pengabdian")
							   ->setSubject("Dupak")
							   ->setDescription("Laporan")
							   ->setKeywords("Pengabdian");

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
		$excel->setActiveSheetIndex(0)->setCellValue('A2', "MELAKSANAKAN PENGABDIAN KEPADA MASYARAKAT");
		$excel->getActiveSheet()->mergeCells('A1:R1');
		$excel->getActiveSheet()->mergeCells('A2:R2');
		$excel->getActiveSheet()->getStyle('A1:A2')->getFont()->setSize(11);
		$excel->getActiveSheet()->getStyle('A1:A2')->getFont()->setBold(TRUE);
		$excel->getActiveSheet()->getStyle('A1:A2')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);

		$excel->setActiveSheetIndex(0)->setCellValue('B4', "Yang bertanda tangan di bawah ini : ");
		$excel->setActiveSheetIndex(0)->setCellValue('B6', "Nama ");
		$excel->setActiveSheetIndex(0)->setCellValue('J6', ":");
		$excel->setActiveSheetIndex(0)->setCellValue('B7', "NIP");
		$excel->setActiveSheetIndex(0)->setCellValue('J7', ":");
		$excel->setActiveSheetIndex(0)->setCellValue('B8', "Pangkat");
		$excel->setActiveSheetIndex(0)->setCellValue('J8', ":");
		$excel->setActiveSheetIndex(0)->setCellValue('B9', "Golongan");
		$excel->setActiveSheetIndex(0)->setCellValue('J9', ":");
		$excel->setActiveSheetIndex(0)->setCellValue('B10', "Jabatan");
		$excel->setActiveSheetIndex(0)->setCellValue('J10', ":");
		$excel->setActiveSheetIndex(0)->setCellValue('B11', "Unit Kerja");
		$excel->setActiveSheetIndex(0)->setCellValue('J11', ":");
		$excel->setActiveSheetIndex(0)->setCellValue('B12', "Menyatakan ");
		$excel->setActiveSheetIndex(0)->setCellValue('B13', "Nama");
		$excel->setActiveSheetIndex(0)->setCellValue('J13', ":");
		$excel->setActiveSheetIndex(0)->setCellValue('B14', "NIP");
		$excel->setActiveSheetIndex(0)->setCellValue('J14', ":");
		$excel->setActiveSheetIndex(0)->setCellValue('B15', "Pangkat");
		$excel->setActiveSheetIndex(0)->setCellValue('J15', ":");
		$excel->setActiveSheetIndex(0)->setCellValue('B16', "Golongan");
		$excel->setActiveSheetIndex(0)->setCellValue('J16', ":");
		$excel->setActiveSheetIndex(0)->setCellValue('B17', "Jabatan");
		$excel->setActiveSheetIndex(0)->setCellValue('J17', ":");
		$excel->setActiveSheetIndex(0)->setCellValue('B18', "Unit Kerja");
		$excel->setActiveSheetIndex(0)->setCellValue('J18', ":");
		$excel->setActiveSheetIndex(0)->setCellValue('B20', "Telah menyatakan pengabdian masyarakat sebagai berikut : ");

		$data_dosen = $this->Dosen->view();

		foreach($data_dosen as $data) {

			$excel->setActiveSheetIndex(0)->setCellValue('K6', $data->nama);
			$excel->setActiveSheetIndex(0)->setCellValue('K7', $data->nip);
			$excel->setActiveSheetIndex(0)->setCellValue('K8', $data->pangkat);
			$excel->setActiveSheetIndex(0)->setCellValue('K9', $data->golongan);
			$excel->setActiveSheetIndex(0)->setCellValue('K10', $data->jabatan);
			$excel->setActiveSheetIndex(0)->setCellValue('K11', $data->unit_kerja);
		}

		$dosen_penunjang = $this->Lektor->view();

		foreach($dosen_penunjang as $data) {

			$excel->setActiveSheetIndex(0)->setCellValue('K13', $data->nama);
			$excel->setActiveSheetIndex(0)->setCellValue('K14', $data->nip);
			$excel->setActiveSheetIndex(0)->setCellValue('K15', $data->pangkat);
			$excel->setActiveSheetIndex(0)->setCellValue('K16', $data->golongan);
			$excel->setActiveSheetIndex(0)->setCellValue('K17', $data->jabatan);
			$excel->setActiveSheetIndex(0)->setCellValue('K18', $data->unit_kerja);
		}




		$excel->setActiveSheetIndex(0)->setCellValue('A22', "No");
		$excel->getActiveSheet()->getStyle('A22')->getAlignment()->setWrapText(TRUE);
		$excel->setActiveSheetIndex(0)->setCellValue('B22', "Uraian Kegiatan");
		$excel->getActiveSheet()->mergeCells('B22:K22');
		$excel->getActiveSheet()->getStyle('B22')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
		$excel->getActiveSheet()->getStyle('B22')->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
		$excel->setActiveSheetIndex(0)->setCellValue('L22', "Tanggal");
		$excel->setActiveSheetIndex(0)->setCellValue('M22', "Satuan Hasil");
		$excel->getActiveSheet()->getStyle('M22')->getAlignment()->setWrapText(TRUE);
		$excel->setActiveSheetIndex(0)->setCellValue('N22', "Jumlah Volume Kegiatan");
		$excel->getActiveSheet()->getStyle('N22')->getAlignment()->setWrapText(TRUE);
		$excel->setActiveSheetIndex(0)->setCellValue('O22', "Angka Kredit");
		$excel->getActiveSheet()->getStyle('O22')->getAlignment()->setWrapText(TRUE);
		$excel->setActiveSheetIndex(0)->setCellValue('P22', "Jumlah Angka Kredit");
		$excel->getActiveSheet()->getStyle('P22')->getAlignment()->setWrapText(TRUE);
		$excel->setActiveSheetIndex(0)->setCellValue('Q22', "Keterangan/Bukti Fisik");
		$excel->getActiveSheet()->mergeCells('Q22:R22');
		$excel->getActiveSheet()->getStyle('Q22')->getAlignment()->setWrapText(TRUE);

		$excel->getActiveSheet()->getStyle('A22:A74')->applyFromArray($style_col);
		$excel->getActiveSheet()->getStyle('L22:L74')->applyFromArray($style_col);
		$excel->getActiveSheet()->getStyle('M22:M74')->applyFromArray($style_col);
		$excel->getActiveSheet()->getStyle('N22:N74')->applyFromArray($style_col);
		$excel->getActiveSheet()->getStyle('O22:O74')->applyFromArray($style_col);
		$excel->getActiveSheet()->getStyle('P22:P74')->applyFromArray($style_col);
		$excel->getActiveSheet()->getStyle('R22:R74')->applyFromArray($style_col);
		$excel->getActiveSheet()->getStyle('A22')->applyFromArray($style_standar);
		$excel->getActiveSheet()->getStyle('B22')->applyFromArray($style_standar);
		$excel->getActiveSheet()->getStyle('C22')->applyFromArray($style_standar);
		$excel->getActiveSheet()->getStyle('D22')->applyFromArray($style_standar);
		$excel->getActiveSheet()->getStyle('E22')->applyFromArray($style_standar);
		$excel->getActiveSheet()->getStyle('F22')->applyFromArray($style_standar);
		$excel->getActiveSheet()->getStyle('G22')->applyFromArray($style_standar);
		$excel->getActiveSheet()->getStyle('H22')->applyFromArray($style_standar);
		$excel->getActiveSheet()->getStyle('I22')->applyFromArray($style_standar);
		$excel->getActiveSheet()->getStyle('J22')->applyFromArray($style_standar);
		$excel->getActiveSheet()->getStyle('K22')->applyFromArray($style_standar);
		$excel->getActiveSheet()->getStyle('L22:Q22')->applyFromArray($style_standar);
		$excel->getActiveSheet()->getStyle('R22')->applyFromArray($style_standar);

		$excel->setActiveSheetIndex(0)->setCellValue('A23', "(1)");
		$excel->setActiveSheetIndex(0)->setCellValue('B23', "(2)");
		$excel->getActiveSheet()->getStyle('B23')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
		$excel->getActiveSheet()->mergeCells('B23:K23');
		$excel->getActiveSheet()->getStyle('A23:Q23')->applyFromArray($style_standar);
		$excel->setActiveSheetIndex(0)->setCellValue('L23', "(3)");
		$excel->setActiveSheetIndex(0)->setCellValue('M23', "(4)");
		$excel->setActiveSheetIndex(0)->setCellValue('N23', "(5)");
		$excel->setActiveSheetIndex(0)->setCellValue('O23', "(6)");
		$excel->setActiveSheetIndex(0)->setCellValue('P23', "(7)");
		$excel->setActiveSheetIndex(0)->setCellValue('R23', "(8)");
		$excel->getActiveSheet()->getStyle('R23')->getAlignment()->setWrapText(TRUE);
		$excel->getActiveSheet()->getStyle('R23')->applyFromArray($style_standar);

		$excel->setActiveSheetIndex(0)->setCellValue('A24', "IV.");
		$excel->getActiveSheet()->getStyle('A24')->getFont()->setBold(TRUE);
		$excel->getActiveSheet()->getStyle('A24')->getFont()->setSize(11);
		$excel->setActiveSheetIndex(0)->setCellValue('B24', "MELAKSANAKAN PENGABDIAN KEPADA MASYARAKAT");
		$excel->getActiveSheet()->mergeCells('B24:G24');
		$excel->getActiveSheet()->getStyle('B24')->getFont()->setBold(TRUE);
		$excel->getActiveSheet()->getStyle('B24')->getFont()->setSize(11);
		$excel->setActiveSheetIndex(0)->setCellValue('B25', "A.");
		$excel->setActiveSheetIndex(0)->setCellValue('C25', "Menduduki jabatan pimpinan.");
		$excel->getActiveSheet()->mergeCells('C25:K25');
		$excel->setActiveSheetIndex(0)->setCellValue('D26', "Menduduki jabatan pimpinan dan lembaga");
		$excel->getActiveSheet()->mergeCells('D26:K26');
		$excel->setActiveSheetIndex(0)->setCellValue('O27', "Jumlah");
		$excel->setActiveSheetIndex(0)->setCellValue('P27', "0,00");
		$excel->getActiveSheet()->getStyle('O27:P27')->getFont()->setBold(TRUE);
		$excel->getActiveSheet()->getStyle('O27:P27')->getFont()->setSize(11);
		$excel->getActiveSheet()->getStyle('A27:R27')->applyFromArray($style_standar);

		$excel->setActiveSheetIndex(0)->setCellValue('B28', "B.");
		$excel->setActiveSheetIndex(0)->setCellValue('C28', "Melaksanakan pengembangan hasil pendidikan dan penelitian.");
		$excel->getActiveSheet()->mergeCells('C28:K28');
		$excel->setActiveSheetIndex(0)->setCellValue('D29', "Melaksanakan pengembangan hasil pendidikan dan penelitian");
		$excel->getActiveSheet()->mergeCells('D29:K29');
		$excel->setActiveSheetIndex(0)->setCellValue('O30', "Jumlah");
		$excel->setActiveSheetIndex(0)->setCellValue('P30', "0,00");
		$excel->getActiveSheet()->getStyle('O30:P30')->getFont()->setBold(TRUE);
		$excel->getActiveSheet()->getStyle('O30:P30')->getFont()->setSize(11);
		$excel->getActiveSheet()->getStyle('A30:R30')->applyFromArray($style_standar);

		$excel->setActiveSheetIndex(0)->setCellValue('B31', "C.");
		$excel->setActiveSheetIndex(0)->setCellValue('C31', "Memberi latihan/penyuluhan/penataran/ceramah kepada masyarakat.");
		$excel->getActiveSheet()->mergeCells('C31:K31');
		$excel->setActiveSheetIndex(0)->setCellValue('C32', "1.");
		$excel->setActiveSheetIndex(0)->setCellValue('D32', "Terjadwal/terprogram");
		$excel->getActiveSheet()->mergeCells('D32:F32');
		$excel->setActiveSheetIndex(0)->setCellValue('D33', "a.");
		$excel->setActiveSheetIndex(0)->setCellValue('E33', "Dalam satu semester atau lebih");
		$excel->getActiveSheet()->mergeCells('E33:K33');
		$excel->setActiveSheetIndex(0)->setCellValue('E34', "1) Tingkat Internasional");
		$excel->setActiveSheetIndex(0)->setCellValue('O35', "Jumlah");
		$excel->setActiveSheetIndex(0)->setCellValue('P35', "0,00");
		$excel->getActiveSheet()->getStyle('O35:P35')->getFont()->setBold(TRUE);
		$excel->getActiveSheet()->getStyle('O35:P35')->getFont()->setSize(11);
		$excel->getActiveSheet()->getStyle('A35:R35')->applyFromArray($style_standar);

		$excel->setActiveSheetIndex(0)->setCellValue('E36', "2) Tingkat Nasional");
		$excel->setActiveSheetIndex(0)->setCellValue('O37', "Jumlah");
		$excel->setActiveSheetIndex(0)->setCellValue('P37', "0,00");
		$excel->getActiveSheet()->getStyle('O37:P37')->getFont()->setBold(TRUE);
		$excel->getActiveSheet()->getStyle('O37:P37')->getFont()->setSize(11);
		$excel->getActiveSheet()->getStyle('A37:R37')->applyFromArray($style_standar);

		$excel->setActiveSheetIndex(0)->setCellValue('E38', "3) Tingkat Lokal");
		$excel->setActiveSheetIndex(0)->setCellValue('O39', "Jumlah");
		$excel->setActiveSheetIndex(0)->setCellValue('P39', "0,00");
		$excel->getActiveSheet()->getStyle('O39:P39')->getFont()->setBold(TRUE);
		$excel->getActiveSheet()->getStyle('O39:P39')->getFont()->setSize(11);
		$excel->getActiveSheet()->getStyle('A39:R39')->applyFromArray($style_standar);

		$excel->setActiveSheetIndex(0)->setCellValue('D40', "b.");
		$excel->setActiveSheetIndex(0)->setCellValue('E40', "Kurang dari satu semester dan minimal satu bulan.");
		$excel->getActiveSheet()->mergeCells('E40:K40');
		$excel->setActiveSheetIndex(0)->setCellValue('E41', "1) Tingkat Internasional");
		$excel->setActiveSheetIndex(0)->setCellValue('O42', "Jumlah");
		$excel->setActiveSheetIndex(0)->setCellValue('P42', "0,00");
		$excel->getActiveSheet()->getStyle('O42:P42')->getFont()->setBold(TRUE);
		$excel->getActiveSheet()->getStyle('O42:P42')->getFont()->setSize(11);
		$excel->getActiveSheet()->getStyle('A42:R42')->applyFromArray($style_standar);

		$excel->setActiveSheetIndex(0)->setCellValue('E43', "2) Tingkat Nasional");
		$excel->setActiveSheetIndex(0)->setCellValue('O44', "Jumlah");
		$excel->setActiveSheetIndex(0)->setCellValue('P44', "0,00");
		$excel->getActiveSheet()->getStyle('O44:P44')->getFont()->setBold(TRUE);
		$excel->getActiveSheet()->getStyle('O44:P44')->getFont()->setSize(11);
		$excel->getActiveSheet()->getStyle('A44:R44')->applyFromArray($style_standar);

		$excel->setActiveSheetIndex(0)->setCellValue('E45', "3) Tingkat Lokal");
		$excel->setActiveSheetIndex(0)->setCellValue('O46', "Jumlah");
		$excel->setActiveSheetIndex(0)->setCellValue('P46', "0,00");
		$excel->getActiveSheet()->getStyle('O46:P46')->getFont()->setBold(TRUE);
		$excel->getActiveSheet()->getStyle('O46:P46')->getFont()->setSize(11);
		$excel->getActiveSheet()->getStyle('A46:R46')->applyFromArray($style_standar);

		$excel->setActiveSheetIndex(0)->setCellValue('C47', "2.");
		$excel->setActiveSheetIndex(0)->setCellValue('D47', "Insidental :");
		$excel->getActiveSheet()->mergeCells('D47:K47');
		$excel->setActiveSheetIndex(0)->setCellValue('C48', "1)");
		$excel->setActiveSheetIndex(0)->setCellValue('D48', "Pelatihan Desain Grafis untuk Usaha Kecil Menengah Desa Wawasan Kecamatan Tanjung Sari Kabupaten Lampung Selatan");
		$excel->setActiveSheetIndex(0)->setCellValue('L48', "23 Agst 2018");
		$excel->setActiveSheetIndex(0)->setCellValue('M48', "Laporan");
		$excel->setActiveSheetIndex(0)->setCellValue('N48', "1,0");
		$excel->setActiveSheetIndex(0)->setCellValue('O48', "1,0");
		$excel->setActiveSheetIndex(0)->setCellValue('P48', "1,0");
		$excel->setActiveSheetIndex(0)->setCellValue('R48', "IV.C.2");
		$excel->setActiveSheetIndex(0)->setCellValue('M49', "Kegiatan");
		$excel->setActiveSheetIndex(0)->setCellValue('R49', "Laporan Kegiatan");

		$excel->setActiveSheetIndex(0)->setCellValue('C53', "2)");
		$excel->setActiveSheetIndex(0)->setCellValue('D53', "Pelatihan Adobe Photosop dan Corel Draw untuk pembuatan alat promosi Sekolah bagi Guru-Guru SMS di Bandar Lampung");
		$excel->setActiveSheetIndex(0)->setCellValue('L53', "29 Nov 2014");
		$excel->setActiveSheetIndex(0)->setCellValue('M53', "Laporan");
		$excel->setActiveSheetIndex(0)->setCellValue('N53', "1,0");
		$excel->setActiveSheetIndex(0)->setCellValue('O53', "1,0");
		$excel->setActiveSheetIndex(0)->setCellValue('P53', "1,0");
		$excel->setActiveSheetIndex(0)->setCellValue('R53', "IV.C.2");
		$excel->setActiveSheetIndex(0)->setCellValue('M54', "Kegiatan");
		$excel->setActiveSheetIndex(0)->setCellValue('R54', "Laporan Kegiatan");

		$excel->setActiveSheetIndex(0)->setCellValue('C58', "3)");
		$excel->setActiveSheetIndex(0)->setCellValue('D58', "Penerapan Media Pembelajaran Interaktif Pengenalan Komputer di SDN 1 Kupang Teba Kota Bandar Lampung");
		$excel->setActiveSheetIndex(0)->setCellValue('L58', "22 Nov 2014");
		$excel->setActiveSheetIndex(0)->setCellValue('M58', "Laporan");
		$excel->setActiveSheetIndex(0)->setCellValue('N58', "1,0");
		$excel->setActiveSheetIndex(0)->setCellValue('O58', "1,0");
		$excel->setActiveSheetIndex(0)->setCellValue('P58', "1,0");
		$excel->setActiveSheetIndex(0)->setCellValue('R58', "IV.C.2");
		$excel->setActiveSheetIndex(0)->setCellValue('M59', "Kegiatan");
		$excel->setActiveSheetIndex(0)->setCellValue('R59', "Laporan Kegiatan");

		$excel->setActiveSheetIndex(0)->setCellValue('C62', "4)");
		$excel->setActiveSheetIndex(0)->setCellValue('D62', "Implementasi Sistem Informasi Akademik di SMUN 1 Gedong Tataan Kabupaten Pesawaran");
		$excel->setActiveSheetIndex(0)->setCellValue('L62', "8-9 Okt 2013");
		$excel->setActiveSheetIndex(0)->setCellValue('M62', "Laporan");
		$excel->setActiveSheetIndex(0)->setCellValue('N62', "1,0");
		$excel->setActiveSheetIndex(0)->setCellValue('O62', "1,0");
		$excel->setActiveSheetIndex(0)->setCellValue('P62', "1,0");
		$excel->setActiveSheetIndex(0)->setCellValue('R62', "IV.C.2");
		$excel->setActiveSheetIndex(0)->setCellValue('M63', "Kegiatan");
		$excel->setActiveSheetIndex(0)->setCellValue('R63', "Laporan Kegiatan");

		$excel->setActiveSheetIndex(0)->setCellValue('O66', "Jumlah");
		$excel->setActiveSheetIndex(0)->setCellValue('P66', "4,00");
		$excel->getActiveSheet()->getStyle('O66:P66')->getFont()->setBold(TRUE);
		$excel->getActiveSheet()->getStyle('O66:P66')->getFont()->setSize(11);
		$excel->getActiveSheet()->getStyle('A66:R66')->applyFromArray($style_standar);


		$excel->setActiveSheetIndex(0)->setCellValue('B67', "D.");
		$excel->setActiveSheetIndex(0)->setCellValue('C67', "Memberi pelayanan kepada masyarakat atau kegiatan lain yang menunjang pelaksanaan tugas umum pemerintah dan pembangunan.");
		$excel->getActiveSheet()->mergeCells('C67:K67');
		$excel->getActiveSheet()->getStyle('C67')->getAlignment()->setWrapText(TRUE);
		$excel->setActiveSheetIndex(0)->setCellValue('C68', "1.");
		$excel->setActiveSheetIndex(0)->setCellValue('D68', "Berdasarkan bidang keahlian.");
		$excel->getActiveSheet()->mergeCells('D68:K68');
		$excel->setActiveSheetIndex(0)->setCellValue('C69', "2.");
		$excel->setActiveSheetIndex(0)->setCellValue('D69', "Berdasarkan penugasan lembaga perguruan tinggi.");
		$excel->getActiveSheet()->mergeCells('D69:K69');
		$excel->setActiveSheetIndex(0)->setCellValue('C70', "3.");
		$excel->setActiveSheetIndex(0)->setCellValue('D70', "Berdasarkan fungsi/jabatan.");
		$excel->getActiveSheet()->mergeCells('D70:K70');

		$excel->setActiveSheetIndex(0)->setCellValue('B71', "E.");
		$excel->setActiveSheetIndex(0)->setCellValue('C71', "Membuat/menulis karya pengabdian.");
		$excel->getActiveSheet()->mergeCells('C71:K71');
		$excel->setActiveSheetIndex(0)->setCellValue('D72', "Membuat/menulis karya pengabdian pada masyarakat yang tidak dipublikasikan.");
		$excel->getActiveSheet()->mergeCells('D72:K72');
		$excel->setActiveSheetIndex(0)->setCellValue('O73', "Jumlah");
		$excel->setActiveSheetIndex(0)->setCellValue('P73', "0,00");
		$excel->getActiveSheet()->getStyle('O73')->getFont()->setBold(TRUE);
		$excel->getActiveSheet()->getStyle('O73')->getFont()->setSize(11);
		$excel->getActiveSheet()->getStyle('A73:R73')->applyFromArray($style_standar);
		$excel->setActiveSheetIndex(0)->setCellValue('A74', "Jumlah Pengabdian kepada Masyarakat");
		$excel->getActiveSheet()->mergeCells('A74:O74');
		$excel->setActiveSheetIndex(0)->setCellValue('P74', "4,00");
		$excel->getActiveSheet()->getStyle('A74:P74')->getFont()->setBold(TRUE);
		$excel->getActiveSheet()->getStyle('A74:P74')->getFont()->setSize(11);
		$excel->getActiveSheet()->getStyle('A74:R74')->applyFromArray($style_standar);

		$excel->setActiveSheetIndex(0)->setCellValue('O76', "Bandar Lampung,  31 Juli 2019");
		$excel->setActiveSheetIndex(0)->setCellValue('O77', "Ketua Jurusan Ilmu Komputer");
		$excel->setActiveSheetIndex(0)->setCellValue('O81', "Dr.Ir. Kurnia Muludi, M.S.Sc");
		$excel->setActiveSheetIndex(0)->setCellValue('O82', "NIP. 19640616 198902 1 001");



		// Set orientasi kertas jadi LANDSCAPE
		$excel->getActiveSheet()->getPageSetup()->setOrientation(PHPExcel_Worksheet_PageSetup::ORIENTATION_LANDSCAPE);

		// Set judul file excel nya
		$excel->getActiveSheet(0)->setTitle("Pengabdian");
		$excel->setActiveSheetIndex(0);

		// Proses file excel
		header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
		header('Content-Disposition: attachment; filename="Pengabdian.xlsx"'); // Set nama file excel nya
		header('Cache-Control: max-age=0');

		$write = PHPExcel_IOFactory::createWriter($excel, 'Excel2007');
		$write->save('php://output');

	}

}
