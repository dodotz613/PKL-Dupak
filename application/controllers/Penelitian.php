<?php
defined('BASEPATH') OR exit('No direct script access allowed');

class Penelitian extends CI_Controller {

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
		$this->load->view('penelitian', $data);

		}


	public function export(){
			// Load plugin PHPExcel nya
		include APPPATH.'third_party/PHPExcel/PHPExcel.php';

			// Panggil class PHPExcel nya
		$excel = new PHPExcel();

		// Settingan awal fil excel
		$excel->getProperties()->setCreator('Awal')
							   ->setLastModifiedBy('Awal')
							   ->setTitle("Penelitian")
							   ->setSubject("Dupak")
							   ->setDescription("Laporan")
							   ->setKeywords("Penelitian");

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
    		$excel->setActiveSheetIndex(0)->setCellValue('A2', "MELAKSANAKAN PENELITIAN");
        $excel->getActiveSheet()->mergeCells('A1:R1');
        $excel->getActiveSheet()->mergeCells('A2:R2');
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

    		$excel->getActiveSheet()->getStyle('A22:A373')->applyFromArray($style_col);
    		$excel->getActiveSheet()->getStyle('L22:L373')->applyFromArray($style_col);
    		$excel->getActiveSheet()->getStyle('M22:M373')->applyFromArray($style_col);
    		$excel->getActiveSheet()->getStyle('N22:N373')->applyFromArray($style_col);
    		$excel->getActiveSheet()->getStyle('O22:O373')->applyFromArray($style_col);
    		$excel->getActiveSheet()->getStyle('P22:P373')->applyFromArray($style_col);
    		$excel->getActiveSheet()->getStyle('R22:R373')->applyFromArray($style_col);
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

        $excel->setActiveSheetIndex(0)->setCellValue('A24', "II.");
    		$excel->getActiveSheet()->getStyle('A24')->getFont()->setBold(TRUE);
    		$excel->setActiveSheetIndex(0)->setCellValue('B24', "MELAKSANAKAN PENELITIAN");
    		$excel->getActiveSheet()->mergeCells('B24:G24');
    		$excel->getActiveSheet()->getStyle('B24')->getFont()->setBold(TRUE);
    		$excel->setActiveSheetIndex(0)->setCellValue('B25', "A.");
        $excel->getActiveSheet()->getStyle('B25')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
    		$excel->setActiveSheetIndex(0)->setCellValue('C25', "Menghasilkan karya ilmiah .");
    		$excel->getActiveSheet()->mergeCells('C25:K25');
        $excel->setActiveSheetIndex(0)->setCellValue('C26', "1");
        $excel->getActiveSheet()->getStyle('C26')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
    		$excel->setActiveSheetIndex(0)->setCellValue('D26', "Hasil penelitian atau pemikiran yang dipublikasikan");
    		$excel->getActiveSheet()->mergeCells('D26:K26');
        $excel->setActiveSheetIndex(0)->setCellValue('D27', "a");
        $excel->getActiveSheet()->getStyle('D27')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
        $excel->setActiveSheetIndex(0)->setCellValue('E27', "Dalam bentuk:");
        $excel->getActiveSheet()->mergeCells('E27:I27');
        $excel->setActiveSheetIndex(0)->setCellValue('E28', "1)");
				$excel->getActiveSheet()->getStyle('E28')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
        $excel->setActiveSheetIndex(0)->setCellValue('F28', "Monograf");
        $excel->getActiveSheet()->mergeCells('F28:I28');
    		$excel->setActiveSheetIndex(0)->setCellValue('O28', "Jumlah");
    		$excel->setActiveSheetIndex(0)->setCellValue('P28', "0,00");
    		$excel->getActiveSheet()->getStyle('O28:P28')->getFont()->setBold(TRUE);
    		$excel->getActiveSheet()->getStyle('A28:R28')->applyFromArray($style_standar);
        $excel->setActiveSheetIndex(0)->setCellValue('E29', "2)");
				$excel->getActiveSheet()->getStyle('E29')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
        $excel->setActiveSheetIndex(0)->setCellValue('F29', "Buku Referensi");
        $excel->setActiveSheetIndex(0)->setCellValue('O30', "Jumlah");
    		$excel->setActiveSheetIndex(0)->setCellValue('P30', "0,00");
    		$excel->getActiveSheet()->getStyle('O30:P30')->getFont()->setBold(TRUE);
    		$excel->getActiveSheet()->getStyle('A30:R30')->applyFromArray($style_standar);
        $excel->setActiveSheetIndex(0)->setCellValue('D31', "b");
        $excel->getActiveSheet()->getStyle('D31')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
        $excel->setActiveSheetIndex(0)->setCellValue('E31', "Jurnal ilmiah:");
        $excel->getActiveSheet()->mergeCells('E31:K31');
        $excel->setActiveSheetIndex(0)->setCellValue('E32', "1)");
        $excel->getActiveSheet()->getStyle('E32')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
        $excel->setActiveSheetIndex(0)->setCellValue('F32', "Internasional");
        $excel->getActiveSheet()->mergeCells('F32:K32');

        $excel->setActiveSheetIndex(0)->setCellValue('C33', "1)");
        $excel->setActiveSheetIndex(0)->setCellValue('D33', "International Journal of Advanced Computer Science and Applications (IJACSA)");
        $excel->setActiveSheetIndex(0)->setCellValue('L33', "April ");
        $excel->setActiveSheetIndex(0)->setCellValue('M33', "Jurnal");
        $excel->setActiveSheetIndex(0)->setCellValue('N33', "1,0");
        $excel->setActiveSheetIndex(0)->setCellValue('O33', "37,3");
        $excel->setActiveSheetIndex(0)->setCellValue('P33', "22,38");
        $excel->setActiveSheetIndex(0)->setCellValue('R33', "https://thesai.org/Downloads/Volume10No4/Paper_27-Comparative_Analysis_of_Cow_Disease_Diagnosis.pdf ");
        $excel->setActiveSheetIndex(0)->setCellValue('L34', "2019 ");
        $excel->setActiveSheetIndex(0)->setCellValue('M34', "Internasional");

        $excel->setActiveSheetIndex(0)->setCellValue('D36', "ISSN (Online) : 2156-5570");
        $excel->setActiveSheetIndex(0)->setCellValue('D37', "Vol. 10, Issue 4, PP 227-235, April 2019");
        $excel->setActiveSheetIndex(0)->setCellValue('D39', "Comparative Analysis of Cow Disease Diagnosis Expert System using Bayesian Network and Dempster-Shafer Method");
        $excel->setActiveSheetIndex(0)->setCellValue('R39', "Terdaftar di LPPM Unila");
        $excel->setActiveSheetIndex(0)->setCellValue('R40', "No. 167/J/B/I/FMIPA/2019");
        $excel->setActiveSheetIndex(0)->setCellValue('R41', "Tanggal 27Juni 2019");
        $excel->setActiveSheetIndex(0)->setCellValue('R43', "Aristoteles, Kusuma Adhianto, Rico Andrian, Yeni Nuhricha Sari");

        $excel->setActiveSheetIndex(0)->setCellValue('C46', "2)");
        $excel->setActiveSheetIndex(0)->setCellValue('D46', "International Journal of Advanced Computer Science and Applications (IJACSA)");
        $excel->setActiveSheetIndex(0)->setCellValue('L46', "November");
        $excel->setActiveSheetIndex(0)->setCellValue('M46', "Jurnal");
        $excel->setActiveSheetIndex(0)->setCellValue('N46', "1,0");
        $excel->setActiveSheetIndex(0)->setCellValue('R46', "https://thesai.org/Downloads/Volume8No11/Paper_21-Expert_System_of_Chili_Plant_Disease_Diagnosis.pdf ");
        $excel->setActiveSheetIndex(0)->setCellValue('L47', "2017");
        $excel->setActiveSheetIndex(0)->setCellValue('M47', "Internasional");

        $excel->setActiveSheetIndex(0)->setCellValue('D49', "ISSN (Online) : 2156-5570");
        $excel->setActiveSheetIndex(0)->setCellValue('D50', "Vol. 8, Issue 11, PP 164-168, November 2017");
        $excel->setActiveSheetIndex(0)->setCellValue('R51', "Terdaftar di LPPM Unila");
        $excel->setActiveSheetIndex(0)->setCellValue('D52', "Expert System of Chili Plant Disease Diagnosis using Forward Chaining Method on Android");
        $excel->setActiveSheetIndex(0)->setCellValue('R52', "No. 168/J/B/I/FMIPA/2019");
        $excel->setActiveSheetIndex(0)->setCellValue('R53', "Tanggal 27 Juni 2019");
        $excel->setActiveSheetIndex(0)->setCellValue('R55', "Aristoteles, Mita Fuljana, Joko Prasetyo, Kurnia Muludi");

        $excel->setActiveSheetIndex(0)->setCellValue('C58', "3)");
        $excel->setActiveSheetIndex(0)->setCellValue('D58', "ARPN Journal of Engineering and Applied Sciences");
        $excel->setActiveSheetIndex(0)->setCellValue('L58', "April");
        $excel->setActiveSheetIndex(0)->setCellValue('M58', "Jurnal");
        $excel->setActiveSheetIndex(0)->setCellValue('N58', "1,0");
        $excel->setActiveSheetIndex(0)->setCellValue('R58', "http://www.arpnjournals.org/jeas/research_papers/rp_2016/jeas_0416_4013.pdf ");
        $excel->setActiveSheetIndex(0)->setCellValue('L59', "2016");
        $excel->setActiveSheetIndex(0)->setCellValue('M59', "Internasional");

        $excel->setActiveSheetIndex(0)->setCellValue('D61', "ISSN (Online) : 1819-6608");
        $excel->setActiveSheetIndex(0)->setCellValue('D62', "Vol. 11, No 7, PP 4713-4719, 2016");
        $excel->setActiveSheetIndex(0)->setCellValue('R62', "Terdaftar di LPPM Unila");
        $excel->setActiveSheetIndex(0)->setCellValue('D64', "Performance Evaluation Of Various Genetic Algorithm Approaches For Knapsack Problem ");
        $excel->setActiveSheetIndex(0)->setCellValue('R64', "No. 165/J/B/I/FMIPA/2019");
        $excel->setActiveSheetIndex(0)->setCellValue('R65', "Tanggal 27 Juni 2019");
        $excel->setActiveSheetIndex(0)->setCellValue('R68', "A. Syarif, Aristoteles, A. Dwiastuti, and R. Malinda");

        $excel->setActiveSheetIndex(0)->setCellValue('C71', "4)");
        $excel->setActiveSheetIndex(0)->setCellValue('D71', "IJCSI International Journal Of Computer Science Issues");
        $excel->setActiveSheetIndex(0)->setCellValue('L71', "Mei");
        $excel->setActiveSheetIndex(0)->setCellValue('M71', "Jurnal");
        $excel->setActiveSheetIndex(0)->setCellValue('N71', "1,0");
        $excel->setActiveSheetIndex(0)->setCellValue('R71', "http://www.ijcsi.org/articles/Chord-identification-using-pitch-class-profile-method-with-fast-fourier-transform-feature-extraction.php ");
        $excel->setActiveSheetIndex(0)->setCellValue('L72', "2014");
        $excel->setActiveSheetIndex(0)->setCellValue('M72', "Internasional");

        $excel->setActiveSheetIndex(0)->setCellValue('D74', "ISSN 1694-0784");
        $excel->setActiveSheetIndex(0)->setCellValue('D75', "Vol. 11, Issue 3, No 1, May 2014, ");
        $excel->setActiveSheetIndex(0)->setCellValue('R76', "Terdaftar di LPPM Unila");
        $excel->setActiveSheetIndex(0)->setCellValue('D77', "Chord Identification Using Pitch Class Profile Method With Fast Fourier Transform Feature Extraction");
        $excel->setActiveSheetIndex(0)->setCellValue('R77', "No. 136/J/B/I/FMIPA/2019");
        $excel->setActiveSheetIndex(0)->setCellValue('R78', "Tanggal 12 Juni 2019");
        $excel->setActiveSheetIndex(0)->setCellValue('R81', "Kurnia Muludi, Aristoteles, Abe Frank SFB Loupatty");

        $excel->setActiveSheetIndex(0)->setCellValue('C84', "5)");
        $excel->setActiveSheetIndex(0)->setCellValue('D84', "International Journal of Computer Science and Telecommunications ");
        $excel->setActiveSheetIndex(0)->setCellValue('L84', "Juli");
        $excel->setActiveSheetIndex(0)->setCellValue('M84', "Jurnal");
        $excel->setActiveSheetIndex(0)->setCellValue('N84', "1,0");
        $excel->setActiveSheetIndex(0)->setCellValue('R84', "http://www.ijcst.org/Volume5/Issue7/p_6_5_7.pdf ");
        $excel->setActiveSheetIndex(0)->setCellValue('L85', "2014");
        $excel->setActiveSheetIndex(0)->setCellValue('M85', "Internasional");
        $excel->setActiveSheetIndex(0)->setCellValue('D87', "ISSN 2047-3338");
        $excel->setActiveSheetIndex(0)->setCellValue('R87', "http://repository.lppm.unila.ac.id/1358/ ");
        $excel->setActiveSheetIndex(0)->setCellValue('D88', "Volume 5, Issue 7, July 2014");
        $excel->setActiveSheetIndex(0)->setCellValue('D90', "Text Feature Weighting for Summarization of Documents Bahasa Indonesia by Using Binary Logistic Regression Algorithm");
        $excel->setActiveSheetIndex(0)->setCellValue('R90', "Terdaftar di LPPM Unila");
        $excel->setActiveSheetIndex(0)->setCellValue('R91', "No. 164/J/B/I/FMIPA/2019");
        $excel->setActiveSheetIndex(0)->setCellValue('R92', "Tanggal 27Juni 2019");
        $excel->setActiveSheetIndex(0)->setCellValue('D94', "Aristoteles, Widiarti and Eko Dwi Wibowo");

        $excel->setActiveSheetIndex(0)->setCellValue('C96', "6)");
        $excel->setActiveSheetIndex(0)->setCellValue('D96', "International Journal Of Computer Applications ");
        $excel->setActiveSheetIndex(0)->setCellValue('L96', "November");
        $excel->setActiveSheetIndex(0)->setCellValue('M96', "Jurnal");
        $excel->setActiveSheetIndex(0)->setCellValue('N96', "1,0");
        $excel->setActiveSheetIndex(0)->setCellValue('R96', "http://www.ijcaonline.org/archives/volume81/number6/14013-2158 ");
        $excel->setActiveSheetIndex(0)->setCellValue('L97', "2013");
        $excel->setActiveSheetIndex(0)->setCellValue('M97', "Internasional");

        $excel->setActiveSheetIndex(0)->setCellValue('D98', "ISSN 0975-8887");
        $excel->setActiveSheetIndex(0)->setCellValue('D99', "Volume 81 - No. 6, November 2013");
        $excel->setActiveSheetIndex(0)->setCellValue('R99', "Terdaftar di LPPM Unila");
        $excel->setActiveSheetIndex(0)->setCellValue('R100', "No. 166/J/B/I/FMIPA/2019");
        $excel->setActiveSheetIndex(0)->setCellValue('R101', "Image Processing For Save Life Predictions Of Tomato Fruit Using RGB Method");
        $excel->setActiveSheetIndex(0)->setCellValue('R101', "Tanggal 27 Juni 2019");
        $excel->setActiveSheetIndex(0)->setCellValue('R104', "Aristoteles, Ossy Dwi Endah W, Dwi Susanto");

        $excel->setActiveSheetIndex(0)->setCellValue('C106', "7)");
        $excel->setActiveSheetIndex(0)->setCellValue('D106', "International Journal Of Computer Applications ");
        $excel->setActiveSheetIndex(0)->setCellValue('L106', "Oktober");
        $excel->setActiveSheetIndex(0)->setCellValue('M106', "Jurnal");
        $excel->setActiveSheetIndex(0)->setCellValue('N106', "1,0");
        $excel->setActiveSheetIndex(0)->setCellValue('R106', "http://www.ijcaonline.org/archives/volume80/number13/13922-1824 ");
        $excel->setActiveSheetIndex(0)->setCellValue('L107', "2013");
        $excel->setActiveSheetIndex(0)->setCellValue('M107', "Internasional");

        $excel->setActiveSheetIndex(0)->setCellValue('D108', "ISSN 0975-8887");
        $excel->setActiveSheetIndex(0)->setCellValue('D109', "Volume 80 – No 13, October 2013, ");
        $excel->setActiveSheetIndex(0)->setCellValue('R109', "Terdaftar di LPPM Unila");
        $excel->setActiveSheetIndex(0)->setCellValue('R110', "No. 135/J/B/I/FMIPA/2019");
        $excel->setActiveSheetIndex(0)->setCellValue('R111', "Implementation Of Multilevel Feedback Queue Algorithm In Restaurant Order Food Application Development For Android And Ios Platforms");
        $excel->setActiveSheetIndex(0)->setCellValue('R111', "Tanggal 12 Juni 2019");
        $excel->setActiveSheetIndex(0)->setCellValue('R114', "Dian Andrian Ginting, Aristoteles, Ossy Dwi Endah");
        $excel->setActiveSheetIndex(0)->setCellValue('O117', "Jumlah");
    		$excel->setActiveSheetIndex(0)->setCellValue('P117', "0,00");
    		$excel->getActiveSheet()->getStyle('O117:P117')->getFont()->setBold(TRUE);
    		$excel->getActiveSheet()->getStyle('A117:R117')->applyFromArray($style_standar);
        $excel->setActiveSheetIndex(0)->setCellValue('E118', "2)");
				$excel->getActiveSheet()->getStyle('E118')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
        $excel->setActiveSheetIndex(0)->setCellValue('F118', "Nasional terakreditasi");
        $excel->getActiveSheet()->mergeCells('F118:I118');
        $excel->setActiveSheetIndex(0)->setCellValue('N120', "Jumlah");
        $excel->setActiveSheetIndex(0)->setCellValue('O120', "0,00");
    		$excel->setActiveSheetIndex(0)->setCellValue('P120', "0,00");
    		$excel->getActiveSheet()->getStyle('N120:O120:P120')->getFont()->setBold(TRUE);
    		$excel->getActiveSheet()->getStyle('A120:R120')->applyFromArray($style_standar);
        $excel->setActiveSheetIndex(0)->setCellValue('E122', "3)");
				$excel->getActiveSheet()->getStyle('E122')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
        $excel->setActiveSheetIndex(0)->setCellValue('F122', "Tidak terakreditasi");
        $excel->getActiveSheet()->mergeCells('F122:I122');

        $excel->setActiveSheetIndex(0)->setCellValue('B124', "1)");
				$excel->getActiveSheet()->getStyle('B124')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
        $excel->setActiveSheetIndex(0)->setCellValue('C124', "Jurnal Komputasi FMIPA Unila ");
        $excel->setActiveSheetIndex(0)->setCellValue('L124', "Oktober");
        $excel->setActiveSheetIndex(0)->setCellValue('M124', "Jurnal");
        $excel->setActiveSheetIndex(0)->setCellValue('N124', "1,00");
        $excel->setActiveSheetIndex(0)->setCellValue('R124', "http://jurnal.fmipa.unila.ac.id/komputasi/article/view/1148 ");
        $excel->setActiveSheetIndex(0)->setCellValue('C125', "ISSN (Online) : 2541-0350");
        $excel->setActiveSheetIndex(0)->setCellValue('L125', "2015");
        $excel->setActiveSheetIndex(0)->setCellValue('M125', "Nasional");
        $excel->setActiveSheetIndex(0)->setCellValue('C126', "Vol. 3 No 2, PP 136-143, Oktober 2015");
        $excel->setActiveSheetIndex(0)->setCellValue('C128', "Implementasi Teknologi Markerless Augmented Reality Berbasis Android untuk Mendeteksi dan Mengetahui Lokasi SPBU Terdekat di Kota Bandar Lampung");
        $excel->getActiveSheet()->mergeCells('C128:K128');
        $excel->getActiveSheet()->getStyle('C128')->getAlignment()->setWrapText(TRUE);
        $excel->getActiveSheet()->getStyle('C128')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);
    		$excel->setActiveSheetIndex(0)->setCellValue('R128', "Terdaftar di LPPM Unila");
        $excel->getActiveSheet()->getStyle('R128')->getAlignment()->setWrapText(TRUE);
        $excel->setActiveSheetIndex(0)->setCellValue('R129', "No. 271/J/B/N/FMIPA/2019");
        $excel->setActiveSheetIndex(0)->setCellValue('R130', "Tanggal 12 Juni 2019");
        $excel->setActiveSheetIndex(0)->setCellValue('C133', "Didik Kurniawan, Aristoteles, M. Fathan Kurniawan");
        $excel->getActiveSheet()->mergeCells('C133:H133');
        $excel->getActiveSheet()->getStyle('C133')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);

        $excel->setActiveSheetIndex(0)->setCellValue('B136', "2)");
				$excel->getActiveSheet()->getStyle('B136')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
        $excel->setActiveSheetIndex(0)->setCellValue('C136', "Jurnal Komputasi FMIPA Unila ");
        $excel->setActiveSheetIndex(0)->setCellValue('L136', "Oktober");
        $excel->setActiveSheetIndex(0)->setCellValue('M136', "Jurnal");
        $excel->setActiveSheetIndex(0)->setCellValue('N136', "1,00");
        $excel->setActiveSheetIndex(0)->setCellValue('R136', "http://jurnal.fmipa.unila.ac.id/komputasi/article/view/1143 ");
        $excel->setActiveSheetIndex(0)->setCellValue('C137', "ISSN (Online) : 2541-0350");
        $excel->setActiveSheetIndex(0)->setCellValue('L137', "2015");
        $excel->setActiveSheetIndex(0)->setCellValue('M137', "Nasional");
        $excel->setActiveSheetIndex(0)->setCellValue('C138', "Vol. 3 No 2, PP 120-128, Oktober 2015");
        $excel->setActiveSheetIndex(0)->setCellValue('C140', "Pengembangan Aplikasi Sistem Pembelajaran Klasifikasi (Taksonomi) dan Tata Nama Ilmiah (Binomial Nomenklatur) pada Kingdom Plantae (Tumbuhan) Berbasis Android");
        $excel->getActiveSheet()->mergeCells('C140:K140');
        $excel->getActiveSheet()->getStyle('C140')->getAlignment()->setWrapText(TRUE);
        $excel->getActiveSheet()->getStyle('C140')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);
        $excel->setActiveSheetIndex(0)->setCellValue('R140', "Terdaftar di LPPM Unila");
        $excel->getActiveSheet()->getStyle('R140')->getAlignment()->setWrapText(TRUE);
        $excel->setActiveSheetIndex(0)->setCellValue('R141', "No. 270/J/B/N/FMIPA/2019");
        $excel->setActiveSheetIndex(0)->setCellValue('R142', "Tanggal 12 Juni 2019");
        $excel->setActiveSheetIndex(0)->setCellValue('C145', "Didik Kurniawan, Aristoteles, Ahmad Amirudin");

        $excel->setActiveSheetIndex(0)->setCellValue('B147', "3)");
				$excel->getActiveSheet()->getStyle('B147')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
        $excel->setActiveSheetIndex(0)->setCellValue('C147', "Jurnal Komputasi FMIPA Unila ");
        $excel->setActiveSheetIndex(0)->setCellValue('L147', "Oktober");
        $excel->setActiveSheetIndex(0)->setCellValue('M147', "Jurnal");
        $excel->setActiveSheetIndex(0)->setCellValue('N147', "1,00");
        $excel->setActiveSheetIndex(0)->setCellValue('R147', "http://jurnal.fmipa.unila.ac.id/komputasi/article/view/1131 ");
        $excel->setActiveSheetIndex(0)->setCellValue('C148', "ISSN (Online) : 2541-0350");
        $excel->setActiveSheetIndex(0)->setCellValue('L148', "2015");
        $excel->setActiveSheetIndex(0)->setCellValue('M148', "Nasional");
        $excel->setActiveSheetIndex(0)->setCellValue('C149', "Vol. 3 No 2, PP 44-52, Oktober 2015");
        $excel->setActiveSheetIndex(0)->setCellValue('C151', "Sistem Informasi Kuliah Kerja Nyata (KKN) dengan Metode Pigeon Hole untuk Menentukan dan Mengelompokkan Peserta KKN Universitas Lampung");
        $excel->getActiveSheet()->mergeCells('C151:K151');
        $excel->getActiveSheet()->getStyle('C151')->getAlignment()->setWrapText(TRUE);
        $excel->getActiveSheet()->getStyle('C151')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);
        $excel->setActiveSheetIndex(0)->setCellValue('R150', "Terdaftar di LPPM Unila");
        $excel->getActiveSheet()->getStyle('R140')->getAlignment()->setWrapText(TRUE);
        $excel->setActiveSheetIndex(0)->setCellValue('R151', "No. 267/J/B/N/FMIPA/2019");
        $excel->setActiveSheetIndex(0)->setCellValue('R152', "Tanggal 12 Juni 2019");
        $excel->setActiveSheetIndex(0)->setCellValue('C156', "Aristoteles, Rico Andrian, Agatha Beny Himawan");
        $excel->getActiveSheet()->mergeCells('C156:K156');
        $excel->getActiveSheet()->getStyle('C156')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);

        $excel->setActiveSheetIndex(0)->setCellValue('B159', "4)");
				$excel->getActiveSheet()->getStyle('B159')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
        $excel->setActiveSheetIndex(0)->setCellValue('C159', "Jurnal Komputer ");
        $excel->setActiveSheetIndex(0)->setCellValue('L159', "Oktober");
        $excel->setActiveSheetIndex(0)->setCellValue('M159', "Jurnal");
        $excel->setActiveSheetIndex(0)->setCellValue('N159', "1,00");
        $excel->setActiveSheetIndex(0)->setCellValue('R159', "http://jurnal.fmipa.unila.ac.id/index.php/komputasi/article/view/1128 ");
        $excel->setActiveSheetIndex(0)->setCellValue('C160', "ISSN :  2541-035");
        $excel->setActiveSheetIndex(0)->setCellValue('L160', "2015");
        $excel->setActiveSheetIndex(0)->setCellValue('M160', "Nasional");
        $excel->setActiveSheetIndex(0)->setCellValue('C161', "Vol 3 No 2, Hal 99-108, Oktober 2015");

        $excel->setActiveSheetIndex(0)->setCellValue('C163', "SISTEM IDENTIFIKASI PENYAKIT TANAMAN PADI DENGAN MENGGUNAKAN METODE FORWARD CHAINING");
        $excel->getActiveSheet()->mergeCells('C163:K163');
        $excel->getActiveSheet()->getStyle('C163')->getAlignment()->setWrapText(TRUE);
        $excel->getActiveSheet()->getStyle('C163')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);
        $excel->setActiveSheetIndex(0)->setCellValue('R163', "Terdaftar di LPPM Unila");
				$excel->getActiveSheet()->getStyle('R163')->getAlignment()->setWrapText(TRUE);
        $excel->setActiveSheetIndex(0)->setCellValue('R164', "No. 265/J/B/N/FMIPA/2019");
        $excel->setActiveSheetIndex(0)->setCellValue('R165', "Tanggal 12 Juni 2019");
        $excel->setActiveSheetIndex(0)->setCellValue('C167', "Aristoteles, Wardiyanto, Ardye Amando Pratama");

        $excel->setActiveSheetIndex(0)->setCellValue('B169', "5)");
				$excel->getActiveSheet()->getStyle('B169')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
        $excel->setActiveSheetIndex(0)->setCellValue('C169', "Jurnal Komputer ");
        $excel->setActiveSheetIndex(0)->setCellValue('L169', "April");
        $excel->setActiveSheetIndex(0)->setCellValue('M169', "Jurnal");
        $excel->setActiveSheetIndex(0)->setCellValue('N169', "1,00");
        $excel->setActiveSheetIndex(0)->setCellValue('R169', "http://jurnal.fmipa.unila.ac.id/index.php/komputasi/article/view/1216 ");
        $excel->setActiveSheetIndex(0)->setCellValue('C170', "ISSN :  2541-035");
        $excel->setActiveSheetIndex(0)->setCellValue('L170', "2016");
        $excel->setActiveSheetIndex(0)->setCellValue('M170', "Nasional");
        $excel->setActiveSheetIndex(0)->setCellValue('C171', "Vol 4 No 1, Hal 9-18, April 2016");

        $excel->setActiveSheetIndex(0)->setCellValue('C173', "SISTEM IDENTIFIKASI PENYAKIT TANAMAN PADI DENGAN MENGGUNAKAN METODE FORWARD CHAINING");
        $excel->getActiveSheet()->mergeCells('C173:K173');
        $excel->getActiveSheet()->getStyle('C173')->getAlignment()->setWrapText(TRUE);
        $excel->getActiveSheet()->getStyle('C173')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);
        $excel->setActiveSheetIndex(0)->setCellValue('R173', "Terdaftar di LPPM Unila");
        $excel->setActiveSheetIndex(0)->setCellValue('R174', "No. 272/J/B/N/FMIPA/2019");
        $excel->setActiveSheetIndex(0)->setCellValue('R175', "Tanggal 12 Juni 2019");
        $excel->setActiveSheetIndex(0)->setCellValue('C177', "Ika Arthalia Wulandari, Aristoteles, Radix Suharjo");

        $excel->setActiveSheetIndex(0)->setCellValue('B179', "6)");
				$excel->getActiveSheet()->getStyle('B179')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
        $excel->setActiveSheetIndex(0)->setCellValue('C179', "Jurnal Komputer ");
        $excel->setActiveSheetIndex(0)->setCellValue('L179', "April");
        $excel->setActiveSheetIndex(0)->setCellValue('M179', "Jurnal");
        $excel->setActiveSheetIndex(0)->setCellValue('N179', "1,00");
        $excel->setActiveSheetIndex(0)->setCellValue('R179', "http://jurnal.fmipa.unila.ac.id/index.php/komputasi/article/view/1164 ");
        $excel->setActiveSheetIndex(0)->setCellValue('C180', "ISSN :  2541-035");
        $excel->setActiveSheetIndex(0)->setCellValue('L180', "2016");
        $excel->setActiveSheetIndex(0)->setCellValue('M180', "Nasional");
        $excel->setActiveSheetIndex(0)->setCellValue('C181', "Vol 4 No 1, Hal 92-98, April 2016");

        $excel->setActiveSheetIndex(0)->setCellValue('C183', "SISTEM PAKAR DIAGNOSA PENYAKIT PADA IKAN BUDIDAYA AIR TAWAR DENGAN METODE FORWARD CHAINING BERBASIS ANDROID ");
        $excel->getActiveSheet()->mergeCells('C183:K183');
        $excel->getActiveSheet()->getStyle('C183')->getAlignment()->setWrapText(TRUE);
        $excel->getActiveSheet()->getStyle('C183')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);
        $excel->setActiveSheetIndex(0)->setCellValue('R183', "Terdaftar di LPPM Unila");
        $excel->setActiveSheetIndex(0)->setCellValue('R184', "No. 301/J/B/N/FMIPA/2019");
        $excel->setActiveSheetIndex(0)->setCellValue('R185', "Tanggal 27 Juni 2019");
        $excel->setActiveSheetIndex(0)->setCellValue('C187', "Ardhika Praseda Ageng Putra, Aristoteles, Rara Diantari");
        $excel->getActiveSheet()->mergeCells('C187:H187');

        $excel->setActiveSheetIndex(0)->setCellValue('B190', "7)");
				$excel->getActiveSheet()->getStyle('B190')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
        $excel->setActiveSheetIndex(0)->setCellValue('C190', "Jurnal Komputer ");
        $excel->setActiveSheetIndex(0)->setCellValue('L190', "April");
        $excel->setActiveSheetIndex(0)->setCellValue('M190', "Jurnal");
        $excel->setActiveSheetIndex(0)->setCellValue('N190', "1,00");
        $excel->setActiveSheetIndex(0)->setCellValue('R190', "http://jurnal.fmipa.unila.ac.id/index.php/komputasi/article/view/1173 ");
        $excel->setActiveSheetIndex(0)->setCellValue('C191', "ISSN :  2541-035");
        $excel->setActiveSheetIndex(0)->setCellValue('L191', "2016");
        $excel->setActiveSheetIndex(0)->setCellValue('M191', "Nasional");
        $excel->setActiveSheetIndex(0)->setCellValue('C192', "Vol 4 No 1, Hal 117-124, April 2016");

        $excel->setActiveSheetIndex(0)->setCellValue('C194', "SISTEM PAKAR DIAGNOSA PENYAKIT PADA IKAN BUDIDAYA AIR TAWAR DENGAN METODE FORWARD CHAINING BERBASIS ANDROID ");
        $excel->getActiveSheet()->mergeCells('C194:K194');
        $excel->getActiveSheet()->getStyle('C194')->getAlignment()->setWrapText(TRUE);
        $excel->getActiveSheet()->getStyle('C194')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);
        $excel->setActiveSheetIndex(0)->setCellValue('R194', "Terdaftar di LPPM Unila");
        $excel->setActiveSheetIndex(0)->setCellValue('R195', "No. 300/J/B/N/FMIPA/2019");
        $excel->setActiveSheetIndex(0)->setCellValue('R196', "Tanggal 27 Juni 2019");
        $excel->setActiveSheetIndex(0)->setCellValue('C197', "Rifki Wardana, Aristoteles, Jani Master");

				$excel->setActiveSheetIndex(0)->setCellValue('B199', "8)");
				$excel->getActiveSheet()->getStyle('B199')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
        $excel->setActiveSheetIndex(0)->setCellValue('C199', "Jurnal Komputer ");
        $excel->setActiveSheetIndex(0)->setCellValue('L199', "April");
        $excel->setActiveSheetIndex(0)->setCellValue('M199', "Jurnal");
        $excel->setActiveSheetIndex(0)->setCellValue('N199', "1,00");
        $excel->setActiveSheetIndex(0)->setCellValue('R199', "http://jurnal.fmipa.unila.ac.id/komputasi/article/view/1191 ");
        $excel->setActiveSheetIndex(0)->setCellValue('C200', "ISSN :  2541-035");
        $excel->setActiveSheetIndex(0)->setCellValue('L200', "2016");
        $excel->setActiveSheetIndex(0)->setCellValue('M200', "Nasional");
        $excel->setActiveSheetIndex(0)->setCellValue('C201', "Vol 4 No 1, Hal 176-1186, April 2016");

        $excel->setActiveSheetIndex(0)->setCellValue('C203', "PENGEMBANGAN SISTEM INFORMASI COMICREADER MENGGUNAKAN KERANGKA KERJA YII");
        $excel->getActiveSheet()->mergeCells('C203:K203');
        $excel->getActiveSheet()->getStyle('C203')->getAlignment()->setWrapText(TRUE);
        $excel->getActiveSheet()->getStyle('C203')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);
        $excel->setActiveSheetIndex(0)->setCellValue('R203', "Terdaftar di LPPM Unila");
        $excel->setActiveSheetIndex(0)->setCellValue('R204', "No. 269/J/B/N/FMIPA/2019");
        $excel->setActiveSheetIndex(0)->setCellValue('R205', "Tanggal 12 Juni 2019");
        $excel->setActiveSheetIndex(0)->setCellValue('C207', "Sabila Rusyda, Aristoteles, Dwi Sakethi, Admi Syarif");
				$excel->getActiveSheet()->mergeCells('C207:H207');

				$excel->setActiveSheetIndex(0)->setCellValue('B210', "9)");
				$excel->getActiveSheet()->getStyle('B210')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
        $excel->setActiveSheetIndex(0)->setCellValue('C210', "Jurnal Komputasi FMIPA Unila");
        $excel->setActiveSheetIndex(0)->setCellValue('L210', "Oktober");
        $excel->setActiveSheetIndex(0)->setCellValue('M210', "Jurnal");
        $excel->setActiveSheetIndex(0)->setCellValue('N210', "1,00");
        $excel->setActiveSheetIndex(0)->setCellValue('R210', "http://jurnal.fmipa.unila.ac.id/komputasi/article/view/1351 ");
        $excel->setActiveSheetIndex(0)->setCellValue('C211', "ISSN (Online) : 2541-0350");
        $excel->setActiveSheetIndex(0)->setCellValue('L211', "2016");
        $excel->setActiveSheetIndex(0)->setCellValue('M211', "Nasional");
        $excel->setActiveSheetIndex(0)->setCellValue('C212', "Vol. 4 No 2, PP 52-66, Oktober 2016");

        $excel->setActiveSheetIndex(0)->setCellValue('C214', "PEMETAAN SEBARAN ASAL SISWA DAN KLASIFIKASI JARAK ASAL SISWA SMA NEGERI DI KABUPATEN PRINGSEWU MENGGUNAKAN METODE NAÏVE BAYES");
        $excel->getActiveSheet()->mergeCells('C214:K214');
        $excel->getActiveSheet()->getStyle('C214')->getAlignment()->setWrapText(TRUE);
        $excel->getActiveSheet()->getStyle('C214')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);
        $excel->setActiveSheetIndex(0)->setCellValue('R214', "Terdaftar di LPPM Unila");
        $excel->setActiveSheetIndex(0)->setCellValue('R215', "No. 299/J/B/N/FMIPA/2019");
        $excel->setActiveSheetIndex(0)->setCellValue('R216', "Tanggal 27 Juni 2019");
        $excel->setActiveSheetIndex(0)->setCellValue('C219', "Riska Aprilia, Kurnia Muludi, Aristoteles");

				$excel->setActiveSheetIndex(0)->setCellValue('B221', "10)");
				$excel->getActiveSheet()->getStyle('B221')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
        $excel->setActiveSheetIndex(0)->setCellValue('C221', "Jurnal Komputasi FMIPA Unila ");
        $excel->setActiveSheetIndex(0)->setCellValue('L221', "April");
        $excel->setActiveSheetIndex(0)->setCellValue('M221', "Jurnal");
        $excel->setActiveSheetIndex(0)->setCellValue('N221', "1,00");
        $excel->setActiveSheetIndex(0)->setCellValue('R221', "http://jurnal.fmipa.unila.ac.id/komputasi/article/view/1402/1220 ");
        $excel->setActiveSheetIndex(0)->setCellValue('C222', "ISSN (Online) : 2541-0350");
        $excel->setActiveSheetIndex(0)->setCellValue('L222', "2017");
        $excel->setActiveSheetIndex(0)->setCellValue('M222', "Nasional");
        $excel->setActiveSheetIndex(0)->setCellValue('C223', "Vol. 5, No 1, PP 8-16, April 2017");

        $excel->setActiveSheetIndex(0)->setCellValue('C225', "PENGEMBANGAN SISTEM PELAPORAN KEGIATAN KKN BERBASIS ANDROID");
        $excel->getActiveSheet()->mergeCells('C225:K225');
        $excel->getActiveSheet()->getStyle('C225')->getAlignment()->setWrapText(TRUE);
        $excel->getActiveSheet()->getStyle('C225')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);
        $excel->setActiveSheetIndex(0)->setCellValue('R225', "Terdaftar di LPPM Unila");
        $excel->setActiveSheetIndex(0)->setCellValue('R226', "No. 299/J/B/N/FMIPA/2019");
        $excel->setActiveSheetIndex(0)->setCellValue('R227', "Tanggal 27 Juni 2019");
        $excel->setActiveSheetIndex(0)->setCellValue('C229', "Danzen Hangga Permana, Aristoteles");

				$excel->setActiveSheetIndex(0)->setCellValue('B231', "11)");
				$excel->getActiveSheet()->getStyle('B231')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
        $excel->setActiveSheetIndex(0)->setCellValue('C231', "Jurnal Komputasi FMIPA Unila ");
        $excel->setActiveSheetIndex(0)->setCellValue('L231', "April");
        $excel->setActiveSheetIndex(0)->setCellValue('M231', "Jurnal");
        $excel->setActiveSheetIndex(0)->setCellValue('N231', "1,00");
        $excel->setActiveSheetIndex(0)->setCellValue('R231', "http://jurnal.fmipa.unila.ac.id/komputasi/article/view/1402/1220 ");
        $excel->setActiveSheetIndex(0)->setCellValue('C232', "ISSN (Online) : 2541-0350");
        $excel->setActiveSheetIndex(0)->setCellValue('L232', "2017");
        $excel->setActiveSheetIndex(0)->setCellValue('M232', "Nasional");
        $excel->setActiveSheetIndex(0)->setCellValue('C233', "Vol. 5, No 1, PP 8-16, April 2017");

        $excel->setActiveSheetIndex(0)->setCellValue('C236', "ANALISIS PENGELOMPOKAN MAHASISWA KKN BERDASARKAN KRITERIA JENIS KELAMIN, FAKULTAS DAN SEKOLAH");
        $excel->getActiveSheet()->mergeCells('C236:K236');
        $excel->getActiveSheet()->getStyle('C236')->getAlignment()->setWrapText(TRUE);
        $excel->getActiveSheet()->getStyle('C236')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);
        $excel->setActiveSheetIndex(0)->setCellValue('R236', "Terdaftar di LPPM Unila");
        $excel->setActiveSheetIndex(0)->setCellValue('R237', "No. 266/J/B/N/FMIPA/2019");
        $excel->setActiveSheetIndex(0)->setCellValue('R238', "Tanggal 12 Juni 2019");
        $excel->setActiveSheetIndex(0)->setCellValue('C240', "Vandu Riski Muwisnawangsa, Aristoteles");

				$excel->setActiveSheetIndex(0)->setCellValue('B242', "12)");
				$excel->getActiveSheet()->getStyle('B242')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
        $excel->setActiveSheetIndex(0)->setCellValue('C242', "Jurnal Komputasi FMIPA Unila ");
        $excel->setActiveSheetIndex(0)->setCellValue('L242', "Oktober");
        $excel->setActiveSheetIndex(0)->setCellValue('M242', "Jurnal");
        $excel->setActiveSheetIndex(0)->setCellValue('N242', "1,00");
        $excel->setActiveSheetIndex(0)->setCellValue('R242', "http://jurnal.fmipa.unila.ac.id/komputasi/article/view/1539/1307 ");
        $excel->setActiveSheetIndex(0)->setCellValue('C243', "ISSN (Online) : 2541-0350");
        $excel->setActiveSheetIndex(0)->setCellValue('L243', "2017");
        $excel->setActiveSheetIndex(0)->setCellValue('M243', "Nasional");
        $excel->setActiveSheetIndex(0)->setCellValue('C244', "Vol. 5, No 2, PP 55-63, Oktober 2017");

        $excel->setActiveSheetIndex(0)->setCellValue('C246', "APLIKASI INFORMASI DOKTER SPESIALIS DI BANDAR LAMPUNG BERBASIS ANDROID DENGAN MENGGUNAKAN TEKNOLOGI LOCATION BASE SERVICE");
        $excel->getActiveSheet()->mergeCells('C246:K246');
        $excel->getActiveSheet()->getStyle('C246')->getAlignment()->setWrapText(TRUE);
        $excel->getActiveSheet()->getStyle('C246')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);
        $excel->setActiveSheetIndex(0)->setCellValue('R246', "Terdaftar di LPPM Unila");
        $excel->setActiveSheetIndex(0)->setCellValue('R247', "No. 274/J/B/N/FMIPA/2019");
        $excel->setActiveSheetIndex(0)->setCellValue('R248', "Tanggal 12 Juni 2019");
        $excel->setActiveSheetIndex(0)->setCellValue('C250', "Nurmayanti, Aristoteles, Astria Hijriani");

				$excel->setActiveSheetIndex(0)->setCellValue('B252', "13)");
				$excel->getActiveSheet()->getStyle('B252')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
        $excel->setActiveSheetIndex(0)->setCellValue('C252', "Jurnal Komputasi FMIPA Unila ");
        $excel->setActiveSheetIndex(0)->setCellValue('L252', "April");
        $excel->setActiveSheetIndex(0)->setCellValue('M252', "Jurnal");
        $excel->setActiveSheetIndex(0)->setCellValue('N252', "1,00");
        $excel->setActiveSheetIndex(0)->setCellValue('R252', "http://jurnal.fmipa.unila.ac.id/komputasi/article/view/1564/1318 ");
        $excel->setActiveSheetIndex(0)->setCellValue('C253', "ISSN (Online) : 2541-0350");
        $excel->setActiveSheetIndex(0)->setCellValue('L253', "2018");
        $excel->setActiveSheetIndex(0)->setCellValue('M253', "Nasional");
        $excel->setActiveSheetIndex(0)->setCellValue('C254', "Vol. 6, No 1, PP 64-74, April 2018");

        $excel->setActiveSheetIndex(0)->setCellValue('C256', "Panduan Lapangan Jenis Kupu-kupu di Lingkungan Universitas Lampung Berbasis Android");
        $excel->getActiveSheet()->mergeCells('C256:K256');
        $excel->getActiveSheet()->getStyle('C256')->getAlignment()->setWrapText(TRUE);
        $excel->getActiveSheet()->getStyle('C256')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);
        $excel->setActiveSheetIndex(0)->setCellValue('R256', "Terdaftar di LPPM Unila");
        $excel->setActiveSheetIndex(0)->setCellValue('R257', "No. 273/J/B/N/FMIPA/2019");
        $excel->setActiveSheetIndex(0)->setCellValue('R258', "Tanggal 12 Juni 2019");
        $excel->setActiveSheetIndex(0)->setCellValue('C260', "Aristoteles, Martinus, Galih Imam Widangga");

				$excel->setActiveSheetIndex(0)->setCellValue('B262', "14)");
				$excel->getActiveSheet()->getStyle('B262')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
        $excel->setActiveSheetIndex(0)->setCellValue('C262', "Jurnal Komputasi FMIPA Unila ");
        $excel->setActiveSheetIndex(0)->setCellValue('L262', "Oktober");
        $excel->setActiveSheetIndex(0)->setCellValue('M262', "Jurnal");
        $excel->setActiveSheetIndex(0)->setCellValue('N262', "1,00");
        $excel->setActiveSheetIndex(0)->setCellValue('R262', "http://jurnal.fmipa.unila.ac.id/komputasi/article/view/1655/1332 ");
        $excel->setActiveSheetIndex(0)->setCellValue('C263', "ISSN (Online) : 2541-0350");
        $excel->setActiveSheetIndex(0)->setCellValue('L263', "2018");
        $excel->setActiveSheetIndex(0)->setCellValue('M263', "Nasional");
        $excel->setActiveSheetIndex(0)->setCellValue('C264', "Vol. 6, No 2, PP 1-10, Oktober 2018");

        $excel->setActiveSheetIndex(0)->setCellValue('C266', "SISTEM INFORMASI KULIAH KERJA NYATA (KKN) BERBASIS ANDROID UNIVERSITAS LAMPUNG");
        $excel->getActiveSheet()->mergeCells('C266:K266');
        $excel->getActiveSheet()->getStyle('C266')->getAlignment()->setWrapText(TRUE);
        $excel->getActiveSheet()->getStyle('C266')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);
        $excel->setActiveSheetIndex(0)->setCellValue('R266', "Terdaftar di LPPM Unila");
        $excel->setActiveSheetIndex(0)->setCellValue('R267', "No. 298/J/B/N/FMIPA/2019");
        $excel->setActiveSheetIndex(0)->setCellValue('R268', "Tanggal 27 Juni 2019");
        $excel->setActiveSheetIndex(0)->setCellValue('C270', "Aristoteles, Nur Efendi, Febi Eka Febriansyah, Wisnu Lukito, Firmansyah");

				$excel->setActiveSheetIndex(0)->setCellValue('B272', "15)");
				$excel->getActiveSheet()->getStyle('B272')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
        $excel->setActiveSheetIndex(0)->setCellValue('C272', "Jurnal Komputasi FMIPA Unila ");
        $excel->setActiveSheetIndex(0)->setCellValue('L272', "Oktober");
        $excel->setActiveSheetIndex(0)->setCellValue('M272', "Jurnal");
        $excel->setActiveSheetIndex(0)->setCellValue('N272', "1,00");
        $excel->setActiveSheetIndex(0)->setCellValue('R272', "http://jurnal.fmipa.unila.ac.id/komputasi/article/view/1693/1339 ");
        $excel->setActiveSheetIndex(0)->setCellValue('C273', "ISSN (Online) : 2541-0350");
        $excel->setActiveSheetIndex(0)->setCellValue('L273', "2018");
        $excel->setActiveSheetIndex(0)->setCellValue('M273', "Nasional");
        $excel->setActiveSheetIndex(0)->setCellValue('C274', "Vol. 6, No 2, PP 64-73, Oktober 2018");

        $excel->setActiveSheetIndex(0)->setCellValue('C276', "Analisis Manajemen Risiko Sistem Informasi KKN Universitas Lampung");
        $excel->getActiveSheet()->mergeCells('C276:K276');
        $excel->getActiveSheet()->getStyle('C276')->getAlignment()->setWrapText(TRUE);
        $excel->getActiveSheet()->getStyle('C276')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);
        $excel->setActiveSheetIndex(0)->setCellValue('R276', "Terdaftar di LPPM Unila");
        $excel->setActiveSheetIndex(0)->setCellValue('R277', "No. 268/J/B/N/FMIPA/2019");
        $excel->setActiveSheetIndex(0)->setCellValue('R278', "Tanggal 12 Juni 2019");
        $excel->setActiveSheetIndex(0)->setCellValue('C280', "Noviyanti, Yunda Heningtyas, Tristiyanto, Aristoteles");
				$excel->getActiveSheet()->getStyle('A282:R282')->applyFromArray($style_standar);
				$excel->setActiveSheetIndex(0)->setCellValue('N283', "Jumlah");
        $excel->setActiveSheetIndex(0)->setCellValue('O283', "0,00");
    		$excel->setActiveSheetIndex(0)->setCellValue('P283', "0,00");
    		$excel->getActiveSheet()->getStyle('N283:P283')->getFont()->setBold(TRUE);
				$excel->getActiveSheet()->getStyle('A283:R283')->applyFromArray($style_standar);

				$excel->setActiveSheetIndex(0)->setCellValue('D284', "c.");
				$excel->getActiveSheet()->getStyle('D284')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
        $excel->setActiveSheetIndex(0)->setCellValue('E284', "Seminar");
        $excel->setActiveSheetIndex(0)->setCellValue('E285', "1)");
				$excel->getActiveSheet()->getStyle('E285')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
        $excel->setActiveSheetIndex(0)->setCellValue('F285', "Disajikan tingkat:");
				$excel->setActiveSheetIndex(0)->setCellValue('N286', "Jumlah");
        $excel->setActiveSheetIndex(0)->setCellValue('O286', "0,00");
    		$excel->setActiveSheetIndex(0)->setCellValue('P286', "0,00");
    		$excel->getActiveSheet()->getStyle('N286:P286')->getFont()->setBold(TRUE);
				$excel->getActiveSheet()->getStyle('A286:R286')->applyFromArray($style_standar);

				$excel->setActiveSheetIndex(0)->setCellValue('F287', "a) Internasional");
				$excel->setActiveSheetIndex(0)->setCellValue('B289', "1)");
				$excel->getActiveSheet()->getStyle('B289')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
				$excel->setActiveSheetIndex(0)->setCellValue('C289', "3rd INTERNATIONAL WILDLIFE SYMPOSIUM");
				$excel->setActiveSheetIndex(0)->setCellValue('L289', "15)");
				$excel->setActiveSheetIndex(0)->setCellValue('M289', "18-20 Oktober");
				$excel->setActiveSheetIndex(0)->setCellValue('N289', "1,00");
				$excel->setActiveSheetIndex(0)->setCellValue('P289', "0,00");
				$excel->setActiveSheetIndex(0)->setCellValue('R289', "http://repository.lppm.unila.ac.id/3816/ ");
				$excel->setActiveSheetIndex(0)->setCellValue('C290', "ISBN 978-602-0860-13-8");
				$excel->setActiveSheetIndex(0)->setCellValue('L290', "2016");

				$excel->setActiveSheetIndex(0)->setCellValue('C292', "An Expert System To Diagnose Chicken Diseases With Certainty Factor Based On Android ");
				$excel->getActiveSheet()->mergeCells('C292:K292');
        $excel->getActiveSheet()->getStyle('C292')->getAlignment()->setWrapText(TRUE);
        $excel->getActiveSheet()->getStyle('C292')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);
				$excel->setActiveSheetIndex(0)->setCellValue('R292', "No. 121/P/B/I/FMIPA/2019");
				$excel->setActiveSheetIndex(0)->setCellValue('R293', "Tanggal 27 Juni 2019");
				$excel->setActiveSheetIndex(0)->setCellValue('C295', "Aristoteles, Kusuma Adhianto,");
				$excel->setActiveSheetIndex(0)->setCellValue('N297', "Jumlah");
        $excel->setActiveSheetIndex(0)->setCellValue('O297', "0,00");
    		$excel->setActiveSheetIndex(0)->setCellValue('P297', "0,00");
    		$excel->getActiveSheet()->getStyle('N297:P297')->getFont()->setBold(TRUE);
				$excel->getActiveSheet()->getStyle('A297:R297')->applyFromArray($style_standar);

				$excel->setActiveSheetIndex(0)->setCellValue('F298', "b) Nasional");
				$excel->setActiveSheetIndex(0)->setCellValue('B300', "2)");
				$excel->getActiveSheet()->getStyle('B300')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
				$excel->setActiveSheetIndex(0)->setCellValue('C300', "Prosiding Sain dan Teknologi VI 2015");
				$excel->setActiveSheetIndex(0)->setCellValue('L300', "3 November");
				$excel->setActiveSheetIndex(0)->setCellValue('M300', "Prosiding");
				$excel->setActiveSheetIndex(0)->setCellValue('N300', "1,00");
				$excel->setActiveSheetIndex(0)->setCellValue('R300', "http://satek.unila.ac.id/wp-content/uploads/2015/08/41-Aldona.pdf ");
				$excel->setActiveSheetIndex(0)->setCellValue('C301', "ISBN  : 978-602-0860-02-2");
				$excel->setActiveSheetIndex(0)->setCellValue('L301', "2015");
				$excel->setActiveSheetIndex(0)->setCellValue('C302', "Hal 485-491, November 2015");

				$excel->setActiveSheetIndex(0)->setCellValue('C304', "Sistem Informasi Pemantauan Potensi Desa dan Pengumpulan Laporan Hasil Kegiatan Kuliah Kerja Nyata (KKN) Universitas Lampung");
				$excel->getActiveSheet()->mergeCells('C304:K304');
        $excel->getActiveSheet()->getStyle('C304')->getAlignment()->setWrapText(TRUE);
        $excel->getActiveSheet()->getStyle('C304')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);
				$excel->setActiveSheetIndex(0)->setCellValue('R304', "No. 117/P/B/N/FMIPA/2019");
				$excel->setActiveSheetIndex(0)->setCellValue('R305', "Tanggal 27 Juni 2019");
				$excel->setActiveSheetIndex(0)->setCellValue('C307', "Aldona Pronika, Aristoteles dan Irwan Adi Pribadi");

				$excel->setActiveSheetIndex(0)->setCellValue('B309', "3)");
				$excel->getActiveSheet()->getStyle('B309')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
				$excel->setActiveSheetIndex(0)->setCellValue('C309', "Prosiding Sain dan Teknologi VI 2015");
				$excel->setActiveSheetIndex(0)->setCellValue('L309', "3 November");
				$excel->setActiveSheetIndex(0)->setCellValue('M309', "Prosiding");
				$excel->setActiveSheetIndex(0)->setCellValue('N309', "1,00");
				$excel->setActiveSheetIndex(0)->setCellValue('R309', "http://satek.unila.ac.id/wp-content/uploads/2015/08/44-Harisa.pdf ");
				$excel->setActiveSheetIndex(0)->setCellValue('C310', "ISBN  : 978-602-0860-02-2");
				$excel->setActiveSheetIndex(0)->setCellValue('L310', "2015");
				$excel->setActiveSheetIndex(0)->setCellValue('C311', "Hal 516-527, November 2015");

				$excel->setActiveSheetIndex(0)->setCellValue('C313', "Pengembangan Sistem Informasi Kuliah Kerja Nyata (KKN) dengan Algortima Greedy Untuk Menentukan Pengelompokan Peserta KKN (Studi Kasus Universitas Lampung)");
				$excel->getActiveSheet()->mergeCells('C313:K313');
        $excel->getActiveSheet()->getStyle('C313')->getAlignment()->setWrapText(TRUE);
        $excel->getActiveSheet()->getStyle('C313')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);
				$excel->setActiveSheetIndex(0)->setCellValue('R313', "No. 120/P/B/N/FMIPA/2019");
				$excel->setActiveSheetIndex(0)->setCellValue('R314', "Tanggal 27 Juni 2019");
				$excel->setActiveSheetIndex(0)->setCellValue('C316', "Harisa Eka Septiarani, Aristoteles  dan Wamiliana");

				$excel->setActiveSheetIndex(0)->setCellValue('B318', "4)");
				$excel->getActiveSheet()->getStyle('B318')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
				$excel->setActiveSheetIndex(0)->setCellValue('C318', "Prosiding Semirata FMIPA Universitas Lampung");
				$excel->setActiveSheetIndex(0)->setCellValue('L318', "10-12 Mei");
				$excel->setActiveSheetIndex(0)->setCellValue('M318', "Prosiding");
				$excel->setActiveSheetIndex(0)->setCellValue('N318', "1,00");
				$excel->setActiveSheetIndex(0)->setCellValue('R318', "http://jurnal.fmipa.unila.ac.id/index.php/semirata/article/view/703 ");
				$excel->setActiveSheetIndex(0)->setCellValue('C319', "ISBN 978-602-985599-2-0");
				$excel->setActiveSheetIndex(0)->setCellValue('L319', "2013");
				$excel->setActiveSheetIndex(0)->setCellValue('C320', "10-12 Mei 2013");

				$excel->setActiveSheetIndex(0)->setCellValue('C322', "Penerapan Algoritma Genetika Pada Peringkasan Teks Dokumen Bahasa Indonesia");
				$excel->getActiveSheet()->mergeCells('C322:K322');
        $excel->getActiveSheet()->getStyle('C322')->getAlignment()->setWrapText(TRUE);
        $excel->getActiveSheet()->getStyle('C322')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);
				$excel->setActiveSheetIndex(0)->setCellValue('R322', "No. 118/P/B/N/FMIPA/2019");
				$excel->setActiveSheetIndex(0)->setCellValue('R323', "Tanggal 27 Juni 2019");
				$excel->setActiveSheetIndex(0)->setCellValue('C325', "Aristoteles");

				$excel->setActiveSheetIndex(0)->setCellValue('B327', "5)");
				$excel->getActiveSheet()->getStyle('B327')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
				$excel->setActiveSheetIndex(0)->setCellValue('C327', "Prosiding Satek V 2013 Unila");
				$excel->setActiveSheetIndex(0)->setCellValue('L327', "30 November");
				$excel->setActiveSheetIndex(0)->setCellValue('M327', "Prosiding");
				$excel->setActiveSheetIndex(0)->setCellValue('N327', "1,00");
				$excel->setActiveSheetIndex(0)->setCellValue('R327', "http://satek.unila.ac.id/wp-content/uploads/2014/03/2-X9.pdf ");
				$excel->setActiveSheetIndex(0)->setCellValue('C328', "ISBN 978-979-8510-71-7");
				$excel->setActiveSheetIndex(0)->setCellValue('L328', "2013");
				$excel->setActiveSheetIndex(0)->setCellValue('C329', "30 November 2013");

				$excel->setActiveSheetIndex(0)->setCellValue('C331', "'Pengembangan E-Commerse T Menggunakan Sistem Database Terdistrubsi (Studi Kasus: Penjualan Dvd Game Terdistribusi)");
				$excel->getActiveSheet()->mergeCells('C331:K331');
        $excel->getActiveSheet()->getStyle('C331')->getAlignment()->setWrapText(TRUE);
        $excel->getActiveSheet()->getStyle('C331')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);
				$excel->setActiveSheetIndex(0)->setCellValue('R331', "No. 116/P/B/N/FMIPA/2019");
				$excel->setActiveSheetIndex(0)->setCellValue('R332', "Tanggal 27 Juni 2019");
				$excel->setActiveSheetIndex(0)->setCellValue('C334', "Favorisen R. Lumbanraja dan Aristoteles");

				$excel->setActiveSheetIndex(0)->setCellValue('B336', "6)");
				$excel->getActiveSheet()->getStyle('B336')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
				$excel->setActiveSheetIndex(0)->setCellValue('C336', "Prosiding SN SMAIP III 2012");
				$excel->setActiveSheetIndex(0)->setCellValue('L336', "Juni");
				$excel->setActiveSheetIndex(0)->setCellValue('M336', "Prosiding");
				$excel->setActiveSheetIndex(0)->setCellValue('N336', "1,00");
				$excel->setActiveSheetIndex(0)->setCellValue('R336', "http://repository.lppm.unila.ac.id/1368/ ");
				$excel->setActiveSheetIndex(0)->setCellValue('C337', "ISBN No. 978-602-98559-1-3");
				$excel->setActiveSheetIndex(0)->setCellValue('L337', "2012");
				$excel->setActiveSheetIndex(0)->setCellValue('C338', "Juni 2012");

				$excel->setActiveSheetIndex(0)->setCellValue('C340', "Implementasi Algoritma Half-Byte Dengan Nilai Parameter 7 Pada Kompresi File Gambar, Teks, Audio, Dan Video");
				$excel->getActiveSheet()->mergeCells('C340:K340');
        $excel->getActiveSheet()->getStyle('C340')->getAlignment()->setWrapText(TRUE);
        $excel->getActiveSheet()->getStyle('C340')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);
				$excel->setActiveSheetIndex(0)->setCellValue('R340', "No. 119/P/B/N/FMIPA/2019");
				$excel->setActiveSheetIndex(0)->setCellValue('R341', "Tanggal 27 Juni 2019");
				$excel->setActiveSheetIndex(0)->setCellValue('C343', "Anggar Bagus Kurniawan, Aristoteles, Machudor");
				$excel->setActiveSheetIndex(0)->setCellValue('N345', "Jumlah");
        $excel->setActiveSheetIndex(0)->setCellValue('O345', "0,00");
    		$excel->setActiveSheetIndex(0)->setCellValue('P345', "0,00");
    		$excel->getActiveSheet()->getStyle('N345:P345')->getFont()->setBold(TRUE);
				$excel->getActiveSheet()->getStyle('A345:R345')->applyFromArray($style_standar);

				$excel->setActiveSheetIndex(0)->setCellValue('E346', "2)");
				$excel->getActiveSheet()->getStyle('E346')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
				$excel->setActiveSheetIndex(0)->setCellValue('F346', "Poster tingkat:	");
				$excel->setActiveSheetIndex(0)->setCellValue('F347', "a) Internasional");
				$excel->setActiveSheetIndex(0)->setCellValue('N348', "Jumlah");
        $excel->setActiveSheetIndex(0)->setCellValue('O348', "0,00");
    		$excel->setActiveSheetIndex(0)->setCellValue('P348', "0,00");
    		$excel->getActiveSheet()->getStyle('N348:P348')->getFont()->setBold(TRUE);
				$excel->getActiveSheet()->getStyle('A348:R348')->applyFromArray($style_standar);
				$excel->setActiveSheetIndex(0)->setCellValue('F349', "b) Nasional");
				$excel->setActiveSheetIndex(0)->setCellValue('N350', "Jumlah");
        $excel->setActiveSheetIndex(0)->setCellValue('O350', "0,00");
    		$excel->setActiveSheetIndex(0)->setCellValue('P350', "0,00");
    		$excel->getActiveSheet()->getStyle('N350:P350')->getFont()->setBold(TRUE);
				$excel->getActiveSheet()->getStyle('A350:R350')->applyFromArray($style_standar);
				$excel->setActiveSheetIndex(0)->setCellValue('D351', "d.");
				$excel->getActiveSheet()->getStyle('D351')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
				$excel->setActiveSheetIndex(0)->setCellValue('E351', "Dalam koran/majalah populer/umum");
				$excel->getActiveSheet()->mergeCells('E351:K351');
        $excel->getActiveSheet()->getStyle('E351')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);
				$excel->setActiveSheetIndex(0)->setCellValue('N352', "Jumlah");
        $excel->setActiveSheetIndex(0)->setCellValue('O352', "0,00");
    		$excel->setActiveSheetIndex(0)->setCellValue('P352', "0,00");
    		$excel->getActiveSheet()->getStyle('N352:P352')->getFont()->setBold(TRUE);
				$excel->getActiveSheet()->getStyle('A352:R352')->applyFromArray($style_standar);
				$excel->setActiveSheetIndex(0)->setCellValue('C353', "2.");
				$excel->getActiveSheet()->getStyle('C353')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
				$excel->setActiveSheetIndex(0)->setCellValue('D353', "Hasil penelitian atau hasil pemikiran yang tidak di publikasikan (tersimpan di perpustakaan perguruan tinggi)");
				$excel->getActiveSheet()->mergeCells('D353:K353');
				$excel->setActiveSheetIndex(0)->setCellValue('N354', "Jumlah");
        $excel->setActiveSheetIndex(0)->setCellValue('O354', "0,00");
    		$excel->setActiveSheetIndex(0)->setCellValue('P354', "0,00");
    		$excel->getActiveSheet()->getStyle('N354:P354')->getFont()->setBold(TRUE);
				$excel->getActiveSheet()->getStyle('A354:R354')->applyFromArray($style_standar);
				$excel->setActiveSheetIndex(0)->setCellValue('B355', "B.");
				$excel->getActiveSheet()->getStyle('B355')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
        $excel->setActiveSheetIndex(0)->setCellValue('C355', "Menerjemahkan / menyadur buku ilmiah");
    		$excel->setActiveSheetIndex(0)->setCellValue('D356', "Diterbitkan dan diedarkan secara nasional.");
				$excel->setActiveSheetIndex(0)->setCellValue('N357', "Jumlah");
        $excel->setActiveSheetIndex(0)->setCellValue('O357', "0,00");
    		$excel->setActiveSheetIndex(0)->setCellValue('P357', "0,00");
    		$excel->getActiveSheet()->getStyle('N357:P357')->getFont()->setBold(TRUE);
				$excel->getActiveSheet()->getStyle('A357:R357')->applyFromArray($style_standar);
				$excel->setActiveSheetIndex(0)->setCellValue('B358', "C.");
				$excel->getActiveSheet()->getStyle('B358')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
        $excel->setActiveSheetIndex(0)->setCellValue('C358', "Mengedit/menyunting karya ilmiah");
    		$excel->setActiveSheetIndex(0)->setCellValue('D359', "Diterbitkan dan diedarkan secara nasional.");
				$excel->setActiveSheetIndex(0)->setCellValue('N360', "Jumlah");
        $excel->setActiveSheetIndex(0)->setCellValue('O360', "0,00");
    		$excel->setActiveSheetIndex(0)->setCellValue('P360', "0,00");
    		$excel->getActiveSheet()->getStyle('N360:P360')->getFont()->setBold(TRUE);
				$excel->getActiveSheet()->getStyle('A360:R360')->applyFromArray($style_standar);
				$excel->setActiveSheetIndex(0)->setCellValue('B361', "D.");
				$excel->getActiveSheet()->getStyle('B361')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
        $excel->setActiveSheetIndex(0)->setCellValue('C361', "Membuat rencana dan karya teknologi yang dipatenkan");
				$excel->setActiveSheetIndex(0)->setCellValue('C362', "1");
				$excel->getActiveSheet()->getStyle('C362')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
    		$excel->setActiveSheetIndex(0)->setCellValue('D362', "Internasional");
				$excel->setActiveSheetIndex(0)->setCellValue('N363', "Jumlah");
        $excel->setActiveSheetIndex(0)->setCellValue('O363', "0,00");
    		$excel->setActiveSheetIndex(0)->setCellValue('P363', "0,00");
    		$excel->getActiveSheet()->getStyle('N363:P363')->getFont()->setBold(TRUE);
				$excel->getActiveSheet()->getStyle('A363:R363')->applyFromArray($style_standar);
				$excel->setActiveSheetIndex(0)->setCellValue('C364', "2");
				$excel->getActiveSheet()->getStyle('C364')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
				$excel->setActiveSheetIndex(0)->setCellValue('D364', "Nasional");
				$excel->setActiveSheetIndex(0)->setCellValue('N365', "Jumlah");
        $excel->setActiveSheetIndex(0)->setCellValue('O365', "0,00");
    		$excel->setActiveSheetIndex(0)->setCellValue('P365', "0,00");
    		$excel->getActiveSheet()->getStyle('N365:P365')->getFont()->setBold(TRUE);
				$excel->getActiveSheet()->getStyle('A365:R365')->applyFromArray($style_standar);
				$excel->setActiveSheetIndex(0)->setCellValue('B366', "E.");
				$excel->getActiveSheet()->getStyle('B366')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
        $excel->setActiveSheetIndex(0)->setCellValue('C366', "Membuat rancangan dan karya teknologi, rancangan dan karya seni monumental/seni pertunjukan/karya sastra ");
				$excel->setActiveSheetIndex(0)->setCellValue('C367', "1");
				$excel->getActiveSheet()->getStyle('C367')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
    		$excel->setActiveSheetIndex(0)->setCellValue('D367', "Tingkat Internasional");
				$excel->setActiveSheetIndex(0)->setCellValue('N368', "Jumlah");
        $excel->setActiveSheetIndex(0)->setCellValue('O368', "0,00");
    		$excel->setActiveSheetIndex(0)->setCellValue('P368', "0,00");
    		$excel->getActiveSheet()->getStyle('N368:P368')->getFont()->setBold(TRUE);
				$excel->getActiveSheet()->getStyle('A368:R368')->applyFromArray($style_standar);
				$excel->setActiveSheetIndex(0)->setCellValue('D369', "Tingkat Nasional");
				$excel->setActiveSheetIndex(0)->setCellValue('N370', "Jumlah");
        $excel->setActiveSheetIndex(0)->setCellValue('O370', "0,00");
    		$excel->setActiveSheetIndex(0)->setCellValue('P370', "0,00");
    		$excel->getActiveSheet()->getStyle('N370:P370')->getFont()->setBold(TRUE);
				$excel->getActiveSheet()->getStyle('A370:R370')->applyFromArray($style_standar);
				$excel->setActiveSheetIndex(0)->setCellValue('D371', "Tingkat Lokal");
				$excel->setActiveSheetIndex(0)->setCellValue('N372', "Jumlah");
        $excel->setActiveSheetIndex(0)->setCellValue('O372', "0,00");
    		$excel->setActiveSheetIndex(0)->setCellValue('P372', "0,00");
    		$excel->getActiveSheet()->getStyle('N372:P372')->getFont()->setBold(TRUE);
				$excel->getActiveSheet()->getStyle('A372:R372')->applyFromArray($style_standar);
				$excel->setActiveSheetIndex(0)->setCellValue('A373', "Jumlah Penelitian");
				$excel->getActiveSheet()->mergeCells('A373:O373');
				$excel->getActiveSheet()->getStyle('A373')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
				$excel->setActiveSheetIndex(0)->setCellValue('P373', "0,00");
				$excel->getActiveSheet()->getStyle('A373:R373')->applyFromArray($style_standar);

				$excel->setActiveSheetIndex(0)->setCellValue('O375', "Bandar Lampung,  31 Juli 2019");
				$excel->setActiveSheetIndex(0)->setCellValue('O376', "Ketua Jurusan Ilmu Komputer");
				$excel->setActiveSheetIndex(0)->setCellValue('O382', "Dr.Ir. Kurnia Muludi, M.S.Sc");
				$excel->setActiveSheetIndex(0)->setCellValue('O383', "NIP. 19640616 198902 1 001");





        // Set orientasi kertas jadi LANDSCAPE
    		$excel->getActiveSheet()->getPageSetup()->setOrientation(PHPExcel_Worksheet_PageSetup::ORIENTATION_LANDSCAPE);

    		// Set judul file excel nya
    		$excel->getActiveSheet(0)->setTitle("Penelitian");
    		$excel->setActiveSheetIndex(0);

    		// Proses file excel
    		header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    		header('Content-Disposition: attachment; filename="Penelitian.xlsx"'); // Set nama file excel nya
    		header('Cache-Control: max-age=0');

    		$write = PHPExcel_IOFactory::createWriter($excel, 'Excel2007');
    		$write->save('php://output');

    	}

    }
