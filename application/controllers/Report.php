<?php
defined('BASEPATH') OR exit('No direct script access allowed');

class Report extends CI_Controller {

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
		$this->load->view('report', $data);
		}

  public function export(){
        // Load plugin PHPExcel nya
      include APPPATH.'third_party/PHPExcel/PHPExcel.php';

        // Panggil class PHPExcel nya
      $excel = new PHPExcel();

      // Settingan awal fil excel
      $excel->getProperties()->setCreator('Sulung')
                   ->setLastModifiedBy('Sulung')
                   ->setTitle("Report")
                   ->setSubject("Dupak")
                   ->setDescription("Laporan")
                   ->setKeywords("Report");

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
        $excel->setActiveSheetIndex(0)->setCellValue('A2', "MELAKSANAKAN PENDIDIKAN");


        $excel->getActiveSheet()->mergeCells('A1:L1'); // Set Merge Cell pada kolom A1 sampai G1
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

				//LAPORAN PENGABDIAN

				$excel->setActiveSheetIndex(0)->setCellValue('A587', "SURAT PERNYATAAN");
				$excel->setActiveSheetIndex(0)->setCellValue('A588', "MELAKSANAKAN PENGABDIAN KEPADA MASYARAKAT");
				$excel->getActiveSheet()->mergeCells('A587:R587');
				$excel->getActiveSheet()->mergeCells('A588:R588');
				$excel->getActiveSheet()->getStyle('A587:A588')->getFont()->setSize(11);
				$excel->getActiveSheet()->getStyle('A587:A588')->getFont()->setBold(TRUE);
				$excel->getActiveSheet()->getStyle('A587:A588')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);

				$excel->setActiveSheetIndex(0)->setCellValue('B590', "Yang bertanda tangan di bawah ini : ");
				$excel->setActiveSheetIndex(0)->setCellValue('B592', "Nama ");
				$excel->setActiveSheetIndex(0)->setCellValue('J592', ":");
				$excel->setActiveSheetIndex(0)->setCellValue('B593', "NIP");
				$excel->setActiveSheetIndex(0)->setCellValue('J593', ":");
				$excel->setActiveSheetIndex(0)->setCellValue('B594', "Pangkat");
				$excel->setActiveSheetIndex(0)->setCellValue('J594', ":");
				$excel->setActiveSheetIndex(0)->setCellValue('B595', "Golongan");
				$excel->setActiveSheetIndex(0)->setCellValue('J595', ":");
				$excel->setActiveSheetIndex(0)->setCellValue('B596', "Jabatan");
				$excel->setActiveSheetIndex(0)->setCellValue('J596', ":");
				$excel->setActiveSheetIndex(0)->setCellValue('B597', "Unit Kerja");
				$excel->setActiveSheetIndex(0)->setCellValue('J597', ":");
				$excel->setActiveSheetIndex(0)->setCellValue('B598', "Menyatakan ");
				$excel->setActiveSheetIndex(0)->setCellValue('B599', "Nama");
				$excel->setActiveSheetIndex(0)->setCellValue('J599', ":");
				$excel->setActiveSheetIndex(0)->setCellValue('B600', "NIP");
				$excel->setActiveSheetIndex(0)->setCellValue('J600', ":");
				$excel->setActiveSheetIndex(0)->setCellValue('B601', "Pangkat");
				$excel->setActiveSheetIndex(0)->setCellValue('J601', ":");
				$excel->setActiveSheetIndex(0)->setCellValue('B602', "Golongan");
				$excel->setActiveSheetIndex(0)->setCellValue('J602', ":");
				$excel->setActiveSheetIndex(0)->setCellValue('B603', "Jabatan");
				$excel->setActiveSheetIndex(0)->setCellValue('J603', ":");
				$excel->setActiveSheetIndex(0)->setCellValue('B604', "Unit Kerja");
				$excel->setActiveSheetIndex(0)->setCellValue('J604', ":");
				$excel->setActiveSheetIndex(0)->setCellValue('B606', "Telah menyatakan pengabdian masyarakat sebagai berikut : ");

				$data_dosen = $this->Dosen->view();

				foreach($data_dosen as $data) {

					$excel->setActiveSheetIndex(0)->setCellValue('K592', $data->nama);
					$excel->setActiveSheetIndex(0)->setCellValue('K593', $data->nip);
					$excel->setActiveSheetIndex(0)->setCellValue('K594', $data->pangkat);
					$excel->setActiveSheetIndex(0)->setCellValue('K595', $data->golongan);
					$excel->setActiveSheetIndex(0)->setCellValue('K596', $data->jabatan);
					$excel->setActiveSheetIndex(0)->setCellValue('K597', $data->unit_kerja);
				}

				$dosen_penunjang = $this->Lektor->view();

				foreach($dosen_penunjang as $data) {

					$excel->setActiveSheetIndex(0)->setCellValue('K599', $data->nama);
					$excel->setActiveSheetIndex(0)->setCellValue('K600', $data->nip);
					$excel->setActiveSheetIndex(0)->setCellValue('K601', $data->pangkat);
					$excel->setActiveSheetIndex(0)->setCellValue('K602', $data->golongan);
					$excel->setActiveSheetIndex(0)->setCellValue('K603', $data->jabatan);
					$excel->setActiveSheetIndex(0)->setCellValue('K604', $data->unit_kerja);
				}


				$excel->setActiveSheetIndex(0)->setCellValue('A608', "No");
				$excel->getActiveSheet()->getStyle('A608')->getAlignment()->setWrapText(TRUE);
				$excel->setActiveSheetIndex(0)->setCellValue('B608', "Uraian Kegiatan");
				$excel->getActiveSheet()->mergeCells('B608:K608');
				$excel->getActiveSheet()->getStyle('B608')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
				$excel->getActiveSheet()->getStyle('B608')->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
				$excel->setActiveSheetIndex(0)->setCellValue('L608', "Tanggal");
				$excel->setActiveSheetIndex(0)->setCellValue('M608', "Satuan Hasil");
				$excel->getActiveSheet()->getStyle('M608')->getAlignment()->setWrapText(TRUE);
				$excel->setActiveSheetIndex(0)->setCellValue('N608', "Jumlah Volume Kegiatan");
				$excel->getActiveSheet()->getStyle('N608')->getAlignment()->setWrapText(TRUE);
				$excel->setActiveSheetIndex(0)->setCellValue('O608', "Angka Kredit");
				$excel->getActiveSheet()->getStyle('O608')->getAlignment()->setWrapText(TRUE);
				$excel->setActiveSheetIndex(0)->setCellValue('P608', "Jumlah Angka Kredit");
				$excel->getActiveSheet()->getStyle('P608')->getAlignment()->setWrapText(TRUE);
				$excel->setActiveSheetIndex(0)->setCellValue('Q608', "Keterangan/Bukti Fisik");
				$excel->getActiveSheet()->mergeCells('Q608:R608');
				$excel->getActiveSheet()->getStyle('Q608')->getAlignment()->setWrapText(TRUE);

				$excel->getActiveSheet()->getStyle('A608:A654')->applyFromArray($style_col);
				$excel->getActiveSheet()->getStyle('L608:L654')->applyFromArray($style_col);
				$excel->getActiveSheet()->getStyle('M608:M654')->applyFromArray($style_col);
				$excel->getActiveSheet()->getStyle('N608:N654')->applyFromArray($style_col);
				$excel->getActiveSheet()->getStyle('O608:O654')->applyFromArray($style_col);
				$excel->getActiveSheet()->getStyle('P608:P654')->applyFromArray($style_col);
				$excel->getActiveSheet()->getStyle('R608:R654')->applyFromArray($style_col);
				$excel->getActiveSheet()->getStyle('A608')->applyFromArray($style_standar);
				$excel->getActiveSheet()->getStyle('B608')->applyFromArray($style_standar);
				$excel->getActiveSheet()->getStyle('C608')->applyFromArray($style_standar);
				$excel->getActiveSheet()->getStyle('D608')->applyFromArray($style_standar);
				$excel->getActiveSheet()->getStyle('E608')->applyFromArray($style_standar);
				$excel->getActiveSheet()->getStyle('F608')->applyFromArray($style_standar);
				$excel->getActiveSheet()->getStyle('G608')->applyFromArray($style_standar);
				$excel->getActiveSheet()->getStyle('H608')->applyFromArray($style_standar);
				$excel->getActiveSheet()->getStyle('I608')->applyFromArray($style_standar);
				$excel->getActiveSheet()->getStyle('J608')->applyFromArray($style_standar);
				$excel->getActiveSheet()->getStyle('K608')->applyFromArray($style_standar);
				$excel->getActiveSheet()->getStyle('L608:Q608')->applyFromArray($style_standar);
				$excel->getActiveSheet()->getStyle('R608')->applyFromArray($style_standar);

				$excel->setActiveSheetIndex(0)->setCellValue('A609', "(1)");
				$excel->setActiveSheetIndex(0)->setCellValue('B609', "(2)");
				$excel->getActiveSheet()->getStyle('B609')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
				$excel->getActiveSheet()->mergeCells('B609:K609');
				$excel->getActiveSheet()->getStyle('A609:Q609')->applyFromArray($style_standar);
				$excel->setActiveSheetIndex(0)->setCellValue('L609', "(3)");
				$excel->setActiveSheetIndex(0)->setCellValue('M609', "(4)");
				$excel->setActiveSheetIndex(0)->setCellValue('N609', "(5)");
				$excel->setActiveSheetIndex(0)->setCellValue('O609', "(6)");
				$excel->setActiveSheetIndex(0)->setCellValue('P609', "(7)");
				$excel->setActiveSheetIndex(0)->setCellValue('R609', "(8)");
				$excel->getActiveSheet()->getStyle('R609')->getAlignment()->setWrapText(TRUE);
				$excel->getActiveSheet()->getStyle('R609')->applyFromArray($style_standar);

				$excel->setActiveSheetIndex(0)->setCellValue('A610', "IV.");
				$excel->getActiveSheet()->getStyle('A610')->getFont()->setBold(TRUE);
				$excel->getActiveSheet()->getStyle('A610')->getFont()->setSize(11);
				$excel->setActiveSheetIndex(0)->setCellValue('B610', "MELAKSANAKAN PENGABDIAN KEPADA MASYARAKAT");
				$excel->getActiveSheet()->mergeCells('B610:G610');
				$excel->getActiveSheet()->getStyle('B610')->getFont()->setBold(TRUE);
				$excel->getActiveSheet()->getStyle('B610')->getFont()->setSize(11);
				$excel->setActiveSheetIndex(0)->setCellValue('B611', "A.");
				$excel->getActiveSheet()->getStyle('B611')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
				$excel->setActiveSheetIndex(0)->setCellValue('C611', "Menduduki jabatan pimpinan.");
				$excel->getActiveSheet()->mergeCells('C611:K611');
				$excel->setActiveSheetIndex(0)->setCellValue('D612', "Menduduki jabatan pimpinan dan lembaga");
				$excel->getActiveSheet()->mergeCells('D612:K612');
				$excel->setActiveSheetIndex(0)->setCellValue('O613', "Jumlah");
				$excel->setActiveSheetIndex(0)->setCellValue('P613', "0,00");
				$excel->getActiveSheet()->getStyle('O613:P613')->getFont()->setBold(TRUE);
				$excel->getActiveSheet()->getStyle('O613:P613')->getFont()->setSize(11);
				$excel->getActiveSheet()->getStyle('A613:R613')->applyFromArray($style_standar);

				$excel->setActiveSheetIndex(0)->setCellValue('B614', "B.");
				$excel->getActiveSheet()->getStyle('B614')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
				$excel->setActiveSheetIndex(0)->setCellValue('C614', "Melaksanakan pengembangan hasil pendidikan dan penelitian.");
				$excel->getActiveSheet()->mergeCells('C614:K614');
				$excel->setActiveSheetIndex(0)->setCellValue('D615', "Melaksanakan pengembangan hasil pendidikan dan penelitian");
				$excel->getActiveSheet()->mergeCells('D615:K615');
				$excel->setActiveSheetIndex(0)->setCellValue('O616', "Jumlah");
				$excel->setActiveSheetIndex(0)->setCellValue('P616', "0,00");
				$excel->getActiveSheet()->getStyle('O616:P616')->getFont()->setBold(TRUE);
				$excel->getActiveSheet()->getStyle('O616:P616')->getFont()->setSize(11);
				$excel->getActiveSheet()->getStyle('A616:R616')->applyFromArray($style_standar);

				$excel->setActiveSheetIndex(0)->setCellValue('B617', "C.");
				$excel->getActiveSheet()->getStyle('B617')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
				$excel->setActiveSheetIndex(0)->setCellValue('C617', "Memberi latihan/penyuluhan/penataran/ceramah kepada masyarakat.");
				$excel->getActiveSheet()->mergeCells('C617:K617');
				$excel->setActiveSheetIndex(0)->setCellValue('C618', "1.");
				$excel->getActiveSheet()->getStyle('C618')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
				$excel->setActiveSheetIndex(0)->setCellValue('D618', "Terjadwal/terprogram");
				$excel->getActiveSheet()->mergeCells('D618:F618');
				$excel->setActiveSheetIndex(0)->setCellValue('D619', "a.");
				$excel->getActiveSheet()->getStyle('D619')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
				$excel->setActiveSheetIndex(0)->setCellValue('E619', "Dalam satu semester atau lebih");
				$excel->getActiveSheet()->mergeCells('E619:K619');
				$excel->setActiveSheetIndex(0)->setCellValue('E620', "1) Tingkat Internasional");
				$excel->setActiveSheetIndex(0)->setCellValue('O621', "Jumlah");
				$excel->setActiveSheetIndex(0)->setCellValue('P621', "0,00");
				$excel->getActiveSheet()->getStyle('O621:P621')->getFont()->setBold(TRUE);
				$excel->getActiveSheet()->getStyle('O621:P621')->getFont()->setSize(11);
				$excel->getActiveSheet()->getStyle('A621:R621')->applyFromArray($style_standar);

				$excel->setActiveSheetIndex(0)->setCellValue('E622', "2) Tingkat Nasional");
				$excel->setActiveSheetIndex(0)->setCellValue('O622', "Jumlah");
				$excel->setActiveSheetIndex(0)->setCellValue('P622', "0,00");
				$excel->getActiveSheet()->getStyle('O622:P622')->getFont()->setBold(TRUE);
				$excel->getActiveSheet()->getStyle('O622:P622')->getFont()->setSize(11);
				$excel->getActiveSheet()->getStyle('A622:R622')->applyFromArray($style_standar);

				$excel->setActiveSheetIndex(0)->setCellValue('E623', "3) Tingkat Lokal");
				$excel->setActiveSheetIndex(0)->setCellValue('O623', "Jumlah");
				$excel->setActiveSheetIndex(0)->setCellValue('P623', "0,00");
				$excel->getActiveSheet()->getStyle('O623:P623')->getFont()->setBold(TRUE);
				$excel->getActiveSheet()->getStyle('O623:P623')->getFont()->setSize(11);
				$excel->getActiveSheet()->getStyle('A623:R623')->applyFromArray($style_standar);

				$excel->setActiveSheetIndex(0)->setCellValue('D624', "b.");
				$excel->getActiveSheet()->getStyle('D624')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
				$excel->setActiveSheetIndex(0)->setCellValue('E624', "Kurang dari satu semester dan minimal satu bulan.");
				$excel->getActiveSheet()->mergeCells('E624:K624');
				$excel->setActiveSheetIndex(0)->setCellValue('E625', "1) Tingkat Internasional");
				$excel->setActiveSheetIndex(0)->setCellValue('O625', "Jumlah");
				$excel->setActiveSheetIndex(0)->setCellValue('P625', "0,00");
				$excel->getActiveSheet()->getStyle('O625:P625')->getFont()->setBold(TRUE);
				$excel->getActiveSheet()->getStyle('O625:P625')->getFont()->setSize(11);
				$excel->getActiveSheet()->getStyle('A625:R625')->applyFromArray($style_standar);

				$excel->setActiveSheetIndex(0)->setCellValue('E626', "2) Tingkat Nasional");
				$excel->setActiveSheetIndex(0)->setCellValue('O626', "Jumlah");
				$excel->setActiveSheetIndex(0)->setCellValue('P626', "0,00");
				$excel->getActiveSheet()->getStyle('O626:P626')->getFont()->setBold(TRUE);
				$excel->getActiveSheet()->getStyle('O626:P626')->getFont()->setSize(11);
				$excel->getActiveSheet()->getStyle('A626:R626')->applyFromArray($style_standar);

				$excel->setActiveSheetIndex(0)->setCellValue('E627', "3) Tingkat Lokal");
				$excel->setActiveSheetIndex(0)->setCellValue('O627', "Jumlah");
				$excel->setActiveSheetIndex(0)->setCellValue('P627', "0,00");
				$excel->getActiveSheet()->getStyle('O627:P627')->getFont()->setBold(TRUE);
				$excel->getActiveSheet()->getStyle('O627:P627')->getFont()->setSize(11);
				$excel->getActiveSheet()->getStyle('A627:R627')->applyFromArray($style_standar);

				$excel->setActiveSheetIndex(0)->setCellValue('C628', "2.");
				$excel->getActiveSheet()->getStyle('C628')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
				$excel->setActiveSheetIndex(0)->setCellValue('D628', "Insidental :");
				$excel->getActiveSheet()->mergeCells('D628:K628');
				$excel->setActiveSheetIndex(0)->setCellValue('C629', "1)");
				$excel->getActiveSheet()->getStyle('C629')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
				$excel->setActiveSheetIndex(0)->setCellValue('D629', "Pelatihan Desain Grafis untuk Usaha Kecil Menengah Desa Wawasan Kecamatan Tanjung Sari Kabupaten Lampung Selatan");
				$excel->setActiveSheetIndex(0)->setCellValue('L629', "23 Agst 2018");
				$excel->setActiveSheetIndex(0)->setCellValue('M629', "Laporan");
				$excel->setActiveSheetIndex(0)->setCellValue('N629', "1,0");
				$excel->setActiveSheetIndex(0)->setCellValue('O629', "1,0");
				$excel->setActiveSheetIndex(0)->setCellValue('P629', "1,0");
				$excel->setActiveSheetIndex(0)->setCellValue('R629', "IV.C.2");
				$excel->setActiveSheetIndex(0)->setCellValue('M629', "Kegiatan");
				$excel->setActiveSheetIndex(0)->setCellValue('R629', "Laporan Kegiatan");

				$excel->setActiveSheetIndex(0)->setCellValue('C633', "2)");
				$excel->getActiveSheet()->getStyle('C633')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
				$excel->setActiveSheetIndex(0)->setCellValue('D633', "Pelatihan Adobe Photosop dan Corel Draw untuk pembuatan alat promosi Sekolah bagi Guru-Guru SMS di Bandar Lampung");
				$excel->setActiveSheetIndex(0)->setCellValue('L633', "29 Nov 2014");
				$excel->setActiveSheetIndex(0)->setCellValue('M633', "Laporan");
				$excel->setActiveSheetIndex(0)->setCellValue('N633', "1,0");
				$excel->setActiveSheetIndex(0)->setCellValue('O633', "1,0");
				$excel->setActiveSheetIndex(0)->setCellValue('P633', "1,0");
				$excel->setActiveSheetIndex(0)->setCellValue('R633', "IV.C.2");
				$excel->setActiveSheetIndex(0)->setCellValue('M634', "Kegiatan");
				$excel->setActiveSheetIndex(0)->setCellValue('R634', "Laporan Kegiatan");

				$excel->setActiveSheetIndex(0)->setCellValue('C638', "3)");
				$excel->getActiveSheet()->getStyle('C638')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
				$excel->setActiveSheetIndex(0)->setCellValue('D638', "Penerapan Media Pembelajaran Interaktif Pengenalan Komputer di SDN 1 Kupang Teba Kota Bandar Lampung");
				$excel->setActiveSheetIndex(0)->setCellValue('L638', "22 Nov 2014");
				$excel->setActiveSheetIndex(0)->setCellValue('M638', "Laporan");
				$excel->setActiveSheetIndex(0)->setCellValue('N638', "1,0");
				$excel->setActiveSheetIndex(0)->setCellValue('O638', "1,0");
				$excel->setActiveSheetIndex(0)->setCellValue('P638', "1,0");
				$excel->setActiveSheetIndex(0)->setCellValue('R638', "IV.C.2");
				$excel->setActiveSheetIndex(0)->setCellValue('M639', "Kegiatan");
				$excel->setActiveSheetIndex(0)->setCellValue('R639', "Laporan Kegiatan");

				$excel->setActiveSheetIndex(0)->setCellValue('C642', "4)");
				$excel->getActiveSheet()->getStyle('C642')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
				$excel->setActiveSheetIndex(0)->setCellValue('D642', "Implementasi Sistem Informasi Akademik di SMUN 1 Gedong Tataan Kabupaten Pesawaran");
				$excel->setActiveSheetIndex(0)->setCellValue('L642', "8-9 Okt 2013");
				$excel->setActiveSheetIndex(0)->setCellValue('M642', "Laporan");
				$excel->setActiveSheetIndex(0)->setCellValue('N642', "1,0");
				$excel->setActiveSheetIndex(0)->setCellValue('O642', "1,0");
				$excel->setActiveSheetIndex(0)->setCellValue('P642', "1,0");
				$excel->setActiveSheetIndex(0)->setCellValue('R642', "IV.C.2");
				$excel->setActiveSheetIndex(0)->setCellValue('M643', "Kegiatan");
				$excel->setActiveSheetIndex(0)->setCellValue('R643', "Laporan Kegiatan");

				$excel->setActiveSheetIndex(0)->setCellValue('O646', "Jumlah");
				$excel->setActiveSheetIndex(0)->setCellValue('P646', "4,00");
				$excel->getActiveSheet()->getStyle('O646:P646')->getFont()->setBold(TRUE);
				$excel->getActiveSheet()->getStyle('O646:P646')->getFont()->setSize(11);
				$excel->getActiveSheet()->getStyle('A646:R646')->applyFromArray($style_standar);


				$excel->setActiveSheetIndex(0)->setCellValue('B647', "D.");
				$excel->getActiveSheet()->getStyle('B647')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
				$excel->setActiveSheetIndex(0)->setCellValue('C647', "Memberi pelayanan kepada masyarakat atau kegiatan lain yang menunjang pelaksanaan tugas umum pemerintah dan pembangunan.");
				$excel->getActiveSheet()->mergeCells('C647:K647');
				$excel->getActiveSheet()->getStyle('C647')->getAlignment()->setWrapText(TRUE);
				$excel->setActiveSheetIndex(0)->setCellValue('C648', "1.");
				$excel->getActiveSheet()->getStyle('C648')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
				$excel->setActiveSheetIndex(0)->setCellValue('D648', "Berdasarkan bidang keahlian.");
				$excel->getActiveSheet()->mergeCells('D648:K648');
				$excel->setActiveSheetIndex(0)->setCellValue('C649', "2.");
				$excel->getActiveSheet()->getStyle('C649')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
				$excel->setActiveSheetIndex(0)->setCellValue('D649', "Berdasarkan penugasan lembaga perguruan tinggi.");
				$excel->getActiveSheet()->mergeCells('D649:K649');
				$excel->setActiveSheetIndex(0)->setCellValue('C650', "3.");
				$excel->getActiveSheet()->getStyle('C650')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
				$excel->setActiveSheetIndex(0)->setCellValue('D650', "Berdasarkan fungsi/jabatan.");
				$excel->getActiveSheet()->mergeCells('D650:K650');

				$excel->setActiveSheetIndex(0)->setCellValue('B651', "E.");
				$excel->getActiveSheet()->getStyle('B651')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
				$excel->setActiveSheetIndex(0)->setCellValue('C651', "Membuat/menulis karya pengabdian.");
				$excel->getActiveSheet()->mergeCells('C651:K651');
				$excel->setActiveSheetIndex(0)->setCellValue('D652', "Membuat/menulis karya pengabdian pada masyarakat yang tidak dipublikasikan.");
				$excel->getActiveSheet()->mergeCells('D652:K652');
				$excel->setActiveSheetIndex(0)->setCellValue('O653', "Jumlah");
				$excel->setActiveSheetIndex(0)->setCellValue('P653', "0,00");
				$excel->getActiveSheet()->getStyle('O653')->getFont()->setBold(TRUE);
				$excel->getActiveSheet()->getStyle('O653')->getFont()->setSize(11);
				$excel->getActiveSheet()->getStyle('A653:R653')->applyFromArray($style_standar);
				$excel->setActiveSheetIndex(0)->setCellValue('A654', "Jumlah Pengabdian kepada Masyarakat");
				$excel->getActiveSheet()->mergeCells('A654:O654');
				$excel->setActiveSheetIndex(0)->setCellValue('P654', "4,00");
				$excel->getActiveSheet()->getStyle('A654:P654')->getFont()->setBold(TRUE);
				$excel->getActiveSheet()->getStyle('A654:P654')->getFont()->setSize(11);
				$excel->getActiveSheet()->getStyle('A654:R654')->applyFromArray($style_standar);

				$excel->setActiveSheetIndex(0)->setCellValue('O656', "Bandar Lampung,  31 Juli 2019");
				$excel->setActiveSheetIndex(0)->setCellValue('O657', "Ketua Jurusan Ilmu Komputer");
				$excel->setActiveSheetIndex(0)->setCellValue('O661', "Dr.Ir. Kurnia Muludi, M.S.Sc");
				$excel->setActiveSheetIndex(0)->setCellValue('O662', "NIP. 19640616 198902 1 001");


				$excel->setActiveSheetIndex(0)->setCellValue('A665', "SURAT PERNYATAAN");
				$excel->setActiveSheetIndex(0)->setCellValue('A666', "MELAKSANAKAN PENUNJANG TUGAS DOSEN");
				$excel->getActiveSheet()->mergeCells('A665:R665');
				$excel->getActiveSheet()->mergeCells('A666:R666');
				$excel->getActiveSheet()->getStyle('A665:A666')->getFont()->setBold(TRUE);
				$excel->getActiveSheet()->getStyle('A665:A666')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);

				$excel->setActiveSheetIndex(0)->setCellValue('B667', "Nama ");
				$excel->setActiveSheetIndex(0)->setCellValue('J667', ":");
				$excel->setActiveSheetIndex(0)->setCellValue('B668', "NIP");
				$excel->setActiveSheetIndex(0)->setCellValue('J668', ":");
				$excel->setActiveSheetIndex(0)->setCellValue('B669', "Pangkat");
				$excel->setActiveSheetIndex(0)->setCellValue('J669', ":");
				$excel->setActiveSheetIndex(0)->setCellValue('B670', "Golongan");
				$excel->setActiveSheetIndex(0)->setCellValue('J670', ":");
				$excel->setActiveSheetIndex(0)->setCellValue('B671', "Jabatan");
				$excel->setActiveSheetIndex(0)->setCellValue('J671', ":");
				$excel->setActiveSheetIndex(0)->setCellValue('B672', "Unit Kerja");
				$excel->setActiveSheetIndex(0)->setCellValue('J672', ":");
				$excel->setActiveSheetIndex(0)->setCellValue('B673', "Menyatakan ");
				$excel->setActiveSheetIndex(0)->setCellValue('B674', "Nama");
				$excel->setActiveSheetIndex(0)->setCellValue('J674', ":");
				$excel->setActiveSheetIndex(0)->setCellValue('B675', "NIP");
				$excel->setActiveSheetIndex(0)->setCellValue('J675', ":");
				$excel->setActiveSheetIndex(0)->setCellValue('B676', "Pangkat");
				$excel->setActiveSheetIndex(0)->setCellValue('J676', ":");
				$excel->setActiveSheetIndex(0)->setCellValue('B677', "Golongan");
				$excel->setActiveSheetIndex(0)->setCellValue('J677', ":");
				$excel->setActiveSheetIndex(0)->setCellValue('B678', "Jabatan");
				$excel->setActiveSheetIndex(0)->setCellValue('J678', ":");
				$excel->setActiveSheetIndex(0)->setCellValue('B679', "Unit Kerja");
				$excel->setActiveSheetIndex(0)->setCellValue('J679', ":");
				$excel->setActiveSheetIndex(0)->setCellValue('B681', "Telah menyatakan pengabdian masyarakat sebagai berikut : ");

				$data_dosen = $this->Dosen->view();

				foreach($data_dosen as $data) {

					$excel->setActiveSheetIndex(0)->setCellValue('K667', $data->nama);
					$excel->setActiveSheetIndex(0)->setCellValue('K668', $data->nip);
					$excel->setActiveSheetIndex(0)->setCellValue('K669', $data->pangkat);
					$excel->setActiveSheetIndex(0)->setCellValue('K670', $data->golongan);
					$excel->setActiveSheetIndex(0)->setCellValue('K671', $data->jabatan);
					$excel->setActiveSheetIndex(0)->setCellValue('K672', $data->unit_kerja);
				}

				$dosen_penunjang = $this->Lektor->view();

				foreach($dosen_penunjang as $data) {

					$excel->setActiveSheetIndex(0)->setCellValue('K674', $data->nama);
					$excel->setActiveSheetIndex(0)->setCellValue('K675', $data->nip);
					$excel->setActiveSheetIndex(0)->setCellValue('K676', $data->pangkat);
					$excel->setActiveSheetIndex(0)->setCellValue('K677', $data->golongan);
					$excel->setActiveSheetIndex(0)->setCellValue('K678', $data->jabatan);
					$excel->setActiveSheetIndex(0)->setCellValue('K679', $data->unit_kerja);
				}


				$excel->setActiveSheetIndex(0)->setCellValue('A683', "No");
				$excel->getActiveSheet()->getStyle('A683')->getAlignment()->setWrapText(TRUE);
				$excel->setActiveSheetIndex(0)->setCellValue('B683', "Uraian Kegiatan");
				$excel->getActiveSheet()->mergeCells('B683:K683');
				$excel->getActiveSheet()->getStyle('B683')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
				$excel->getActiveSheet()->getStyle('B683')->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
				$excel->setActiveSheetIndex(0)->setCellValue('L683', "Tanggal");
				$excel->setActiveSheetIndex(0)->setCellValue('M683', "Satuan Hasil");
				$excel->getActiveSheet()->getStyle('M683')->getAlignment()->setWrapText(TRUE);
				$excel->setActiveSheetIndex(0)->setCellValue('N683', "Jumlah Volume Kegiatan");
				$excel->getActiveSheet()->getStyle('N683')->getAlignment()->setWrapText(TRUE);
				$excel->setActiveSheetIndex(0)->setCellValue('O683', "Angka Kredit");
				$excel->getActiveSheet()->getStyle('O683')->getAlignment()->setWrapText(TRUE);
				$excel->setActiveSheetIndex(0)->setCellValue('P683', "Jumlah Angka Kredit");
				$excel->getActiveSheet()->getStyle('P683')->getAlignment()->setWrapText(TRUE);
				$excel->setActiveSheetIndex(0)->setCellValue('Q683', "Keterangan/Bukti Fisik");
				$excel->getActiveSheet()->mergeCells('Q683:R683');
				$excel->getActiveSheet()->getStyle('Q683')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
				$excel->getActiveSheet()->getStyle('Q683')->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
				$excel->getActiveSheet()->getStyle('Q683')->getAlignment()->setWrapText(TRUE);

				$excel->getActiveSheet()->getStyle('A683:A863')->applyFromArray($style_col);
				$excel->getActiveSheet()->getStyle('L683:L863')->applyFromArray($style_col);
				$excel->getActiveSheet()->getStyle('M683:M863')->applyFromArray($style_col);
				$excel->getActiveSheet()->getStyle('N683:N863')->applyFromArray($style_col);
				$excel->getActiveSheet()->getStyle('O683:O863')->applyFromArray($style_col);
				$excel->getActiveSheet()->getStyle('P683:P863')->applyFromArray($style_col);
				$excel->getActiveSheet()->getStyle('R684:R863')->applyFromArray($style_col);
				$excel->getActiveSheet()->getStyle('A683')->applyFromArray($style_standar);
				$excel->getActiveSheet()->getStyle('B683')->applyFromArray($style_standar);
				$excel->getActiveSheet()->getStyle('C683')->applyFromArray($style_standar);
				$excel->getActiveSheet()->getStyle('D683')->applyFromArray($style_standar);
				$excel->getActiveSheet()->getStyle('E683')->applyFromArray($style_standar);
				$excel->getActiveSheet()->getStyle('F683')->applyFromArray($style_standar);
				$excel->getActiveSheet()->getStyle('G683')->applyFromArray($style_standar);
				$excel->getActiveSheet()->getStyle('H683')->applyFromArray($style_standar);
				$excel->getActiveSheet()->getStyle('I683')->applyFromArray($style_standar);
				$excel->getActiveSheet()->getStyle('J683')->applyFromArray($style_standar);
				$excel->getActiveSheet()->getStyle('K683')->applyFromArray($style_standar);
				$excel->getActiveSheet()->getStyle('L683:Q683')->applyFromArray($style_standar);
				$excel->getActiveSheet()->getStyle('R683')->applyFromArray($style_standar);

				$excel->setActiveSheetIndex(0)->setCellValue('A684', "(1)");
				$excel->setActiveSheetIndex(0)->setCellValue('B684', "(2)");
				$excel->getActiveSheet()->getStyle('B684')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
				$excel->getActiveSheet()->mergeCells('B684:K684');
				$excel->getActiveSheet()->getStyle('A684:Q684')->applyFromArray($style_standar);
				$excel->setActiveSheetIndex(0)->setCellValue('L684', "(3)");
				$excel->setActiveSheetIndex(0)->setCellValue('M684', "(4)");
				$excel->setActiveSheetIndex(0)->setCellValue('N684', "(5)");
				$excel->setActiveSheetIndex(0)->setCellValue('O684', "(6)");
				$excel->setActiveSheetIndex(0)->setCellValue('P684', "(7)");
				$excel->setActiveSheetIndex(0)->setCellValue('R684', "(8)");
				$excel->getActiveSheet()->getStyle('R684')->getAlignment()->setWrapText(TRUE);
				$excel->getActiveSheet()->getStyle('R684')->applyFromArray($style_standar);

				$excel->setActiveSheetIndex(0)->setCellValue('A685', "IV.");
				$excel->getActiveSheet()->getStyle('A685')->getFont()->setBold(TRUE);
				$excel->getActiveSheet()->getStyle('A685')->getFont()->setSize(11);
				$excel->setActiveSheetIndex(0)->setCellValue('B685', "PENUNJANG TUGAS");
				$excel->getActiveSheet()->mergeCells('B685:K685');
				$excel->getActiveSheet()->getStyle('B685')->getFont()->setBold(TRUE);
				$excel->getActiveSheet()->getStyle('B685')->getFont()->setSize(11);
				$excel->getActiveSheet()->getStyle('A685:R685')->applyFromArray($style_standar);

				$excel->setActiveSheetIndex(0)->setCellValue('A686', "A");
				$excel->setActiveSheetIndex(0)->setCellValue('B686', "Menjadi anggota dalam suatu Panitia/Badan");
				$excel->getActiveSheet()->mergeCells('B686:K686');
				$excel->setActiveSheetIndex(0)->setCellValue('B687', "2.");
				$excel->getActiveSheet()->getStyle('B687')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
				$excel->setActiveSheetIndex(0)->setCellValue('C687', "Sebagai anggota");
				$excel->setActiveSheetIndex(0)->setCellValue('Q687', "VI.A.2");
				$excel->setActiveSheetIndex(0)->setCellValue('C688', "1)");
				$excel->getActiveSheet()->getStyle('C688')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
				$excel->setActiveSheetIndex(0)->setCellValue('D688', "Anggota Panitia Tim Penjamin Mutu Program Sarjana (S1) FMIPA Unila ");
				$excel->setActiveSheetIndex(0)->setCellValue('L688', "Sem Ganjil ");
				$excel->setActiveSheetIndex(0)->setCellValue('M688', "1 Semester");
				$excel->setActiveSheetIndex(0)->setCellValue('N688', "1,0");
				$excel->setActiveSheetIndex(0)->setCellValue('O688', "1,0");
				$excel->setActiveSheetIndex(0)->setCellValue('P688', "1,0");
				$excel->setActiveSheetIndex(0)->setCellValue('Q688', "1");
				$excel->setActiveSheetIndex(0)->setCellValue('R688', "SK Dekan FMIPA Unila");
				$excel->setActiveSheetIndex(0)->setCellValue('L689', "2013/2014");
				$excel->setActiveSheetIndex(0)->setCellValue('R689', "No 473/UN26/7/DT/2013");
				$excel->setActiveSheetIndex(0)->setCellValue('R690', "Tanggal 4 Januari 2013");

				$excel->setActiveSheetIndex(0)->setCellValue('C692', "2)");
				$excel->getActiveSheet()->getStyle('C692')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
				$excel->setActiveSheetIndex(0)->setCellValue('D692', "Anggota panitia Seminar dan Rapat Tahunan Bidang Ilmu MIPA (Semirata BKS PTN Barat)");
				$excel->setActiveSheetIndex(0)->setCellValue('L692', "Sem Ganjil ");
				$excel->setActiveSheetIndex(0)->setCellValue('M692', "1 Semester");
				$excel->setActiveSheetIndex(0)->setCellValue('N692', "1,0");
				$excel->setActiveSheetIndex(0)->setCellValue('O692', "1,0");
				$excel->setActiveSheetIndex(0)->setCellValue('P692', "1,0");
				$excel->setActiveSheetIndex(0)->setCellValue('Q692', "1");
				$excel->setActiveSheetIndex(0)->setCellValue('R692', "SK Dekan FMIPA Unila");
				$excel->setActiveSheetIndex(0)->setCellValue('L693', "2013/2014");
				$excel->setActiveSheetIndex(0)->setCellValue('R693', "4517/UN26/7/DT/2013");
				$excel->setActiveSheetIndex(0)->setCellValue('R694', "Tanggal 20 Maret 2013");

				$excel->setActiveSheetIndex(0)->setCellValue('C696', "3)");
				$excel->getActiveSheet()->getStyle('C696')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
				$excel->setActiveSheetIndex(0)->setCellValue('D696', "Anggota Panitia Tim Penjamin Mutu Program Sarjana (S1) FMIPA Unila ");
				$excel->setActiveSheetIndex(0)->setCellValue('L696', "Sem Ganjil ");
				$excel->setActiveSheetIndex(0)->setCellValue('M696', "1 Semester");
				$excel->setActiveSheetIndex(0)->setCellValue('N696', "1,0");
				$excel->setActiveSheetIndex(0)->setCellValue('O696', "1,0");
				$excel->setActiveSheetIndex(0)->setCellValue('P696', "1,0");
				$excel->setActiveSheetIndex(0)->setCellValue('Q696', "1");
				$excel->setActiveSheetIndex(0)->setCellValue('R696', "SK Dekan FMIPA Unila");
				$excel->setActiveSheetIndex(0)->setCellValue('L697', "2013/2014");
				$excel->setActiveSheetIndex(0)->setCellValue('R697', "No 1462a/UN26/7/DT/2013");
				$excel->setActiveSheetIndex(0)->setCellValue('R698', "Tanggal 1 Oktober 2013");

				$excel->setActiveSheetIndex(0)->setCellValue('C700', "4)");
				$excel->getActiveSheet()->getStyle('C700')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
				$excel->setActiveSheetIndex(0)->setCellValue('D700', "Anggota Tim Verifikasi Rekayasa Perangkat Lunak Tugas Akhir D3 MI Semester Ganjil ");
				$excel->setActiveSheetIndex(0)->setCellValue('L700', "Sem Ganjil ");
				$excel->setActiveSheetIndex(0)->setCellValue('M700', "1 Semester");
				$excel->setActiveSheetIndex(0)->setCellValue('N700', "1,0");
				$excel->setActiveSheetIndex(0)->setCellValue('O700', "1,0");
				$excel->setActiveSheetIndex(0)->setCellValue('P700', "1,0");
				$excel->setActiveSheetIndex(0)->setCellValue('Q700', "1");
				$excel->setActiveSheetIndex(0)->setCellValue('R700', "SK Dekan FMIPA Unila");
				$excel->setActiveSheetIndex(0)->setCellValue('L701', "2013/2014");
				$excel->setActiveSheetIndex(0)->setCellValue('R701', "No 2604/UN26/7/DT/2013");
				$excel->setActiveSheetIndex(0)->setCellValue('R702', "Tanggal 7 Oktober 2013");

				$excel->setActiveSheetIndex(0)->setCellValue('C704', "5)");
				$excel->getActiveSheet()->getStyle('C704')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
				$excel->setActiveSheetIndex(0)->setCellValue('D704', "Tim Penilai Sertifikasi Dosen FMIPA Unila");
				$excel->setActiveSheetIndex(0)->setCellValue('L704', "Sem Ganjil ");
				$excel->setActiveSheetIndex(0)->setCellValue('M704', "1 Semester");
				$excel->setActiveSheetIndex(0)->setCellValue('N704', "1,0");
				$excel->setActiveSheetIndex(0)->setCellValue('O704', "1,0");
				$excel->setActiveSheetIndex(0)->setCellValue('P704', "1,0");
				$excel->setActiveSheetIndex(0)->setCellValue('Q704', "1");
				$excel->setActiveSheetIndex(0)->setCellValue('R704', "SK Dekan FMIPA Unila");
				$excel->setActiveSheetIndex(0)->setCellValue('L705', "2013/2014");
				$excel->setActiveSheetIndex(0)->setCellValue('R705', "No 2679/UN26/7/DT/2013");
				$excel->setActiveSheetIndex(0)->setCellValue('R706', "Tanggal 18 Oktober 2013");

				$excel->setActiveSheetIndex(0)->setCellValue('C708', "6)");
				$excel->getActiveSheet()->getStyle('C708')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
				$excel->setActiveSheetIndex(0)->setCellValue('D708', "Anggota Panitia Seminar Nasional Sain dan Teknologi (SATEK) V ");
				$excel->setActiveSheetIndex(0)->setCellValue('L708', "Sem Ganjil ");
				$excel->setActiveSheetIndex(0)->setCellValue('M708', "1 Semester");
				$excel->setActiveSheetIndex(0)->setCellValue('N708', "1,0");
				$excel->setActiveSheetIndex(0)->setCellValue('O708', "1,0");
				$excel->setActiveSheetIndex(0)->setCellValue('P708', "1,0");
				$excel->setActiveSheetIndex(0)->setCellValue('Q708', "1");
				$excel->setActiveSheetIndex(0)->setCellValue('R708', "SK Rektor Unila");
				$excel->setActiveSheetIndex(0)->setCellValue('L709', "2013/2014");
				$excel->setActiveSheetIndex(0)->setCellValue('R709', "No 767/UN26/LP/2013");
				$excel->setActiveSheetIndex(0)->setCellValue('R710', "Tanggal November 2013");

				$excel->setActiveSheetIndex(0)->setCellValue('C712', "7)");
				$excel->getActiveSheet()->getStyle('C712')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
				$excel->setActiveSheetIndex(0)->setCellValue('D712', "Anggota Tim Verifikasi Rekayasa Perangkat Lunak Tugas Akhir D3 MI Semester Genap");
				$excel->setActiveSheetIndex(0)->setCellValue('L712', "Sem Genap");
				$excel->setActiveSheetIndex(0)->setCellValue('M712', "1 Semester");
				$excel->setActiveSheetIndex(0)->setCellValue('N712', "1,0");
				$excel->setActiveSheetIndex(0)->setCellValue('O712', "1,0");
				$excel->setActiveSheetIndex(0)->setCellValue('P712', "1,0");
				$excel->setActiveSheetIndex(0)->setCellValue('Q712', "1");
				$excel->setActiveSheetIndex(0)->setCellValue('R712', "SK Dekan FMIPA Unila");
				$excel->setActiveSheetIndex(0)->setCellValue('L713', "2013/2014");
				$excel->setActiveSheetIndex(0)->setCellValue('R713', "No 135a/UN26/7/DT/2014");
				$excel->setActiveSheetIndex(0)->setCellValue('R714', "Tanggal 14 Januari 2014");

				$excel->setActiveSheetIndex(0)->setCellValue('C717', "8)");
				$excel->getActiveSheet()->getStyle('C717')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
				$excel->setActiveSheetIndex(0)->setCellValue('D717', "Tim Audit Internal ISO 9001:2008 Jurusan Ilmu Komputer FMIPA Unila");
				$excel->setActiveSheetIndex(0)->setCellValue('L717', "Sem Genap");
				$excel->setActiveSheetIndex(0)->setCellValue('M717', "1 Semester");
				$excel->setActiveSheetIndex(0)->setCellValue('N717', "1,0");
				$excel->setActiveSheetIndex(0)->setCellValue('O717', "1,0");
				$excel->setActiveSheetIndex(0)->setCellValue('P717', "1,0");
				$excel->setActiveSheetIndex(0)->setCellValue('Q717', "1");
				$excel->setActiveSheetIndex(0)->setCellValue('R717', "SK Dekan FMIPA UNILA");
				$excel->setActiveSheetIndex(0)->setCellValue('L718', "2014/2015");
				$excel->setActiveSheetIndex(0)->setCellValue('R718', "No. 476a/UN26/7/KM/2014");
				$excel->setActiveSheetIndex(0)->setCellValue('R719', "Tanggal 3 Maret 2014");

				$excel->setActiveSheetIndex(0)->setCellValue('C721', "9)");
				$excel->getActiveSheet()->getStyle('C721')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
				$excel->setActiveSheetIndex(0)->setCellValue('D721', "Tim Jaminan Mutu ISO 9001:2008 Jurusan Ilmu Komputer FMIPA Unila");
				$excel->setActiveSheetIndex(0)->setCellValue('L721', "Sem Genap");
				$excel->setActiveSheetIndex(0)->setCellValue('M721', "1 Semester");
				$excel->setActiveSheetIndex(0)->setCellValue('N721', "1,0");
				$excel->setActiveSheetIndex(0)->setCellValue('O721', "1,0");
				$excel->setActiveSheetIndex(0)->setCellValue('P721', "1,0");
				$excel->setActiveSheetIndex(0)->setCellValue('Q721', "1");
				$excel->setActiveSheetIndex(0)->setCellValue('R721', "SK Dekan FMIPA UNILA");
				$excel->setActiveSheetIndex(0)->setCellValue('L722', "2014/2015");
				$excel->setActiveSheetIndex(0)->setCellValue('R722', "No. 483a/UN26/7/KM/2014");
				$excel->setActiveSheetIndex(0)->setCellValue('R723', "Tanggal 4 Maret 2014");

				$excel->setActiveSheetIndex(0)->setCellValue('C725', "10)");
				$excel->getActiveSheet()->getStyle('C725')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
				$excel->setActiveSheetIndex(0)->setCellValue('D725', "Anggota Tim Verifikasi Rekayasa Perangkat Lunak Tugas Akhir D3 MI Semester Ganjil");
				$excel->setActiveSheetIndex(0)->setCellValue('L725', "Sem Ganjil");
				$excel->setActiveSheetIndex(0)->setCellValue('M725', "1 Semester");
				$excel->setActiveSheetIndex(0)->setCellValue('N725', "1,0");
				$excel->setActiveSheetIndex(0)->setCellValue('O725', "1,0");
				$excel->setActiveSheetIndex(0)->setCellValue('P725', "1,0");
				$excel->setActiveSheetIndex(0)->setCellValue('Q725', "1");
				$excel->setActiveSheetIndex(0)->setCellValue('R725', "SK Dekan FMIPA UNILA");
				$excel->setActiveSheetIndex(0)->setCellValue('L726', "2014/2015");
				$excel->setActiveSheetIndex(0)->setCellValue('R726', "No 2012/UN26/7/DT/2014");
				$excel->setActiveSheetIndex(0)->setCellValue('R727', "Tanggal 9 Oktober 2014");

				$excel->setActiveSheetIndex(0)->setCellValue('C730', "11)");
				$excel->getActiveSheet()->getStyle('C730')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
				$excel->setActiveSheetIndex(0)->setCellValue('D730', "Anggota tim audit pembelajaran program sarjana Jurusan Ilmu Komputer FMIPA Unila");
				$excel->setActiveSheetIndex(0)->setCellValue('L730', "Smstr Genap");
				$excel->setActiveSheetIndex(0)->setCellValue('M730', "1 Semester");
				$excel->setActiveSheetIndex(0)->setCellValue('N730', "1,0");
				$excel->setActiveSheetIndex(0)->setCellValue('O730', "1,0");
				$excel->setActiveSheetIndex(0)->setCellValue('P730', "1,0");
				$excel->setActiveSheetIndex(0)->setCellValue('Q730', "1");
				$excel->setActiveSheetIndex(0)->setCellValue('R730', "SK Dekan FMIPA UNILA");
				$excel->setActiveSheetIndex(0)->setCellValue('L731', "2014/2015");
				$excel->setActiveSheetIndex(0)->setCellValue('R731', "No. 16a/UN26/7/DT/2015");
				$excel->setActiveSheetIndex(0)->setCellValue('R732', "Tanggal 7 Januari 2015");

				$excel->setActiveSheetIndex(0)->setCellValue('C733', "12)");
				$excel->getActiveSheet()->getStyle('C733')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
				$excel->setActiveSheetIndex(0)->setCellValue('D733', "Pengurus Badan Pelaksana Kuliah Kerja Nyata (BP-KKN) Unila");
				$excel->setActiveSheetIndex(0)->setCellValue('L733', "Smstr Genap");
				$excel->setActiveSheetIndex(0)->setCellValue('M733', "1 Semester");
				$excel->setActiveSheetIndex(0)->setCellValue('N733', "1,0");
				$excel->setActiveSheetIndex(0)->setCellValue('O733', "1,0");
				$excel->setActiveSheetIndex(0)->setCellValue('P733', "1,0");
				$excel->setActiveSheetIndex(0)->setCellValue('Q733', "1");
				$excel->setActiveSheetIndex(0)->setCellValue('R733', "SK Rektor Unila");
				$excel->setActiveSheetIndex(0)->setCellValue('L734', "2014/2015");
				$excel->setActiveSheetIndex(0)->setCellValue('R734', "No 140/UN26/KP/2015");
				$excel->setActiveSheetIndex(0)->setCellValue('R735', "Tanggal 24 Maret 2015");

				$excel->setActiveSheetIndex(0)->setCellValue('C738', "13)");
				$excel->getActiveSheet()->getStyle('C738')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
				$excel->setActiveSheetIndex(0)->setCellValue('D738', "Tim Auditor ISO 9001:2008 Jurusan Ilmu Komputer FMIPA Unila");
				$excel->setActiveSheetIndex(0)->setCellValue('L738', "Smstr Genap");
				$excel->setActiveSheetIndex(0)->setCellValue('M738', "1 Semester");
				$excel->setActiveSheetIndex(0)->setCellValue('N738', "1,0");
				$excel->setActiveSheetIndex(0)->setCellValue('O738', "1,0");
				$excel->setActiveSheetIndex(0)->setCellValue('P738', "1,0");
				$excel->setActiveSheetIndex(0)->setCellValue('Q738', "1");
				$excel->setActiveSheetIndex(0)->setCellValue('R738', "SK Dekan FMIPA UNILA");
				$excel->setActiveSheetIndex(0)->setCellValue('L739', "2014/2015");
				$excel->setActiveSheetIndex(0)->setCellValue('R739', "No. 1637/UN26/7/KM/2014");
				$excel->setActiveSheetIndex(0)->setCellValue('R740', "Tanggal 4 Mei 2015");

				$excel->setActiveSheetIndex(0)->setCellValue('C742', "14)");
				$excel->getActiveSheet()->getStyle('C742')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
				$excel->setActiveSheetIndex(0)->setCellValue('D742', "Anggota Panitia Tim Penjamin Mutu Program Sarjana (S1) FMIPA Unila ");
				$excel->setActiveSheetIndex(0)->setCellValue('L742', "Smstr Genap");
				$excel->setActiveSheetIndex(0)->setCellValue('M742', "1 Semester");
				$excel->setActiveSheetIndex(0)->setCellValue('N742', "1,0");
				$excel->setActiveSheetIndex(0)->setCellValue('O742', "1,0");
				$excel->setActiveSheetIndex(0)->setCellValue('P742', "1,0");
				$excel->setActiveSheetIndex(0)->setCellValue('Q742', "1");
				$excel->setActiveSheetIndex(0)->setCellValue('R742', "SK Dekan FMIPA UNILA");
				$excel->setActiveSheetIndex(0)->setCellValue('L743', "2014/2015");
				$excel->setActiveSheetIndex(0)->setCellValue('R743', "No 1642a/UN26/7/DT/2015");
				$excel->setActiveSheetIndex(0)->setCellValue('R744', "Tanggal 5 Mei 2015");

				$excel->setActiveSheetIndex(0)->setCellValue('C746', "15)");
				$excel->getActiveSheet()->getStyle('C746')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
				$excel->setActiveSheetIndex(0)->setCellValue('D746', "Anggota tim penyusun akreditasi program sarjana Jurusan Ilmu Komputer FMIPA Unila");
				$excel->setActiveSheetIndex(0)->setCellValue('L746', "Smstr Genap");
				$excel->setActiveSheetIndex(0)->setCellValue('M746', "1 Semester");
				$excel->setActiveSheetIndex(0)->setCellValue('N746', "1,0");
				$excel->setActiveSheetIndex(0)->setCellValue('O746', "1,0");
				$excel->setActiveSheetIndex(0)->setCellValue('P746', "1,0");
				$excel->setActiveSheetIndex(0)->setCellValue('Q746', "1");
				$excel->setActiveSheetIndex(0)->setCellValue('R746', "SK Dekan FMIPA UNILA");
				$excel->setActiveSheetIndex(0)->setCellValue('L747', "2014/2015");
				$excel->setActiveSheetIndex(0)->setCellValue('R747', "No. 1941/UN26/7/DT/2015");
				$excel->setActiveSheetIndex(0)->setCellValue('R748', "Tanggal 15 Juni 2015");

				$excel->setActiveSheetIndex(0)->setCellValue('C750', "16)");
				$excel->getActiveSheet()->getStyle('C750')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
				$excel->setActiveSheetIndex(0)->setCellValue('D750', "Juri Mahasiswa Berprestasi FMIPA Unila");
				$excel->setActiveSheetIndex(0)->setCellValue('L750', "Smstr Genap");
				$excel->setActiveSheetIndex(0)->setCellValue('M750', "1 Semester");
				$excel->setActiveSheetIndex(0)->setCellValue('N750', "1,0");
				$excel->setActiveSheetIndex(0)->setCellValue('O750', "1,0");
				$excel->setActiveSheetIndex(0)->setCellValue('P750', "1,0");
				$excel->setActiveSheetIndex(0)->setCellValue('Q750', "1");
				$excel->setActiveSheetIndex(0)->setCellValue('R750', "SK Dekan FMIPA UNILA");
				$excel->setActiveSheetIndex(0)->setCellValue('L751', "2015/2016");
				$excel->setActiveSheetIndex(0)->setCellValue('R751', "No. 636/UN26/7/KM/2016");
				$excel->setActiveSheetIndex(0)->setCellValue('R752', "Tanggal 21 Maret 2016");

				$excel->setActiveSheetIndex(0)->setCellValue('C754', "17)");
				$excel->getActiveSheet()->getStyle('C754')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
				$excel->setActiveSheetIndex(0)->setCellValue('D754', "Anggota Tim Verifikasi Rekayasa Perangkat Lunak Tugas Akhir D3 MI Semester Genap");
				$excel->setActiveSheetIndex(0)->setCellValue('L754', "Smstr Genap");
				$excel->setActiveSheetIndex(0)->setCellValue('M754', "1 Semester");
				$excel->setActiveSheetIndex(0)->setCellValue('N754', "1,0");
				$excel->setActiveSheetIndex(0)->setCellValue('O754', "1,0");
				$excel->setActiveSheetIndex(0)->setCellValue('P754', "1,0");
				$excel->setActiveSheetIndex(0)->setCellValue('Q754', "1");
				$excel->setActiveSheetIndex(0)->setCellValue('R754', "SK Dekan FMIPA UNILA");
				$excel->setActiveSheetIndex(0)->setCellValue('L755', "2015/2016");
				$excel->setActiveSheetIndex(0)->setCellValue('R755', "No 771/UN26/7/DT/2014");
				$excel->setActiveSheetIndex(0)->setCellValue('R756', "Tanggal 7 April 2016");

				$excel->setActiveSheetIndex(0)->setCellValue('C758', "18)");
				$excel->getActiveSheet()->getStyle('C758')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
				$excel->setActiveSheetIndex(0)->setCellValue('D758', "Anggota tim penyusun akreditasi program sarjana Jurusan Ilmu Komputer FMIPA Unila");
				$excel->setActiveSheetIndex(0)->setCellValue('L758', "Smstr Genap");
				$excel->setActiveSheetIndex(0)->setCellValue('M758', "1 Semester");
				$excel->setActiveSheetIndex(0)->setCellValue('N758', "1,0");
				$excel->setActiveSheetIndex(0)->setCellValue('O758', "1,0");
				$excel->setActiveSheetIndex(0)->setCellValue('P758', "1,0");
				$excel->setActiveSheetIndex(0)->setCellValue('Q758', "1");
				$excel->setActiveSheetIndex(0)->setCellValue('R758', "SK Dekan FMIPA UNILA");
				$excel->setActiveSheetIndex(0)->setCellValue('L759', "2015/2016");
				$excel->setActiveSheetIndex(0)->setCellValue('R759', "No. 814/UN26/7/DT/2016");
				$excel->setActiveSheetIndex(0)->setCellValue('R760', "Tanggal 13 April 2016");

				$excel->setActiveSheetIndex(0)->setCellValue('C762', "19)");
				$excel->getActiveSheet()->getStyle('C762')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
				$excel->setActiveSheetIndex(0)->setCellValue('D762', "Anggota Tim Pengelola Lokakarya Revisi Kurikulum PS S1 Ilmu Komputer FMIPA Unila");
				$excel->setActiveSheetIndex(0)->setCellValue('L762', "Smstr Genap");
				$excel->setActiveSheetIndex(0)->setCellValue('M762', "1 Semester");
				$excel->setActiveSheetIndex(0)->setCellValue('N762', "1,0");
				$excel->setActiveSheetIndex(0)->setCellValue('O762', "1,0");
				$excel->setActiveSheetIndex(0)->setCellValue('P762', "1,0");
				$excel->setActiveSheetIndex(0)->setCellValue('Q762', "1");
				$excel->setActiveSheetIndex(0)->setCellValue('R762', "SK Dekan FMIPA UNILA");
				$excel->setActiveSheetIndex(0)->setCellValue('L763', "2015/2016");
				$excel->setActiveSheetIndex(0)->setCellValue('R763', "No 961/UN26/7/DT/2016");
				$excel->setActiveSheetIndex(0)->setCellValue('R764', "Tanggal 26 April 2016");

				$excel->setActiveSheetIndex(0)->setCellValue('C766', "20)");
				$excel->getActiveSheet()->getStyle('C766')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
				$excel->setActiveSheetIndex(0)->setCellValue('D766', "Anggota Tim Pengelola Lokakarya Revisi Kurikulum PS D3 Manajemen Informatika FMIPA Unila");
				$excel->setActiveSheetIndex(0)->setCellValue('L766', "Smstr Genap");
				$excel->setActiveSheetIndex(0)->setCellValue('M766', "1 Semester");
				$excel->setActiveSheetIndex(0)->setCellValue('N766', "1,0");
				$excel->setActiveSheetIndex(0)->setCellValue('O766', "1,0");
				$excel->setActiveSheetIndex(0)->setCellValue('P766', "1,0");
				$excel->setActiveSheetIndex(0)->setCellValue('Q766', "1");
				$excel->setActiveSheetIndex(0)->setCellValue('R766', "SK Dekan FMIPA UNILA");
				$excel->setActiveSheetIndex(0)->setCellValue('L767', "2015/2016");
				$excel->setActiveSheetIndex(0)->setCellValue('R767', "No 963/UN26/7/DT/2016");
				$excel->setActiveSheetIndex(0)->setCellValue('R768', "Tanggal 28 April 2016");

				$excel->setActiveSheetIndex(0)->setCellValue('C770', "21)");
				$excel->getActiveSheet()->getStyle('C770')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
				$excel->setActiveSheetIndex(0)->setCellValue('D770', "Tim Audit Internal ISO 9001:2008 Jurusan Ilmu Komputer FMIPA Unila");
				$excel->setActiveSheetIndex(0)->setCellValue('L770', "Smstr Ganjil");
				$excel->setActiveSheetIndex(0)->setCellValue('M770', "1 Semester");
				$excel->setActiveSheetIndex(0)->setCellValue('N770', "1,0");
				$excel->setActiveSheetIndex(0)->setCellValue('O770', "1,0");
				$excel->setActiveSheetIndex(0)->setCellValue('P770', "1,0");
				$excel->setActiveSheetIndex(0)->setCellValue('Q770', "1");
				$excel->setActiveSheetIndex(0)->setCellValue('R770', "SK Dekan FMIPA UNILA");
				$excel->setActiveSheetIndex(0)->setCellValue('L771', "2016/2017");
				$excel->setActiveSheetIndex(0)->setCellValue('R771', "No. 2597/UN26/7/DT/2016");
				$excel->setActiveSheetIndex(0)->setCellValue('R772', "Tanggal 24 Oktober 2016");

				$excel->setActiveSheetIndex(0)->setCellValue('C774', "22)");
				$excel->getActiveSheet()->getStyle('C774')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
				$excel->setActiveSheetIndex(0)->setCellValue('D774', "Pengurus Badan Pelaksana Kuliah Kerja Nyata (BP-KKN) Unila");
				$excel->setActiveSheetIndex(0)->setCellValue('L774', "Smstr Genap");
				$excel->setActiveSheetIndex(0)->setCellValue('M774', "1 Semester");
				$excel->setActiveSheetIndex(0)->setCellValue('N774', "1,0");
				$excel->setActiveSheetIndex(0)->setCellValue('O774', "1,0");
				$excel->setActiveSheetIndex(0)->setCellValue('P774', "1,0");
				$excel->setActiveSheetIndex(0)->setCellValue('Q774', "1");
				$excel->setActiveSheetIndex(0)->setCellValue('R774', "SK Rektor Unila");
				$excel->setActiveSheetIndex(0)->setCellValue('L775', "2016/2017");
				$excel->setActiveSheetIndex(0)->setCellValue('R775', "No. 129/UN26/KP/2017");
				$excel->setActiveSheetIndex(0)->setCellValue('R776', "Tanggal 01 Februari 2017");

				$excel->setActiveSheetIndex(0)->setCellValue('C779', "23)");
				$excel->getActiveSheet()->getStyle('C779')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
				$excel->setActiveSheetIndex(0)->setCellValue('D779', "Anggota Panitia Seleksi Dosen Kontrak FMIPA Unila");
				$excel->setActiveSheetIndex(0)->setCellValue('L779', "Smstr Ganjil");
				$excel->setActiveSheetIndex(0)->setCellValue('M779', "1 Semester");
				$excel->setActiveSheetIndex(0)->setCellValue('N779', "1,0");
				$excel->setActiveSheetIndex(0)->setCellValue('O779', "1,0");
				$excel->setActiveSheetIndex(0)->setCellValue('P779', "1,0");
				$excel->setActiveSheetIndex(0)->setCellValue('Q779', "1");
				$excel->setActiveSheetIndex(0)->setCellValue('R779', "SK Dekan FMIPA Unila");
				$excel->setActiveSheetIndex(0)->setCellValue('L780', "2017/2018");
				$excel->setActiveSheetIndex(0)->setCellValue('R780', "No 2723/UN26/7/KP/2017");
				$excel->setActiveSheetIndex(0)->setCellValue('R781', "Tanggal 15 Juni 2017");

				$excel->setActiveSheetIndex(0)->setCellValue('C782', "24)");
				$excel->getActiveSheet()->getStyle('C782')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
				$excel->setActiveSheetIndex(0)->setCellValue('D782', "Anggota Tim Pelaksana Sie Koreksi Tugas Mahasiswa Kegiatan Program Pengenalan Kehidupan Kampus bagi Mahasiswa baru Unila");
				$excel->setActiveSheetIndex(0)->setCellValue('L782', "Smstr Genap");
				$excel->setActiveSheetIndex(0)->setCellValue('M782', "1 Semester");
				$excel->setActiveSheetIndex(0)->setCellValue('N782', "1,0");
				$excel->setActiveSheetIndex(0)->setCellValue('O782', "1,0");
				$excel->setActiveSheetIndex(0)->setCellValue('P782', "1,0");
				$excel->setActiveSheetIndex(0)->setCellValue('Q782', "1");
				$excel->setActiveSheetIndex(0)->setCellValue('R782', "SK Rektor Unila");
				$excel->setActiveSheetIndex(0)->setCellValue('L783', "2017/2018");
				$excel->setActiveSheetIndex(0)->setCellValue('R783', "No 952/UN26/DT/2017");
				$excel->setActiveSheetIndex(0)->setCellValue('R784', "Tanggal 07 Agustus 2017");

				$excel->setActiveSheetIndex(0)->setCellValue('C786', "25)");
				$excel->getActiveSheet()->getStyle('C786')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
				$excel->setActiveSheetIndex(0)->setCellValue('D786', "Anggota Tim Verifikasi Rekayasa Perangkat Lunak Tugas Akhir D3 MI Semester Ganjil");
				$excel->setActiveSheetIndex(0)->setCellValue('L786', "Smstr Ganjil");
				$excel->setActiveSheetIndex(0)->setCellValue('M786', "1 Semester");
				$excel->setActiveSheetIndex(0)->setCellValue('N786', "1,0");
				$excel->setActiveSheetIndex(0)->setCellValue('O786', "1,0");
				$excel->setActiveSheetIndex(0)->setCellValue('P786', "1,0");
				$excel->setActiveSheetIndex(0)->setCellValue('Q786', "1");
				$excel->setActiveSheetIndex(0)->setCellValue('R786', "SK Dekan FMIPA Unila");
				$excel->setActiveSheetIndex(0)->setCellValue('L787', "2017/2018");
				$excel->setActiveSheetIndex(0)->setCellValue('R787', "No 4216/UN26/7/DT/2017");
				$excel->setActiveSheetIndex(0)->setCellValue('R788', "Tanggal 27 Oktober 2017");

				$excel->setActiveSheetIndex(0)->setCellValue('C790', "26)");
				$excel->getActiveSheet()->getStyle('C790')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
				$excel->setActiveSheetIndex(0)->setCellValue('D790', "Anggota IT Kuliah Kerja Nyata Kebangsaan ");
				$excel->setActiveSheetIndex(0)->setCellValue('L790', "Smstr Genap");
				$excel->setActiveSheetIndex(0)->setCellValue('M790', "1 Semester");
				$excel->setActiveSheetIndex(0)->setCellValue('N790', "1,0");
				$excel->setActiveSheetIndex(0)->setCellValue('O790', "1,0");
				$excel->setActiveSheetIndex(0)->setCellValue('P790', "1,0");
				$excel->setActiveSheetIndex(0)->setCellValue('Q790', "1");
				$excel->setActiveSheetIndex(0)->setCellValue('R790', "SK Rektor Unila");
				$excel->setActiveSheetIndex(0)->setCellValue('L791', "2017/2018");
				$excel->setActiveSheetIndex(0)->setCellValue('R791', "No 1247/UN26/PM.03/2018");
				$excel->setActiveSheetIndex(0)->setCellValue('R792', "Tanggal 08 Juni 2018");

				$excel->setActiveSheetIndex(0)->setCellValue('C794', "27)");
				$excel->getActiveSheet()->getStyle('C794')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
				$excel->setActiveSheetIndex(0)->setCellValue('D794', "Kepala Divisi pada UPT Bahasa Unila");
				$excel->setActiveSheetIndex(0)->setCellValue('L794', "Smstr Ganjil");
				$excel->setActiveSheetIndex(0)->setCellValue('M794', "1 Semester");
				$excel->setActiveSheetIndex(0)->setCellValue('N794', "1,0");
				$excel->setActiveSheetIndex(0)->setCellValue('O794', "1,0");
				$excel->setActiveSheetIndex(0)->setCellValue('P794', "1,0");
				$excel->setActiveSheetIndex(0)->setCellValue('Q794', "1");
				$excel->setActiveSheetIndex(0)->setCellValue('R794', "SK Rektor Unila");
				$excel->setActiveSheetIndex(0)->setCellValue('L795', "2018/2019");
				$excel->setActiveSheetIndex(0)->setCellValue('R795', "No 1903/UN26/KP/2018");
				$excel->setActiveSheetIndex(0)->setCellValue('R796', "Tanggal 26 Oktober 2018");

				$excel->setActiveSheetIndex(0)->setCellValue('O798', "Jumlah");
				$excel->setActiveSheetIndex(0)->setCellValue('P798', "27,00");
				$excel->getActiveSheet()->getStyle('O798:P798')->getFont()->setBold(TRUE);
				$excel->getActiveSheet()->getStyle('O798:P798')->getFont()->setSize(11);
				$excel->getActiveSheet()->getStyle('A798:R798')->applyFromArray($style_standar);

				$excel->setActiveSheetIndex(0)->setCellValue('A799', "B");
				$excel->setActiveSheetIndex(0)->setCellValue('B799', "Menjadi anggota panitia/badan pada lembaga pemerintah");
				$excel->setActiveSheetIndex(0)->setCellValue('B800', "1.");
				$excel->getActiveSheet()->getStyle('B800')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
				$excel->setActiveSheetIndex(0)->setCellValue('C800', "Panitia pusat");
				$excel->setActiveSheetIndex(0)->setCellValue('C801', "a.");
				$excel->getActiveSheet()->getStyle('C801')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
				$excel->setActiveSheetIndex(0)->setCellValue('D801', "Ketua/Wakil Ketua");
				$excel->setActiveSheetIndex(0)->setCellValue('C802', "b.");
				$excel->getActiveSheet()->getStyle('C802')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
				$excel->setActiveSheetIndex(0)->setCellValue('D802', "Anggota");
				$excel->setActiveSheetIndex(0)->setCellValue('B803', "2.");
				$excel->getActiveSheet()->getStyle('B803')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
				$excel->setActiveSheetIndex(0)->setCellValue('C803', "Panitia daerah");
				$excel->setActiveSheetIndex(0)->setCellValue('C804', "a.");
				$excel->getActiveSheet()->getStyle('C804')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
				$excel->setActiveSheetIndex(0)->setCellValue('D804', "Ketua/Wakil Ketua");
				$excel->setActiveSheetIndex(0)->setCellValue('C805', "b.");
				$excel->getActiveSheet()->getStyle('C805')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
				$excel->setActiveSheetIndex(0)->setCellValue('D805', "Anggota");
				$excel->getActiveSheet()->getStyle('A805:R805')->applyFromArray($style_standar);

				$excel->setActiveSheetIndex(0)->setCellValue('A806', "C");
				$excel->setActiveSheetIndex(0)->setCellValue('B806', "Menjadi anggota organisasi profesi");
				$excel->setActiveSheetIndex(0)->setCellValue('B807', "1.");
				$excel->getActiveSheet()->getStyle('B807')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
				$excel->setActiveSheetIndex(0)->setCellValue('C807', "Tingkat Internasional");
				$excel->getActiveSheet()->mergeCells('C807:F807');
				$excel->setActiveSheetIndex(0)->setCellValue('C808', "a");
				$excel->getActiveSheet()->getStyle('C808')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
				$excel->setActiveSheetIndex(0)->setCellValue('D808', "Pengurus");
				$excel->setActiveSheetIndex(0)->setCellValue('C809', "b");
				$excel->getActiveSheet()->getStyle('C809')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
				$excel->setActiveSheetIndex(0)->setCellValue('D809', "Anggota atas permintaan");
				$excel->setActiveSheetIndex(0)->setCellValue('C810', "c");
				$excel->getActiveSheet()->getStyle('C810')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
				$excel->setActiveSheetIndex(0)->setCellValue('D810', "Anggota");
				$excel->setActiveSheetIndex(0)->setCellValue('B811', "2.");
				$excel->getActiveSheet()->getStyle('B811')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
				$excel->setActiveSheetIndex(0)->setCellValue('C811', "Tingkat Nasional");
				$excel->getActiveSheet()->mergeCells('C811:F811');
				$excel->setActiveSheetIndex(0)->setCellValue('C812', "a");
				$excel->getActiveSheet()->getStyle('C812')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
				$excel->setActiveSheetIndex(0)->setCellValue('D812', "Pengurus");
				$excel->setActiveSheetIndex(0)->setCellValue('C813', "b");
				$excel->getActiveSheet()->getStyle('C813')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
				$excel->setActiveSheetIndex(0)->setCellValue('D813', "Anggota atas permintaan");
				$excel->setActiveSheetIndex(0)->setCellValue('C814', "c");
				$excel->getActiveSheet()->getStyle('C814')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
				$excel->setActiveSheetIndex(0)->setCellValue('D814', "Anggota");
				$excel->setActiveSheetIndex(0)->setCellValue('Q814', "VI.C2.c");
				$excel->setActiveSheetIndex(0)->setCellValue('D815', "Anggota Ikatan Ahli Informatika");
				$excel->setActiveSheetIndex(0)->setCellValue('L815', "Periode");
				$excel->setActiveSheetIndex(0)->setCellValue('M815', "Setiap");
				$excel->setActiveSheetIndex(0)->setCellValue('N815', "1");
				$excel->setActiveSheetIndex(0)->setCellValue('O815', "0,50");
				$excel->setActiveSheetIndex(0)->setCellValue('P815', "0,50");
				$excel->setActiveSheetIndex(0)->setCellValue('R815', "Kartu Anggota IAII");
				$excel->setActiveSheetIndex(0)->setCellValue('D816', "Indonesia");
				$excel->setActiveSheetIndex(0)->setCellValue('L816', "2016-2018");
				$excel->setActiveSheetIndex(0)->setCellValue('M816', "Periode");
				$excel->setActiveSheetIndex(0)->setCellValue('R816', "No. 16.10.10002");
				$excel->setActiveSheetIndex(0)->setCellValue('O817', "Jumlah");
				$excel->setActiveSheetIndex(0)->setCellValue('P817', "0,50");
				$excel->getActiveSheet()->getStyle('O817:P817')->getFont()->setBold(TRUE);
				$excel->getActiveSheet()->getStyle('O817:P817')->getFont()->setSize(11);
				$excel->getActiveSheet()->getStyle('A817:R817')->applyFromArray($style_standar);

				$excel->setActiveSheetIndex(0)->setCellValue('A818', "D");
				$excel->setActiveSheetIndex(0)->setCellValue('B818', "Mewakili perguruan tinggi/lembaga pemerintah");
				$excel->getActiveSheet()->mergeCells('B818:K818');
				$excel->setActiveSheetIndex(0)->setCellValue('C819', "Mewakili perguruan tinggi/lembaga pemerintah duduk dalam panitia antar lembaga");
				$excel->getActiveSheet()->mergeCells('C819:K819');
				$excel->getActiveSheet()->getStyle('A819:R819')->applyFromArray($style_standar);

				$excel->setActiveSheetIndex(0)->setCellValue('A820', "E");
				$excel->setActiveSheetIndex(0)->setCellValue('B820', "Menjadi anggota delegasi nasional ke pertemuan internasional");
				$excel->setActiveSheetIndex(0)->setCellValue('B821', "1.");
				$excel->getActiveSheet()->getStyle('B821')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
				$excel->setActiveSheetIndex(0)->setCellValue('C821', "Sebagai ketua delegasi");
				$excel->getActiveSheet()->mergeCells('C821:K821');
				$excel->setActiveSheetIndex(0)->setCellValue('B822', "2.");
				$excel->getActiveSheet()->getStyle('B822')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
				$excel->setActiveSheetIndex(0)->setCellValue('C822', "Sebagai anggota delegasi");
				$excel->getActiveSheet()->mergeCells('C822:K822');
				$excel->getActiveSheet()->getStyle('A822:R822')->applyFromArray($style_standar);

				$excel->setActiveSheetIndex(0)->setCellValue('A823', "F");
				$excel->setActiveSheetIndex(0)->setCellValue('B823', "Berperan serta aktif dalam pertemuan ilmiah");
				$excel->getActiveSheet()->mergeCells('B823:K823');
				$excel->setActiveSheetIndex(0)->setCellValue('B824', "1.");
				$excel->getActiveSheet()->getStyle('B824')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
				$excel->setActiveSheetIndex(0)->setCellValue('C824', "Tingkat internasional/nasional/regional sebagai :");
				$excel->getActiveSheet()->mergeCells('C824:K824');
				$excel->setActiveSheetIndex(0)->setCellValue('C825', "a.");
				$excel->getActiveSheet()->getStyle('C825')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
				$excel->setActiveSheetIndex(0)->setCellValue('D825', "Ketua");
				$excel->getActiveSheet()->mergeCells('D825:F825');
				$excel->getActiveSheet()->getStyle('A826:R826')->applyFromArray($style_standar);
				$excel->setActiveSheetIndex(0)->setCellValue('C827', "b.");
				$excel->getActiveSheet()->getStyle('C827')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
				$excel->setActiveSheetIndex(0)->setCellValue('D827', "Anggota");
				$excel->getActiveSheet()->mergeCells('D827:F827');
				$excel->setActiveSheetIndex(0)->setCellValue('Q827', "VI.F1.b");
				$excel->getActiveSheet()->getStyle('A827:R827')->applyFromArray($style_standar);
				$excel->getActiveSheet()->getStyle('A828:R828')->applyFromArray($style_standar);
				$excel->setActiveSheetIndex(0)->setCellValue('B829', "2.");
				$excel->getActiveSheet()->getStyle('B829')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
				$excel->setActiveSheetIndex(0)->setCellValue('C829', "Di lingkungan perguruan tinggi sebagai :");
				$excel->setActiveSheetIndex(0)->setCellValue('C830', "a.");
				$excel->getActiveSheet()->getStyle('C830')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
				$excel->setActiveSheetIndex(0)->setCellValue('D830', "Ketua");
				$excel->getActiveSheet()->mergeCells('D830:F830');
				$excel->setActiveSheetIndex(0)->setCellValue('O831', "Jumlah");
				$excel->setActiveSheetIndex(0)->setCellValue('P831', "0,00");
				$excel->getActiveSheet()->getStyle('O831:P831')->getFont()->setBold(TRUE);
				$excel->getActiveSheet()->getStyle('O831:P831')->getFont()->setSize(11);
				$excel->getActiveSheet()->getStyle('A831:R831')->applyFromArray($style_standar);
				$excel->setActiveSheetIndex(0)->setCellValue('C832', "b.");
				$excel->getActiveSheet()->getStyle('C832')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
				$excel->setActiveSheetIndex(0)->setCellValue('D832', "Anggota");
				$excel->getActiveSheet()->mergeCells('D832:F832');
				$excel->setActiveSheetIndex(0)->setCellValue('O834', "Jumlah");
				$excel->setActiveSheetIndex(0)->setCellValue('P834', "0,00");
				$excel->getActiveSheet()->getStyle('O834:P834')->getFont()->setBold(TRUE);
				$excel->getActiveSheet()->getStyle('O834:P834')->getFont()->setSize(11);
				$excel->getActiveSheet()->getStyle('A834:R834')->applyFromArray($style_standar);
				$excel->setActiveSheetIndex(0)->setCellValue('A835', "G");
				$excel->setActiveSheetIndex(0)->setCellValue('B835', "Mendapat penghargaan/ tanda jasa");
				$excel->getActiveSheet()->mergeCells('B835:K835');
				$excel->setActiveSheetIndex(0)->setCellValue('B836', "1.");
				$excel->getActiveSheet()->getStyle('B836')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
				$excel->setActiveSheetIndex(0)->setCellValue('C836', "Penghargaan/tanda jasa Satya Lancana Karya Satya");
				$excel->getActiveSheet()->mergeCells('C836:K836');
				$excel->setActiveSheetIndex(0)->setCellValue('C837', "a");
				$excel->getActiveSheet()->getStyle('C837')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
				$excel->setActiveSheetIndex(0)->setCellValue('D837', "30 (tiga puluh) tahun");
				$excel->getActiveSheet()->mergeCells('D837:K837');
				$excel->getActiveSheet()->getStyle('A838:R838')->applyFromArray($style_standar);
				$excel->setActiveSheetIndex(0)->setCellValue('C839', "b");
				$excel->getActiveSheet()->getStyle('C839')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
				$excel->setActiveSheetIndex(0)->setCellValue('D839', "20 (dua puluh) tahun");
				$excel->getActiveSheet()->mergeCells('D839:K839');
				$excel->getActiveSheet()->getStyle('A839:R839')->applyFromArray($style_standar);
				$excel->setActiveSheetIndex(0)->setCellValue('O840', "Jumlah");
				$excel->setActiveSheetIndex(0)->setCellValue('P840', "0,00");
				$excel->getActiveSheet()->getStyle('O840:P840')->getFont()->setBold(TRUE);
				$excel->getActiveSheet()->getStyle('O840:P840')->getFont()->setSize(11);
				$excel->getActiveSheet()->getStyle('A840:R840')->applyFromArray($style_standar);
				$excel->setActiveSheetIndex(0)->setCellValue('C841', "c");
				$excel->getActiveSheet()->getStyle('C841')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
				$excel->setActiveSheetIndex(0)->setCellValue('D841', "10 (sepuluh) tahun");
				$excel->getActiveSheet()->mergeCells('D841:K841');
				$excel->setActiveSheetIndex(0)->setCellValue('B842', "2.");
				$excel->getActiveSheet()->getStyle('B842')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
				$excel->setActiveSheetIndex(0)->setCellValue('C842', "Memperoleh penghargaan lainnya");
				$excel->getActiveSheet()->mergeCells('C842:K842');
				$excel->setActiveSheetIndex(0)->setCellValue('C843', "a");
				$excel->getActiveSheet()->getStyle('C843')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
				$excel->setActiveSheetIndex(0)->setCellValue('D843', "Tingkat internasional");
				$excel->setActiveSheetIndex(0)->setCellValue('C844', "b");
				$excel->getActiveSheet()->getStyle('C844')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
				$excel->setActiveSheetIndex(0)->setCellValue('D844', "Tingkat nasional");
				$excel->setActiveSheetIndex(0)->setCellValue('C845', "c");
				$excel->getActiveSheet()->getStyle('C845')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
				$excel->setActiveSheetIndex(0)->setCellValue('D845', "Tingkat provinsi");
				$excel->setActiveSheetIndex(0)->setCellValue('O847', "Jumlah");
				$excel->setActiveSheetIndex(0)->setCellValue('P847', "0,00");
				$excel->getActiveSheet()->getStyle('O847:P847')->getFont()->setBold(TRUE);
				$excel->getActiveSheet()->getStyle('O847:P847')->getFont()->setSize(11);
				$excel->getActiveSheet()->getStyle('A847:R847')->applyFromArray($style_standar);
				$excel->setActiveSheetIndex(0)->setCellValue('A848', "H");
				$excel->setActiveSheetIndex(0)->setCellValue('B848', "Menulis buku pelajaran SLTA ke bawah yang diterbitkan dan diedarkan secara nasional");
				$excel->getActiveSheet()->mergeCells('B848:K848');
				$excel->setActiveSheetIndex(0)->setCellValue('B849', "1");
				$excel->getActiveSheet()->getStyle('B849')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
				$excel->setActiveSheetIndex(0)->setCellValue('C849', "Buku SLTA atau setingkat");
				$excel->getActiveSheet()->mergeCells('C849:F849');
				$excel->setActiveSheetIndex(0)->setCellValue('B850', "2");
				$excel->getActiveSheet()->getStyle('B850')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
				$excel->setActiveSheetIndex(0)->setCellValue('C850', "Buku SLTP atau setingkat");
				$excel->getActiveSheet()->mergeCells('C850:F850');
				$excel->setActiveSheetIndex(0)->setCellValue('B851', "3");
				$excel->getActiveSheet()->getStyle('B851')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
				$excel->setActiveSheetIndex(0)->setCellValue('C851', "Buku SD atau setingkat");
				$excel->getActiveSheet()->mergeCells('C851:F851');
				$excel->getActiveSheet()->getStyle('A851:R851')->applyFromArray($style_standar);
				$excel->setActiveSheetIndex(0)->setCellValue('A852', "I");
				$excel->setActiveSheetIndex(0)->setCellValue('B852', "Mempunyai prestasi di bidang olahraga/-humaniora");
				$excel->getActiveSheet()->mergeCells('B852:K852');
				$excel->setActiveSheetIndex(0)->setCellValue('B853', "1.");
				$excel->getActiveSheet()->getStyle('B853')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
				$excel->setActiveSheetIndex(0)->setCellValue('C853', "Tingkat internasional");
				$excel->setActiveSheetIndex(0)->setCellValue('B854', "2.");
				$excel->getActiveSheet()->getStyle('B854')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
				$excel->setActiveSheetIndex(0)->setCellValue('C854', "Tingkat nasional");
				$excel->setActiveSheetIndex(0)->setCellValue('B855', "3.");
				$excel->getActiveSheet()->getStyle('B855')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
				$excel->setActiveSheetIndex(0)->setCellValue('C855', "Tingkat daerah/lokal");
				$excel->getActiveSheet()->getStyle('A855:R855')->applyFromArray($style_standar);
				$excel->setActiveSheetIndex(0)->setCellValue('A856', "J");
				$excel->setActiveSheetIndex(0)->setCellValue('B856', "Keanggotaan dalam tim penilaian ");
				$excel->getActiveSheet()->mergeCells('B856:K856');
				$excel->setActiveSheetIndex(0)->setCellValue('C857', "Menjadi anggota tim penilaian  jabatan Akademik Dosen");
				$excel->getActiveSheet()->mergeCells('C857:K857');
				$excel->getActiveSheet()->getStyle('A857:R857')->applyFromArray($style_standar);
				$excel->setActiveSheetIndex(0)->setCellValue('O862', "Jumlah");
				$excel->setActiveSheetIndex(0)->setCellValue('P862', "0,00");
				$excel->getActiveSheet()->getStyle('O862:P862')->getFont()->setBold(TRUE);
				$excel->getActiveSheet()->getStyle('O862:P862')->getFont()->setSize(11);
				$excel->getActiveSheet()->getStyle('A862:R862')->applyFromArray($style_standar);
				$excel->setActiveSheetIndex(0)->setCellValue('L863', "Total Penunjang");
				$excel->setActiveSheetIndex(0)->setCellValue('P863', "27,50");
				$excel->getActiveSheet()->getStyle('P863')->getFont()->setBold(TRUE);
				$excel->getActiveSheet()->getStyle('P863')->getFont()->setSize(11);
				$excel->getActiveSheet()->getStyle('A863:R863')->applyFromArray($style_standar);

				$excel->setActiveSheetIndex(0)->setCellValue('O864', "Bandar Lampung,  31 Juli 2019");
				$excel->setActiveSheetIndex(0)->setCellValue('O865', "Ketua Jurusan Ilmu Komputer");
				$excel->setActiveSheetIndex(0)->setCellValue('O869', "Dr.Ir. Kurnia Muludi, M.S.Sc");
				$excel->setActiveSheetIndex(0)->setCellValue('O870', "NIP. 19640616 198902 1 001");


				//PENELITIAN

				$excel->setActiveSheetIndex(0)->setCellValue('A875', "SURAT PERNYATAAN");
    		$excel->setActiveSheetIndex(0)->setCellValue('A876', "MELAKSANAKAN PENELITIAN");
        $excel->getActiveSheet()->mergeCells('A875:R875');
        $excel->getActiveSheet()->mergeCells('A876:R876');
        $excel->getActiveSheet()->getStyle('A875:A876')->getFont()->setBold(TRUE);
    		$excel->getActiveSheet()->getStyle('A875:A876')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);

        $excel->setActiveSheetIndex(0)->setCellValue('B878', "Yang bertanda tangan di bawah ini : ");
    		$excel->setActiveSheetIndex(0)->setCellValue('B880', "Nama ");
    		$excel->setActiveSheetIndex(0)->setCellValue('J880', ":");
    		$excel->setActiveSheetIndex(0)->setCellValue('B881', "NIP");
    		$excel->setActiveSheetIndex(0)->setCellValue('J881', ":");
    		$excel->setActiveSheetIndex(0)->setCellValue('B882', "Pangkat");
    		$excel->setActiveSheetIndex(0)->setCellValue('J882', ":");
    		$excel->setActiveSheetIndex(0)->setCellValue('B883', "Golongan");
    		$excel->setActiveSheetIndex(0)->setCellValue('J883', ":");
    		$excel->setActiveSheetIndex(0)->setCellValue('B884', "Jabatan");
    		$excel->setActiveSheetIndex(0)->setCellValue('J884', ":");
    		$excel->setActiveSheetIndex(0)->setCellValue('B885', "Unit Kerja");
    		$excel->setActiveSheetIndex(0)->setCellValue('J885', ":");
    		$excel->setActiveSheetIndex(0)->setCellValue('B886', "Menyatakan ");
    		$excel->setActiveSheetIndex(0)->setCellValue('B887', "Nama");
    		$excel->setActiveSheetIndex(0)->setCellValue('J887', ":");
    		$excel->setActiveSheetIndex(0)->setCellValue('B888', "NIP");
    		$excel->setActiveSheetIndex(0)->setCellValue('J888', ":");
    		$excel->setActiveSheetIndex(0)->setCellValue('B889', "Pangkat");
    		$excel->setActiveSheetIndex(0)->setCellValue('J889', ":");
    		$excel->setActiveSheetIndex(0)->setCellValue('B890', "Golongan");
    		$excel->setActiveSheetIndex(0)->setCellValue('J890', ":");
    		$excel->setActiveSheetIndex(0)->setCellValue('B891', "Jabatan");
    		$excel->setActiveSheetIndex(0)->setCellValue('J891', ":");
    		$excel->setActiveSheetIndex(0)->setCellValue('B891', "Unit Kerja");
    		$excel->setActiveSheetIndex(0)->setCellValue('J892', ":");
    		$excel->setActiveSheetIndex(0)->setCellValue('B894', "Telah menyatakan pengabdian masyarakat sebagai berikut : ");

        $data_dosen = $this->Dosen->view();

    		foreach($data_dosen as $data) {

    			$excel->setActiveSheetIndex(0)->setCellValue('K880', $data->nama);
    			$excel->setActiveSheetIndex(0)->setCellValue('K881', $data->nip);
    			$excel->setActiveSheetIndex(0)->setCellValue('K882', $data->pangkat);
    			$excel->setActiveSheetIndex(0)->setCellValue('K883', $data->golongan);
    			$excel->setActiveSheetIndex(0)->setCellValue('K884', $data->jabatan);
    			$excel->setActiveSheetIndex(0)->setCellValue('K885', $data->unit_kerja);
    		}

    		$dosen_penunjang = $this->Lektor->view();

    		foreach($dosen_penunjang as $data) {

    			$excel->setActiveSheetIndex(0)->setCellValue('K887', $data->nama);
    			$excel->setActiveSheetIndex(0)->setCellValue('K888', $data->nip);
    			$excel->setActiveSheetIndex(0)->setCellValue('K889', $data->pangkat);
    			$excel->setActiveSheetIndex(0)->setCellValue('K890', $data->golongan);
    			$excel->setActiveSheetIndex(0)->setCellValue('K891', $data->jabatan);
    			$excel->setActiveSheetIndex(0)->setCellValue('K892', $data->unit_kerja);
    		}

        $excel->setActiveSheetIndex(0)->setCellValue('A896', "No");
    		$excel->getActiveSheet()->getStyle('A896')->getAlignment()->setWrapText(TRUE);
    		$excel->setActiveSheetIndex(0)->setCellValue('B896', "Uraian Kegiatan");
    		$excel->getActiveSheet()->mergeCells('B896:K896');
    		$excel->getActiveSheet()->getStyle('B896')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
    		$excel->getActiveSheet()->getStyle('B896')->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
    		$excel->setActiveSheetIndex(0)->setCellValue('L896', "Tanggal");
    		$excel->setActiveSheetIndex(0)->setCellValue('M896', "Satuan Hasil");
    		$excel->getActiveSheet()->getStyle('M896')->getAlignment()->setWrapText(TRUE);
    		$excel->setActiveSheetIndex(0)->setCellValue('N896', "Jumlah Volume Kegiatan");
    		$excel->getActiveSheet()->getStyle('N896')->getAlignment()->setWrapText(TRUE);
    		$excel->setActiveSheetIndex(0)->setCellValue('O896', "Angka Kredit");
    		$excel->getActiveSheet()->getStyle('O896')->getAlignment()->setWrapText(TRUE);
    		$excel->setActiveSheetIndex(0)->setCellValue('P896', "Jumlah Angka Kredit");
    		$excel->getActiveSheet()->getStyle('P896')->getAlignment()->setWrapText(TRUE);
    		$excel->setActiveSheetIndex(0)->setCellValue('Q896', "Keterangan/Bukti Fisik");
    		$excel->getActiveSheet()->mergeCells('Q896:R896');
    		$excel->getActiveSheet()->getStyle('Q896')->getAlignment()->setWrapText(TRUE);

    		$excel->getActiveSheet()->getStyle('A896:A1259')->applyFromArray($style_col);
    		$excel->getActiveSheet()->getStyle('L896:L1259')->applyFromArray($style_col);
    		$excel->getActiveSheet()->getStyle('M896:M1259')->applyFromArray($style_col);
    		$excel->getActiveSheet()->getStyle('N896:N1259')->applyFromArray($style_col);
    		$excel->getActiveSheet()->getStyle('O896:O1259')->applyFromArray($style_col);
    		$excel->getActiveSheet()->getStyle('P896:P1259')->applyFromArray($style_col);
    		$excel->getActiveSheet()->getStyle('R896:R1259')->applyFromArray($style_col);
    		$excel->getActiveSheet()->getStyle('A896')->applyFromArray($style_standar);
    		$excel->getActiveSheet()->getStyle('B896')->applyFromArray($style_standar);
    		$excel->getActiveSheet()->getStyle('C896')->applyFromArray($style_standar);
    		$excel->getActiveSheet()->getStyle('D896')->applyFromArray($style_standar);
    		$excel->getActiveSheet()->getStyle('E896')->applyFromArray($style_standar);
    		$excel->getActiveSheet()->getStyle('F896')->applyFromArray($style_standar);
    		$excel->getActiveSheet()->getStyle('G896')->applyFromArray($style_standar);
    		$excel->getActiveSheet()->getStyle('H896')->applyFromArray($style_standar);
    		$excel->getActiveSheet()->getStyle('I896')->applyFromArray($style_standar);
    		$excel->getActiveSheet()->getStyle('J896')->applyFromArray($style_standar);
    		$excel->getActiveSheet()->getStyle('K896')->applyFromArray($style_standar);
    		$excel->getActiveSheet()->getStyle('L896:Q896')->applyFromArray($style_standar);
    		$excel->getActiveSheet()->getStyle('R896')->applyFromArray($style_standar);

    		$excel->setActiveSheetIndex(0)->setCellValue('A897', "(1)");
    		$excel->setActiveSheetIndex(0)->setCellValue('B897', "(2)");
    		$excel->getActiveSheet()->getStyle('B897')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
    		$excel->getActiveSheet()->mergeCells('B897:K897');
    		$excel->getActiveSheet()->getStyle('A897:Q897')->applyFromArray($style_standar);
    		$excel->setActiveSheetIndex(0)->setCellValue('L897', "(3)");
    		$excel->setActiveSheetIndex(0)->setCellValue('M897', "(4)");
    		$excel->setActiveSheetIndex(0)->setCellValue('N897', "(5)");
    		$excel->setActiveSheetIndex(0)->setCellValue('O897', "(6)");
    		$excel->setActiveSheetIndex(0)->setCellValue('P897', "(7)");
    		$excel->setActiveSheetIndex(0)->setCellValue('R897', "(8)");
    		$excel->getActiveSheet()->getStyle('R897')->getAlignment()->setWrapText(TRUE);
    		$excel->getActiveSheet()->getStyle('R897')->applyFromArray($style_standar);

        $excel->setActiveSheetIndex(0)->setCellValue('A898', "II.");
    		$excel->getActiveSheet()->getStyle('A898')->getFont()->setBold(TRUE);
    		$excel->setActiveSheetIndex(0)->setCellValue('B898', "MELAKSANAKAN PENELITIAN");
    		$excel->getActiveSheet()->mergeCells('B898:G898');
    		$excel->getActiveSheet()->getStyle('B898')->getFont()->setBold(TRUE);
    		$excel->setActiveSheetIndex(0)->setCellValue('B899', "A.");
        $excel->getActiveSheet()->getStyle('B899')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
    		$excel->setActiveSheetIndex(0)->setCellValue('C899', "Menghasilkan karya ilmiah .");
    		$excel->getActiveSheet()->mergeCells('C899:K899');
        $excel->setActiveSheetIndex(0)->setCellValue('C900', "1");
        $excel->getActiveSheet()->getStyle('C900')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
    		$excel->setActiveSheetIndex(0)->setCellValue('D900', "Hasil penelitian atau pemikiran yang dipublikasikan");
    		$excel->getActiveSheet()->mergeCells('D900:K900');
        $excel->setActiveSheetIndex(0)->setCellValue('D901', "a");
        $excel->getActiveSheet()->getStyle('D901')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
        $excel->setActiveSheetIndex(0)->setCellValue('E901', "Dalam bentuk:");
        $excel->getActiveSheet()->mergeCells('E901:K901');
        $excel->setActiveSheetIndex(0)->setCellValue('E902', "1)");
				$excel->getActiveSheet()->getStyle('E902')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
        $excel->setActiveSheetIndex(0)->setCellValue('F902', "Monograf");
        $excel->getActiveSheet()->mergeCells('F902:K902');
    		$excel->setActiveSheetIndex(0)->setCellValue('O902', "Jumlah");
    		$excel->setActiveSheetIndex(0)->setCellValue('P902', "0,00");
    		$excel->getActiveSheet()->getStyle('O902:P902')->getFont()->setBold(TRUE);
    		$excel->getActiveSheet()->getStyle('A902:R902')->applyFromArray($style_standar);
        $excel->setActiveSheetIndex(0)->setCellValue('E903', "2)");
				$excel->getActiveSheet()->getStyle('E903')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
        $excel->setActiveSheetIndex(0)->setCellValue('F903', "Buku Referensi");
				$excel->getActiveSheet()->mergeCells('F903:K903');
        $excel->setActiveSheetIndex(0)->setCellValue('O904', "Jumlah");
    		$excel->setActiveSheetIndex(0)->setCellValue('P904', "0,00");
    		$excel->getActiveSheet()->getStyle('O904:P904')->getFont()->setBold(TRUE);
    		$excel->getActiveSheet()->getStyle('A904:R904')->applyFromArray($style_standar);
        $excel->setActiveSheetIndex(0)->setCellValue('D905', "b");
        $excel->getActiveSheet()->getStyle('D905')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
        $excel->setActiveSheetIndex(0)->setCellValue('E905', "Jurnal ilmiah:");
        $excel->getActiveSheet()->mergeCells('E905:K905');
        $excel->setActiveSheetIndex(0)->setCellValue('E906', "1)");
        $excel->getActiveSheet()->getStyle('E906')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
        $excel->setActiveSheetIndex(0)->setCellValue('F906', "Internasional");
        $excel->getActiveSheet()->mergeCells('F906:K906');

        $excel->setActiveSheetIndex(0)->setCellValue('C907', "1)");
				$excel->getActiveSheet()->getStyle('C907')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
        $excel->setActiveSheetIndex(0)->setCellValue('D907', "International Journal of Advanced Computer Science and Applications (IJACSA)");
				$excel->getActiveSheet()->mergeCells('D907:K907');
        $excel->setActiveSheetIndex(0)->setCellValue('L907', "April ");
        $excel->setActiveSheetIndex(0)->setCellValue('M907', "Jurnal");
        $excel->setActiveSheetIndex(0)->setCellValue('N907', "1,0");
        $excel->setActiveSheetIndex(0)->setCellValue('O907', "37,3");
        $excel->setActiveSheetIndex(0)->setCellValue('P907', "22,38");
        $excel->setActiveSheetIndex(0)->setCellValue('R907', "https://thesai.org/Downloads/Volume10No4/Paper_27-Comparative_Analysis_of_Cow_Disease_Diagnosis.pdf ");
        $excel->setActiveSheetIndex(0)->setCellValue('L908', "2019 ");
        $excel->setActiveSheetIndex(0)->setCellValue('M908', "Internasional");

        $excel->setActiveSheetIndex(0)->setCellValue('D910', "ISSN (Online) : 2156-5570");
        $excel->setActiveSheetIndex(0)->setCellValue('D911', "Vol. 10, Issue 4, PP 227-235, April 2019");
        $excel->setActiveSheetIndex(0)->setCellValue('D913', "Comparative Analysis of Cow Disease Diagnosis Expert System using Bayesian Network and Dempster-Shafer Method");
				$excel->getActiveSheet()->mergeCells('D913:K913');
				$excel->getActiveSheet()->getStyle('D913')->getAlignment()->setWrapText(TRUE);
				$excel->getActiveSheet()->getStyle('D913')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
        $excel->setActiveSheetIndex(0)->setCellValue('R913', "Terdaftar di LPPM Unila");
        $excel->setActiveSheetIndex(0)->setCellValue('R914', "No. 167/J/B/I/FMIPA/2019");
        $excel->setActiveSheetIndex(0)->setCellValue('R915', "Tanggal 27Juni 2019");
        $excel->setActiveSheetIndex(0)->setCellValue('D917', "Aristoteles, Kusuma Adhianto, Rico Andrian, Yeni Nuhricha Sari");
				$excel->getActiveSheet()->mergeCells('D917:K917');

        $excel->setActiveSheetIndex(0)->setCellValue('C920', "2)");
				$excel->getActiveSheet()->getStyle('C920')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
        $excel->setActiveSheetIndex(0)->setCellValue('D920', "International Journal of Advanced Computer Science and Applications (IJACSA)");
				$excel->getActiveSheet()->mergeCells('D920:K920');
        $excel->setActiveSheetIndex(0)->setCellValue('L920', "November");
        $excel->setActiveSheetIndex(0)->setCellValue('M920', "Jurnal");
        $excel->setActiveSheetIndex(0)->setCellValue('N920', "1,0");
        $excel->setActiveSheetIndex(0)->setCellValue('R920', "https://thesai.org/Downloads/Volume8No11/Paper_21-Expert_System_of_Chili_Plant_Disease_Diagnosis.pdf ");
        $excel->setActiveSheetIndex(0)->setCellValue('L921', "2017");
        $excel->setActiveSheetIndex(0)->setCellValue('M921', "Internasional");

        $excel->setActiveSheetIndex(0)->setCellValue('D923', "ISSN (Online) : 2156-5570");
        $excel->setActiveSheetIndex(0)->setCellValue('D924', "Vol. 8, Issue 11, PP 164-168, November 2017");
        $excel->setActiveSheetIndex(0)->setCellValue('R926', "Terdaftar di LPPM Unila");
        $excel->setActiveSheetIndex(0)->setCellValue('D926', "Expert System of Chili Plant Disease Diagnosis using Forward Chaining Method on Android");
				$excel->getActiveSheet()->mergeCells('D926:K926');
				$excel->getActiveSheet()->getStyle('D920')->getAlignment()->setWrapText(TRUE);
				$excel->getActiveSheet()->getStyle('D920')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
        $excel->setActiveSheetIndex(0)->setCellValue('R927', "No. 168/J/B/I/FMIPA/2019");
        $excel->setActiveSheetIndex(0)->setCellValue('R928', "Tanggal 27 Juni 2019");
        $excel->setActiveSheetIndex(0)->setCellValue('D930', "Aristoteles, Mita Fuljana, Joko Prasetyo, Kurnia Muludi");
				$excel->getActiveSheet()->mergeCells('D930:K930');

        $excel->setActiveSheetIndex(0)->setCellValue('C933', "3)");
				$excel->getActiveSheet()->getStyle('C933')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
        $excel->setActiveSheetIndex(0)->setCellValue('D933', "ARPN Journal of Engineering and Applied Sciences");
				$excel->getActiveSheet()->mergeCells('D933:K933');
        $excel->setActiveSheetIndex(0)->setCellValue('L933', "April");
        $excel->setActiveSheetIndex(0)->setCellValue('M933', "Jurnal");
        $excel->setActiveSheetIndex(0)->setCellValue('N933', "1,0");
        $excel->setActiveSheetIndex(0)->setCellValue('R933', "http://www.arpnjournals.org/jeas/research_papers/rp_2016/jeas_0416_4013.pdf ");
        $excel->setActiveSheetIndex(0)->setCellValue('L934', "2016");
        $excel->setActiveSheetIndex(0)->setCellValue('M934', "Internasional");

        $excel->setActiveSheetIndex(0)->setCellValue('D936', "ISSN (Online) : 1819-6608");
        $excel->setActiveSheetIndex(0)->setCellValue('D937', "Vol. 11, No 7, PP 4713-4719, 2016");
        $excel->setActiveSheetIndex(0)->setCellValue('R939', "Terdaftar di LPPM Unila");
        $excel->setActiveSheetIndex(0)->setCellValue('D939', "Performance Evaluation Of Various Genetic Algorithm Approaches For Knapsack Problem ");
				$excel->getActiveSheet()->mergeCells('D939:K939');
				$excel->getActiveSheet()->getStyle('D939')->getAlignment()->setWrapText(TRUE);
				$excel->getActiveSheet()->getStyle('D939')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
        $excel->setActiveSheetIndex(0)->setCellValue('R940', "No. 165/J/B/I/FMIPA/2019");
        $excel->setActiveSheetIndex(0)->setCellValue('R941', "Tanggal 27 Juni 2019");
        $excel->setActiveSheetIndex(0)->setCellValue('D943', "A. Syarif, Aristoteles, A. Dwiastuti, and R. Malinda");
				$excel->getActiveSheet()->mergeCells('D943:K943');

        $excel->setActiveSheetIndex(0)->setCellValue('C946', "4)");
				$excel->getActiveSheet()->getStyle('C946')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
        $excel->setActiveSheetIndex(0)->setCellValue('D946', "IJCSI International Journal Of Computer Science Issues");
				$excel->getActiveSheet()->mergeCells('D946:K946');
        $excel->setActiveSheetIndex(0)->setCellValue('L946', "Mei");
        $excel->setActiveSheetIndex(0)->setCellValue('M946', "Jurnal");
        $excel->setActiveSheetIndex(0)->setCellValue('N946', "1,0");
        $excel->setActiveSheetIndex(0)->setCellValue('R946', "http://www.ijcsi.org/articles/Chord-identification-using-pitch-class-profile-method-with-fast-fourier-transform-feature-extraction.php ");
        $excel->setActiveSheetIndex(0)->setCellValue('L947', "2014");
        $excel->setActiveSheetIndex(0)->setCellValue('M947', "Internasional");

        $excel->setActiveSheetIndex(0)->setCellValue('D949', "ISSN 1694-0784");
        $excel->setActiveSheetIndex(0)->setCellValue('D950', "Vol. 11, Issue 3, No 1, May 2014, ");
        $excel->setActiveSheetIndex(0)->setCellValue('R952', "Terdaftar di LPPM Unila");
        $excel->setActiveSheetIndex(0)->setCellValue('D952', "Chord Identification Using Pitch Class Profile Method With Fast Fourier Transform Feature Extraction");
				$excel->getActiveSheet()->mergeCells('D952:K952');
				$excel->getActiveSheet()->getStyle('D952')->getAlignment()->setWrapText(TRUE);
				$excel->getActiveSheet()->getStyle('D952')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
        $excel->setActiveSheetIndex(0)->setCellValue('R953', "No. 136/J/B/I/FMIPA/2019");
        $excel->setActiveSheetIndex(0)->setCellValue('R954', "Tanggal 12 Juni 2019");
        $excel->setActiveSheetIndex(0)->setCellValue('D956', "Kurnia Muludi, Aristoteles, Abe Frank SFB Loupatty");
				$excel->getActiveSheet()->mergeCells('D956:K956');

        $excel->setActiveSheetIndex(0)->setCellValue('C959', "5)");
				$excel->getActiveSheet()->getStyle('C959')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
        $excel->setActiveSheetIndex(0)->setCellValue('D959', "International Journal of Computer Science and Telecommunications ");
				$excel->getActiveSheet()->mergeCells('D959:K959');
        $excel->setActiveSheetIndex(0)->setCellValue('L959', "Juli");
        $excel->setActiveSheetIndex(0)->setCellValue('M959', "Jurnal");
        $excel->setActiveSheetIndex(0)->setCellValue('N959', "1,0");
        $excel->setActiveSheetIndex(0)->setCellValue('R959', "http://www.ijcst.org/Volume5/Issue7/p_6_5_7.pdf ");
        $excel->setActiveSheetIndex(0)->setCellValue('L960', "2014");
        $excel->setActiveSheetIndex(0)->setCellValue('M960', "Internasional");

        $excel->setActiveSheetIndex(0)->setCellValue('D962', "ISSN 2047-3338");
        $excel->setActiveSheetIndex(0)->setCellValue('R962', "http://repository.lppm.unila.ac.id/1358/ ");
        $excel->setActiveSheetIndex(0)->setCellValue('D963', "Volume 5, Issue 7, July 2014");
        $excel->setActiveSheetIndex(0)->setCellValue('D965', "Text Feature Weighting for Summarization of Documents Bahasa Indonesia by Using Binary Logistic Regression Algorithm");
				$excel->getActiveSheet()->mergeCells('D965:K965');
				$excel->getActiveSheet()->getStyle('D965')->getAlignment()->setWrapText(TRUE);
				$excel->getActiveSheet()->getStyle('D965')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
        $excel->setActiveSheetIndex(0)->setCellValue('R965', "Terdaftar di LPPM Unila");
        $excel->setActiveSheetIndex(0)->setCellValue('R966', "No. 164/J/B/I/FMIPA/2019");
        $excel->setActiveSheetIndex(0)->setCellValue('R967', "Tanggal 27Juni 2019");
        $excel->setActiveSheetIndex(0)->setCellValue('D969', "Aristoteles, Widiarti and Eko Dwi Wibowo");
				$excel->getActiveSheet()->mergeCells('D969:K969');

        $excel->setActiveSheetIndex(0)->setCellValue('C972', "6)");
				$excel->getActiveSheet()->getStyle('C972')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
        $excel->setActiveSheetIndex(0)->setCellValue('D972', "International Journal Of Computer Applications ");
        $excel->setActiveSheetIndex(0)->setCellValue('L972', "November");
        $excel->setActiveSheetIndex(0)->setCellValue('M972', "Jurnal");
        $excel->setActiveSheetIndex(0)->setCellValue('N972', "1,0");
        $excel->setActiveSheetIndex(0)->setCellValue('R972', "http://www.ijcaonline.org/archives/volume81/number6/14013-2158 ");
        $excel->setActiveSheetIndex(0)->setCellValue('L973', "2013");
        $excel->setActiveSheetIndex(0)->setCellValue('M973', "Internasional");

        $excel->setActiveSheetIndex(0)->setCellValue('D975', "ISSN 0975-8887");
        $excel->setActiveSheetIndex(0)->setCellValue('D976', "Volume 81 - No. 6, November 2013");
				$excel->setActiveSheetIndex(0)->setCellValue('D978', "Image Processing For Save Life Predictions Of Tomato Fruit Using RGB Method");
				$excel->getActiveSheet()->mergeCells('D978:K978');
				$excel->getActiveSheet()->getStyle('D978')->getAlignment()->setWrapText(TRUE);
				$excel->getActiveSheet()->getStyle('D978')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
        $excel->setActiveSheetIndex(0)->setCellValue('R978', "Terdaftar di LPPM Unila");
        $excel->setActiveSheetIndex(0)->setCellValue('R979', "No. 166/J/B/I/FMIPA/2019");
        $excel->setActiveSheetIndex(0)->setCellValue('R980', "Tanggal 27 Juni 2019");
        $excel->setActiveSheetIndex(0)->setCellValue('D982', "Aristoteles, Ossy Dwi Endah W, Dwi Susanto");
				$excel->getActiveSheet()->mergeCells('D982:K982');

        $excel->setActiveSheetIndex(0)->setCellValue('C985', "7)");
				$excel->getActiveSheet()->getStyle('C985')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
        $excel->setActiveSheetIndex(0)->setCellValue('D985', "International Journal Of Computer Applications ");
        $excel->setActiveSheetIndex(0)->setCellValue('L985', "Oktober");
        $excel->setActiveSheetIndex(0)->setCellValue('M985', "Jurnal");
        $excel->setActiveSheetIndex(0)->setCellValue('N985', "1,0");
        $excel->setActiveSheetIndex(0)->setCellValue('R985', "http://www.ijcaonline.org/archives/volume80/number13/13922-1824 ");
        $excel->setActiveSheetIndex(0)->setCellValue('L986', "2013");
        $excel->setActiveSheetIndex(0)->setCellValue('M986', "Internasional");

        $excel->setActiveSheetIndex(0)->setCellValue('D988', "ISSN 0975-8887");
        $excel->setActiveSheetIndex(0)->setCellValue('D989', "Volume 80 – No 13, October 2013, ");
        $excel->setActiveSheetIndex(0)->setCellValue('R991', "Terdaftar di LPPM Unila");
        $excel->setActiveSheetIndex(0)->setCellValue('R992', "No. 135/J/B/I/FMIPA/2019");
        $excel->setActiveSheetIndex(0)->setCellValue('D991', "Implementation Of Multilevel Feedback Queue Algorithm In Restaurant Order Food Application Development For Android And Ios Platforms");
        $excel->setActiveSheetIndex(0)->setCellValue('R993', "Tanggal 12 Juni 2019");
        $excel->setActiveSheetIndex(0)->setCellValue('D995', "Dian Andrian Ginting, Aristoteles, Ossy Dwi Endah");
				$excel->getActiveSheet()->mergeCells('D995:K995');
        $excel->setActiveSheetIndex(0)->setCellValue('O998', "Jumlah");
    		$excel->setActiveSheetIndex(0)->setCellValue('P998', "0,00");
    		$excel->getActiveSheet()->getStyle('O998:P998')->getFont()->setBold(TRUE);
    		$excel->getActiveSheet()->getStyle('A998:R998')->applyFromArray($style_standar);
        $excel->setActiveSheetIndex(0)->setCellValue('E999', "2)");
				$excel->getActiveSheet()->getStyle('E999')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
        $excel->setActiveSheetIndex(0)->setCellValue('F999', "Nasional terakreditasi");
        $excel->getActiveSheet()->mergeCells('F999:K999');
        $excel->setActiveSheetIndex(0)->setCellValue('N1001', "Jumlah");
        $excel->setActiveSheetIndex(0)->setCellValue('O1001', "0,00");
    		$excel->setActiveSheetIndex(0)->setCellValue('P1001', "0,00");
    		$excel->getActiveSheet()->getStyle('N1001:P1001')->getFont()->setBold(TRUE);
    		$excel->getActiveSheet()->getStyle('A1001:R1001')->applyFromArray($style_standar);
        $excel->setActiveSheetIndex(0)->setCellValue('E1003', "3)");
				$excel->getActiveSheet()->getStyle('E1003')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
        $excel->setActiveSheetIndex(0)->setCellValue('F1003', "Tidak terakreditasi");
        $excel->getActiveSheet()->mergeCells('F1003:K1003');

        $excel->setActiveSheetIndex(0)->setCellValue('B1005', "1)");
				$excel->getActiveSheet()->getStyle('B1005')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
        $excel->setActiveSheetIndex(0)->setCellValue('C1005', "Jurnal Komputasi FMIPA Unila ");
        $excel->setActiveSheetIndex(0)->setCellValue('L1005', "Oktober");
        $excel->setActiveSheetIndex(0)->setCellValue('M1005', "Jurnal");
        $excel->setActiveSheetIndex(0)->setCellValue('N1005', "1,00");
        $excel->setActiveSheetIndex(0)->setCellValue('R1005', "http://jurnal.fmipa.unila.ac.id/komputasi/article/view/1148 ");
        $excel->setActiveSheetIndex(0)->setCellValue('C1006', "ISSN (Online) : 2541-0350");
        $excel->setActiveSheetIndex(0)->setCellValue('L1006', "2015");
        $excel->setActiveSheetIndex(0)->setCellValue('M1006', "Nasional");
        $excel->setActiveSheetIndex(0)->setCellValue('C1007', "Vol. 3 No 2, PP 136-143, Oktober 2015");
        $excel->setActiveSheetIndex(0)->setCellValue('C1009', "Implementasi Teknologi Markerless Augmented Reality Berbasis Android untuk Mendeteksi dan Mengetahui Lokasi SPBU Terdekat di Kota Bandar Lampung");
        $excel->getActiveSheet()->mergeCells('C1009:K1009');
        $excel->getActiveSheet()->getStyle('C1009')->getAlignment()->setWrapText(TRUE);
        $excel->getActiveSheet()->getStyle('C1009')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);
    		$excel->setActiveSheetIndex(0)->setCellValue('R1009', "Terdaftar di LPPM Unila");
        $excel->getActiveSheet()->getStyle('R1009')->getAlignment()->setWrapText(TRUE);
        $excel->setActiveSheetIndex(0)->setCellValue('R1010', "No. 271/J/B/N/FMIPA/2019");
        $excel->setActiveSheetIndex(0)->setCellValue('R1011', "Tanggal 12 Juni 2019");
        $excel->setActiveSheetIndex(0)->setCellValue('C1013', "Didik Kurniawan, Aristoteles, M. Fathan Kurniawan");
        $excel->getActiveSheet()->mergeCells('C1013:K1013');
        $excel->getActiveSheet()->getStyle('C1011')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);

        $excel->setActiveSheetIndex(0)->setCellValue('B1016', "2)");
				$excel->getActiveSheet()->getStyle('B1016')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
        $excel->setActiveSheetIndex(0)->setCellValue('C1016', "Jurnal Komputasi FMIPA Unila ");
        $excel->setActiveSheetIndex(0)->setCellValue('L1016', "Oktober");
        $excel->setActiveSheetIndex(0)->setCellValue('M1016', "Jurnal");
        $excel->setActiveSheetIndex(0)->setCellValue('N1016', "1,00");
        $excel->setActiveSheetIndex(0)->setCellValue('R1016', "http://jurnal.fmipa.unila.ac.id/komputasi/article/view/1143 ");
        $excel->setActiveSheetIndex(0)->setCellValue('C1017', "ISSN (Online) : 2541-0350");
        $excel->setActiveSheetIndex(0)->setCellValue('L1017', "2015");
        $excel->setActiveSheetIndex(0)->setCellValue('M1017', "Nasional");
        $excel->setActiveSheetIndex(0)->setCellValue('C1018', "Vol. 3 No 2, PP 120-128, Oktober 2015");
        $excel->setActiveSheetIndex(0)->setCellValue('C1020', "Pengembangan Aplikasi Sistem Pembelajaran Klasifikasi (Taksonomi) dan Tata Nama Ilmiah (Binomial Nomenklatur) pada Kingdom Plantae (Tumbuhan) Berbasis Android");
        $excel->getActiveSheet()->mergeCells('C1020:K1020');
        $excel->getActiveSheet()->getStyle('C1020')->getAlignment()->setWrapText(TRUE);
        $excel->getActiveSheet()->getStyle('C1020')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);
        $excel->setActiveSheetIndex(0)->setCellValue('R1020', "Terdaftar di LPPM Unila");
        $excel->getActiveSheet()->getStyle('R1020')->getAlignment()->setWrapText(TRUE);
        $excel->setActiveSheetIndex(0)->setCellValue('R1021', "No. 270/J/B/N/FMIPA/2019");
        $excel->setActiveSheetIndex(0)->setCellValue('R1022', "Tanggal 12 Juni 2019");
        $excel->setActiveSheetIndex(0)->setCellValue('C1024', "Didik Kurniawan, Aristoteles, Ahmad Amirudin");
				$excel->getActiveSheet()->mergeCells('C1024:K1024');

        $excel->setActiveSheetIndex(0)->setCellValue('B1027', "3)");
				$excel->getActiveSheet()->getStyle('B1027')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
        $excel->setActiveSheetIndex(0)->setCellValue('C1027', "Jurnal Komputasi FMIPA Unila ");
        $excel->setActiveSheetIndex(0)->setCellValue('L1027', "Oktober");
        $excel->setActiveSheetIndex(0)->setCellValue('M1027', "Jurnal");
        $excel->setActiveSheetIndex(0)->setCellValue('N1027', "1,00");
        $excel->setActiveSheetIndex(0)->setCellValue('R1027', "http://jurnal.fmipa.unila.ac.id/komputasi/article/view/1131 ");
        $excel->setActiveSheetIndex(0)->setCellValue('C1028', "ISSN (Online) : 2541-0350");
        $excel->setActiveSheetIndex(0)->setCellValue('L1028', "2015");
        $excel->setActiveSheetIndex(0)->setCellValue('M1028', "Nasional");
        $excel->setActiveSheetIndex(0)->setCellValue('C1029', "Vol. 3 No 2, PP 44-52, Oktober 2015");
        $excel->setActiveSheetIndex(0)->setCellValue('C1031', "Sistem Informasi Kuliah Kerja Nyata (KKN) dengan Metode Pigeon Hole untuk Menentukan dan Mengelompokkan Peserta KKN Universitas Lampung");
        $excel->getActiveSheet()->mergeCells('C1031:K1031');
        $excel->getActiveSheet()->getStyle('C1031')->getAlignment()->setWrapText(TRUE);
        $excel->getActiveSheet()->getStyle('C1031')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);
        $excel->setActiveSheetIndex(0)->setCellValue('R1031', "Terdaftar di LPPM Unila");
        $excel->getActiveSheet()->getStyle('R1031')->getAlignment()->setWrapText(TRUE);
        $excel->setActiveSheetIndex(0)->setCellValue('R1032', "No. 267/J/B/N/FMIPA/2019");
        $excel->setActiveSheetIndex(0)->setCellValue('R1033', "Tanggal 12 Juni 2019");
        $excel->setActiveSheetIndex(0)->setCellValue('C1035', "Aristoteles, Rico Andrian, Agatha Beny Himawan");
        $excel->getActiveSheet()->mergeCells('C1035:K1035');
        $excel->getActiveSheet()->getStyle('C1035')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);

        $excel->setActiveSheetIndex(0)->setCellValue('B1038', "4)");
				$excel->getActiveSheet()->getStyle('B1038')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
        $excel->setActiveSheetIndex(0)->setCellValue('C1038', "Jurnal Komputer ");
        $excel->setActiveSheetIndex(0)->setCellValue('L1038', "Oktober");
        $excel->setActiveSheetIndex(0)->setCellValue('M1038', "Jurnal");
        $excel->setActiveSheetIndex(0)->setCellValue('N1038', "1,00");
        $excel->setActiveSheetIndex(0)->setCellValue('R1038', "http://jurnal.fmipa.unila.ac.id/index.php/komputasi/article/view/1128 ");
        $excel->setActiveSheetIndex(0)->setCellValue('C1039', "ISSN :  2541-035");
        $excel->setActiveSheetIndex(0)->setCellValue('L1039', "2015");
        $excel->setActiveSheetIndex(0)->setCellValue('M1039', "Nasional");
        $excel->setActiveSheetIndex(0)->setCellValue('C1040', "Vol 3 No 2, Hal 99-108, Oktober 2015");

        $excel->setActiveSheetIndex(0)->setCellValue('C1042', "SISTEM IDENTIFIKASI PENYAKIT TANAMAN PADI DENGAN MENGGUNAKAN METODE FORWARD CHAINING");
        $excel->getActiveSheet()->mergeCells('C1042:K1042');
        $excel->getActiveSheet()->getStyle('C1042')->getAlignment()->setWrapText(TRUE);
        $excel->getActiveSheet()->getStyle('C1042')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);
        $excel->setActiveSheetIndex(0)->setCellValue('R1042', "Terdaftar di LPPM Unila");
				$excel->getActiveSheet()->getStyle('R1042')->getAlignment()->setWrapText(TRUE);
        $excel->setActiveSheetIndex(0)->setCellValue('R1043', "No. 265/J/B/N/FMIPA/2019");
        $excel->setActiveSheetIndex(0)->setCellValue('R1045', "Tanggal 12 Juni 2019");
        $excel->setActiveSheetIndex(0)->setCellValue('C1047', "Aristoteles, Wardiyanto, Ardye Amando Pratama");
				$excel->getActiveSheet()->mergeCells('C1047:K1047');

        $excel->setActiveSheetIndex(0)->setCellValue('B1050', "5)");
				$excel->getActiveSheet()->getStyle('B1050')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
        $excel->setActiveSheetIndex(0)->setCellValue('C1050', "Jurnal Komputer ");
        $excel->setActiveSheetIndex(0)->setCellValue('L1050', "April");
        $excel->setActiveSheetIndex(0)->setCellValue('M1050', "Jurnal");
        $excel->setActiveSheetIndex(0)->setCellValue('N1050', "1,00");
        $excel->setActiveSheetIndex(0)->setCellValue('R1050', "http://jurnal.fmipa.unila.ac.id/index.php/komputasi/article/view/1216 ");
        $excel->setActiveSheetIndex(0)->setCellValue('C1051', "ISSN :  2541-035");
        $excel->setActiveSheetIndex(0)->setCellValue('L1051', "2016");
        $excel->setActiveSheetIndex(0)->setCellValue('M1051', "Nasional");
        $excel->setActiveSheetIndex(0)->setCellValue('C1052', "Vol 4 No 1, Hal 9-18, April 2016");

        $excel->setActiveSheetIndex(0)->setCellValue('C1054', "SISTEM IDENTIFIKASI PENYAKIT TANAMAN PADI DENGAN MENGGUNAKAN METODE FORWARD CHAINING");
        $excel->getActiveSheet()->mergeCells('C1054:K1054');
        $excel->getActiveSheet()->getStyle('C1054')->getAlignment()->setWrapText(TRUE);
        $excel->getActiveSheet()->getStyle('C1054')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);
        $excel->setActiveSheetIndex(0)->setCellValue('R1054', "Terdaftar di LPPM Unila");
				$excel->getActiveSheet()->getStyle('R1054')->getAlignment()->setWrapText(TRUE);
        $excel->setActiveSheetIndex(0)->setCellValue('R1055', "No. 272/J/B/N/FMIPA/2019");
        $excel->setActiveSheetIndex(0)->setCellValue('R1056', "Tanggal 12 Juni 2019");
        $excel->setActiveSheetIndex(0)->setCellValue('C1058', "Ika Arthalia Wulandari, Aristoteles, Radix Suharjo");
				$excel->getActiveSheet()->mergeCells('C1058:K1058');


        $excel->setActiveSheetIndex(0)->setCellValue('B1061', "6)");
				$excel->getActiveSheet()->getStyle('B1061')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
        $excel->setActiveSheetIndex(0)->setCellValue('C1061', "Jurnal Komputer ");
        $excel->setActiveSheetIndex(0)->setCellValue('L1061', "April");
        $excel->setActiveSheetIndex(0)->setCellValue('M1061', "Jurnal");
        $excel->setActiveSheetIndex(0)->setCellValue('N1061', "1,00");
        $excel->setActiveSheetIndex(0)->setCellValue('R1061', "http://jurnal.fmipa.unila.ac.id/index.php/komputasi/article/view/1164 ");
        $excel->setActiveSheetIndex(0)->setCellValue('C1062', "ISSN :  2541-035");
        $excel->setActiveSheetIndex(0)->setCellValue('L1062', "2016");
        $excel->setActiveSheetIndex(0)->setCellValue('M1062', "Nasional");
        $excel->setActiveSheetIndex(0)->setCellValue('C1063', "Vol 4 No 1, Hal 92-98, April 2016");

        $excel->setActiveSheetIndex(0)->setCellValue('C1065', "SISTEM PAKAR DIAGNOSA PENYAKIT PADA IKAN BUDIDAYA AIR TAWAR DENGAN METODE FORWARD CHAINING BERBASIS ANDROID ");
        $excel->getActiveSheet()->mergeCells('C1065:K1065');
        $excel->getActiveSheet()->getStyle('C1065')->getAlignment()->setWrapText(TRUE);
        $excel->getActiveSheet()->getStyle('C1065')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);
        $excel->setActiveSheetIndex(0)->setCellValue('R1065', "Terdaftar di LPPM Unila");
				$excel->getActiveSheet()->getStyle('R1065')->getAlignment()->setWrapText(TRUE);
        $excel->setActiveSheetIndex(0)->setCellValue('R1065', "No. 301/J/B/N/FMIPA/2019");
        $excel->setActiveSheetIndex(0)->setCellValue('R1065', "Tanggal 27 Juni 2019");
        $excel->setActiveSheetIndex(0)->setCellValue('C1065', "Ardhika Praseda Ageng Putra, Aristoteles, Rara Diantari");
        $excel->getActiveSheet()->mergeCells('C1065:K1065');

        $excel->setActiveSheetIndex(0)->setCellValue('B1068', "7)");
				$excel->getActiveSheet()->getStyle('B1068')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
        $excel->setActiveSheetIndex(0)->setCellValue('C1068', "Jurnal Komputer ");
        $excel->setActiveSheetIndex(0)->setCellValue('L1068', "April");
        $excel->setActiveSheetIndex(0)->setCellValue('M1068', "Jurnal");
        $excel->setActiveSheetIndex(0)->setCellValue('N1068', "1,00");
        $excel->setActiveSheetIndex(0)->setCellValue('R1068', "http://jurnal.fmipa.unila.ac.id/index.php/komputasi/article/view/1173 ");
        $excel->setActiveSheetIndex(0)->setCellValue('C1069', "ISSN :  2541-035");
        $excel->setActiveSheetIndex(0)->setCellValue('L1069', "2016");
        $excel->setActiveSheetIndex(0)->setCellValue('M1069', "Nasional");
        $excel->setActiveSheetIndex(0)->setCellValue('C1070', "Vol 4 No 1, Hal 117-124, April 2016");

        $excel->setActiveSheetIndex(0)->setCellValue('C1072', "SISTEM PAKAR DIAGNOSA PENYAKIT PADA IKAN BUDIDAYA AIR TAWAR DENGAN METODE FORWARD CHAINING BERBASIS ANDROID ");
        $excel->getActiveSheet()->mergeCells('C1072:K1072');
        $excel->getActiveSheet()->getStyle('C1072')->getAlignment()->setWrapText(TRUE);
        $excel->getActiveSheet()->getStyle('C1072')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);
        $excel->setActiveSheetIndex(0)->setCellValue('R1072', "Terdaftar di LPPM Unila");
        $excel->setActiveSheetIndex(0)->setCellValue('R1073', "No. 300/J/B/N/FMIPA/2019");
        $excel->setActiveSheetIndex(0)->setCellValue('R1074', "Tanggal 27 Juni 2019");
        $excel->setActiveSheetIndex(0)->setCellValue('C1076', "Rifki Wardana, Aristoteles, Jani Master");
				$excel->getActiveSheet()->mergeCells('C1076:K1076');

				$excel->setActiveSheetIndex(0)->setCellValue('B1079', "8)");
				$excel->getActiveSheet()->getStyle('B1079')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
        $excel->setActiveSheetIndex(0)->setCellValue('C1079', "Jurnal Komputer ");
        $excel->setActiveSheetIndex(0)->setCellValue('L1079', "April");
        $excel->setActiveSheetIndex(0)->setCellValue('M1079', "Jurnal");
        $excel->setActiveSheetIndex(0)->setCellValue('N1079', "1,00");
        $excel->setActiveSheetIndex(0)->setCellValue('R1079', "http://jurnal.fmipa.unila.ac.id/komputasi/article/view/1191 ");
        $excel->setActiveSheetIndex(0)->setCellValue('C1080', "ISSN :  2541-035");
        $excel->setActiveSheetIndex(0)->setCellValue('L1080', "2016");
        $excel->setActiveSheetIndex(0)->setCellValue('M1080', "Nasional");
        $excel->setActiveSheetIndex(0)->setCellValue('C1081', "Vol 4 No 1, Hal 176-1186, April 2016");

        $excel->setActiveSheetIndex(0)->setCellValue('C1083', "PENGEMBANGAN SISTEM INFORMASI COMICREADER MENGGUNAKAN KERANGKA KERJA YII");
        $excel->getActiveSheet()->mergeCells('C1083:K1083');
        $excel->getActiveSheet()->getStyle('C1083')->getAlignment()->setWrapText(TRUE);
        $excel->getActiveSheet()->getStyle('C1083')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);
        $excel->setActiveSheetIndex(0)->setCellValue('R1083', "Terdaftar di LPPM Unila");
				$excel->getActiveSheet()->getStyle('R1083')->getAlignment()->setWrapText(TRUE);
        $excel->setActiveSheetIndex(0)->setCellValue('R1084', "No. 269/J/B/N/FMIPA/2019");
        $excel->setActiveSheetIndex(0)->setCellValue('R1085', "Tanggal 12 Juni 2019");
        $excel->setActiveSheetIndex(0)->setCellValue('C1087', "Sabila Rusyda, Aristoteles, Dwi Sakethi, Admi Syarif");
				$excel->getActiveSheet()->mergeCells('C1087:K1087');

				$excel->setActiveSheetIndex(0)->setCellValue('B1090', "9)");
				$excel->getActiveSheet()->getStyle('B1090')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
        $excel->setActiveSheetIndex(0)->setCellValue('C1090', "Jurnal Komputasi FMIPA Unila");
        $excel->setActiveSheetIndex(0)->setCellValue('L1090', "Oktober");
        $excel->setActiveSheetIndex(0)->setCellValue('M1090', "Jurnal");
        $excel->setActiveSheetIndex(0)->setCellValue('N1090', "1,00");
        $excel->setActiveSheetIndex(0)->setCellValue('R1090', "http://jurnal.fmipa.unila.ac.id/komputasi/article/view/1351 ");
        $excel->setActiveSheetIndex(0)->setCellValue('C1091', "ISSN (Online) : 2541-0350");
        $excel->setActiveSheetIndex(0)->setCellValue('L1091', "2016");
        $excel->setActiveSheetIndex(0)->setCellValue('M1091', "Nasional");
        $excel->setActiveSheetIndex(0)->setCellValue('C1092', "Vol. 4 No 2, PP 52-66, Oktober 2016");

        $excel->setActiveSheetIndex(0)->setCellValue('C1094', "PEMETAAN SEBARAN ASAL SISWA DAN KLASIFIKASI JARAK ASAL SISWA SMA NEGERI DI KABUPATEN PRINGSEWU MENGGUNAKAN METODE NAÏVE BAYES");
        $excel->getActiveSheet()->mergeCells('C1094:K1094');
        $excel->getActiveSheet()->getStyle('C1094')->getAlignment()->setWrapText(TRUE);
        $excel->getActiveSheet()->getStyle('C1094')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);
        $excel->setActiveSheetIndex(0)->setCellValue('R1094', "Terdaftar di LPPM Unila");
				$excel->getActiveSheet()->getStyle('R1094')->getAlignment()->setWrapText(TRUE);
        $excel->setActiveSheetIndex(0)->setCellValue('R1095', "No. 299/J/B/N/FMIPA/2019");
        $excel->setActiveSheetIndex(0)->setCellValue('R1096', "Tanggal 27 Juni 2019");
        $excel->setActiveSheetIndex(0)->setCellValue('C1099', "Riska Aprilia, Kurnia Muludi, Aristoteles");
				$excel->getActiveSheet()->mergeCells('C1099:K1099');

				$excel->setActiveSheetIndex(0)->setCellValue('B1102', "10)");
				$excel->getActiveSheet()->getStyle('B1102')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
        $excel->setActiveSheetIndex(0)->setCellValue('C1102', "Jurnal Komputasi FMIPA Unila ");
        $excel->setActiveSheetIndex(0)->setCellValue('L1102', "April");
        $excel->setActiveSheetIndex(0)->setCellValue('M1102', "Jurnal");
        $excel->setActiveSheetIndex(0)->setCellValue('N1102', "1,00");
        $excel->setActiveSheetIndex(0)->setCellValue('R1102', "http://jurnal.fmipa.unila.ac.id/komputasi/article/view/1402/1220 ");
        $excel->setActiveSheetIndex(0)->setCellValue('C1103', "ISSN (Online) : 2541-0350");
        $excel->setActiveSheetIndex(0)->setCellValue('L1103', "2017");
        $excel->setActiveSheetIndex(0)->setCellValue('M1103', "Nasional");
        $excel->setActiveSheetIndex(0)->setCellValue('C1104', "Vol. 5, No 1, PP 8-16, April 2017");

        $excel->setActiveSheetIndex(0)->setCellValue('C1106', "PENGEMBANGAN SISTEM PELAPORAN KEGIATAN KKN BERBASIS ANDROID");
        $excel->getActiveSheet()->mergeCells('C1106:K1106');
        $excel->getActiveSheet()->getStyle('C1106')->getAlignment()->setWrapText(TRUE);
        $excel->getActiveSheet()->getStyle('C1106')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);
        $excel->setActiveSheetIndex(0)->setCellValue('R1106', "Terdaftar di LPPM Unila");
				$excel->getActiveSheet()->getStyle('R1106')->getAlignment()->setWrapText(TRUE);
        $excel->setActiveSheetIndex(0)->setCellValue('R1107', "No. 299/J/B/N/FMIPA/2019");
        $excel->setActiveSheetIndex(0)->setCellValue('R1108', "Tanggal 27 Juni 2019");
        $excel->setActiveSheetIndex(0)->setCellValue('C1110', "Danzen Hangga Permana, Aristoteles");
				$excel->getActiveSheet()->mergeCells('C1110:K1110');

				$excel->setActiveSheetIndex(0)->setCellValue('B1113', "11)");
				$excel->getActiveSheet()->getStyle('B1113')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
        $excel->setActiveSheetIndex(0)->setCellValue('C1113', "Jurnal Komputasi FMIPA Unila ");
        $excel->setActiveSheetIndex(0)->setCellValue('L1113', "April");
        $excel->setActiveSheetIndex(0)->setCellValue('M1113', "Jurnal");
        $excel->setActiveSheetIndex(0)->setCellValue('N1113', "1,00");
        $excel->setActiveSheetIndex(0)->setCellValue('R1113', "http://jurnal.fmipa.unila.ac.id/komputasi/article/view/1402/1220 ");
        $excel->setActiveSheetIndex(0)->setCellValue('C1114', "ISSN (Online) : 2541-0350");
        $excel->setActiveSheetIndex(0)->setCellValue('L1114', "2017");
        $excel->setActiveSheetIndex(0)->setCellValue('M1114', "Nasional");
        $excel->setActiveSheetIndex(0)->setCellValue('C1115', "Vol. 5, No 1, PP 8-16, April 2017");

        $excel->setActiveSheetIndex(0)->setCellValue('C1117', "ANALISIS PENGELOMPOKAN MAHASISWA KKN BERDASARKAN KRITERIA JENIS KELAMIN, FAKULTAS DAN SEKOLAH");
        $excel->getActiveSheet()->mergeCells('C1117:K1117');
        $excel->getActiveSheet()->getStyle('C1117')->getAlignment()->setWrapText(TRUE);
        $excel->getActiveSheet()->getStyle('C1117')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);
        $excel->setActiveSheetIndex(0)->setCellValue('R1117', "Terdaftar di LPPM Unila");
				$excel->getActiveSheet()->getStyle('R1117')->getAlignment()->setWrapText(TRUE);
        $excel->setActiveSheetIndex(0)->setCellValue('R1118', "No. 266/J/B/N/FMIPA/2019");
        $excel->setActiveSheetIndex(0)->setCellValue('R1119', "Tanggal 12 Juni 2019");
        $excel->setActiveSheetIndex(0)->setCellValue('C1121', "Vandu Riski Muwisnawangsa, Aristoteles");
				$excel->getActiveSheet()->mergeCells('C1121:K1121');

				$excel->setActiveSheetIndex(0)->setCellValue('B1124', "12)");
				$excel->getActiveSheet()->getStyle('B1124')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
        $excel->setActiveSheetIndex(0)->setCellValue('C1124', "Jurnal Komputasi FMIPA Unila ");
        $excel->setActiveSheetIndex(0)->setCellValue('L1124', "Oktober");
        $excel->setActiveSheetIndex(0)->setCellValue('M1124', "Jurnal");
        $excel->setActiveSheetIndex(0)->setCellValue('N1124', "1,00");
        $excel->setActiveSheetIndex(0)->setCellValue('R1124', "http://jurnal.fmipa.unila.ac.id/komputasi/article/view/1539/1307 ");
        $excel->setActiveSheetIndex(0)->setCellValue('C1125', "ISSN (Online) : 2541-0350");
        $excel->setActiveSheetIndex(0)->setCellValue('L1125', "2017");
        $excel->setActiveSheetIndex(0)->setCellValue('M1125', "Nasional");
        $excel->setActiveSheetIndex(0)->setCellValue('C1126', "Vol. 5, No 2, PP 55-63, Oktober 2017");

        $excel->setActiveSheetIndex(0)->setCellValue('C1128', "APLIKASI INFORMASI DOKTER SPESIALIS DI BANDAR LAMPUNG BERBASIS ANDROID DENGAN MENGGUNAKAN TEKNOLOGI LOCATION BASE SERVICE");
        $excel->getActiveSheet()->mergeCells('C1128:K1128');
        $excel->getActiveSheet()->getStyle('C1128')->getAlignment()->setWrapText(TRUE);
        $excel->getActiveSheet()->getStyle('C1128')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);
        $excel->setActiveSheetIndex(0)->setCellValue('R1128', "Terdaftar di LPPM Unila");
				$excel->getActiveSheet()->getStyle('R1128')->getAlignment()->setWrapText(TRUE);
        $excel->setActiveSheetIndex(0)->setCellValue('R1129', "No. 274/J/B/N/FMIPA/2019");
        $excel->setActiveSheetIndex(0)->setCellValue('R1130', "Tanggal 12 Juni 2019");
        $excel->setActiveSheetIndex(0)->setCellValue('C1132', "Nurmayanti, Aristoteles, Astria Hijriani");
				$excel->getActiveSheet()->mergeCells('C1132:K1132');

				$excel->setActiveSheetIndex(0)->setCellValue('B1135', "13)");
				$excel->getActiveSheet()->getStyle('B1135')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
        $excel->setActiveSheetIndex(0)->setCellValue('C1135', "Jurnal Komputasi FMIPA Unila ");
        $excel->setActiveSheetIndex(0)->setCellValue('L1135', "April");
        $excel->setActiveSheetIndex(0)->setCellValue('M1135', "Jurnal");
        $excel->setActiveSheetIndex(0)->setCellValue('N1135', "1,00");
        $excel->setActiveSheetIndex(0)->setCellValue('R1135', "http://jurnal.fmipa.unila.ac.id/komputasi/article/view/1564/1318 ");
        $excel->setActiveSheetIndex(0)->setCellValue('C1136', "ISSN (Online) : 2541-0350");
        $excel->setActiveSheetIndex(0)->setCellValue('L1136', "2018");
        $excel->setActiveSheetIndex(0)->setCellValue('M1136', "Nasional");
        $excel->setActiveSheetIndex(0)->setCellValue('C1137', "Vol. 6, No 1, PP 64-74, April 2018");

        $excel->setActiveSheetIndex(0)->setCellValue('C1139', "Panduan Lapangan Jenis Kupu-kupu di Lingkungan Universitas Lampung Berbasis Android");
        $excel->getActiveSheet()->mergeCells('C1139:K1139');
        $excel->getActiveSheet()->getStyle('C1139')->getAlignment()->setWrapText(TRUE);
        $excel->getActiveSheet()->getStyle('C1139')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);
        $excel->setActiveSheetIndex(0)->setCellValue('R1139', "Terdaftar di LPPM Unila");
				$excel->getActiveSheet()->getStyle('R1139')->getAlignment()->setWrapText(TRUE);
        $excel->setActiveSheetIndex(0)->setCellValue('R1140', "No. 273/J/B/N/FMIPA/2019");
        $excel->setActiveSheetIndex(0)->setCellValue('R1141', "Tanggal 12 Juni 2019");
        $excel->setActiveSheetIndex(0)->setCellValue('C1143', "Aristoteles, Martinus, Galih Imam Widangga");
				$excel->getActiveSheet()->mergeCells('C1143:K1143');

				$excel->setActiveSheetIndex(0)->setCellValue('B1146', "14)");
				$excel->getActiveSheet()->getStyle('B1146')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
        $excel->setActiveSheetIndex(0)->setCellValue('C1146', "Jurnal Komputasi FMIPA Unila ");
        $excel->setActiveSheetIndex(0)->setCellValue('L1146', "Oktober");
        $excel->setActiveSheetIndex(0)->setCellValue('M1146', "Jurnal");
        $excel->setActiveSheetIndex(0)->setCellValue('N1146', "1,00");
        $excel->setActiveSheetIndex(0)->setCellValue('R1146', "http://jurnal.fmipa.unila.ac.id/komputasi/article/view/1655/1332 ");
        $excel->setActiveSheetIndex(0)->setCellValue('C1147', "ISSN (Online) : 2541-0350");
        $excel->setActiveSheetIndex(0)->setCellValue('L1147', "2018");
        $excel->setActiveSheetIndex(0)->setCellValue('M1147', "Nasional");
        $excel->setActiveSheetIndex(0)->setCellValue('C1148', "Vol. 6, No 2, PP 1-10, Oktober 2018");

        $excel->setActiveSheetIndex(0)->setCellValue('C1150', "SISTEM INFORMASI KULIAH KERJA NYATA (KKN) BERBASIS ANDROID UNIVERSITAS LAMPUNG");
        $excel->getActiveSheet()->mergeCells('C1150:K1150');
        $excel->getActiveSheet()->getStyle('C1150')->getAlignment()->setWrapText(TRUE);
        $excel->getActiveSheet()->getStyle('C1150')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);
        $excel->setActiveSheetIndex(0)->setCellValue('R1150', "Terdaftar di LPPM Unila");
				$excel->getActiveSheet()->getStyle('R1150')->getAlignment()->setWrapText(TRUE);
        $excel->setActiveSheetIndex(0)->setCellValue('R1151', "No. 298/J/B/N/FMIPA/2019");
        $excel->setActiveSheetIndex(0)->setCellValue('R1152', "Tanggal 27 Juni 2019");
        $excel->setActiveSheetIndex(0)->setCellValue('C1154', "Aristoteles, Nur Efendi, Febi Eka Febriansyah, Wisnu Lukito, Firmansyah");
				$excel->getActiveSheet()->mergeCells('C1154:K1154');

				$excel->setActiveSheetIndex(0)->setCellValue('B1157', "15)");
				$excel->getActiveSheet()->getStyle('B1157')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
        $excel->setActiveSheetIndex(0)->setCellValue('C1157', "Jurnal Komputasi FMIPA Unila ");
        $excel->setActiveSheetIndex(0)->setCellValue('L1157', "Oktober");
        $excel->setActiveSheetIndex(0)->setCellValue('M1157', "Jurnal");
        $excel->setActiveSheetIndex(0)->setCellValue('N1157', "1,00");
        $excel->setActiveSheetIndex(0)->setCellValue('R1157', "http://jurnal.fmipa.unila.ac.id/komputasi/article/view/1693/1339 ");
        $excel->setActiveSheetIndex(0)->setCellValue('C1158', "ISSN (Online) : 2541-0350");
        $excel->setActiveSheetIndex(0)->setCellValue('L1158', "2018");
        $excel->setActiveSheetIndex(0)->setCellValue('M1158', "Nasional");
        $excel->setActiveSheetIndex(0)->setCellValue('C1159', "Vol. 6, No 2, PP 64-73, Oktober 2018");

        $excel->setActiveSheetIndex(0)->setCellValue('C1161', "Analisis Manajemen Risiko Sistem Informasi KKN Universitas Lampung");
        $excel->getActiveSheet()->mergeCells('C1161:K1161');
        $excel->getActiveSheet()->getStyle('C1161')->getAlignment()->setWrapText(TRUE);
        $excel->getActiveSheet()->getStyle('C1161')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);
        $excel->setActiveSheetIndex(0)->setCellValue('R1161', "Terdaftar di LPPM Unila");
				$excel->getActiveSheet()->getStyle('R1161')->getAlignment()->setWrapText(TRUE);
        $excel->setActiveSheetIndex(0)->setCellValue('R1162', "No. 268/J/B/N/FMIPA/2019");
        $excel->setActiveSheetIndex(0)->setCellValue('R1163', "Tanggal 12 Juni 2019");
        $excel->setActiveSheetIndex(0)->setCellValue('C1165', "Noviyanti, Yunda Heningtyas, Tristiyanto, Aristoteles");
				$excel->getActiveSheet()->mergeCells('C1165:K1165');
				$excel->setActiveSheetIndex(0)->setCellValue('N1168', "Jumlah");
        $excel->setActiveSheetIndex(0)->setCellValue('O1168', "0,00");
    		$excel->setActiveSheetIndex(0)->setCellValue('P1168', "0,00");
    		$excel->getActiveSheet()->getStyle('N1168:P1168')->getFont()->setBold(TRUE);
				$excel->getActiveSheet()->getStyle('A1168:R1168')->applyFromArray($style_standar);

				$excel->setActiveSheetIndex(0)->setCellValue('D1169', "c.");
				$excel->getActiveSheet()->getStyle('D1169')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
        $excel->setActiveSheetIndex(0)->setCellValue('E1169', "Seminar");
        $excel->setActiveSheetIndex(0)->setCellValue('E1170', "1)");
				$excel->getActiveSheet()->getStyle('E1170')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
        $excel->setActiveSheetIndex(0)->setCellValue('F1170', "Disajikan tingkat:");
				$excel->setActiveSheetIndex(0)->setCellValue('N1171', "Jumlah");
        $excel->setActiveSheetIndex(0)->setCellValue('O1171', "0,00");
    		$excel->setActiveSheetIndex(0)->setCellValue('P1171', "0,00");
    		$excel->getActiveSheet()->getStyle('N1171:P1171')->getFont()->setBold(TRUE);
				$excel->getActiveSheet()->getStyle('A1171:R1171')->applyFromArray($style_standar);

				$excel->setActiveSheetIndex(0)->setCellValue('F1172', "a) Internasional");
				$excel->setActiveSheetIndex(0)->setCellValue('B1174', "1)");
				$excel->getActiveSheet()->getStyle('B1174')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
				$excel->setActiveSheetIndex(0)->setCellValue('C1174', "3rd INTERNATIONAL WILDLIFE SYMPOSIUM");
				$excel->setActiveSheetIndex(0)->setCellValue('L1174', "15)");
				$excel->setActiveSheetIndex(0)->setCellValue('M1174', "18-20 Oktober");
				$excel->setActiveSheetIndex(0)->setCellValue('N1174', "1,00");
				$excel->setActiveSheetIndex(0)->setCellValue('P1174', "0,00");
				$excel->setActiveSheetIndex(0)->setCellValue('R1174', "http://repository.lppm.unila.ac.id/3816/ ");
				$excel->setActiveSheetIndex(0)->setCellValue('C1175', "ISBN 978-602-0860-13-8");
				$excel->setActiveSheetIndex(0)->setCellValue('L1175', "2016");

				$excel->setActiveSheetIndex(0)->setCellValue('C1177', "An Expert System To Diagnose Chicken Diseases With Certainty Factor Based On Android ");
				$excel->getActiveSheet()->mergeCells('C1177:K1177');
        $excel->getActiveSheet()->getStyle('C1177')->getAlignment()->setWrapText(TRUE);
        $excel->getActiveSheet()->getStyle('C1177')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);
				$excel->setActiveSheetIndex(0)->setCellValue('R1177', "No. 121/P/B/I/FMIPA/2019");
				$excel->getActiveSheet()->getStyle('R1177')->getAlignment()->setWrapText(TRUE);
				$excel->setActiveSheetIndex(0)->setCellValue('R1178', "Tanggal 27 Juni 2019");
				$excel->setActiveSheetIndex(0)->setCellValue('C1180', "Aristoteles, Kusuma Adhianto,");
				$excel->getActiveSheet()->mergeCells('C1180:K1180');
				$excel->setActiveSheetIndex(0)->setCellValue('N1182', "Jumlah");
        $excel->setActiveSheetIndex(0)->setCellValue('O1182', "0,00");
    		$excel->setActiveSheetIndex(0)->setCellValue('P1182', "0,00");
    		$excel->getActiveSheet()->getStyle('N1182:P1182')->getFont()->setBold(TRUE);
				$excel->getActiveSheet()->getStyle('A1182:R1182')->applyFromArray($style_standar);

				$excel->setActiveSheetIndex(0)->setCellValue('F1183', "b) Nasional");
				$excel->setActiveSheetIndex(0)->setCellValue('B1185', "2)");
				$excel->getActiveSheet()->getStyle('B1185')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
				$excel->setActiveSheetIndex(0)->setCellValue('C1185', "Prosiding Sain dan Teknologi VI 2015");
				$excel->setActiveSheetIndex(0)->setCellValue('L1185', "3 November");
				$excel->setActiveSheetIndex(0)->setCellValue('M1185', "Prosiding");
				$excel->setActiveSheetIndex(0)->setCellValue('N1185', "1,00");
				$excel->setActiveSheetIndex(0)->setCellValue('R1185', "http://satek.unila.ac.id/wp-content/uploads/2015/08/41-Aldona.pdf ");
				$excel->setActiveSheetIndex(0)->setCellValue('C1186', "ISBN  : 978-602-0860-02-2");
				$excel->setActiveSheetIndex(0)->setCellValue('L1186', "2015");
				$excel->setActiveSheetIndex(0)->setCellValue('C1187', "Hal 485-491, November 2015");

				$excel->setActiveSheetIndex(0)->setCellValue('C1189', "Sistem Informasi Pemantauan Potensi Desa dan Pengumpulan Laporan Hasil Kegiatan Kuliah Kerja Nyata (KKN) Universitas Lampung");
				$excel->getActiveSheet()->mergeCells('C1189:K1189');
        $excel->getActiveSheet()->getStyle('C1189')->getAlignment()->setWrapText(TRUE);
        $excel->getActiveSheet()->getStyle('C1189')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);
				$excel->setActiveSheetIndex(0)->setCellValue('R1189', "No. 117/P/B/N/FMIPA/2019");
				$excel->getActiveSheet()->getStyle('R1189')->getAlignment()->setWrapText(TRUE);
				$excel->setActiveSheetIndex(0)->setCellValue('R1190', "Tanggal 27 Juni 2019");
				$excel->setActiveSheetIndex(0)->setCellValue('C1192', "Aldona Pronika, Aristoteles dan Irwan Adi Pribadi");
				$excel->getActiveSheet()->mergeCells('C1192:K1192');

				$excel->setActiveSheetIndex(0)->setCellValue('B1195', "3)");
				$excel->getActiveSheet()->getStyle('B1195')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
				$excel->setActiveSheetIndex(0)->setCellValue('C1195', "Prosiding Sain dan Teknologi VI 2015");
				$excel->setActiveSheetIndex(0)->setCellValue('L1195', "3 November");
				$excel->setActiveSheetIndex(0)->setCellValue('M1195', "Prosiding");
				$excel->setActiveSheetIndex(0)->setCellValue('N1195', "1,00");
				$excel->setActiveSheetIndex(0)->setCellValue('R1195', "http://satek.unila.ac.id/wp-content/uploads/2015/08/44-Harisa.pdf ");
				$excel->setActiveSheetIndex(0)->setCellValue('C1196', "ISBN  : 978-602-0860-02-2");
				$excel->setActiveSheetIndex(0)->setCellValue('L1196', "2015");
				$excel->setActiveSheetIndex(0)->setCellValue('C1197', "Hal 516-527, November 2015");

				$excel->setActiveSheetIndex(0)->setCellValue('C1199', "Pengembangan Sistem Informasi Kuliah Kerja Nyata (KKN) dengan Algortima Greedy Untuk Menentukan Pengelompokan Peserta KKN (Studi Kasus Universitas Lampung)");
				$excel->getActiveSheet()->mergeCells('C1199:K1199');
        $excel->getActiveSheet()->getStyle('C1199')->getAlignment()->setWrapText(TRUE);
        $excel->getActiveSheet()->getStyle('C1199')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);
				$excel->setActiveSheetIndex(0)->setCellValue('R1199', "No. 120/P/B/N/FMIPA/2019");
				$excel->getActiveSheet()->getStyle('R1199')->getAlignment()->setWrapText(TRUE);
				$excel->setActiveSheetIndex(0)->setCellValue('R1200', "Tanggal 27 Juni 2019");
				$excel->setActiveSheetIndex(0)->setCellValue('C1202', "Harisa Eka Septiarani, Aristoteles  dan Wamiliana");
				$excel->getActiveSheet()->mergeCells('C1202:K1202');

				$excel->setActiveSheetIndex(0)->setCellValue('B1205', "4)");
				$excel->getActiveSheet()->getStyle('B1205')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
				$excel->setActiveSheetIndex(0)->setCellValue('C1205', "Prosiding Semirata FMIPA Universitas Lampung");
				$excel->setActiveSheetIndex(0)->setCellValue('L1205', "10-12 Mei");
				$excel->setActiveSheetIndex(0)->setCellValue('M1205', "Prosiding");
				$excel->setActiveSheetIndex(0)->setCellValue('N1205', "1,00");
				$excel->setActiveSheetIndex(0)->setCellValue('R1205', "http://jurnal.fmipa.unila.ac.id/index.php/semirata/article/view/703 ");
				$excel->setActiveSheetIndex(0)->setCellValue('C1206', "ISBN 978-602-985599-2-0");
				$excel->setActiveSheetIndex(0)->setCellValue('L1206', "2013");
				$excel->setActiveSheetIndex(0)->setCellValue('C1207', "10-12 Mei 2013");

				$excel->setActiveSheetIndex(0)->setCellValue('C1209', "Penerapan Algoritma Genetika Pada Peringkasan Teks Dokumen Bahasa Indonesia");
				$excel->getActiveSheet()->mergeCells('C1209:K1209');
        $excel->getActiveSheet()->getStyle('C1209')->getAlignment()->setWrapText(TRUE);
        $excel->getActiveSheet()->getStyle('C1209')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);
				$excel->setActiveSheetIndex(0)->setCellValue('R1209', "No. 118/P/B/N/FMIPA/2019");
				$excel->getActiveSheet()->getStyle('R1209')->getAlignment()->setWrapText(TRUE);
				$excel->setActiveSheetIndex(0)->setCellValue('R1210', "Tanggal 27 Juni 2019");
				$excel->setActiveSheetIndex(0)->setCellValue('C1212', "Aristoteles");
				$excel->getActiveSheet()->mergeCells('C1212:K1212');

				$excel->setActiveSheetIndex(0)->setCellValue('B1214', "5)");
				$excel->getActiveSheet()->getStyle('B1214')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
				$excel->setActiveSheetIndex(0)->setCellValue('C1214', "Prosiding Satek V 2013 Unila");
				$excel->setActiveSheetIndex(0)->setCellValue('L1214', "30 November");
				$excel->setActiveSheetIndex(0)->setCellValue('M1214', "Prosiding");
				$excel->setActiveSheetIndex(0)->setCellValue('N1214', "1,00");
				$excel->setActiveSheetIndex(0)->setCellValue('R1214', "http://satek.unila.ac.id/wp-content/uploads/2014/03/2-X9.pdf ");
				$excel->setActiveSheetIndex(0)->setCellValue('C1215', "ISBN 978-979-8510-71-7");
				$excel->setActiveSheetIndex(0)->setCellValue('L1215', "2013");
				$excel->setActiveSheetIndex(0)->setCellValue('C1216', "30 November 2013");

				$excel->setActiveSheetIndex(0)->setCellValue('C1218', "'Pengembangan E-Commerse T Menggunakan Sistem Database Terdistrubsi (Studi Kasus: Penjualan Dvd Game Terdistribusi)");
				$excel->getActiveSheet()->mergeCells('C1218:K1218');
        $excel->getActiveSheet()->getStyle('C1218')->getAlignment()->setWrapText(TRUE);
        $excel->getActiveSheet()->getStyle('C1218')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);
				$excel->setActiveSheetIndex(0)->setCellValue('R1218', "No. 116/P/B/N/FMIPA/2019");
				$excel->getActiveSheet()->getStyle('R1218')->getAlignment()->setWrapText(TRUE);
				$excel->setActiveSheetIndex(0)->setCellValue('R1219', "Tanggal 27 Juni 2019");
				$excel->setActiveSheetIndex(0)->setCellValue('C1221', "Favorisen R. Lumbanraja dan Aristoteles");
				$excel->getActiveSheet()->mergeCells('C1221:K1221');

				$excel->setActiveSheetIndex(0)->setCellValue('B1223', "6)");
				$excel->getActiveSheet()->getStyle('B1223')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
				$excel->setActiveSheetIndex(0)->setCellValue('C1223', "Prosiding SN SMAIP III 2012");
				$excel->setActiveSheetIndex(0)->setCellValue('L1223', "Juni");
				$excel->setActiveSheetIndex(0)->setCellValue('M1223', "Prosiding");
				$excel->setActiveSheetIndex(0)->setCellValue('N1223', "1,00");
				$excel->setActiveSheetIndex(0)->setCellValue('R1223', "http://repository.lppm.unila.ac.id/1368/ ");
				$excel->setActiveSheetIndex(0)->setCellValue('C1224', "ISBN No. 978-602-98559-1-3");
				$excel->setActiveSheetIndex(0)->setCellValue('L1224', "2012");
				$excel->setActiveSheetIndex(0)->setCellValue('C1224', "Juni 2012");

				$excel->setActiveSheetIndex(0)->setCellValue('C1226', "Implementasi Algoritma Half-Byte Dengan Nilai Parameter 7 Pada Kompresi File Gambar, Teks, Audio, Dan Video");
				$excel->getActiveSheet()->mergeCells('C1226:K1226');
        $excel->getActiveSheet()->getStyle('C1226')->getAlignment()->setWrapText(TRUE);
        $excel->getActiveSheet()->getStyle('C1226')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);
				$excel->setActiveSheetIndex(0)->setCellValue('R1226', "No. 119/P/B/N/FMIPA/2019");
				$excel->getActiveSheet()->getStyle('R1226')->getAlignment()->setWrapText(TRUE);
				$excel->setActiveSheetIndex(0)->setCellValue('R1227', "Tanggal 27 Juni 2019");
				$excel->setActiveSheetIndex(0)->setCellValue('C1229', "Anggar Bagus Kurniawan, Aristoteles, Machudor");
				$excel->getActiveSheet()->mergeCells('C1229:K1229');
				$excel->setActiveSheetIndex(0)->setCellValue('N1231', "Jumlah");
        $excel->setActiveSheetIndex(0)->setCellValue('O1231', "0,00");
    		$excel->setActiveSheetIndex(0)->setCellValue('P1231', "0,00");
    		$excel->getActiveSheet()->getStyle('N1231:P1231')->getFont()->setBold(TRUE);
				$excel->getActiveSheet()->getStyle('A1231:R1231')->applyFromArray($style_standar);

				$excel->setActiveSheetIndex(0)->setCellValue('E1232', "2)");
				$excel->getActiveSheet()->getStyle('E1232')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
				$excel->setActiveSheetIndex(0)->setCellValue('F1232', "Poster tingkat:	");
				$excel->setActiveSheetIndex(0)->setCellValue('F1233', "a) Internasional");
				$excel->setActiveSheetIndex(0)->setCellValue('N1234', "Jumlah");
        $excel->setActiveSheetIndex(0)->setCellValue('O1234', "0,00");
    		$excel->setActiveSheetIndex(0)->setCellValue('P1234', "0,00");
    		$excel->getActiveSheet()->getStyle('N1234:P1234')->getFont()->setBold(TRUE);
				$excel->getActiveSheet()->getStyle('A1234:R1234')->applyFromArray($style_standar);
				$excel->setActiveSheetIndex(0)->setCellValue('F1235', "b) Nasional");
				$excel->setActiveSheetIndex(0)->setCellValue('N1236', "Jumlah");
        $excel->setActiveSheetIndex(0)->setCellValue('O1236', "0,00");
    		$excel->setActiveSheetIndex(0)->setCellValue('P1236', "0,00");
    		$excel->getActiveSheet()->getStyle('N1236:P1236')->getFont()->setBold(TRUE);
				$excel->getActiveSheet()->getStyle('A1236:R1236')->applyFromArray($style_standar);
				$excel->setActiveSheetIndex(0)->setCellValue('D1237', "d.");
				$excel->getActiveSheet()->getStyle('D1237')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
				$excel->setActiveSheetIndex(0)->setCellValue('E1237', "Dalam koran/majalah populer/umum");
				$excel->getActiveSheet()->mergeCells('E1237:K1237');
        $excel->getActiveSheet()->getStyle('E1237')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);
				$excel->setActiveSheetIndex(0)->setCellValue('N1238', "Jumlah");
        $excel->setActiveSheetIndex(0)->setCellValue('O1238', "0,00");
    		$excel->setActiveSheetIndex(0)->setCellValue('P1238', "0,00");
    		$excel->getActiveSheet()->getStyle('N1238:P1238')->getFont()->setBold(TRUE);
				$excel->getActiveSheet()->getStyle('A1238:R1238')->applyFromArray($style_standar);

				$excel->setActiveSheetIndex(0)->setCellValue('C1239', "2.");
				$excel->getActiveSheet()->getStyle('C1239')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
				$excel->setActiveSheetIndex(0)->setCellValue('D1239', "Hasil penelitian atau hasil pemikiran yang tidak di publikasikan (tersimpan di perpustakaan perguruan tinggi)");
				$excel->getActiveSheet()->mergeCells('D1239:K1239');
				$excel->setActiveSheetIndex(0)->setCellValue('N1240', "Jumlah");
        $excel->setActiveSheetIndex(0)->setCellValue('O1240', "0,00");
    		$excel->setActiveSheetIndex(0)->setCellValue('P1240', "0,00");
    		$excel->getActiveSheet()->getStyle('N1240:P1240')->getFont()->setBold(TRUE);
				$excel->getActiveSheet()->getStyle('A240:R240')->applyFromArray($style_standar);

				$excel->setActiveSheetIndex(0)->setCellValue('B1241', "B.");
				$excel->getActiveSheet()->getStyle('B1241')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
        $excel->setActiveSheetIndex(0)->setCellValue('C1241', "Menerjemahkan / menyadur buku ilmiah");
    		$excel->setActiveSheetIndex(0)->setCellValue('D1242', "Diterbitkan dan diedarkan secara nasional.");
				$excel->setActiveSheetIndex(0)->setCellValue('N1243', "Jumlah");
        $excel->setActiveSheetIndex(0)->setCellValue('O1243', "0,00");
    		$excel->setActiveSheetIndex(0)->setCellValue('P1243', "0,00");
    		$excel->getActiveSheet()->getStyle('N1243:P1243')->getFont()->setBold(TRUE);
				$excel->getActiveSheet()->getStyle('A1243:R1243')->applyFromArray($style_standar);

				$excel->setActiveSheetIndex(0)->setCellValue('B1244', "C.");
				$excel->getActiveSheet()->getStyle('B1244')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
        $excel->setActiveSheetIndex(0)->setCellValue('C1244', "Mengedit/menyunting karya ilmiah");
    		$excel->setActiveSheetIndex(0)->setCellValue('D1245', "Diterbitkan dan diedarkan secara nasional.");
				$excel->setActiveSheetIndex(0)->setCellValue('N1246', "Jumlah");
        $excel->setActiveSheetIndex(0)->setCellValue('O1246', "0,00");
    		$excel->setActiveSheetIndex(0)->setCellValue('P1246', "0,00");
    		$excel->getActiveSheet()->getStyle('N1246:P1246')->getFont()->setBold(TRUE);
				$excel->getActiveSheet()->getStyle('A1246:R1246')->applyFromArray($style_standar);

				$excel->setActiveSheetIndex(0)->setCellValue('B1247', "D.");
				$excel->getActiveSheet()->getStyle('B1247')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
        $excel->setActiveSheetIndex(0)->setCellValue('C1247', "Membuat rencana dan karya teknologi yang dipatenkan");
				$excel->setActiveSheetIndex(0)->setCellValue('C1248', "1");
				$excel->getActiveSheet()->getStyle('C1248')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
    		$excel->setActiveSheetIndex(0)->setCellValue('D1248', "Internasional");
				$excel->setActiveSheetIndex(0)->setCellValue('N1249', "Jumlah");
        $excel->setActiveSheetIndex(0)->setCellValue('O1249', "0,00");
    		$excel->setActiveSheetIndex(0)->setCellValue('P1249', "0,00");
    		$excel->getActiveSheet()->getStyle('N1249:P1249')->getFont()->setBold(TRUE);
				$excel->getActiveSheet()->getStyle('A1249:R1249')->applyFromArray($style_standar);
				$excel->setActiveSheetIndex(0)->setCellValue('C1250', "2");
				$excel->getActiveSheet()->getStyle('C1250')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
				$excel->setActiveSheetIndex(0)->setCellValue('D1251', "Nasional");
				$excel->setActiveSheetIndex(0)->setCellValue('N1251', "Jumlah");
        $excel->setActiveSheetIndex(0)->setCellValue('O1251', "0,00");
    		$excel->setActiveSheetIndex(0)->setCellValue('P1251', "0,00");
    		$excel->getActiveSheet()->getStyle('N1251:P1251')->getFont()->setBold(TRUE);
				$excel->getActiveSheet()->getStyle('A1251:R1251')->applyFromArray($style_standar);

				$excel->setActiveSheetIndex(0)->setCellValue('B1252', "E.");
				$excel->getActiveSheet()->getStyle('B1252')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
        $excel->setActiveSheetIndex(0)->setCellValue('C1252', "Membuat rancangan dan karya teknologi, rancangan dan karya seni monumental/seni pertunjukan/karya sastra ");
				$excel->getActiveSheet()->mergeCells('C1252:K1252');
        $excel->getActiveSheet()->getStyle('C1252')->getAlignment()->setWrapText(TRUE);
        $excel->getActiveSheet()->getStyle('C1252')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);
				$excel->setActiveSheetIndex(0)->setCellValue('C1253', "1");
				$excel->getActiveSheet()->getStyle('C1253')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
    		$excel->setActiveSheetIndex(0)->setCellValue('D1253', "Tingkat Internasional");
				$excel->setActiveSheetIndex(0)->setCellValue('N1254', "Jumlah");
        $excel->setActiveSheetIndex(0)->setCellValue('O1254', "0,00");
    		$excel->setActiveSheetIndex(0)->setCellValue('P1254', "0,00");
    		$excel->getActiveSheet()->getStyle('N1254:P1254')->getFont()->setBold(TRUE);
				$excel->getActiveSheet()->getStyle('A1254:R1254')->applyFromArray($style_standar);
				$excel->setActiveSheetIndex(0)->setCellValue('C1255', "1");
				$excel->setActiveSheetIndex(0)->setCellValue('D1255', "Tingkat Nasional");
				$excel->setActiveSheetIndex(0)->setCellValue('N1256', "Jumlah");
        $excel->setActiveSheetIndex(0)->setCellValue('O1256', "0,00");
    		$excel->setActiveSheetIndex(0)->setCellValue('P1256', "0,00");
    		$excel->getActiveSheet()->getStyle('N1256:P1256')->getFont()->setBold(TRUE);
				$excel->getActiveSheet()->getStyle('A1256:R1256')->applyFromArray($style_standar);
				$excel->setActiveSheetIndex(0)->setCellValue('C1257', "1");
				$excel->setActiveSheetIndex(0)->setCellValue('D1257', "Tingkat Lokal");
				$excel->setActiveSheetIndex(0)->setCellValue('N1258', "Jumlah");
        $excel->setActiveSheetIndex(0)->setCellValue('O1258', "0,00");
    		$excel->setActiveSheetIndex(0)->setCellValue('P1258', "0,00");
    		$excel->getActiveSheet()->getStyle('N1258:P1258')->getFont()->setBold(TRUE);
				$excel->getActiveSheet()->getStyle('A1258:R1258')->applyFromArray($style_standar);
				$excel->setActiveSheetIndex(0)->setCellValue('A1259', "Jumlah Penelitian");
				$excel->getActiveSheet()->mergeCells('A1259:O1259');
				$excel->getActiveSheet()->getStyle('A1259')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
				$excel->setActiveSheetIndex(0)->setCellValue('P1259', "0,00");
				$excel->getActiveSheet()->getStyle('A1259:R1259')->applyFromArray($style_standar);

				$excel->setActiveSheetIndex(0)->setCellValue('O1261', "Bandar Lampung,  31 Juli 2019");
				$excel->setActiveSheetIndex(0)->setCellValue('O1262', "Ketua Jurusan Ilmu Komputer");
				$excel->setActiveSheetIndex(0)->setCellValue('O1268', "Dr.Ir. Kurnia Muludi, M.S.Sc");
				$excel->setActiveSheetIndex(0)->setCellValue('O1269', "NIP. 19640616 198902 1 001");



				$excel->getActiveSheet()->getDefaultRowDimension()->setRowHeight(-1);
        // Set orientasi kertas jadi LANDSCAPE
        $excel->getActiveSheet()->getPageSetup()->setOrientation(PHPExcel_Worksheet_PageSetup::ORIENTATION_LANDSCAPE);

        // Set judul file excel nya
        $excel->getActiveSheet(0)->setTitle("Report");
        $excel->setActiveSheetIndex(0);

        // Proses file excel
        header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        header('Content-Disposition: attachment; filename="Report.xlsx"'); // Set nama file excel nya
        header('Cache-Control: max-age=0');

        $write = PHPExcel_IOFactory::createWriter($excel, 'Excel2007');
        $write->save('php://output');
    	}
}
