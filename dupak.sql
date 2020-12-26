-- phpMyAdmin SQL Dump
-- version 4.9.0.1
-- https://www.phpmyadmin.net/
--
-- Host: 127.0.0.1
-- Generation Time: Nov 07, 2020 at 12:44 AM
-- Server version: 10.4.6-MariaDB
-- PHP Version: 7.3.8

SET SQL_MODE = "NO_AUTO_VALUE_ON_ZERO";
SET AUTOCOMMIT = 0;
START TRANSACTION;
SET time_zone = "+00:00";


/*!40101 SET @OLD_CHARACTER_SET_CLIENT=@@CHARACTER_SET_CLIENT */;
/*!40101 SET @OLD_CHARACTER_SET_RESULTS=@@CHARACTER_SET_RESULTS */;
/*!40101 SET @OLD_COLLATION_CONNECTION=@@COLLATION_CONNECTION */;
/*!40101 SET NAMES utf8mb4 */;

--
-- Database: `dupak`
--

-- --------------------------------------------------------

--
-- Table structure for table `data_dosen`
--

CREATE TABLE `data_dosen` (
  `id_dosen` int(15) NOT NULL,
  `nama` varchar(100) NOT NULL,
  `nip` varchar(25) NOT NULL,
  `pangkat` varchar(50) NOT NULL,
  `golongan` varchar(10) NOT NULL,
  `jabatan` varchar(200) NOT NULL,
  `unit_kerja` varchar(250) NOT NULL
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

--
-- Dumping data for table `data_dosen`
--

INSERT INTO `data_dosen` (`id_dosen`, `nama`, `nip`, `pangkat`, `golongan`, `jabatan`, `unit_kerja`) VALUES
(2, 'Dr. Ir. Kurnia Muludi, M.S.Sc', '196406161989021001', 'Pembina', 'IV.a', 'Ketua Jurusan Ilmu Komputer', 'Fakultas MIPA Universitas Lampung');

-- --------------------------------------------------------

--
-- Table structure for table `dosen_penunjang`
--

CREATE TABLE `dosen_penunjang` (
  `id_dosen` int(10) NOT NULL,
  `nama` varchar(100) NOT NULL,
  `nip` varchar(100) NOT NULL,
  `pangkat` varchar(50) NOT NULL,
  `golongan` varchar(25) NOT NULL,
  `jabatan` varchar(50) NOT NULL,
  `unit_kerja` varchar(100) NOT NULL
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

--
-- Dumping data for table `dosen_penunjang`
--

INSERT INTO `dosen_penunjang` (`id_dosen`, `nama`, `nip`, `pangkat`, `golongan`, `jabatan`, `unit_kerja`) VALUES
(1, 'Aristoteles, M.Si', '198105212006040000', 'Penata', 'III.c', 'Lektor', 'Jurusan Ilmu Komputer, Fakultas MIPA Universitas Lampung');

-- --------------------------------------------------------

--
-- Table structure for table `login`
--

CREATE TABLE `login` (
  `id` int(5) NOT NULL,
  `username` varchar(100) NOT NULL,
  `password` varchar(50) NOT NULL
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

--
-- Dumping data for table `login`
--

INSERT INTO `login` (`id`, `username`, `password`) VALUES
(1, 'Rizky Prabowo', 'e889a217392286c7fbaec4f07aa7caea'),
(2, 'Aristoteles', 'd80d050ae43358fdfbb8fc04e7088f60'),
(3, 'Kurnia Muludi', 'abf59ed6f0dd0bfb049b850713046a1c');

--
-- Indexes for dumped tables
--

--
-- Indexes for table `data_dosen`
--
ALTER TABLE `data_dosen`
  ADD PRIMARY KEY (`id_dosen`);

--
-- Indexes for table `dosen_penunjang`
--
ALTER TABLE `dosen_penunjang`
  ADD PRIMARY KEY (`id_dosen`);

--
-- Indexes for table `login`
--
ALTER TABLE `login`
  ADD PRIMARY KEY (`id`);
COMMIT;

/*!40101 SET CHARACTER_SET_CLIENT=@OLD_CHARACTER_SET_CLIENT */;
/*!40101 SET CHARACTER_SET_RESULTS=@OLD_CHARACTER_SET_RESULTS */;
/*!40101 SET COLLATION_CONNECTION=@OLD_COLLATION_CONNECTION */;
