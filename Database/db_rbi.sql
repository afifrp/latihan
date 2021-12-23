-- phpMyAdmin SQL Dump
-- version 5.1.0
-- https://www.phpmyadmin.net/
--
-- Host: 127.0.0.1
-- Generation Time: May 02, 2021 at 09:07 AM
-- Server version: 10.4.18-MariaDB
-- PHP Version: 8.0.3

SET SQL_MODE = "NO_AUTO_VALUE_ON_ZERO";
START TRANSACTION;
SET time_zone = "+00:00";


/*!40101 SET @OLD_CHARACTER_SET_CLIENT=@@CHARACTER_SET_CLIENT */;
/*!40101 SET @OLD_CHARACTER_SET_RESULTS=@@CHARACTER_SET_RESULTS */;
/*!40101 SET @OLD_COLLATION_CONNECTION=@@COLLATION_CONNECTION */;
/*!40101 SET NAMES utf8mb4 */;

--
-- Database: `db_rbi`
--

-- --------------------------------------------------------

--
-- Table structure for table `tbl_fluida`
--

CREATE TABLE `tbl_fluida` (
  `units` varchar(6) COLLATE latin1_general_ci NOT NULL,
  `kode_fluida` varchar(50) COLLATE latin1_general_ci NOT NULL,
  `nama_fluida` varchar(50) COLLATE latin1_general_ci NOT NULL,
  `density` double NOT NULL,
  `nbp` double NOT NULL,
  `mw` double NOT NULL,
  `ait` double NOT NULL,
  `liquid_flow_velocity` double NOT NULL,
  `viscosity` double NOT NULL,
  `igc_a` double NOT NULL,
  `igc_b` double NOT NULL,
  `igc_c` double NOT NULL,
  `igc_d` double NOT NULL,
  `igc_e` double NOT NULL,
  `fluid_type` varchar(10) COLLATE latin1_general_ci NOT NULL,
  `ambient_state` varchar(10) COLLATE latin1_general_ci NOT NULL,
  `liquid_dynamic_viscosity` double NOT NULL
) ENGINE=InnoDB DEFAULT CHARSET=latin1 COLLATE=latin1_general_ci;

--
-- Dumping data for table `tbl_fluida`
--

INSERT INTO `tbl_fluida` (`units`, `kode_fluida`, `nama_fluida`, `density`, `nbp`, `mw`, `ait`, `liquid_flow_velocity`, `viscosity`, `igc_a`, `igc_b`, `igc_c`, `igc_d`, `igc_e`, `fluid_type`, `ambient_state`, `liquid_dynamic_viscosity`) VALUES
('SI', 'A01', 'Ammonia', 233.44, 333.2, 33, 45, 0.45, 0.5, 0.1, 0.11, 0.111, 0.111, 0.11233, 'TYPE 0', 'LIQUID', 0.94);

-- --------------------------------------------------------

--
-- Table structure for table `tbl_material`
--

CREATE TABLE `tbl_material` (
  `units` varchar(6) COLLATE latin1_general_ci NOT NULL,
  `kode_material` varchar(20) COLLATE latin1_general_ci NOT NULL,
  `nama_material` varchar(50) COLLATE latin1_general_ci NOT NULL,
  `yield_strength` double NOT NULL,
  `tensile_strength` double NOT NULL,
  `design_pressure` double NOT NULL,
  `max_operating_pressure` double NOT NULL,
  `design_temperature` double NOT NULL,
  `max_design_temperature` double NOT NULL,
  `cost_factor` double NOT NULL
) ENGINE=InnoDB DEFAULT CHARSET=latin1 COLLATE=latin1_general_ci;

--
-- Dumping data for table `tbl_material`
--

INSERT INTO `tbl_material` (`units`, `kode_material`, `nama_material`, `yield_strength`, `tensile_strength`, `design_pressure`, `max_operating_pressure`, `design_temperature`, `max_design_temperature`, `cost_factor`) VALUES
('SI', 'SS459', 'Carbon Steel', 445, 540, 345.66, 345.77, 300, 690.55, 500000),
('SI', 'ST42', 'Stainless Steel', 340, 450, 234.66, 455.5, 600.77, 970.65, 700000);

--
-- Indexes for dumped tables
--

--
-- Indexes for table `tbl_fluida`
--
ALTER TABLE `tbl_fluida`
  ADD PRIMARY KEY (`kode_fluida`);

--
-- Indexes for table `tbl_material`
--
ALTER TABLE `tbl_material`
  ADD PRIMARY KEY (`kode_material`);
COMMIT;

/*!40101 SET CHARACTER_SET_CLIENT=@OLD_CHARACTER_SET_CLIENT */;
/*!40101 SET CHARACTER_SET_RESULTS=@OLD_CHARACTER_SET_RESULTS */;
/*!40101 SET COLLATION_CONNECTION=@OLD_COLLATION_CONNECTION */;
