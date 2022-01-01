-- phpMyAdmin SQL Dump
-- version 5.1.0
-- https://www.phpmyadmin.net/
--
-- Host: 127.0.0.1
-- Generation Time: Jan 01, 2022 at 10:55 AM
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
-- Table structure for table `tbl_coflevel1`
--

CREATE TABLE `tbl_coflevel1` (
  `equipment_code` varchar(20) NOT NULL,
  `store` varchar(10) NOT NULL,
  `storage_condition` varchar(10) NOT NULL,
  `consequence_area_type` varchar(30) NOT NULL,
  `fluid_mass_component` double NOT NULL,
  `fluid_mass_inventory` double NOT NULL,
  `detection_system` double NOT NULL,
  `isolation_system` double NOT NULL
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;

-- --------------------------------------------------------

--
-- Table structure for table `tbl_equipment`
--

CREATE TABLE `tbl_equipment` (
  `equipment_code` varchar(20) NOT NULL,
  `equipment_description` varchar(100) NOT NULL,
  `equipment_type` varchar(50) NOT NULL,
  `component_type` varchar(50) NOT NULL,
  `industry_name` varchar(50) NOT NULL
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;

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
-- Table structure for table `tbl_fms`
--

CREATE TABLE `tbl_fms` (
  `Code of Equipment` varchar(30) NOT NULL,
  `Leadershipadmin1` tinyint(1) NOT NULL,
  `Leadership admin 2a` tinyint(1) NOT NULL,
  `Leadershipadmin2b` tinyint(1) NOT NULL,
  `Leadershipadmin2c` tinyint(1) NOT NULL,
  `Leadershipadmin2d` tinyint(1) NOT NULL,
  `Leadershipadmin2e` tinyint(1) NOT NULL,
  `Leadershipadmin` varchar(200) NOT NULL,
  `Leadershipadmin3` tinyint(1) NOT NULL,
  `Leadershipadmin4` tinyint(1) NOT NULL,
  `Leadershipadmin5` double NOT NULL,
  `Leadershipadmin6` tinyint(1) NOT NULL,
  `Leadershipadmin6a` tinyint(1) NOT NULL,
  `Leadershipadmin 6b` tinyint(1) NOT NULL,
  `Processsafety1` tinyint(1) NOT NULL,
  `Processsafety1a` tinyint(1) NOT NULL,
  `Processsafety1b` tinyint(1) NOT NULL,
  `Processsafety1c` tinyint(1) NOT NULL,
  `Processsafety2` tinyint(1) NOT NULL,
  `Processsafety3a` tinyint(1) NOT NULL,
  `Processsafety3b` tinyint(1) NOT NULL,
  `Processsafety3c` tinyint(1) NOT NULL,
  `Processsafety4` tinyint(1) NOT NULL,
  `Processsafety5` tinyint(1) NOT NULL,
  `Processsafety6` tinyint(1) NOT NULL,
  `Processsafety7a` tinyint(1) NOT NULL,
  `Processsafety7b` tinyint(1) NOT NULL,
  `Processsafety8a` tinyint(1) NOT NULL,
  `Processsafety8b` tinyint(1) NOT NULL,
  `Processsafety8c` tinyint(1) NOT NULL,
  `Processsafety8d` tinyint(1) NOT NULL,
  `Processsafety8e` tinyint(1) NOT NULL,
  `Processsafety8f` tinyint(1) NOT NULL,
  `Processsafety9` tinyint(1) NOT NULL,
  `Processsafety10` tinyint(1) NOT NULL,
  `Process hazard 1` double NOT NULL,
  `Processhazard2` tinyint(1) NOT NULL,
  `Processhazard2a` tinyint(1) NOT NULL,
  `Processhazard2b` tinyint(1) NOT NULL,
  `Process hazard 2c` tinyint(1) NOT NULL,
  `Process hazard 2d` tinyint(1) NOT NULL,
  `Process hazard 2e` tinyint(1) NOT NULL,
  `Processhazard3a` tinyint(1) NOT NULL,
  `Processhazard3b` tinyint(1) NOT NULL,
  `Processhazard3c` tinyint(1) NOT NULL,
  `Processhazard3d` tinyint(1) NOT NULL,
  `Processhazard3e` tinyint(1) NOT NULL,
  `Processhazard3f` tinyint(1) NOT NULL,
  `Processhazard3g` tinyint(1) NOT NULL,
  `Processhazard4a` tinyint(1) NOT NULL,
  `Processhazard4b` tinyint(1) NOT NULL,
  `Processhazard4c` tinyint(1) NOT NULL,
  `Processhazard4d` tinyint(1) NOT NULL,
  `Processhazard4e` tinyint(1) NOT NULL,
  `Processhazard5` tinyint(1) NOT NULL,
  `Processhazard5a` tinyint(1) NOT NULL,
  `Processhazard5b` tinyint(1) NOT NULL,
  `Processhazard6` tinyint(1) NOT NULL,
  `Processhazard7` tinyint(1) NOT NULL,
  `Processhazard8` tinyint(1) NOT NULL,
  `Processhazard9` tinyint(1) NOT NULL,
  `Managementchange1a` tinyint(1) NOT NULL,
  `Managementchange1b` tinyint(1) NOT NULL,
  `Managementchange2a` tinyint(1) NOT NULL,
  `Managementchange2b` tinyint(1) NOT NULL,
  `Managementchange2c` tinyint(1) NOT NULL,
  `Managementchange2d` tinyint(1) NOT NULL,
  `Managementchange3` tinyint(1) NOT NULL,
  `Managementchange3a` tinyint(1) NOT NULL,
  `Managementchange3b` tinyint(1) NOT NULL,
  `Managementchange4a` tinyint(1) NOT NULL,
  `Managementchange4b` tinyint(1) NOT NULL,
  `Managementchange4c` tinyint(1) NOT NULL,
  `Managementchange4d` tinyint(1) NOT NULL,
  `Managementchange4e` tinyint(1) NOT NULL,
  `Managementchange4f` tinyint(1) NOT NULL,
  `Managementchange4g` tinyint(1) NOT NULL,
  `Managementchange5` tinyint(1) NOT NULL,
  `Managementchange6` tinyint(1) NOT NULL,
  `Operatingprocedures1a` tinyint(1) NOT NULL,
  `Operatingprocedures1b` tinyint(1) NOT NULL,
  `Operatingprocedures2a` tinyint(1) NOT NULL,
  `Operatingprocedures2b` tinyint(1) NOT NULL,
  `Operatingprocedures2c` tinyint(1) NOT NULL,
  `Operatingprocedures2d` tinyint(1) NOT NULL,
  `Operatingprocedures2e` tinyint(1) NOT NULL,
  `Operatingprocedures2f` tinyint(1) NOT NULL,
  `Operatingprocedures2g` tinyint(1) NOT NULL,
  `Operatingprocedures2h` tinyint(1) NOT NULL,
  `Operatingprocedures3a` tinyint(1) NOT NULL,
  `Operatingprocedures3b` tinyint(1) NOT NULL,
  `Operatingprocedures3c` tinyint(1) NOT NULL,
  `Operatingprocedures4` tinyint(1) NOT NULL,
  `Operatingprocedures5` tinyint(1) NOT NULL,
  `Operatingprocedures6a` tinyint(1) NOT NULL,
  `Operatingprocedures6b` tinyint(1) NOT NULL,
  `Operatingprocedures6c` tinyint(1) NOT NULL,
  `Operatingprocedures6d` tinyint(1) NOT NULL,
  `Operatingprocedures7a` tinyint(1) NOT NULL,
  `Operatingprocedures7b` tinyint(1) NOT NULL,
  `Operatingprocedures7c` tinyint(1) NOT NULL,
  `Operatingprocedures7d` tinyint(1) NOT NULL,
  `Safework1a` tinyint(1) NOT NULL,
  `Safework1b` tinyint(1) NOT NULL,
  `Safework1c` tinyint(1) NOT NULL,
  `Safework1d` tinyint(1) NOT NULL,
  `Safework1e` tinyint(1) NOT NULL,
  `Safework1f` tinyint(1) NOT NULL,
  `Safework1g` tinyint(1) NOT NULL,
  `Safework1h` tinyint(1) NOT NULL,
  `Safework1i` tinyint(1) NOT NULL,
  `Safework1j` tinyint(1) NOT NULL,
  `Safework2` tinyint(1) NOT NULL,
  `Safework3a` tinyint(1) NOT NULL,
  `Safework3b` tinyint(1) NOT NULL,
  `Safework3c` tinyint(1) NOT NULL,
  `Safework3d` tinyint(1) NOT NULL,
  `Safework3e` tinyint(1) NOT NULL,
  `Safework4` tinyint(1) NOT NULL,
  `Safework5` tinyint(1) NOT NULL,
  `Safework6a` tinyint(1) NOT NULL,
  `Safework6b` tinyint(1) NOT NULL,
  `Safework6c` tinyint(1) NOT NULL,
  `Safework6d` tinyint(1) NOT NULL,
  `Safework7a` tinyint(1) NOT NULL,
  `Safework7b` tinyint(1) NOT NULL,
  `Safework8a` tinyint(1) NOT NULL,
  `Safework8b` tinyint(1) NOT NULL,
  `Training1` tinyint(1) NOT NULL,
  `Training2` tinyint(1) NOT NULL,
  `Training3a` tinyint(1) NOT NULL,
  `Training3b` tinyint(1) NOT NULL,
  `Training3c` tinyint(1) NOT NULL,
  `Training3d` tinyint(1) NOT NULL,
  `Training3e` tinyint(1) NOT NULL,
  `Training3f` tinyint(1) NOT NULL,
  `Training4a` tinyint(1) NOT NULL,
  `Training4b` tinyint(1) NOT NULL,
  `Training4c` tinyint(1) NOT NULL,
  `Training4d` tinyint(1) NOT NULL,
  `Training5a` tinyint(1) NOT NULL,
  `Training5b` tinyint(1) NOT NULL,
  `Training5c` tinyint(1) NOT NULL,
  `Training6a` tinyint(1) NOT NULL,
  `Training6b` tinyint(1) NOT NULL,
  `Training6c` tinyint(1) NOT NULL,
  `Training6d` tinyint(1) NOT NULL,
  `Training6e` tinyint(1) NOT NULL,
  `Training7` tinyint(1) NOT NULL,
  `Training7a` tinyint(1) NOT NULL,
  `Training7b` tinyint(1) NOT NULL,
  `Training8a` tinyint(1) NOT NULL,
  `Training8b` tinyint(1) NOT NULL,
  `Training8c` tinyint(1) NOT NULL,
  `Training8d` tinyint(1) NOT NULL,
  `Mechanicalintegrity1a` tinyint(1) NOT NULL,
  `Mechanicalintegrity1b` tinyint(1) NOT NULL,
  `Mechanicalintegrity1c` tinyint(1) NOT NULL,
  `Mechanicalintegrity1d` tinyint(1) NOT NULL,
  `Mechanicalintegrity1e` tinyint(1) NOT NULL,
  `Mechanicalintegrity2` tinyint(1) NOT NULL,
  `Mechanicalintegrity2a` tinyint(1) NOT NULL,
  `Mechanicalintegrity2b` tinyint(1) NOT NULL,
  `Mechanicalintegrity2c` tinyint(1) NOT NULL,
  `Mechanicalintegrity3` tinyint(1) NOT NULL,
  `Mechanicalintegrity4` tinyint(1) NOT NULL,
  `Mechanicalintegrity4a` tinyint(1) NOT NULL,
  `Mechanicalintegrity4b` tinyint(1) NOT NULL,
  `Mechanicalintegrity5` tinyint(1) NOT NULL,
  `Mechanicalintegrity5a1` tinyint(1) NOT NULL,
  `Mechanicalintegrity5a2` tinyint(1) NOT NULL,
  `Mechanicalintegrity5b` tinyint(1) NOT NULL,
  `Mechanicalintegrity5c` tinyint(1) NOT NULL,
  `Mechanicalintegrity5d` tinyint(1) NOT NULL,
  `Mechanicalintegrity6a` tinyint(1) NOT NULL,
  `Mechanicalintegrity6b` tinyint(1) NOT NULL,
  `Mechanicalintegrity7` tinyint(1) NOT NULL,
  `Mechanicalintegrity8a` tinyint(1) NOT NULL,
  `Mechanicalintegrity8b` tinyint(1) NOT NULL,
  `Mechanicalintegrity9` tinyint(1) NOT NULL,
  `Mechanicalintegrity9a` tinyint(1) NOT NULL,
  `Mechanicalintegrity10` tinyint(1) NOT NULL,
  `Mechanicalintegrity10a` tinyint(1) NOT NULL,
  `Mechanicalintegrity10b` tinyint(1) NOT NULL,
  `Mechanicalintegrity11a` tinyint(1) NOT NULL,
  `Mechanicalintegrity11b` tinyint(1) NOT NULL,
  `Mechanicalintegrity12` tinyint(1) NOT NULL,
  `Mechanicalintegrity13a` tinyint(1) NOT NULL,
  `Mechanicalintegrity13b` tinyint(1) NOT NULL,
  `Mechanicalintegrity14` tinyint(1) NOT NULL,
  `Mechanicalintegrity15` tinyint(1) NOT NULL,
  `Mechanicalintegrity16` tinyint(1) NOT NULL,
  `Mechanicalintegrity16a` tinyint(1) NOT NULL,
  `Mechanicalintegrity16b` tinyint(1) NOT NULL,
  `Mechanicalintegrity16c` tinyint(1) NOT NULL,
  `Mechanicalintegrity17a` tinyint(1) NOT NULL,
  `Mechanicalintegrity17b` tinyint(1) NOT NULL,
  `Mechanicalintegrity17c` tinyint(1) NOT NULL,
  `Mechanicalintegrity17d` tinyint(1) NOT NULL,
  `Mechanicalintegrity17e` tinyint(1) NOT NULL,
  `Mechanicalintegrity18a` tinyint(1) NOT NULL,
  `Mechanicalintegrity18b` tinyint(1) NOT NULL,
  `Mechanicalintegrity18c` tinyint(1) NOT NULL,
  `Mechanicalintegrity18d` tinyint(1) NOT NULL,
  `Mechanicalintegrity18e` tinyint(1) NOT NULL,
  `Mechanicalintegrity19` tinyint(1) NOT NULL,
  `Mechanicalintegrity20` tinyint(1) NOT NULL,
  `Prestartupsafety1` tinyint(1) NOT NULL,
  `Prestartupsafety2` tinyint(1) NOT NULL,
  `Prestartupsafety3` tinyint(1) NOT NULL,
  `Prestartupsafety3a` tinyint(1) NOT NULL,
  `Prestartupsafety3b` tinyint(1) NOT NULL,
  `Prestartupsafety4a` tinyint(1) NOT NULL,
  `Prestartupsafety4b` tinyint(1) NOT NULL,
  `Prestartupsafety4c` tinyint(1) NOT NULL,
  `Prestartupsafety5` tinyint(1) NOT NULL,
  `Emergencyresponse1` tinyint(1) NOT NULL,
  `Emergencyresponse2` tinyint(1) NOT NULL,
  `Emergencyresponse2a` tinyint(1) NOT NULL,
  `Emergencyresponse2b` tinyint(1) NOT NULL,
  `Emergencyresponse3a` tinyint(1) NOT NULL,
  `Emergencyresponse3b` tinyint(1) NOT NULL,
  `Emergencyresponse3c` tinyint(1) NOT NULL,
  `Emergencyresponse3d` tinyint(1) NOT NULL,
  `Emergencyresponse3e` tinyint(1) NOT NULL,
  `Emergencyresponse3f` tinyint(1) NOT NULL,
  `Emergencyresponse3g` tinyint(1) NOT NULL,
  `Emergencyresponse3h` tinyint(1) NOT NULL,
  `Emergencyresponse3i` tinyint(1) NOT NULL,
  `Emergencyresponse4` tinyint(1) NOT NULL,
  `Emergency response 4a` tinyint(1) NOT NULL,
  `Emergencyresponse4b` tinyint(1) NOT NULL,
  `Emergencyresponse4c` tinyint(1) NOT NULL,
  `Emergencyresponse5a` tinyint(1) NOT NULL,
  `Emergencyresponse5b` tinyint(1) NOT NULL,
  `Emergencyresponse6` tinyint(1) NOT NULL,
  `Incidentinvestigation1a` tinyint(1) NOT NULL,
  `Incidentinvestigation1b` tinyint(1) NOT NULL,
  `Incidentinvestigation2a` tinyint(1) NOT NULL,
  `Incidentinvestigation2b` tinyint(1) NOT NULL,
  `Incidentinvestigation3a` tinyint(1) NOT NULL,
  `Incidentinvestigation3b` tinyint(1) NOT NULL,
  `Incidentinvestigation3c` tinyint(1) NOT NULL,
  `Incidentinvestigation3d` tinyint(1) NOT NULL,
  `Incidentinvestigation3e` tinyint(1) NOT NULL,
  `Incidentinvestigation4a` tinyint(1) NOT NULL,
  `Incidentinvestigation4b` tinyint(1) NOT NULL,
  `Incidentinvestigation4c` tinyint(1) NOT NULL,
  `Incidentinvestigation4d` tinyint(1) NOT NULL,
  `Incidentinvestigation4e` tinyint(1) NOT NULL,
  `Incidentinvestigation4f` tinyint(1) NOT NULL,
  `Incidentinvestigation5` tinyint(1) NOT NULL,
  `Incidentinvestigation6` tinyint(1) NOT NULL,
  `Incidentinvestigation7` tinyint(1) NOT NULL,
  `Incidentinvestigation8` tinyint(1) NOT NULL,
  `Incidentinvestigation9` tinyint(1) NOT NULL,
  `Contractors1a` tinyint(1) NOT NULL,
  `Contractors1b` tinyint(1) NOT NULL,
  `Contractors1c` tinyint(1) NOT NULL,
  `Contractors2a` tinyint(1) NOT NULL,
  `Contractors2b` tinyint(1) NOT NULL,
  `Contractors2c` tinyint(1) NOT NULL,
  `Contractors2d` tinyint(1) NOT NULL,
  `Contractors3` tinyint(1) NOT NULL,
  `Contractors4` tinyint(1) NOT NULL,
  `Contractors5` tinyint(1) NOT NULL,
  `Audits1a` tinyint(1) NOT NULL,
  `Audits1b` tinyint(1) NOT NULL,
  `Audits1c` tinyint(1) NOT NULL,
  `Audits2` tinyint(1) NOT NULL,
  `Audits3a` tinyint(1) NOT NULL,
  `Audits3b` tinyint(1) NOT NULL,
  `Audits4` tinyint(1) NOT NULL,
  `Totalfms` double NOT NULL
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;

-- --------------------------------------------------------

--
-- Table structure for table `tbl_generaldata`
--

CREATE TABLE `tbl_generaldata` (
  `equipment_code` varchar(20) NOT NULL,
  `rbi_cal_date` date DEFAULT current_timestamp(),
  `installation_date` date NOT NULL DEFAULT current_timestamp(),
  `last_inspection_date` date NOT NULL DEFAULT current_timestamp(),
  `material_type` varchar(50) NOT NULL,
  `material` varchar(50) NOT NULL,
  `fluid` varchar(50) NOT NULL,
  `furnished_thickness` double NOT NULL,
  `measured_thickness` double NOT NULL,
  `internal_design_pressure` double NOT NULL,
  `operating_pressure` double NOT NULL,
  `operating_temperature` double NOT NULL,
  `corrosion_allowance` double NOT NULL,
  `final_corrosion_rate` double NOT NULL,
  `weld_joint-efficiency` double NOT NULL,
  `allowable_stree` double NOT NULL,
  `component_diameter` double NOT NULL,
  `screening_1` tinyint(1) NOT NULL,
  `screening_2` tinyint(1) NOT NULL,
  `screening_3` tinyint(1) NOT NULL,
  `screening_4` tinyint(1) NOT NULL,
  `screening_5` tinyint(1) NOT NULL,
  `screening_6` tinyint(1) NOT NULL,
  `screening_7` tinyint(1) NOT NULL,
  `screening_8` tinyint(1) NOT NULL,
  `screening_9` tinyint(1) NOT NULL,
  `screening_10` tinyint(1) NOT NULL,
  `screening_11` tinyint(1) NOT NULL,
  `screening_12` tinyint(1) NOT NULL,
  `screening_13` tinyint(1) NOT NULL,
  `screening_14` tinyint(1) NOT NULL,
  `screening_15` tinyint(1) NOT NULL,
  `screening_16` tinyint(1) NOT NULL,
  `screening_17` tinyint(1) NOT NULL,
  `screening_18` tinyint(1) NOT NULL,
  `screening_19` tinyint(1) NOT NULL,
  `screening_20` tinyint(1) NOT NULL,
  `screening_21` tinyint(1) NOT NULL,
  `screening_22` tinyint(1) NOT NULL,
  `screening_23` tinyint(1) NOT NULL,
  `screening_24` tinyint(1) NOT NULL,
  `screening_25` tinyint(1) NOT NULL
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;

-- --------------------------------------------------------

--
-- Table structure for table `tbl_industry`
--

CREATE TABLE `tbl_industry` (
  `industry_code` varchar(5) NOT NULL,
  `industry_name` varchar(50) NOT NULL
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;

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

-- --------------------------------------------------------

--
-- Table structure for table `tbl_reportmetodegff`
--

CREATE TABLE `tbl_reportmetodegff` (
  `equipment_code` varchar(20) NOT NULL,
  `equipment_description` varchar(100) NOT NULL,
  `gff_total` double NOT NULL,
  `msf` double NOT NULL,
  `thinning` double NOT NULL,
  `ext_damage` double NOT NULL,
  `brittle_fracture` double NOT NULL,
  `stress_corr_crack` double NOT NULL,
  `htha` double NOT NULL,
  `mechanical_fatigue` double NOT NULL,
  `total_df` double NOT NULL,
  `pof` double NOT NULL,
  `pof_cat` varchar(3) NOT NULL,
  `flammable_cd` double NOT NULL,
  `flammable_pi` double NOT NULL,
  `toxic_pi` double NOT NULL,
  `non_pi` double NOT NULL,
  `final_area_cof` double NOT NULL,
  `area_cof_cat` varchar(3) NOT NULL,
  `final_financial_cof` double NOT NULL,
  `financial_cof_cat` varchar(3) NOT NULL,
  `area_risk_cat` varchar(3) NOT NULL,
  `financial_risk_cat` varchar(3) NOT NULL,
  `area_risk_ranking` varchar(20) NOT NULL,
  `financial_risk_ranking` varchar(20) NOT NULL
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

--
-- Indexes for dumped tables
--

--
-- Indexes for table `tbl_equipment`
--
ALTER TABLE `tbl_equipment`
  ADD PRIMARY KEY (`equipment_code`);

--
-- Indexes for table `tbl_fluida`
--
ALTER TABLE `tbl_fluida`
  ADD PRIMARY KEY (`kode_fluida`);

--
-- Indexes for table `tbl_industry`
--
ALTER TABLE `tbl_industry`
  ADD PRIMARY KEY (`industry_code`);

--
-- Indexes for table `tbl_material`
--
ALTER TABLE `tbl_material`
  ADD PRIMARY KEY (`kode_material`);

--
-- Indexes for table `tbl_reportmetodegff`
--
ALTER TABLE `tbl_reportmetodegff`
  ADD PRIMARY KEY (`equipment_code`);
COMMIT;

/*!40101 SET CHARACTER_SET_CLIENT=@OLD_CHARACTER_SET_CLIENT */;
/*!40101 SET CHARACTER_SET_RESULTS=@OLD_CHARACTER_SET_RESULTS */;
/*!40101 SET COLLATION_CONNECTION=@OLD_COLLATION_CONNECTION */;
