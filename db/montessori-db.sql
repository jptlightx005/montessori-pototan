-- phpMyAdmin SQL Dump
-- version 4.5.1
-- http://www.phpmyadmin.net
--
-- Host: 127.0.0.1
-- Generation Time: Nov 14, 2016 at 05:04 PM
-- Server version: 10.1.9-MariaDB
-- PHP Version: 5.6.15

SET SQL_MODE = "NO_AUTO_VALUE_ON_ZERO";
SET time_zone = "+00:00";


/*!40101 SET @OLD_CHARACTER_SET_CLIENT=@@CHARACTER_SET_CLIENT */;
/*!40101 SET @OLD_CHARACTER_SET_RESULTS=@@CHARACTER_SET_RESULTS */;
/*!40101 SET @OLD_COLLATION_CONNECTION=@@COLLATION_CONNECTION */;
/*!40101 SET NAMES utf8mb4 */;

--
-- Database: `montessori-db`
--

-- --------------------------------------------------------

--
-- Table structure for table `montessori_accounts`
--

CREATE TABLE `montessori_accounts` (
  `Student_ID` int(4) UNSIGNED ZEROFILL NOT NULL,
  `school_year` text NOT NULL,
  `total_matriculation` int(11) NOT NULL,
  `total_payment` int(11) NOT NULL,
  `latest_payment` date NOT NULL
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

--
-- Dumping data for table `montessori_accounts`
--

INSERT INTO `montessori_accounts` (`Student_ID`, `school_year`, `total_matriculation`, `total_payment`, `latest_payment`) VALUES
(0001, '2016-2017', 25000, 7000, '2016-11-14');

-- --------------------------------------------------------

--
-- Table structure for table `montessori_admin`
--

CREATE TABLE `montessori_admin` (
  `ID` int(2) UNSIGNED ZEROFILL NOT NULL,
  `usrn` text NOT NULL,
  `pssw` text NOT NULL,
  `role` text NOT NULL,
  `login_count` int(4) NOT NULL DEFAULT '0'
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

--
-- Dumping data for table `montessori_admin`
--

INSERT INTO `montessori_admin` (`ID`, `usrn`, `pssw`, `role`, `login_count`) VALUES
(01, 'pmontessori', 'pmontessori', 'master', 0),
(02, 'admin1', 'admin1pssw', 'admin', 49),
(03, 'admin2', 'admin2pssw', 'admin', 0),
(04, 'admin3', 'admin3pssw', 'admin', 0),
(05, 'registraros', 'regpssw', 'registrar', 395),
(06, 'acctos', 'accpssw', 'accountant', 159);

-- --------------------------------------------------------

--
-- Table structure for table `montessori_queue`
--

CREATE TABLE `montessori_queue` (
  `Student_ID` int(4) UNSIGNED ZEROFILL NOT NULL,
  `usrn` text NOT NULL,
  `rf_ip` text NOT NULL,
  `is_new` int(11) NOT NULL,
  `status` text NOT NULL,
  `date_registered` datetime NOT NULL DEFAULT CURRENT_TIMESTAMP
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

--
-- Dumping data for table `montessori_queue`
--

INSERT INTO `montessori_queue` (`Student_ID`, `usrn`, `rf_ip`, `is_new`, `status`, `date_registered`) VALUES
(0001, 'admin1', '127.0.0.1', 0, 'enrolled', '2016-11-14 23:08:27');

-- --------------------------------------------------------

--
-- Table structure for table `montessori_records`
--

CREATE TABLE `montessori_records` (
  `ID` int(4) UNSIGNED ZEROFILL NOT NULL,
  `current_grade` text NOT NULL,
  `last_name` text NOT NULL,
  `first_name` text NOT NULL,
  `middle_name` text NOT NULL,
  `gender` text NOT NULL,
  `date_of_birth` date NOT NULL,
  `place_of_birth` text NOT NULL,
  `fathers_name` text NOT NULL,
  `father_occupation` text NOT NULL,
  `mothers_name` text NOT NULL,
  `mother_occupation` text NOT NULL,
  `home_address_brgy` text NOT NULL,
  `home_address_city` text NOT NULL,
  `home_address_province` text NOT NULL,
  `home_number` text NOT NULL,
  `guardian_name` text NOT NULL,
  `guardian_relation` text NOT NULL,
  `guardian_address_brgy` text NOT NULL,
  `guardian_address_city` text NOT NULL,
  `guardian_address_province` text NOT NULL,
  `guardian_number` text NOT NULL,
  `last_school_attended` text NOT NULL,
  `religion` text NOT NULL,
  `is_baptized` int(1) NOT NULL,
  `first_communion` int(1) NOT NULL,
  `date_enrolled` datetime NOT NULL DEFAULT CURRENT_TIMESTAMP
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

--
-- Dumping data for table `montessori_records`
--

INSERT INTO `montessori_records` (`ID`, `current_grade`, `last_name`, `first_name`, `middle_name`, `gender`, `date_of_birth`, `place_of_birth`, `fathers_name`, `father_occupation`, `mothers_name`, `mother_occupation`, `home_address_brgy`, `home_address_city`, `home_address_province`, `home_number`, `guardian_name`, `guardian_relation`, `guardian_address_brgy`, `guardian_address_city`, `guardian_address_province`, `guardian_number`, `last_school_attended`, `religion`, `is_baptized`, `first_communion`, `date_enrolled`) VALUES
(0001, 'grade3', 'Tokunaga', 'Hideaki', 'Ito', 'Male', '2009-06-05', '', '', '', '', '', 'Ichiban', 'Daijina', 'Mono Ga', '', '', '', '', '', '', '', '', '', 1, 0, '2016-11-14 23:08:27');

-- --------------------------------------------------------

--
-- Table structure for table `montessori_transactions`
--

CREATE TABLE `montessori_transactions` (
  `ID` int(4) UNSIGNED ZEROFILL NOT NULL,
  `Student_ID` int(4) UNSIGNED ZEROFILL NOT NULL,
  `payment` int(11) NOT NULL,
  `date_of_payment` datetime NOT NULL DEFAULT CURRENT_TIMESTAMP
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

--
-- Dumping data for table `montessori_transactions`
--

INSERT INTO `montessori_transactions` (`ID`, `Student_ID`, `payment`, `date_of_payment`) VALUES
(0001, 0001, 5000, '2016-11-14 23:37:41'),
(0002, 0001, 2000, '2016-11-14 23:39:26');

--
-- Indexes for dumped tables
--

--
-- Indexes for table `montessori_accounts`
--
ALTER TABLE `montessori_accounts`
  ADD PRIMARY KEY (`Student_ID`),
  ADD KEY `Student_ID` (`Student_ID`);

--
-- Indexes for table `montessori_admin`
--
ALTER TABLE `montessori_admin`
  ADD PRIMARY KEY (`ID`);

--
-- Indexes for table `montessori_queue`
--
ALTER TABLE `montessori_queue`
  ADD PRIMARY KEY (`Student_ID`);

--
-- Indexes for table `montessori_records`
--
ALTER TABLE `montessori_records`
  ADD PRIMARY KEY (`ID`);

--
-- Indexes for table `montessori_transactions`
--
ALTER TABLE `montessori_transactions`
  ADD PRIMARY KEY (`ID`),
  ADD KEY `Student_ID` (`Student_ID`);

--
-- AUTO_INCREMENT for dumped tables
--

--
-- AUTO_INCREMENT for table `montessori_admin`
--
ALTER TABLE `montessori_admin`
  MODIFY `ID` int(2) UNSIGNED ZEROFILL NOT NULL AUTO_INCREMENT, AUTO_INCREMENT=7;
--
-- AUTO_INCREMENT for table `montessori_records`
--
ALTER TABLE `montessori_records`
  MODIFY `ID` int(4) UNSIGNED ZEROFILL NOT NULL AUTO_INCREMENT, AUTO_INCREMENT=2;
--
-- AUTO_INCREMENT for table `montessori_transactions`
--
ALTER TABLE `montessori_transactions`
  MODIFY `ID` int(4) UNSIGNED ZEROFILL NOT NULL AUTO_INCREMENT, AUTO_INCREMENT=3;
--
-- Constraints for dumped tables
--

--
-- Constraints for table `montessori_accounts`
--
ALTER TABLE `montessori_accounts`
  ADD CONSTRAINT `montessori_accounts_ibfk_1` FOREIGN KEY (`Student_ID`) REFERENCES `montessori_records` (`ID`) ON DELETE CASCADE;

--
-- Constraints for table `montessori_queue`
--
ALTER TABLE `montessori_queue`
  ADD CONSTRAINT `montessori_queue_ibfk_1` FOREIGN KEY (`Student_ID`) REFERENCES `montessori_records` (`ID`) ON DELETE CASCADE;

--
-- Constraints for table `montessori_transactions`
--
ALTER TABLE `montessori_transactions`
  ADD CONSTRAINT `montessori_transactions_ibfk_1` FOREIGN KEY (`Student_ID`) REFERENCES `montessori_records` (`ID`);

/*!40101 SET CHARACTER_SET_CLIENT=@OLD_CHARACTER_SET_CLIENT */;
/*!40101 SET CHARACTER_SET_RESULTS=@OLD_CHARACTER_SET_RESULTS */;
/*!40101 SET COLLATION_CONNECTION=@OLD_COLLATION_CONNECTION */;