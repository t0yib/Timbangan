/*
SQLyog Ultimate - MySQL GUI v8.22 
MySQL - 5.1.41 : Database - perpus
*********************************************************************
*/

/*!40101 SET NAMES utf8 */;

/*!40101 SET SQL_MODE=''*/;

/*!40014 SET @OLD_UNIQUE_CHECKS=@@UNIQUE_CHECKS, UNIQUE_CHECKS=0 */;
/*!40014 SET @OLD_FOREIGN_KEY_CHECKS=@@FOREIGN_KEY_CHECKS, FOREIGN_KEY_CHECKS=0 */;
/*!40101 SET @OLD_SQL_MODE=@@SQL_MODE, SQL_MODE='NO_AUTO_VALUE_ON_ZERO' */;
/*!40111 SET @OLD_SQL_NOTES=@@SQL_NOTES, SQL_NOTES=0 */;
CREATE DATABASE /*!32312 IF NOT EXISTS*/`perpus` /*!40100 DEFAULT CHARACTER SET latin1 */;

USE `perpus`;

/*Table structure for table `anggota` */

DROP TABLE IF EXISTS `anggota`;

CREATE TABLE `anggota` (
  `no_anggota` varchar(6) DEFAULT NULL,
  `nis` varchar(30) DEFAULT NULL,
  `nama` varchar(30) DEFAULT NULL,
  `alamat` tinytext,
  `notelpon` varchar(30) DEFAULT NULL,
  `status` enum('A','N') DEFAULT 'A'
) ENGINE=MyISAM DEFAULT CHARSET=latin1;

/*Data for the table `anggota` */

insert  into `anggota`(`no_anggota`,`nis`,`nama`,`alamat`,`notelpon`,`status`) values ('A0002','9951007648','Ajeng Ayu ','Desa Pilang sari Rt08/V Demak','','A'),('A0003','9957977903','Ajib Zamzuri','Desa Sriwulan Rt03/III\r\nDemal','','A'),('A0004','9956934638','Auliya Alfiyani','Desa Sidorawoh Rt3/II','','A'),('A0005','9956934609','Atik Dina ','Pondok Raden Patah W1/28 Demak','','A');

/*!40101 SET SQL_MODE=@OLD_SQL_MODE */;
/*!40014 SET FOREIGN_KEY_CHECKS=@OLD_FOREIGN_KEY_CHECKS */;
/*!40014 SET UNIQUE_CHECKS=@OLD_UNIQUE_CHECKS */;
/*!40111 SET SQL_NOTES=@OLD_SQL_NOTES */;
