/*
SQLyog Community Edition- MySQL GUI v7.01 
MySQL - 5.6.15 : Database - jobcraftingtool
*********************************************************************
*/

/*!40101 SET NAMES utf8 */;

/*!40101 SET SQL_MODE=''*/;

/*!40014 SET @OLD_FOREIGN_KEY_CHECKS=@@FOREIGN_KEY_CHECKS, FOREIGN_KEY_CHECKS=0 */;
/*!40101 SET @OLD_SQL_MODE=@@SQL_MODE, SQL_MODE='NO_AUTO_VALUE_ON_ZERO' */;

CREATE DATABASE /*!32312 IF NOT EXISTS*/`jobcraftingtool` /*!40100 DEFAULT CHARACTER SET utf8 */;

USE `jobcraftingtool`;

/*Table structure for table `jct_action_plan` */

DROP TABLE IF EXISTS `jct_action_plan`;

CREATE TABLE `jct_action_plan` (
  `jct_status_id` int(11) DEFAULT NULL,
  `jct_action_plan_id` int(11) NOT NULL,
  `jct_jobref_no` varchar(50) DEFAULT NULL,
  `jct_action_plan_flg` varchar(1) DEFAULT NULL,
  `jct_action_plan_desc` varchar(100) DEFAULT NULL,
  `version` int(11) DEFAULT NULL,
  `jct_as_created_ts` timestamp NOT NULL DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
  `jct_as_lastmodified_ts` timestamp NOT NULL DEFAULT '0000-00-00 00:00:00',
  `jct_as_lastmodified_by` varchar(50) DEFAULT NULL,
  `jct_as_created_by` varchar(50) DEFAULT NULL,
  PRIMARY KEY (`jct_action_plan_id`),
  KEY `FK_jct_action_plan` (`jct_status_id`),
  CONSTRAINT `FK_jct_action_plan` FOREIGN KEY (`jct_status_id`) REFERENCES `jct_status` (`jct_status_id`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8;

/*Data for the table `jct_action_plan` */

/*Table structure for table `jct_additional_info` */

DROP TABLE IF EXISTS `jct_additional_info`;

CREATE TABLE `jct_additional_info` (
  `jct_status_id` int(11) DEFAULT NULL,
  `jct_additional_info_id` int(11) NOT NULL,
  `jct_jobref_no` varchar(50) DEFAULT NULL,
  `jct_job_desc` varchar(50) DEFAULT NULL,
  `jct_yr_current_job` varchar(50) DEFAULT NULL,
  `jct_current_org_desc` varchar(50) DEFAULT NULL,
  `jct_yr_current_org` varchar(50) DEFAULT NULL,
  `jct_current_industry` varchar(50) DEFAULT NULL,
  `jct_year_of_birth` varchar(10) DEFAULT NULL,
  `jct_higest_edu_level` varchar(50) DEFAULT NULL,
  `jct_gender` varchar(10) DEFAULT NULL,
  `jct_nationality` varchar(10) DEFAULT NULL,
  `jct_email` varchar(50) DEFAULT NULL,
  `jct_shared_disgram` tinyint(1) DEFAULT NULL,
  `version` int(11) DEFAULT NULL,
  `jct_as_created_ts` timestamp NOT NULL DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
  `jct_as_lastmodified_ts` timestamp NOT NULL DEFAULT '0000-00-00 00:00:00',
  `jct_as_lastmodified_by` varchar(50) DEFAULT NULL,
  `jct_as_created_by` varchar(50) DEFAULT NULL,
  PRIMARY KEY (`jct_additional_info_id`),
  KEY `FK_jct_additional_info` (`jct_status_id`),
  CONSTRAINT `FK_jct_additional_info` FOREIGN KEY (`jct_status_id`) REFERENCES `jct_status` (`jct_status_id`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8;

/*Data for the table `jct_additional_info` */

/*Table structure for table `jct_after_sketch` */

DROP TABLE IF EXISTS `jct_after_sketch`;

CREATE TABLE `jct_after_sketch` (
  `jct_as_id` int(11) NOT NULL,
  `jct_as_jobref_no` varchar(50) DEFAULT NULL,
  `jct_as_role_id` int(11) NOT NULL,
  `jct_as_element_code` varchar(5) DEFAULT NULL,
  `jct_as_element_id` int(11) DEFAULT NULL,
  `jct_as_element_energy` int(11) DEFAULT NULL,
  `jct_as_element_time` int(11) DEFAULT NULL,
  `jct_as_element_desc` varchar(200) DEFAULT NULL,
  `jct_as_role_desc` varchar(200) DEFAULT NULL,
  `jct_as_status_id` int(10) DEFAULT NULL,
  `jct_as_position` varchar(50) DEFAULT NULL,
  `jct_as_json` mediumtext,
  `version` int(11) DEFAULT NULL,
  `jct_as_created_ts` timestamp NOT NULL DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
  `jct_as_lastmodified_ts` timestamp NOT NULL DEFAULT '0000-00-00 00:00:00',
  `jct_as_lastmodified_by` varchar(50) DEFAULT NULL,
  `jct_as_created_by` varchar(50) DEFAULT NULL,
  PRIMARY KEY (`jct_as_id`),
  KEY `FK_jct_after_sketch` (`jct_as_status_id`),
  KEY `FK_jct_after_sketch_job_attribute` (`jct_as_element_id`),
  CONSTRAINT `FK_jct_after_sketch` FOREIGN KEY (`jct_as_status_id`) REFERENCES `jct_status` (`jct_status_id`),
  CONSTRAINT `FK_jct_after_sketch_job_attribute` FOREIGN KEY (`jct_as_element_id`) REFERENCES `jct_job_attribute` (`jct_job_attribute_id`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8;

/*Data for the table `jct_after_sketch` */

/*Table structure for table `jct_before_sketch_details` */

DROP TABLE IF EXISTS `jct_before_sketch_details`;

CREATE TABLE `jct_before_sketch_details` (
  `jct_bs_task_id` int(11) NOT NULL,
  `jct_bs_jobref_no` varchar(50) DEFAULT NULL,
  `jct_bs_energy` int(11) DEFAULT NULL,
  `jct_bs_time` int(11) DEFAULT NULL,
  `jct_bs_task_desc` varchar(200) DEFAULT NULL,
  `jct_bs_status_id` int(10) DEFAULT NULL,
  `jct_bs_position` varchar(50) DEFAULT NULL,
  `version` int(11) DEFAULT NULL,
  `jct_ds_created_ts` timestamp NOT NULL DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
  `jct_bs_lastmodified_ts` timestamp NOT NULL DEFAULT '0000-00-00 00:00:00',
  `jct_bs_lastmodified_by` varchar(50) DEFAULT NULL,
  `jct_bs_created_by` varchar(50) DEFAULT NULL,
  `jct_bs_header_id` int(11) NOT NULL,
  PRIMARY KEY (`jct_bs_task_id`),
  KEY `FK_jct_before_sketch` (`jct_bs_status_id`),
  KEY `FK_jct_before_sketch_details_id` (`jct_bs_header_id`),
  CONSTRAINT `FK_jct_before_sketch_details_id` FOREIGN KEY (`jct_bs_header_id`) REFERENCES `jct_before_sketch_header` (`jct_bs_header_id`),
  CONSTRAINT `FK_jct_before_sketch` FOREIGN KEY (`jct_bs_status_id`) REFERENCES `jct_status` (`jct_status_id`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8;

/*Data for the table `jct_before_sketch_details` */

/*Table structure for table `jct_before_sketch_header` */

DROP TABLE IF EXISTS `jct_before_sketch_header`;

CREATE TABLE `jct_before_sketch_header` (
  `jct_bs_header_id` int(11) NOT NULL,
  `jct_bs_jobref_no` varchar(50) DEFAULT NULL,
  `jct_bs_status_id` int(10) DEFAULT NULL,
  `jct_bs_json` mediumtext,
  `jct_bs_time_spent` int(11) DEFAULT NULL,
  `version` int(11) DEFAULT NULL,
  `jct_ds_created_ts` timestamp NOT NULL DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
  `jct_bs_lastmodified_ts` timestamp NOT NULL DEFAULT '0000-00-00 00:00:00',
  `jct_bs_lastmodified_by` varchar(50) DEFAULT NULL,
  `jct_bs_created_by` varchar(50) DEFAULT NULL,
  PRIMARY KEY (`jct_bs_header_id`),
  KEY `FK_jct_before_sketch` (`jct_bs_status_id`),
  CONSTRAINT `FK_jct_before_sketch_header` FOREIGN KEY (`jct_bs_status_id`) REFERENCES `jct_status` (`jct_status_id`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8;

/*Data for the table `jct_before_sketch_header` */

/*Table structure for table `jct_before_sketch_question` */

DROP TABLE IF EXISTS `jct_before_sketch_question`;

CREATE TABLE `jct_before_sketch_question` (
  `jct_bs_question_id` int(11) NOT NULL,
  `jct_bs_jobref_no` varchar(50) DEFAULT NULL,
  `jct_bs_question_desc` varchar(200) DEFAULT NULL,
  `jct_bs_status_id` int(10) DEFAULT NULL,
  `version` int(11) DEFAULT NULL,
  `jct_ds_created_ts` timestamp NOT NULL DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
  `jct_bs_lastmodified_ts` timestamp NOT NULL DEFAULT '0000-00-00 00:00:00',
  `jct_bs_lastmodified_by` varchar(50) DEFAULT NULL,
  `jct_bs_created_by` varchar(50) DEFAULT NULL,
  PRIMARY KEY (`jct_bs_question_id`),
  KEY `FK_jct_before_sketch_question` (`jct_bs_status_id`),
  CONSTRAINT `FK_jct_before_sketch_question` FOREIGN KEY (`jct_bs_status_id`) REFERENCES `jct_status` (`jct_status_id`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8;

/*Data for the table `jct_before_sketch_question` */

/*Table structure for table `jct_id_gen_table` */

DROP TABLE IF EXISTS `jct_id_gen_table`;

CREATE TABLE `jct_id_gen_table` (
  `jct_id_gen_id` int(11) NOT NULL,
  `jct_id_name` varchar(30) NOT NULL,
  `jct_id_val` int(11) NOT NULL,
  PRIMARY KEY (`jct_id_gen_id`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8;

/*Data for the table `jct_id_gen_table` */

insert  into `jct_id_gen_table`(`jct_id_gen_id`,`jct_id_name`,`jct_id_val`) values (1,'jct_action_plan_id',1),(2,'jct_additional_info_id',1),(3,'jct_after_sketch_id',1),(4,'jct_before_sketch_id',1),(5,'jct_before_sketch_question_id',1),(6,'jct_job_attribute_id',1),(7,'jct_status_id',1),(8,'jct_user_id',49),(9,'jct_user_login_info_id',1);

/*Table structure for table `jct_job_attribute` */

DROP TABLE IF EXISTS `jct_job_attribute`;

CREATE TABLE `jct_job_attribute` (
  `jct_job_attribute_id` int(11) NOT NULL,
  `jct_job_code` varchar(20) DEFAULT NULL,
  `jct_job_attribute_desc` varchar(100) DEFAULT NULL,
  `jct_status_id` int(11) DEFAULT NULL,
  `version` int(11) DEFAULT NULL,
  `jct_as_created_ts` timestamp NOT NULL DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
  `jct_as_lastmodified_ts` timestamp NOT NULL DEFAULT '0000-00-00 00:00:00',
  `jct_as_lastmodified_by` varchar(50) DEFAULT NULL,
  `jct_as_created_by` varchar(50) DEFAULT NULL,
  PRIMARY KEY (`jct_job_attribute_id`),
  KEY `FK_jct_job_attribute` (`jct_status_id`),
  CONSTRAINT `FK_jct_job_attribute` FOREIGN KEY (`jct_status_id`) REFERENCES `jct_status` (`jct_status_id`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8;

/*Data for the table `jct_job_attribute` */

/*Table structure for table `jct_status` */

DROP TABLE IF EXISTS `jct_status`;

CREATE TABLE `jct_status` (
  `jct_status_id` int(11) NOT NULL,
  `jct_status_desc` varchar(50) DEFAULT NULL,
  `version` int(11) DEFAULT NULL,
  `jct_ds_created_ts` timestamp NOT NULL DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
  `jct_bs_lastmodified_ts` timestamp NOT NULL DEFAULT '0000-00-00 00:00:00',
  `jct_bs_lastmodified_by` varchar(50) DEFAULT NULL,
  `jct_bs_created_by` varchar(50) DEFAULT NULL,
  PRIMARY KEY (`jct_status_id`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8;

/*Data for the table `jct_status` */

insert  into `jct_status`(`jct_status_id`,`jct_status_desc`,`version`,`jct_ds_created_ts`,`jct_bs_lastmodified_ts`,`jct_bs_lastmodified_by`,`jct_bs_created_by`) values (1,'USER CREATED',1,'2014-01-10 14:18:30','2014-01-10 11:59:00','ADMIN','ADMIN'),(2,'USER_ACTIVE',1,'2014-01-10 14:18:34','2014-01-10 11:59:00','ADMIN','ADMIN'),(3,'USER_DEACTIVATED',1,'2014-01-10 14:18:40','2014-01-10 11:59:00','ADMIN','ADMIN');

/*Table structure for table `jct_user` */

DROP TABLE IF EXISTS `jct_user`;

CREATE TABLE `jct_user` (
  `jct_user_id` int(11) NOT NULL,
  `jct_first_name` varchar(20) NOT NULL,
  `jct_last_name` varchar(20) NOT NULL,
  `jct_user_name` varchar(20) NOT NULL,
  `jct_password` varchar(40) NOT NULL,
  `jct_user_email` varchar(50) DEFAULT NULL,
  `jct_user_phone` varchar(50) DEFAULT NULL,
  `jct_gender` char(1) DEFAULT NULL,
  `jct_date_of_birth` date DEFAULT NULL,
  `jct_profile_picture` longblob,
  `jct_active_yn` int(1) NOT NULL,
  `jct_version` int(11) DEFAULT NULL,
  `jct_created_ts` timestamp NULL DEFAULT NULL,
  `jct_lastmodified_by` varchar(50) DEFAULT NULL,
  `jct_created_by` varchar(50) DEFAULT NULL,
  `lastmodified_ts` timestamp NULL DEFAULT NULL,
  `jct_base_string` varchar(50) DEFAULT NULL,
  `jct_role_id` int(5) DEFAULT '1',
  PRIMARY KEY (`jct_user_id`),
  KEY `FK_User_Id` (`jct_role_id`),
  CONSTRAINT `FK_User_Id` FOREIGN KEY (`jct_role_id`) REFERENCES `jct_user_role` (`jct_role_id`) ON DELETE CASCADE ON UPDATE CASCADE
) ENGINE=InnoDB DEFAULT CHARSET=utf8;

/*Data for the table `jct_user` */

/*Table structure for table `jct_user_login_info` */

DROP TABLE IF EXISTS `jct_user_login_info`;

CREATE TABLE `jct_user_login_info` (
  `jct_user_login_info_id` int(11) NOT NULL,
  `jct_jobref_no` varchar(50) DEFAULT NULL,
  `jct_user_id` int(11) NOT NULL,
  `jct_start_time` timestamp NOT NULL DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
  `jct_end_time` timestamp NOT NULL DEFAULT '0000-00-00 00:00:00',
  `jct_status_id` int(10) DEFAULT NULL,
  `jct_page_info` varchar(50) DEFAULT NULL,
  `version` int(11) DEFAULT NULL,
  `jct_as_created_ts` timestamp NOT NULL DEFAULT '0000-00-00 00:00:00',
  `jct_as_lastmodified_ts` timestamp NOT NULL DEFAULT '0000-00-00 00:00:00',
  `jct_as_lastmodified_by` varchar(50) DEFAULT NULL,
  `jct_as_created_by` varchar(50) DEFAULT NULL,
  PRIMARY KEY (`jct_user_login_info_id`),
  KEY `FK_jct_user_login_info` (`jct_status_id`),
  CONSTRAINT `FK_jct_user_login_info` FOREIGN KEY (`jct_status_id`) REFERENCES `jct_status` (`jct_status_id`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8;

/*Data for the table `jct_user_login_info` */

/*Table structure for table `jct_user_role` */

DROP TABLE IF EXISTS `jct_user_role`;

CREATE TABLE `jct_user_role` (
  `jct_role_id` int(5) NOT NULL,
  `jct_role_code` varchar(15) NOT NULL,
  `jct_role_desc` varchar(40) DEFAULT NULL,
  PRIMARY KEY (`jct_role_id`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8;

/*Data for the table `jct_user_role` */

insert  into `jct_user_role`(`jct_role_id`,`jct_role_code`,`jct_role_desc`) values (1,'USER',NULL),(2,'ADMIN',NULL);

/*!40101 SET SQL_MODE=@OLD_SQL_MODE */;
/*!40014 SET FOREIGN_KEY_CHECKS=@OLD_FOREIGN_KEY_CHECKS */;
