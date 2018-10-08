/*
SQLyog Trial v12.2.1 (64 bit)
MySQL - 5.6.28-log : Database - jobcraftingtoolpublicversion
*********************************************************************
*/


/*!40101 SET NAMES utf8 */;

/*!40101 SET SQL_MODE=''*/;

/*!40014 SET @OLD_UNIQUE_CHECKS=@@UNIQUE_CHECKS, UNIQUE_CHECKS=0 */;
/*!40014 SET @OLD_FOREIGN_KEY_CHECKS=@@FOREIGN_KEY_CHECKS, FOREIGN_KEY_CHECKS=0 */;
/*!40101 SET @OLD_SQL_MODE=@@SQL_MODE, SQL_MODE='NO_AUTO_VALUE_ON_ZERO' */;
/*!40111 SET @OLD_SQL_NOTES=@@SQL_NOTES, SQL_NOTES=0 */;

USE `ebdb`;

/*Table structure for table `jct_action_plan` */

DROP TABLE IF EXISTS `jct_action_plan`;

CREATE TABLE `jct_action_plan` (
  `jct_action_plan_id` int(11) NOT NULL AUTO_INCREMENT,
  `jct_jobref_no` varchar(50) DEFAULT NULL,
  `jct_action_plan_flg` varchar(1) DEFAULT NULL,
  `jct_action_plan_desc` varchar(100) DEFAULT NULL,
  `jct_as_answar_desc` varchar(2000) DEFAULT NULL,
  `jct_as_question_desc` varchar(200) DEFAULT NULL,
  `jct_as_question_sub_desc` varchar(200) DEFAULT NULL,
  `jct_status_id` int(11) DEFAULT NULL,
  `version` int(11) DEFAULT NULL,
  `jct_as_created_ts` timestamp DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
  `jct_as_lastmodified_ts` timestamp NULL DEFAULT '0000-00-00 00:00:00',
  `jct_as_lastmodified_by` varchar(50) DEFAULT NULL,
  `jct_as_created_by` varchar(50) DEFAULT NULL,
  `jct_action_plan_soft_delete` int(1) DEFAULT '0',
  `jct_as_header_id` int(11) NOT NULL,
  `jct_action_plan_user_id` int(11) NOT NULL,
  PRIMARY KEY (`jct_action_plan_id`),
  KEY `FK_jct_action_plan` (`jct_status_id`),
  KEY `NewIndex1` (`jct_jobref_no`,`jct_status_id`,`jct_action_plan_soft_delete`),
  KEY `FK_jct_action_plan_user_id` (`jct_action_plan_user_id`),
  CONSTRAINT `FK_jct_action_plan` FOREIGN KEY (`jct_status_id`) REFERENCES `jct_status` (`jct_status_id`),
  CONSTRAINT `FK_jct_action_plan_user_id` FOREIGN KEY (`jct_action_plan_user_id`) REFERENCES `jct_user` (`jct_user_id`)
) ENGINE=InnoDB AUTO_INCREMENT=4816 DEFAULT CHARSET=utf8;

/*Data for the table `jct_action_plan` */

/*Table structure for table `jct_additional_info` */

DROP TABLE IF EXISTS `jct_additional_info`;

CREATE TABLE `jct_additional_info` (
  `jct_status_id` int(11) DEFAULT NULL,
  `jct_additional_info_id` int(11) NOT NULL AUTO_INCREMENT,
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
  `jct_as_created_ts` timestamp DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
  `jct_as_lastmodified_ts` timestamp NULL DEFAULT '0000-00-00 00:00:00',
  `jct_as_lastmodified_by` varchar(50) DEFAULT NULL,
  `jct_as_created_by` varchar(50) DEFAULT NULL,
  PRIMARY KEY (`jct_additional_info_id`),
  KEY `FK_jct_additional_info` (`jct_status_id`),
  CONSTRAINT `FK_jct_additional_info` FOREIGN KEY (`jct_status_id`) REFERENCES `jct_status` (`jct_status_id`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8;

/*Data for the table `jct_additional_info` */

/*Table structure for table `jct_admin_menu` */

DROP TABLE IF EXISTS `jct_admin_menu`;

CREATE TABLE `jct_admin_menu` (
  `jct_menu_id` int(11) NOT NULL AUTO_INCREMENT,
  `jct_main_menu` varchar(100) DEFAULT NULL,
  `jct_sub_menu` varchar(100) DEFAULT NULL,
  `jct_role_id` int(2) DEFAULT NULL,
  `jct_soft_delete` smallint(1) DEFAULT '0',
  `jct_sub_menu_linked_page` varchar(100) DEFAULT NULL,
  `jct_created_by` varchar(100) DEFAULT NULL,
  `jct_created_ts` timestamp NULL DEFAULT '0000-00-00 00:00:00',
  `jct_last_modified_by` varchar(100) DEFAULT NULL,
  `jct_last_modified_ts` timestamp NULL DEFAULT '0000-00-00 00:00:00',
  `jct_main_menu_order` int(4) DEFAULT '1',
  `jct_sub_menu_order` int(4) DEFAULT '1',
  PRIMARY KEY (`jct_menu_id`)
) ENGINE=InnoDB AUTO_INCREMENT=38 DEFAULT CHARSET=utf8;

/*Data for the table `jct_admin_menu` */

insert  into `jct_admin_menu`(`jct_menu_id`,`jct_main_menu`,`jct_sub_menu`,`jct_role_id`,`jct_soft_delete`,`jct_sub_menu_linked_page`,`jct_created_by`,`jct_created_ts`,`jct_last_modified_by`,`jct_last_modified_ts`,`jct_main_menu_order`,`jct_sub_menu_order`) values 

(1,'Manage<br />Users','Create User Profile',2,0,'userProfile.jsp','Admin','2014-09-01 10:00:00','Admin','2014-09-01 10:00:00',1,1),

(2,'Manage<br />Users','Create User Group',2,0,'userGroup.jsp','Admin','2014-09-01 10:00:00','Admin','2014-09-01 10:00:00',1,1),

(3,'Manage<br />Users','Create User List',2,1,'addUser.jsp','Admin','2014-09-01 10:00:00','Admin','2014-09-01 10:00:00',1,1),

(4,'Application<br />Settings','Application Settings',2,0,'appSettings.jsp','Admin','2014-09-01 10:00:00','Admin','2014-09-01 10:00:00',2,1),

(5,'Content<br />Configuration','Reflection Question / Action Plan',2,0,'reflectionQtn.jsp','Admin','2014-09-01 10:00:00','Admin','2014-09-01 10:00:00',3,4),

(6,'Content<br />Configuration','User Role',2,1,'userRole.jsp','Admin','2014-09-01 10:00:00','Admin','2014-09-01 10:00:00',3,1),

(7,'Content<br />Configuration','Mapping Values',2,0,'mappingPage.jsp','Admin','2014-09-01 10:00:00','Admin','2014-09-01 10:00:00',3,5),

(8,'Content<br />Configuration','Region',2,1,'region.jsp','Admin','2014-09-01 10:00:00','Admin','2014-09-01 10:00:00',3,1),

(9,'Content<br />Configuration','Instructions',2,0,'instruction.jsp','Admin','2014-09-01 10:00:00','Admin','2014-09-01 10:00:00',3,2),

(10,'Content<br />Configuration','Configure Survey Questions',2,0,'configureSurveyQuestions.jsp','Admin','2014-09-01 10:00:00','Admin','2014-09-01 10:00:00',1,1),

(11,'Content<br />Configuration','O*Net Integration',2,0,'onetIntegration.jsp','Admin','2014-09-01 10:00:00','Admin','2014-09-01 10:00:00',3,6),

(12,'My<br />Account','My Account',2,0,'myAccount.jsp','Admin','2014-09-01 10:00:00','Admin','2014-09-01 10:00:00',4,1),

(13,'Reports','Before Sketch',2,0,'reportsBeforeSketch.jsp','Admin','2014-09-01 10:00:00','Admin','2014-09-01 10:00:00',5,1),

(14,'Reports','After Diagram',2,0,'reportsAfterSketch.jsp','Admin','2014-09-01 10:00:00','Admin','2014-09-01 10:00:00',5,1),

(15,'Reports','Before Sketch To After Diagram',2,0,'reportsBeforeSketchToAfterSketch.jsp','Admin','2014-09-01 10:00:00','Admin','2014-09-01 10:00:00',5,1),

(16,'Reports','Action Plan',2,0,'reportsActionPlan.jsp','Admin','2014-09-01 10:00:00','Admin','2014-09-01 10:00:00',5,1),

(17,'Reports','Merged Report',2,0,'reportsAll.jsp','Admin','2014-09-01 10:00:00','Admin','2014-09-01 10:00:00',5,1),

(18,'Reports','Misc Report',2,0,'reportsMisc.jsp','Admin','2014-09-01 10:00:00','Admin','2014-09-01 10:00:00',5,1),

(19,'Reports','Group Profile Reports',2,0,'reportsGroupProfile.jsp','Admin','2014-09-01 10:00:00','Admin','2014-09-01 10:00:00',5,1),

(20,'Reports','Facilitator / Individual Report',2,0,'facilitatorIndividualReport.jsp','Admin','2014-09-01 10:00:00','Admin','2014-09-01 10:00:00',5,1),

(21,'Reports','Survey Report',2,0,'surveyReport.jsp','Admin','2014-09-01 10:00:00','Admin','2014-09-01 10:00:00',5,1),

(22,'Refund<br />Request','Refund Request',2,0,'refundRequest.jsp','Admin','2014-09-01 10:00:00','Admin','2014-09-01 10:00:00',6,1),

(23,'Manage<br />Users','Create / Edit User List',3,0,'createUserFacilitator.jsp','Admin','2014-09-01 10:00:00','Admin','2014-09-01 10:00:00',1,2),

(24,'Manage<br />Users','Renew Individual User Subscription',3,1,'renewSubscription.jsp','Admin','2014-09-01 10:00:00','Admin','2014-09-01 10:00:00',1,3),

(25,'Manage<br />Users','Assign New Facilitator',3,1,'changeIndividualUserRole.jsp','Admin','2014-09-01 10:00:00','Admin','2014-09-01 10:00:00',1,4),

(26,'My<br />Account','My Account',3,0,'facilitatorAccount.jsp','Admin','2014-09-01 10:00:00','Admin','2014-09-01 10:00:00',2,1),

(27,'Preview<br />Tool','Preview',3,0,'NA','Admin','2014-09-01 10:00:00','Admin','2014-09-01 10:00:00',3,1),

(28,'Track<br />Users','User Progress Report',3,0,'progressReport.jsp','Admin','2014-09-01 10:00:00','Admin','2014-09-01 10:00:00',4,1),

(29,'Content<br />Configuration','Remove Diagram',2,0,'removeDiagram.jsp','Admin','2014-09-01 10:00:00','Admin','2014-09-01 10:00:00',1,7),

(30,'Content<br />Configuration','Popup Instruction',2,0,'popupInfoInstruction.jsp','Admin','2014-09-01 10:00:00','Admin','2014-09-01 10:00:00',1,3),

(31,'Manage<br />Users','Subscribe / Renew More Users',2,1,'manageFacilitator.jsp','Admin','2015-03-04 11:48:00','Admin','2015-03-04 11:48:00',1,1),

(32,'Manage<br />Users','Check Realization',2,0,'chequeRealization.jsp','Admin','2015-03-04 11:48:00','Admin','2015-03-04 11:48:00',1,1),

(33,'Manage<br />Users','Create User Group',3,1,'facilitatorUserGroup.jsp','Admin','2015-03-04 11:48:00','Admin','2015-03-04 11:48:00',1,1),

(34,'Reports','Payment Report',2,0,'paymentReport.jsp','Admin','2015-03-07 11:48:00','Admin','2015-03-07 11:48:00',1,1),

(35,'Manage<br />Users','Buy More Subscriptions',3,1,'buyMoreSubscription.jsp','Admin','2015-03-07 11:48:00','Admin','2015-03-07 11:48:00',1,5),

(36,'Content<br />Configuration','Terms & Conditions',2,0,'termsAndConditions.jsp','Admin','2015-07-13 11:45:00','Admin','2015-07-13 11:45:00',3,8),

(37,'Manage<br />Users','Add/Renew users',2,0,'addRenewUsers.jsp','Admin','2015-07-21 11:45:00','Admin','2015-07-21 11:45:00',1,1);

/*Table structure for table `jct_after_sketch_finalpage_details` */

DROP TABLE IF EXISTS `jct_after_sketch_finalpage_details`;

CREATE TABLE `jct_after_sketch_finalpage_details` (
  `jct_as_finalpage_details_id` int(11) NOT NULL AUTO_INCREMENT,
  `jct_as_jobref_no` varchar(50) DEFAULT NULL,
  `jct_as_element_code` varchar(5) DEFAULT NULL,
  `jct_as_element_desc` varchar(50) DEFAULT NULL,
  `jct_as_position` varchar(50) DEFAULT NULL,
  `jct_as_role_desc` varchar(130) DEFAULT NULL,
  `jct_as_task_energy` int(11) DEFAULT NULL,
  `jct_as_task_id` int(11) DEFAULT NULL,
  `jct_as_task_desc` varchar(200) DEFAULT NULL,
  `jct_as_status_id` int(10) DEFAULT NULL,
  `jct_as_soft_delete` tinyint(1) NULL DEFAULT '0',
  `version` int(11) DEFAULT NULL,
  `jct_as_created_ts` timestamp NULL DEFAULT '0000-00-00 00:00:00',
  `jct_as_lastmodified_ts` timestamp NULL DEFAULT '0000-00-00 00:00:00',
  `jct_as_lastmodified_by` varchar(50) DEFAULT NULL,
  `jct_as_created_by` varchar(50) DEFAULT NULL,
  `jct_as_header_id` int(11) NOT NULL,
  `jct_as_additional_desc` varchar(200) DEFAULT NULL,
  PRIMARY KEY (`jct_as_finalpage_details_id`),
  KEY `FK_jct_as_finalpage_details_id` (`jct_as_header_id`),
  KEY `FK_jct_after_sketch_job_attribute` (`jct_as_element_desc`),
  KEY `NewIndex1` (`jct_as_jobref_no`,`jct_as_task_id`,`jct_as_soft_delete`),
  CONSTRAINT `FK_jct_as_finalpage_details_id` FOREIGN KEY (`jct_as_header_id`) REFERENCES `jct_after_sketch_header` (`jct_as_header_id`)
) ENGINE=InnoDB AUTO_INCREMENT=1444 DEFAULT CHARSET=utf8;

/*Data for the table `jct_after_sketch_finalpage_details` */

/*Table structure for table `jct_after_sketch_finalpage_details_temp` */

DROP TABLE IF EXISTS `jct_after_sketch_finalpage_details_temp`;

CREATE TABLE `jct_after_sketch_finalpage_details_temp` (
  `jct_as_finalpage_details_id` int(11) NOT NULL AUTO_INCREMENT,
  `jct_as_jobref_no` varchar(50) DEFAULT NULL,
  `jct_as_element_code` varchar(5) DEFAULT NULL,
  `jct_as_element_desc` varchar(50) DEFAULT NULL,
  `jct_as_position` varchar(50) DEFAULT NULL,
  `jct_as_role_desc` varchar(130) DEFAULT NULL,
  `jct_as_task_energy` int(11) DEFAULT NULL,
  `jct_as_task_id` int(11) DEFAULT NULL,
  `jct_as_task_desc` varchar(200) DEFAULT NULL,
  `jct_as_status_id` int(10) DEFAULT NULL,
  `jct_as_soft_delete` tinyint(1) NULL DEFAULT '0',
  `version` int(11) DEFAULT NULL,
  `jct_as_created_ts` timestamp NULL DEFAULT '0000-00-00 00:00:00',
  `jct_as_lastmodified_ts` timestamp NULL DEFAULT '0000-00-00 00:00:00',
  `jct_as_lastmodified_by` varchar(50) DEFAULT NULL,
  `jct_as_created_by` varchar(50) DEFAULT NULL,
  `jct_as_header_id` int(11) NOT NULL,
  `jct_as_additional_desc` varchar(200) DEFAULT NULL,
  PRIMARY KEY (`jct_as_finalpage_details_id`),
  KEY `FK_jct_as_finalpage_details_id_temp` (`jct_as_header_id`),
  KEY `FK_jct_after_sketch_job_attribute_temp` (`jct_as_element_desc`),
  CONSTRAINT `FK_jct_as_finalpage_details_id_temp` FOREIGN KEY (`jct_as_header_id`) REFERENCES `jct_after_sketch_header_temp` (`jct_as_header_id`)
) ENGINE=InnoDB AUTO_INCREMENT=42 DEFAULT CHARSET=utf8;

/*Data for the table `jct_after_sketch_finalpage_details_temp` */

/*Table structure for table `jct_after_sketch_header` */

DROP TABLE IF EXISTS `jct_after_sketch_header`;

CREATE TABLE `jct_after_sketch_header` (
  `jct_as_header_id` int(11) NOT NULL AUTO_INCREMENT,
  `jct_as_jobref_no` varchar(50) DEFAULT NULL,
  `jct_as_status_id` int(10) DEFAULT NULL,
  `jct_as_pagene_json` mediumtext,
  `jct_as_finalpage_json` mediumtext,
  `jct_as_pageone_time_spent` int(11) DEFAULT NULL,
  `jct_as_finalpage_time_spent` int(11) DEFAULT NULL,
  `version` int(11) DEFAULT NULL,
  `jct_as_created_ts` timestamp DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
  `jct_as_lastmodified_ts` timestamp NULL DEFAULT '0000-00-00 00:00:00',
  `jct_as_lastmodified_by` varchar(50) DEFAULT NULL,
  `jct_as_created_by` varchar(50) DEFAULT NULL,
  `jct_as_user_job_title` varchar(30) DEFAULT NULL,
  `jct_as_pageone_snapshot` longblob,
  `jct_as_finalpage_snapshot` longblob,
  `jct_as_finalpage_initial_json_value` mediumtext,
  `jct_as_pageone_checked_value` mediumtext,
  `jct_as_pageone_snapshot_string` longtext,
  `jct_as_finalpage_snapshot_string` longtext,
  `jct_final_page_div_height` varchar(25) DEFAULT NULL,
  `jct_final_page_div_width` varchar(25) DEFAULT NULL,
  `jct_checked_strength` varchar(7) DEFAULT NULL,
  `jct_checked_passion` varchar(7) DEFAULT NULL,
  `jct_checked_value` varchar(7) DEFAULT NULL,
  `jct_as_soft_delete` tinyint(1) DEFAULT '0',
  `jct_as_total_count` int(11) DEFAULT NULL,
  `jct_as_insertion_sequence` int(11) DEFAULT NULL,
  `jct_as_user_id` int(11) NOT NULL,
  PRIMARY KEY (`jct_as_header_id`),
  KEY `FK_jct_after_sketch_header` (`jct_as_status_id`),
  KEY `NewIndex1` (`jct_as_jobref_no`,`jct_as_status_id`,`jct_as_soft_delete`),
  KEY `FK_jct_after_sketch_header_user_id` (`jct_as_user_id`),
  CONSTRAINT `FK_jct_after_sketch_header` FOREIGN KEY (`jct_as_status_id`) REFERENCES `jct_status` (`jct_status_id`),
  CONSTRAINT `FK_jct_after_sketch_header_user_id` FOREIGN KEY (`jct_as_user_id`) REFERENCES `jct_user` (`jct_user_id`)
) ENGINE=InnoDB AUTO_INCREMENT=538 DEFAULT CHARSET=utf8;

/*Data for the table `jct_after_sketch_header` */

/*Table structure for table `jct_after_sketch_header_temp` */

DROP TABLE IF EXISTS `jct_after_sketch_header_temp`;

CREATE TABLE `jct_after_sketch_header_temp` (
  `jct_as_header_id` int(11) NOT NULL AUTO_INCREMENT,
  `jct_as_jobref_no` varchar(50) DEFAULT NULL,
  `jct_as_status_id` int(10) DEFAULT NULL,
  `jct_as_pagene_json` mediumtext,
  `jct_as_finalpage_json` mediumtext,
  `jct_as_pageone_time_spent` int(11) DEFAULT NULL,
  `jct_as_finalpage_time_spent` int(11) DEFAULT NULL,
  `version` int(11) DEFAULT NULL,
  `jct_as_created_ts` timestamp DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
  `jct_as_lastmodified_ts` timestamp NULL DEFAULT '0000-00-00 00:00:00',
  `jct_as_lastmodified_by` varchar(50) DEFAULT NULL,
  `jct_as_created_by` varchar(50) DEFAULT NULL,
  `jct_as_user_job_title` varchar(30) DEFAULT NULL,
  `jct_as_pageone_snapshot` longblob,
  `jct_as_finalpage_snapshot` longblob,
  `jct_as_finalpage_initial_json_value` mediumtext,
  `jct_as_pageone_checked_value` mediumtext,
  `jct_as_pageone_snapshot_string` longtext,
  `jct_as_finalpage_snapshot_string` longtext,
  `jct_final_page_div_height` varchar(25) DEFAULT NULL,
  `jct_final_page_div_width` varchar(25) DEFAULT NULL,
  `jct_checked_strength` varchar(7) DEFAULT NULL,
  `jct_checked_passion` varchar(7) DEFAULT NULL,
  `jct_checked_value` varchar(7) DEFAULT NULL,
  `jct_as_soft_delete` tinyint(1) DEFAULT '0',
  `jct_as_total_count` int(11) DEFAULT NULL,
  `jct_as_insertion_sequence` int(11) DEFAULT NULL,
  `jct_as_temp_user_id` int(11) NOT NULL,
  PRIMARY KEY (`jct_as_header_id`),
  KEY `FK_jct_after_sketch_header_temp` (`jct_as_status_id`),
  KEY `FK_jct_after_sketch_header_temp_user_id` (`jct_as_temp_user_id`),
  CONSTRAINT `FK_jct_after_sketch_header_temp` FOREIGN KEY (`jct_as_status_id`) REFERENCES `jct_status` (`jct_status_id`),
  CONSTRAINT `FK_jct_after_sketch_header_temp_user_id` FOREIGN KEY (`jct_as_temp_user_id`) REFERENCES `jct_user` (`jct_user_id`)
) ENGINE=InnoDB AUTO_INCREMENT=16 DEFAULT CHARSET=utf8;

/*Data for the table `jct_after_sketch_header_temp` */

/*Table structure for table `jct_after_sketch_pageone_details` */

DROP TABLE IF EXISTS `jct_after_sketch_pageone_details`;

CREATE TABLE `jct_after_sketch_pageone_details` (
  `jct_as_pageone_details_id` int(11) NOT NULL AUTO_INCREMENT,
  `jct_as_jobref_no` varchar(50) DEFAULT NULL,
  `jct_as_element_code` varchar(5) DEFAULT NULL,
  `jct_as_element_id` int(11) DEFAULT NULL,
  `jct_as_position` varchar(50) DEFAULT NULL,
  `jct_as_soft_delete` tinyint(1) NULL DEFAULT '0',
  `version` int(11) DEFAULT NULL,
  `jct_as_created_ts` timestamp DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
  `jct_as_lastmodified_ts` timestamp NULL DEFAULT '0000-00-00 00:00:00',
  `jct_as_lastmodified_by` varchar(50) DEFAULT NULL,
  `jct_as_created_by` varchar(50) DEFAULT NULL,
  `jct_as_header_id` int(11) NOT NULL,
  `jct_as_value_count` int(11) DEFAULT NULL,
  `jct_as_passion_count` int(11) DEFAULT NULL,
  `jct_as_strength_count` int(11) DEFAULT NULL,
  PRIMARY KEY (`jct_as_pageone_details_id`),
  KEY `FK_jct_as_pageone_details_id` (`jct_as_header_id`),
  CONSTRAINT `FK_jct_as_pageone_details_id` FOREIGN KEY (`jct_as_header_id`) REFERENCES `jct_after_sketch_header` (`jct_as_header_id`)
) ENGINE=InnoDB AUTO_INCREMENT=2105 DEFAULT CHARSET=utf8;

/*Data for the table `jct_after_sketch_pageone_details` */

/*Table structure for table `jct_after_sketch_pageone_details_temp` */

DROP TABLE IF EXISTS `jct_after_sketch_pageone_details_temp`;

CREATE TABLE `jct_after_sketch_pageone_details_temp` (
  `jct_as_pageone_details_id` int(11) NOT NULL AUTO_INCREMENT,
  `jct_as_jobref_no` varchar(50) DEFAULT NULL,
  `jct_as_element_code` varchar(5) DEFAULT NULL,
  `jct_as_element_id` int(11) DEFAULT NULL,
  `jct_as_position` varchar(50) DEFAULT NULL,
  `jct_as_soft_delete` tinyint(1) NULL DEFAULT '0',
  `version` int(11) DEFAULT NULL,
  `jct_as_created_ts` timestamp DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
  `jct_as_lastmodified_ts` timestamp NULL DEFAULT '0000-00-00 00:00:00',
  `jct_as_lastmodified_by` varchar(50) DEFAULT NULL,
  `jct_as_created_by` varchar(50) DEFAULT NULL,
  `jct_as_header_id` int(11) NOT NULL,
  `jct_as_value_count` int(11) DEFAULT NULL,
  `jct_as_passion_count` int(11) DEFAULT NULL,
  `jct_as_strength_count` int(11) DEFAULT NULL,
  PRIMARY KEY (`jct_as_pageone_details_id`),
  KEY `FK_jct_as_pageone_details_id_temp` (`jct_as_header_id`),
  CONSTRAINT `FK_jct_as_pageone_details_id_temp` FOREIGN KEY (`jct_as_header_id`) REFERENCES `jct_after_sketch_header_temp` (`jct_as_header_id`)
) ENGINE=InnoDB AUTO_INCREMENT=43 DEFAULT CHARSET=utf8;

/*Data for the table `jct_after_sketch_pageone_details_temp` */

/*Table structure for table `jct_app_settings_master` */

DROP TABLE IF EXISTS `jct_app_settings_master`;

CREATE TABLE `jct_app_settings_master` (
  `jct_app_settings_id` int(11) NOT NULL AUTO_INCREMENT,
  `jct_app_header_color` varchar(10) DEFAULT NULL,
  `jct_app_footer_color` varchar(10) DEFAULT NULL,
  `jct_app_sub_header_color` varchar(10) DEFAULT NULL,
  `jct_app_ins_panel_color` varchar(10) DEFAULT NULL,
  `jct_app_ins_panel_txt_color` varchar(10) DEFAULT NULL,
  `jct_app_soft_delete` smallint(1) DEFAULT '0',
  `jct_app_last_updated_by` varchar(50) DEFAULT NULL,
  `jct_app_last_updated_ts` timestamp NULL DEFAULT '0000-00-00 00:00:00',
  `jct_app_image` longblob,
  PRIMARY KEY (`jct_app_settings_id`),
  KEY `NewIndex1` (`jct_app_soft_delete`)
) ENGINE=InnoDB AUTO_INCREMENT=37 DEFAULT CHARSET=utf8;

/*Data for the table `jct_app_settings_master` */

insert  into `jct_app_settings_master`(`jct_app_settings_id`,`jct_app_header_color`,`jct_app_footer_color`,`jct_app_sub_header_color`,`jct_app_ins_panel_color`,`jct_app_ins_panel_txt_color`,`jct_app_soft_delete`,`jct_app_last_updated_by`,`jct_app_last_updated_ts`,`jct_app_image`) values 

(1,'#006990','#006990','#88cbdf','#810123','#fff',0,'admin','2014-08-06 11:32:00',NULL);

/*Table structure for table `jct_before_sketch_details` */

DROP TABLE IF EXISTS `jct_before_sketch_details`;

CREATE TABLE `jct_before_sketch_details` (
  `jct_bs_detail_id` int(11) NOT NULL AUTO_INCREMENT,
  `jct_bs_task_id` int(11) DEFAULT NULL,
  `jct_bs_jobref_no` varchar(50) DEFAULT NULL,
  `jct_bs_energy` int(11) DEFAULT NULL,
  `jct_bs_time` int(11) DEFAULT NULL,
  `jct_bs_task_desc` varchar(200) DEFAULT NULL,
  `jct_bs_status_id` int(10) DEFAULT NULL,
  `jct_bs_position` varchar(50) DEFAULT NULL,
  `jct_bs_soft_delete` tinyint(1) NULL DEFAULT '0',
  `version` int(11) DEFAULT NULL,
  `jct_ds_created_ts` timestamp DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
  `jct_bs_lastmodified_ts` timestamp NULL DEFAULT '0000-00-00 00:00:00',
  `jct_bs_lastmodified_by` varchar(50) DEFAULT NULL,
  `jct_bs_created_by` varchar(50) DEFAULT NULL,
  `jct_bs_header_id` int(11) NOT NULL,
  PRIMARY KEY (`jct_bs_detail_id`),
  KEY `FK_jct_before_sketch` (`jct_bs_status_id`),
  KEY `FK_jct_before_sketch_details_id` (`jct_bs_header_id`),
  KEY `NewIndex1` (`jct_bs_task_id`,`jct_bs_jobref_no`,`jct_bs_soft_delete`),
  CONSTRAINT `FK_jct_before_sketch` FOREIGN KEY (`jct_bs_status_id`) REFERENCES `jct_status` (`jct_status_id`),
  CONSTRAINT `FK_jct_before_sketch_details_id` FOREIGN KEY (`jct_bs_header_id`) REFERENCES `jct_before_sketch_header` (`jct_bs_header_id`)
) ENGINE=InnoDB AUTO_INCREMENT=445 DEFAULT CHARSET=utf8;

/*Data for the table `jct_before_sketch_details` */

/*Table structure for table `jct_before_sketch_details_temp` */

DROP TABLE IF EXISTS `jct_before_sketch_details_temp`;

CREATE TABLE `jct_before_sketch_details_temp` (
  `jct_bs_detail_id` int(11) NOT NULL AUTO_INCREMENT,
  `jct_bs_task_id` int(11) DEFAULT NULL,
  `jct_bs_jobref_no` varchar(50) DEFAULT NULL,
  `jct_bs_energy` int(11) DEFAULT NULL,
  `jct_bs_time` int(11) DEFAULT NULL,
  `jct_bs_task_desc` varchar(200) DEFAULT NULL,
  `jct_bs_status_id` int(10) DEFAULT NULL,
  `jct_bs_position` varchar(50) DEFAULT NULL,
  `jct_bs_soft_delete` tinyint(1) NULL DEFAULT '0',
  `version` int(11) DEFAULT NULL,
  `jct_ds_created_ts` timestamp DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
  `jct_bs_lastmodified_ts` timestamp NULL DEFAULT '0000-00-00 00:00:00',
  `jct_bs_lastmodified_by` varchar(50) DEFAULT NULL,
  `jct_bs_created_by` varchar(50) DEFAULT NULL,
  `jct_bs_header_id` int(11) NOT NULL,
  PRIMARY KEY (`jct_bs_detail_id`),
  KEY `FK_jct_before_sketch_temp` (`jct_bs_status_id`),
  KEY `FK_jct_before_sketch_details_id_temp` (`jct_bs_header_id`),
  CONSTRAINT `FK_jct_before_sketch_details_id_temp` FOREIGN KEY (`jct_bs_header_id`) REFERENCES `jct_before_sketch_header_temp` (`jct_bs_header_id`),
  CONSTRAINT `FK_jct_before_sketch_temp` FOREIGN KEY (`jct_bs_status_id`) REFERENCES `jct_status` (`jct_status_id`)
) ENGINE=InnoDB AUTO_INCREMENT=20 DEFAULT CHARSET=utf8;

/*Data for the table `jct_before_sketch_details_temp` */

/*Table structure for table `jct_before_sketch_header` */

DROP TABLE IF EXISTS `jct_before_sketch_header`;

CREATE TABLE `jct_before_sketch_header` (
  `jct_bs_header_id` int(11) NOT NULL AUTO_INCREMENT,
  `jct_bs_jobref_no` varchar(50) DEFAULT NULL,
  `jct_bs_status_id` int(10) DEFAULT NULL,
  `jct_bs_json` mediumtext,
  `jct_bs_time_spent` int(11) DEFAULT '0',
  `version` int(11) DEFAULT NULL,
  `jct_ds_created_ts` timestamp DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
  `jct_bs_lastmodified_ts` timestamp NULL DEFAULT '0000-00-00 00:00:00',
  `jct_bs_lastmodified_by` varchar(50) DEFAULT NULL,
  `jct_bs_created_by` varchar(50) DEFAULT NULL,
  `jct_bs_user_job_title` varchar(30) DEFAULT NULL,
  `jct_bs_snapshot` longblob NOT NULL,
  `jct_bs_initial_json_value` mediumtext NOT NULL,
  `jct_bs_snapshot_string` longtext NOT NULL,
  `jct_bs_soft_delete` tinyint(1) DEFAULT '0',
  `jct_bs_insertion_sequence` int(11) DEFAULT NULL,
  `jct_bs_user_id` int(11) NOT NULL,
  PRIMARY KEY (`jct_bs_header_id`),
  KEY `FK_jct_before_sketch` (`jct_bs_status_id`),
  KEY `NewIndex1` (`jct_bs_jobref_no`,`jct_bs_status_id`,`jct_bs_soft_delete`),
  KEY `FK_jct_before_sketch_header_user_id` (`jct_bs_user_id`),
  CONSTRAINT `FK_jct_before_sketch_header` FOREIGN KEY (`jct_bs_status_id`) REFERENCES `jct_status` (`jct_status_id`),
  CONSTRAINT `FK_jct_before_sketch_header_user_id` FOREIGN KEY (`jct_bs_user_id`) REFERENCES `jct_user` (`jct_user_id`)
) ENGINE=InnoDB AUTO_INCREMENT=284 DEFAULT CHARSET=utf8;

/*Data for the table `jct_before_sketch_header` */

/*Table structure for table `jct_before_sketch_header_temp` */

DROP TABLE IF EXISTS `jct_before_sketch_header_temp`;

CREATE TABLE `jct_before_sketch_header_temp` (
  `jct_bs_header_id` int(11) NOT NULL AUTO_INCREMENT,
  `jct_bs_jobref_no` varchar(50) DEFAULT NULL,
  `jct_bs_status_id` int(10) DEFAULT NULL,
  `jct_bs_json` mediumtext,
  `jct_bs_time_spent` int(11) DEFAULT NULL,
  `version` int(11) DEFAULT NULL,
  `jct_ds_created_ts` timestamp DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
  `jct_bs_lastmodified_ts` timestamp NULL DEFAULT '0000-00-00 00:00:00',
  `jct_bs_lastmodified_by` varchar(50) DEFAULT NULL,
  `jct_bs_created_by` varchar(50) DEFAULT NULL,
  `jct_bs_user_job_title` varchar(30) DEFAULT NULL,
  `jct_bs_snapshot` longblob NOT NULL,
  `jct_bs_initial_json_value` mediumtext NOT NULL,
  `jct_bs_snapshot_string` longtext NOT NULL,
  `jct_bs_soft_delete` tinyint(1) DEFAULT '0',
  `jct_bs_insertion_sequence` int(11) DEFAULT NULL,
  `jct_bs_temp_user_id` int(11) NOT NULL,
  PRIMARY KEY (`jct_bs_header_id`),
  KEY `FK_jct_before_sketch` (`jct_bs_status_id`),
  KEY `FK_jct_before_sketch_header_temp_user_id` (`jct_bs_temp_user_id`),
  CONSTRAINT `FK_jct_before_sketch_header_temp` FOREIGN KEY (`jct_bs_status_id`) REFERENCES `jct_status` (`jct_status_id`),
  CONSTRAINT `FK_jct_before_sketch_header_temp_user_id` FOREIGN KEY (`jct_bs_temp_user_id`) REFERENCES `jct_user` (`jct_user_id`)
) ENGINE=InnoDB AUTO_INCREMENT=19 DEFAULT CHARSET=utf8;

/*Data for the table `jct_before_sketch_header_temp` */

/*Table structure for table `jct_before_sketch_question` */

DROP TABLE IF EXISTS `jct_before_sketch_question`;

CREATE TABLE `jct_before_sketch_question` (
  `jct_bs_answar_desc` varchar(2000) DEFAULT NULL,
  `jct_bs_question_id` int(11) NOT NULL AUTO_INCREMENT,
  `jct_bs_jobref_no` varchar(50) DEFAULT NULL,
  `jct_bs_question_desc` varchar(500) DEFAULT NULL,
  `jct_bs_status_id` int(10) DEFAULT NULL,
  `version` int(11) DEFAULT NULL,
  `jct_ds_created_ts` timestamp NULL DEFAULT '0000-00-00 00:00:00',
  `jct_bs_lastmodified_ts` timestamp NULL DEFAULT '0000-00-00 00:00:00',
  `jct_bs_lastmodified_by` varchar(50) DEFAULT NULL,
  `jct_bs_created_by` varchar(50) DEFAULT NULL,
  `jct_bs_soft_delete` tinyint(1) NULL DEFAULT '0',
  `jct_bs_header_id` int(11) NOT NULL,
  `jct_bs_sub_question` varchar(2000) DEFAULT NULL,
  `jct_bs_user_id` int(11) NOT NULL,
  PRIMARY KEY (`jct_bs_question_id`),
  KEY `FK_jct_before_sketch_question` (`jct_bs_status_id`),
  KEY `NewIndex1` (`jct_bs_jobref_no`,`jct_bs_status_id`,`jct_bs_soft_delete`),
  KEY `FK_jct_before_sketch_question_user_id` (`jct_bs_user_id`),
  CONSTRAINT `FK_jct_before_sketch_question` FOREIGN KEY (`jct_bs_status_id`) REFERENCES `jct_status` (`jct_status_id`),
  CONSTRAINT `FK_jct_before_sketch_question_user_id` FOREIGN KEY (`jct_bs_user_id`) REFERENCES `jct_user` (`jct_user_id`)
) ENGINE=InnoDB AUTO_INCREMENT=678 DEFAULT CHARSET=utf8;

/*Data for the table `jct_before_sketch_question` */

/*Table structure for table `jct_bs_to_as` */

DROP TABLE IF EXISTS `jct_bs_to_as`;

CREATE TABLE `jct_bs_to_as` (
  `jct_bs_to_as_id` int(11) NOT NULL AUTO_INCREMENT,
  `jct_jobref_no` varchar(50) DEFAULT NULL,
  `jct_task_id` int(11) DEFAULT NULL,
  `jct_bs_energy` int(11) DEFAULT NULL,
  `jct_bs_time` int(11) DEFAULT NULL,
  `jct_bs_task_desc` varchar(200) DEFAULT NULL,
  `jct_as_energy` int(11) DEFAULT NULL,
  `jct_as_time` int(11) DEFAULT NULL,
  `jct_as_task_desc` varchar(200) DEFAULT NULL,
  `jct_diff_status` varchar(20) DEFAULT NULL,
  `jct_diff_energy` varchar(11) DEFAULT NULL,
  `jct_diff_time` varchar(11) DEFAULT NULL,
  `jct_diff_task_desc` varchar(200) DEFAULT NULL,
  `version` int(11) DEFAULT NULL,
  `jct_bs_to_as_created_ts` timestamp DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
  `jct_bs_to_as_lastmodified_ts` timestamp NULL DEFAULT '0000-00-00 00:00:00',
  `jct_bs_to_as_lastmodified_by` varchar(50) DEFAULT NULL,
  `jct_bs_to_as_created_by` varchar(50) DEFAULT NULL,
  `jct_bs_to_as_task_desc_diff` int(1) DEFAULT NULL,
  `jct_soft_delete` int(1) DEFAULT '0',
  `jct_bs_to_as_user_id` int(11) DEFAULT NULL,
  PRIMARY KEY (`jct_bs_to_as_id`)
) ENGINE=InnoDB AUTO_INCREMENT=694 DEFAULT CHARSET=utf8;

/*Data for the table `jct_bs_to_as` */

/*Table structure for table `jct_check_payment_user_details` */

DROP TABLE IF EXISTS `jct_check_payment_user_details`;

CREATE TABLE `jct_check_payment_user_details` (
  `jct_check_payment_user_id` int(11) NOT NULL AUTO_INCREMENT,
  `jct_check_payment_user_name` varchar(50) DEFAULT NULL,
  `jct_check_payment_customer_id` varchar(200) DEFAULT NULL,
  `jct_check_payment_role_id` int(5) DEFAULT NULL,
  `jct_check_payment_check_no` varchar(100) DEFAULT NULL,
  `jct_check_payment_check_date` timestamp DEFAULT CURRENT_TIMESTAMP,
  `jct_check_payment_header_id` int(11) NOT NULL,
  `jct_check_payment_details_id` int(11) NOT NULL,
  `jct_check_payment_is_honored` tinyint(1) DEFAULT '0',
  `jct_check_payment_created_by` varchar(50) DEFAULT NULL,
  `jct_check_payment_created_ts` timestamp NULL DEFAULT '0000-00-00 00:00:00',
  `jct_check_payment_modified_by` varchar(50) DEFAULT NULL,
  `jct_check_payment_modified_ts` timestamp NULL DEFAULT '0000-00-00 00:00:00',
  `jct_check_payment_user_type` varchar(50) DEFAULT 'NEW_USER',
  `jct_check_payment_soft_delete` tinyint(5) DEFAULT '0',
  `jct_pmt_dtls_pmt_trans_nos` varchar(100) DEFAULT NULL,
  `jct_user_id` int(11) DEFAULT NULL,
  `jct_user_expiry_duration` int(11) DEFAULT '0',
  PRIMARY KEY (`jct_check_payment_user_id`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8;

/*Data for the table `jct_check_payment_user_details` */

/*Table structure for table `jct_completion_status` */

DROP TABLE IF EXISTS `jct_completion_status`;

CREATE TABLE `jct_completion_status` (
  `jct_completion_id` int(11) NOT NULL AUTO_INCREMENT,
  `jct_job_reference_no` varchar(50) NOT NULL,
  `jct_start_date` timestamp DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
  `jct_end_date` timestamp NULL DEFAULT '0000-00-00 00:00:00',
  `jct_completion_counter` int(11) DEFAULT NULL,
  `jct_option_selected` int(1) DEFAULT '0',
  PRIMARY KEY (`jct_completion_id`)
) ENGINE=InnoDB AUTO_INCREMENT=51 DEFAULT CHARSET=utf8;

/*Data for the table `jct_completion_status` */

/*Table structure for table `jct_email_details` */

DROP TABLE IF EXISTS `jct_email_details`;

CREATE TABLE `jct_email_details` (
  `jct_registration_id` int(11) NOT NULL,
  `jct_registered_mail` varchar(50) NOT NULL,
  `jct_registered_password` varchar(50) NOT NULL,
  `jct_mail_dispatched` int(1) NULL DEFAULT '0',
  `jct_account_created_date` timestamp NULL DEFAULT '0000-00-00 00:00:00',
  `jct_account_created_by` varchar(50) NOT NULL,
  `jct_mail_sent_date` timestamp NULL DEFAULT '0000-00-00 00:00:00',
  `jct_user_group_desc` varchar(50) DEFAULT NULL,
  `jct_user_profile_desc` varchar(50) DEFAULT NULL,
  PRIMARY KEY (`jct_registration_id`),
  KEY `NewIndex1` (`jct_mail_dispatched`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8;

/*Data for the table `jct_email_details` */

/*Table structure for table `jct_facilitator_details` */

DROP TABLE IF EXISTS `jct_facilitator_details`;

CREATE TABLE `jct_facilitator_details` (
  `jct_fac_mstr_id` int(11) NOT NULL AUTO_INCREMENT,
  `jct_fac_user_id` int(11) NOT NULL,
  `jct_fac_user_name` varchar(100) NOT NULL,
  `jct_fac_total_limit` int(11) NOT NULL,
  `jct_fac_subscribe_limit` int(11) NULL DEFAULT '0',
  `jct_fac_type` varchar(10) NOT NULL,
  `jct_fac_subscription_dt` timestamp NULL DEFAULT '0000-00-00 00:00:00',
  `jct_fac_created_by` varchar(100) NOT NULL,
  `jct_fac_created_ts` timestamp NULL DEFAULT '0000-00-00 00:00:00',
  `jct_pmt_hdr_id` int(11) NOT NULL,
  `jct_user_customer_id` varchar(200) NOT NULL,
  PRIMARY KEY (`jct_fac_mstr_id`),
  KEY `FK_jct_facilitator_details` (`jct_pmt_hdr_id`),
  CONSTRAINT `FK_jct_facilitator_details` FOREIGN KEY (`jct_pmt_hdr_id`) REFERENCES `jct_payment_header` (`jct_pmt_hdr_id`)
) ENGINE=InnoDB AUTO_INCREMENT=38 DEFAULT CHARSET=utf8;

/*Data for the table `jct_facilitator_details` */

/*Table structure for table `jct_functions` */

DROP TABLE IF EXISTS `jct_functions`;

CREATE TABLE `jct_functions` (
  `jct_functions_id` int(11) NOT NULL,
  `jct_functions_desc` varchar(50) DEFAULT NULL,
  `jct_functions_soft_delete` tinyint(1) NULL DEFAULT '0',
  `version` int(11) DEFAULT NULL,
  `jct_ds_created_ts` timestamp DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
  `jct_bs_lastmodified_ts` timestamp NULL DEFAULT '0000-00-00 00:00:00',
  `jct_bs_lastmodified_by` varchar(50) DEFAULT NULL,
  `jct_bs_created_by` varchar(50) DEFAULT NULL,
  `jct_function_name` varchar(50) DEFAULT NULL,
  `jct_function_order` int(11) DEFAULT NULL,
  PRIMARY KEY (`jct_functions_id`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8;

/*Data for the table `jct_functions` */

/*Table structure for table `jct_global_profile_history` */

DROP TABLE IF EXISTS `jct_global_profile_history`;

CREATE TABLE `jct_global_profile_history` (
  `jct_global_profile_id` int(11) NOT NULL AUTO_INCREMENT,
  `jct_global_profile_bs_sub_qtn_original` varchar(500) DEFAULT NULL,
  `jct_global_profile_bs_sub_qtn_changed` varchar(500) DEFAULT NULL,
  `jct_global_profile_ap_sub_qtn_original` varchar(500) DEFAULT NULL,
  `jct_global_profile_ap_sub_qtn_changed` varchar(500) DEFAULT NULL,
  `jct_global_profile_str_original` varchar(500) DEFAULT NULL,
  `jct_global_profile_str_changed` varchar(500) DEFAULT NULL,
  `jct_global_profile_str_desc_original` varchar(1000) DEFAULT NULL,
  `jct_global_profile_str_desc_changed` varchar(1000) DEFAULT NULL,
  `jct_global_profile_val_original` varchar(500) DEFAULT NULL,
  `jct_global_profile_val_changed` varchar(500) DEFAULT NULL,
  `jct_global_profile_val_desc_original` varchar(1000) DEFAULT NULL,
  `jct_global_profile_val_desc_changed` varchar(1000) DEFAULT NULL,
  `jct_global_profile_pas_original` varchar(500) DEFAULT NULL,
  `jct_global_profile_pas_changed` varchar(500) DEFAULT NULL,
  `jct_global_profile_pas_desc_original` varchar(1000) DEFAULT NULL,
  `jct_global_profile_pas_desc_changed` varchar(1000) DEFAULT NULL,
  `jct_global_profile_soft_delete` tinyint(1) DEFAULT '0',
  `jct_global_profile_created_timestamp` timestamp DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
  `jct_global_profile_bs_main_qtn` varchar(1000) DEFAULT NULL,
  `jct_global_profile_ap_main_qtn` varchar(1000) DEFAULT NULL,
  `jct_global_profile_is_attribute` varchar(4) NULL DEFAULT 'N',
  `jct_global_profile_attribute_type` varchar(10) DEFAULT NULL,
  PRIMARY KEY (`jct_global_profile_id`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8;

/*Data for the table `jct_global_profile_history` */

/*Table structure for table `jct_id_gen_table` */

DROP TABLE IF EXISTS `jct_id_gen_table`;

CREATE TABLE `jct_id_gen_table` (
  `jct_id_gen_id` int(11) NOT NULL,
  `jct_id_name` varchar(50) NOT NULL,
  `jct_id_val` int(11) NOT NULL,
  PRIMARY KEY (`jct_id_gen_id`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8;

/*Data for the table `jct_id_gen_table` */

insert  into `jct_id_gen_table`(`jct_id_gen_id`,`jct_id_name`,`jct_id_val`) values 

(2,'jct_additional_info_id',300),

(3,'jct_after_sketch_id',300),

(4,'jct_before_sketch_details_id',300),

(5,'jct_before_sketch_question_id',300),

(7,'jct_status_id',300),

(8,'jct_user_id',454),

(9,'jct_user_login_info_id',300),

(10,'jct_job_reference_no',366),

(11,'jct_before_sketch_header_id',300),

(12,'jct_after_sketch_header_id',300),

(13,'jct_after_sketch_page_one_id',300),

(14,'jct_after_sketch_final_page_id',300),

(15,'jct_bs_to_as',300),

(16,'jct_action_plan',300),

(17,'jct_pdf_records',300),

(18,'jct_user_profile',300),

(19,'jct_user_group',320),

(20,'jct_email_dtls_id',300),

(21,'jct_user_info_id',300),

(22,'jct_instruction',323),

(23,'jct_questions',308),

(24,'jct_region',300),

(25,'jct_job_attribute',314),

(26,'jct_function_group',300),

(27,'jct_job_level',300),

(28,'jct_before_sketch_header_temp_id',300),

(29,'jct_before_sketch_details_temp_id',300),

(30,'jct_completion_status_id',300),

(31,'jct_after_sketch_header_temp_id',300),

(32,'jct_after_sketch_page_one_temp_id',300),

(33,'jct_after_sketch_final_page_temp_id',300),

(34,'jct_status_search_id',300),

(35,'jct_testimonial_id',300),

(36,'jct_user_facilitator_id',3),

(37,'jct_user_customer_id',20),

(38,'jct_ins_video_file_id',306),

(39,'jct_payment_transaction_id',392);

/*Table structure for table `jct_instruction_bar` */

DROP TABLE IF EXISTS `jct_instruction_bar`;

CREATE TABLE `jct_instruction_bar` (
  `jct_instruction_bar_id` int(11) NOT NULL,
  `jct_page_details` varchar(100) DEFAULT NULL,
  `jct_instruction_bar_desc` mediumtext,
  `version` int(11) DEFAULT NULL,
  `jct_instruction_bar_created_ts` timestamp DEFAULT CURRENT_TIMESTAMP,
  `jct_instruction_bar_lastmodified_ts` timestamp NULL DEFAULT '0000-00-00 00:00:00',
  `jct_instruction_bar_lastmodified_by` varchar(50) DEFAULT NULL,
  `jct_instruction_bar_created_by` varchar(50) DEFAULT NULL,
  `jct_instruction_bar_soft_delete` tinyint(1) DEFAULT '0',
  `jct_profile_id` int(11) DEFAULT NULL,
  `jct_profile_desc` varchar(50) DEFAULT NULL,
  `jct_instruction_type` varchar(200) DEFAULT NULL,
  `jct_video_path` varchar(200) DEFAULT NULL,
  PRIMARY KEY (`jct_instruction_bar_id`),
  KEY `FK_jct_instruction_bar` (`jct_profile_id`),
  CONSTRAINT `FK_jct_instruction_bar` FOREIGN KEY (`jct_profile_id`) REFERENCES `jct_user_profile` (`jct_user_profile`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8;

/*Data for the table `jct_instruction_bar` */

/*Table structure for table `jct_job_attribute` */

DROP TABLE IF EXISTS `jct_job_attribute`;

CREATE TABLE `jct_job_attribute` (
  `jct_job_attribute_id` int(11) NOT NULL,
  `jct_job_attribute_code` varchar(20) DEFAULT NULL,
  `jct_job_attribute_name` varchar(50) DEFAULT NULL,
  `version` int(11) DEFAULT NULL,
  `jct_job_attribute_created_ts` timestamp DEFAULT CURRENT_TIMESTAMP,
  `jct_job_attribute_lastmodified_ts` timestamp NULL DEFAULT '0000-00-00 00:00:00',
  `jct_job_attribute_lastmodified_by` varchar(50) DEFAULT NULL,
  `jct_job_attribute_created_by` varchar(50) DEFAULT NULL,
  `jct_job_attribute_soft_delete` tinyint(1) DEFAULT '0',
  `jct_profile_id` int(11) DEFAULT NULL,
  `jct_profile_desc` varchar(50) DEFAULT NULL,
  `jct_job_attribute_desc` varchar(100) DEFAULT NULL,
  `jct_job_attribute_order` int(11) DEFAULT NULL,
  PRIMARY KEY (`jct_job_attribute_id`),
  KEY `FK_jct_user_profile` (`jct_profile_id`),
  CONSTRAINT `FK_jct_job_attribute` FOREIGN KEY (`jct_profile_id`) REFERENCES `jct_user_profile` (`jct_user_profile`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8;

/*Data for the table `jct_job_attribute` */

/*Table structure for table `jct_job_attribute_frozen` */

DROP TABLE IF EXISTS `jct_job_attribute_frozen`;

CREATE TABLE `jct_job_attribute_frozen` (
  `jct_job_attribute_id` int(11) NOT NULL AUTO_INCREMENT,
  `jct_job_attribute_code` varchar(50) NOT NULL,
  `jct_job_attribute_name` varchar(100) NOT NULL,
  `jct_job_attribute_profile_id` int(11) NOT NULL,
  `jct_job_attribute_profile_desc` varchar(100) DEFAULT NULL,
  `jct_job_attribute_desc` varchar(200) DEFAULT NULL,
  `jct_job_attribute_order` int(11) DEFAULT NULL,
  `jct_job_attribute_soft_delete` tinyint(1) NULL DEFAULT '0',
  `jct_job_attribute_created_by` varchar(50) DEFAULT NULL,
  `jct_job_attribute_created_ts` timestamp NULL DEFAULT '0000-00-00 00:00:00',
  `jct_job_attribute_modified_by` varchar(50) DEFAULT NULL,
  `jct_job_attribute_modified_ts` timestamp NULL DEFAULT '0000-00-00 00:00:00',
  `jct_job_attribute_created_for` varchar(100) NOT NULL,
  `jct_job_attribute_user_id` int(11) NOT NULL,
  PRIMARY KEY (`jct_job_attribute_id`)
) ENGINE=InnoDB AUTO_INCREMENT=2087 DEFAULT CHARSET=utf8;

/*Data for the table `jct_job_attribute_frozen` */

/*Table structure for table `jct_levels` */

DROP TABLE IF EXISTS `jct_levels`;

CREATE TABLE `jct_levels` (
  `jct_levels_id` int(11) NOT NULL,
  `jct_levels_desc` varchar(50) DEFAULT NULL,
  `jct_levels_soft_delete` tinyint(1) NULL DEFAULT '0',
  `version` int(11) DEFAULT NULL,
  `jct_ds_created_ts` timestamp DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
  `jct_bs_lastmodified_ts` timestamp NULL DEFAULT '0000-00-00 00:00:00',
  `jct_bs_lastmodified_by` varchar(50) DEFAULT NULL,
  `jct_bs_created_by` varchar(50) DEFAULT NULL,
  `jct_level_name` varchar(50) DEFAULT NULL,
  `jct_level_order` int(11) DEFAULT NULL,
  PRIMARY KEY (`jct_levels_id`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8;

/*Data for the table `jct_levels` */

/*Table structure for table `jct_onet_occupation_list` */

DROP TABLE IF EXISTS `jct_onet_occupation_list`;

CREATE TABLE `jct_onet_occupation_list` (
  `jct_onet_occupation_id` int(11) NOT NULL AUTO_INCREMENT,
  `jct_onet_occupation_code` varchar(50) DEFAULT NULL,
  `jct_onet_occupation_title` varchar(800) DEFAULT NULL,
  `jct_onet_occupation_soft_delete` tinyint(1) DEFAULT '0',
  `jct_onet_occupation_created_ts` timestamp DEFAULT CURRENT_TIMESTAMP,
  `jct_onet_occupation_created_by` varchar(50) DEFAULT NULL,
  `jct_onet_occupation_description` mediumtext,
  PRIMARY KEY (`jct_onet_occupation_id`)
) ENGINE=InnoDB AUTO_INCREMENT=1111 DEFAULT CHARSET=utf8;

/*Data for the table `jct_onet_occupation_list` */

/*Table structure for table `jct_payment_details` */

DROP TABLE IF EXISTS `jct_payment_details`;

CREATE TABLE `jct_payment_details` (
  `jct_pmt_dtls_id` int(11) NOT NULL AUTO_INCREMENT,
  `jct_pmt_dtls_date` timestamp DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
  `jct_pmt_dtls_pmt_typ_id` int(11) DEFAULT NULL,
  `jct_pmt_dtls_pmt_typ_desc` varchar(100) DEFAULT NULL,
  `jct_pmt_dtls_pmt_trans_nos` varchar(100) DEFAULT NULL,
  `jct_pmt_dtls_cheque_nos` varchar(100) DEFAULT NULL,
  `jct_pmt_dtls_created_by` varchar(100) DEFAULT NULL,
  `jct_pmt_dtls_created_ts` timestamp NULL DEFAULT '0000-00-00 00:00:00',
  `jct_pmt_dtls_modified_by` varchar(100) DEFAULT NULL,
  `jct_pmt_dtls_modified_ts` timestamp NULL DEFAULT '0000-00-00 00:00:00',
  `jct_pmt_dtls_soft_delete` int(1) NULL DEFAULT '0',
  `jct_pmt_dtls_bank_name` varchar(100) DEFAULT NULL,
  `jct_pmt_hdr_id` int(11) NOT NULL,
  PRIMARY KEY (`jct_pmt_dtls_id`)
) ENGINE=InnoDB AUTO_INCREMENT=59 DEFAULT CHARSET=utf8;

/*Data for the table `jct_payment_details` */

/*Table structure for table `jct_payment_header` */

DROP TABLE IF EXISTS `jct_payment_header`;

CREATE TABLE `jct_payment_header` (
  `jct_pmt_hdr_id` int(11) NOT NULL AUTO_INCREMENT,
  `jct_pmt_hdr_email_id` varchar(100) NOT NULL,
  `jct_pmt_hdr_user_id` int(11) NOT NULL,
  `jct_pmt_hdr_total_amt` double NULL DEFAULT '0',
  `jct_pmt_hdr_created_by` varchar(100) DEFAULT 'Admin',
  `jct_pmt_hdr_created_ts` timestamp NULL DEFAULT NULL,
  `jct_pmt_hdr_modified_by` varchar(100) DEFAULT 'Admin',
  `jct_pmt_hdr_modified_ts` timestamp NULL DEFAULT NULL,
  `jct_pmt_hdr_soft_delete` int(1) DEFAULT '0',
  `jct_user_customer_id` varchar(200) NOT NULL,
  PRIMARY KEY (`jct_pmt_hdr_id`)
) ENGINE=InnoDB AUTO_INCREMENT=339 DEFAULT CHARSET=utf8;

/*Data for the table `jct_payment_header` */

/*Table structure for table `jct_pdf_records` */

DROP TABLE IF EXISTS `jct_pdf_records`;

CREATE TABLE `jct_pdf_records` (
  `jct_file_id` int(11) NOT NULL AUTO_INCREMENT,
  `jct_file_location` varchar(50) NOT NULL,
  `jct_job_reference_no` varchar(50) NOT NULL,
  `jct_file_name` varchar(50) NOT NULL,
  `jct_created_timestamp` timestamp NULL DEFAULT '0000-00-00 00:00:00',
  `jct_created_by` varchar(50) DEFAULT NULL,
  `jct_user_id` int(11) NOT NULL,
  `jct_show_pdf` tinyint(1) DEFAULT '0',
  PRIMARY KEY (`jct_file_id`),
  KEY `FK_jct_pdf_records` (`jct_user_id`),
  CONSTRAINT `FK_jct_pdf_records` FOREIGN KEY (`jct_user_id`) REFERENCES `jct_user` (`jct_user_id`)
) ENGINE=InnoDB AUTO_INCREMENT=93 DEFAULT CHARSET=utf8;

/*Data for the table `jct_pdf_records` */

/*Table structure for table `jct_popup_instruction` */

DROP TABLE IF EXISTS `jct_popup_instruction`;

CREATE TABLE `jct_popup_instruction` (
  `jct_popup_instruction_id` int(11) NOT NULL AUTO_INCREMENT,
  `jct_popup_instruction_type` varchar(20) NOT NULL,
  `jct_popup_instruction_page` varchar(2) NOT NULL,
  `jct_popup_instruction_text_before_video` mediumtext,
  `jct_popup_instruction_soft_delete` tinyint(1) DEFAULT '0',
  `jct_popup_instruction_creation_date` timestamp NULL DEFAULT '0000-00-00 00:00:00',
  `jct_popup_instruction_last_updation_date` timestamp NULL DEFAULT '0000-00-00 00:00:00',
  `jct_popup_instruction_created_by` varchar(200) DEFAULT NULL,
  `jct_popup_instruction_updated_by` varchar(200) DEFAULT NULL,
  `jct_popup_instruction_video_name` varchar(200) DEFAULT NULL,
  `jct_popup_instruction_profile_id` int(11) NOT NULL,
  `jct_popup_instruction_text_after_video` mediumtext,
  PRIMARY KEY (`jct_popup_instruction_id`)
) ENGINE=InnoDB AUTO_INCREMENT=7 DEFAULT CHARSET=utf8;

/*Data for the table `jct_popup_instruction` */

/*Table structure for table `jct_questions` */

DROP TABLE IF EXISTS `jct_questions`;

CREATE TABLE `jct_questions` (
  `jct_questions_id` int(11) NOT NULL,
  `jct_questions_desc` varchar(200) DEFAULT NULL,
  `jct_profile_id` int(11) NOT NULL,
  `jct_profile_desc` varchar(50) NOT NULL,
  `jct_questions_soft_delete` tinyint(1) NULL DEFAULT '0',
  `version` int(11) DEFAULT NULL,
  `jct_ds_created_ts` timestamp DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
  `jct_bs_lastmodified_ts` timestamp NULL DEFAULT '0000-00-00 00:00:00',
  `jct_bs_lastmodified_by` varchar(50) DEFAULT NULL,
  `jct_bs_created_by` varchar(50) DEFAULT NULL,
  `jct_question_sub_desc` varchar(200) DEFAULT NULL,
  `jct_question_bsas` varchar(2) DEFAULT NULL,
  `jct_question_order` int(11) DEFAULT NULL,
  `jct_sub_question_order` int(11) DEFAULT NULL,
  PRIMARY KEY (`jct_questions_id`),
  KEY `FK_jct_questions` (`jct_profile_id`),
  KEY `NewIndex1` (`jct_profile_id`,`jct_questions_soft_delete`,`jct_question_bsas`),
  CONSTRAINT `FK_jct_questions` FOREIGN KEY (`jct_profile_id`) REFERENCES `jct_user_profile` (`jct_user_profile`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8;

/*Data for the table `jct_questions` */

/*Table structure for table `jct_region` */

DROP TABLE IF EXISTS `jct_region`;

CREATE TABLE `jct_region` (
  `jct_region_id` int(11) NOT NULL,
  `jct_region_desc` varchar(50) DEFAULT NULL,
  `jct_region_soft_delete` tinyint(1) NULL DEFAULT '0',
  `version` int(11) DEFAULT NULL,
  `jct_ds_created_ts` timestamp DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
  `jct_bs_lastmodified_ts` timestamp NULL DEFAULT '0000-00-00 00:00:00',
  `jct_bs_lastmodified_by` varchar(50) DEFAULT NULL,
  `jct_bs_created_by` varchar(50) DEFAULT NULL,
  `jct_region_name` varchar(50) DEFAULT NULL,
  `jct_region_order` int(11) DEFAULT NULL,
  PRIMARY KEY (`jct_region_id`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8;

/*Data for the table `jct_region` */

/*Table structure for table `jct_status` */

DROP TABLE IF EXISTS `jct_status`;

CREATE TABLE `jct_status` (
  `jct_status_id` int(11) NOT NULL,
  `jct_status_desc` varchar(50) DEFAULT NULL,
  `version` int(11) DEFAULT NULL,
  `jct_ds_created_ts` timestamp DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
  `jct_bs_lastmodified_ts` timestamp NULL DEFAULT '0000-00-00 00:00:00',
  `jct_bs_lastmodified_by` varchar(50) DEFAULT NULL,
  `jct_bs_created_by` varchar(50) DEFAULT NULL,
  PRIMARY KEY (`jct_status_id`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8;

/*Data for the table `jct_status` */

insert  into `jct_status`(`jct_status_id`,`jct_status_desc`,`version`,`jct_ds_created_ts`,`jct_bs_lastmodified_ts`,`jct_bs_lastmodified_by`,`jct_bs_created_by`) values 

(1,'USER CREATED',1,'2014-01-10 14:18:30','2014-01-10 11:59:00','ADMIN','ADMIN'),

(2,'USER_ACTIVE',1,'2014-01-10 14:18:34','2014-01-10 11:59:00','ADMIN','ADMIN'),

(3,'USER_DEACTIVATED',1,'2014-01-10 14:18:40','2014-01-10 11:59:00','ADMIN','ADMIN'),

(4,'TASK_PENDING',1,'2014-01-23 10:54:40','2014-01-10 11:59:00',NULL,'ADMIN'),

(5,'TASK_COMPLETED',1,'2014-01-23 10:54:46','2014-01-10 11:59:00',NULL,'ADMIN');

/*Table structure for table `jct_status_search` */

DROP TABLE IF EXISTS `jct_status_search`;

CREATE TABLE `jct_status_search` (
  `jct_status_search_id` int(11) NOT NULL AUTO_INCREMENT,
  `jct_jobref_no` varchar(50) DEFAULT NULL,
  `jct_status_search_flg` varchar(2) DEFAULT NULL,
  `jct_as_snapshot` longblob,
  `jct_bs_snapshot` longblob,
  `jct_bs_snapshot_string` longtext,
  `jct_as_snapshot_string` longtext,
  `jct_status_id` int(11) DEFAULT NULL,
  `jct_soft_delete` tinyint(1) NULL DEFAULT '0',
  `version` int(11) DEFAULT NULL,
  `jct_created_ts` timestamp NULL DEFAULT '0000-00-00 00:00:00',
  `jct_lastmodified_ts` timestamp NULL DEFAULT '0000-00-00 00:00:00',
  `jct_lastmodified_by` varchar(50) DEFAULT NULL,
  `jct_created_by` varchar(50) DEFAULT NULL,
  `jct_user_levels` varchar(50) DEFAULT NULL,
  `jct_is_inactive` int(1) NULL DEFAULT '0',
  `jct_status_decision` tinyint(1) DEFAULT '0',
  `occupation_code` varchar(100) DEFAULT NULL,
  PRIMARY KEY (`jct_status_search_id`),
  KEY `FK_jct_status_search` (`jct_status_id`),
  CONSTRAINT `FK_jct_status_search` FOREIGN KEY (`jct_status_id`) REFERENCES `jct_status` (`jct_status_id`)
) ENGINE=InnoDB AUTO_INCREMENT=42 DEFAULT CHARSET=utf8;

/*Data for the table `jct_status_search` */

/*Table structure for table `jct_survey_question_master` */

DROP TABLE IF EXISTS `jct_survey_question_master`;

CREATE TABLE `jct_survey_question_master` (
  `jct_survey_qtn_master_id` int(11) NOT NULL AUTO_INCREMENT,
  `jct_survey_qtn_master_user_type` int(2) NOT NULL,
  `jct_survey_qtn_master_ans_type` int(2) NOT NULL,
  `jct_survey_qtn_master_main_qtn` varchar(2000) NOT NULL,
  `jct_survey_qtn_master_sub_qtn` varchar(2000) DEFAULT NULL,
  `jct_survey_qtn_master_soft_del` tinyint(11) NULL DEFAULT '0',
  `jct_survey_qtn_master_created_by` varchar(50) DEFAULT NULL,
  `jct_survey_qtn_master_created_ts` timestamp NULL DEFAULT '0000-00-00 00:00:00',
  `jct_survey_qtn_master_modified_by` varchar(50) DEFAULT NULL,
  `jct_survey_qtn_master_modified_ts` timestamp NULL DEFAULT '0000-00-00 00:00:00',
  `jct_survey_qtn_master_mandatory` varchar(2) NOT NULL,
  `jct_survey_qtn_master_profile_id` int(11) DEFAULT NULL,
  `jct_survey_qtn_master_profile_name` varchar(500) DEFAULT NULL,
  `jct_survey_qtn_master_order` tinyint(11) DEFAULT '0',
  PRIMARY KEY (`jct_survey_qtn_master_id`)
) ENGINE=InnoDB AUTO_INCREMENT=7 DEFAULT CHARSET=utf8;

/*Data for the table `jct_survey_question_master` */

/*Table structure for table `jct_survey_questions` */

DROP TABLE IF EXISTS `jct_survey_questions`;

CREATE TABLE `jct_survey_questions` (
  `jct_survey_question_id` int(11) NOT NULL AUTO_INCREMENT,
  `jct_survey_user_type` int(11) NOT NULL,
  `jct_survey_taken_by` varchar(200) NOT NULL,
  `jct_survey_taken_by_user_id` int(11) NOT NULL,
  `jct_survey_answer_type` int(11) NOT NULL,
  `jct_survey_main_qtn` varchar(2000) NOT NULL,
  `jct_survey_sub_qtn` varchar(2000) NOT NULL,
  `jct_survey_answer` varchar(2000) NOT NULL,
  `jct_survey_created_ts` timestamp DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
  `jct_survey_soft_delete` int(11) NULL DEFAULT '0',
  PRIMARY KEY (`jct_survey_question_id`)
) ENGINE=InnoDB AUTO_INCREMENT=247 DEFAULT CHARSET=utf8;

/*Data for the table `jct_survey_questions` */

/*Table structure for table `jct_terms_and_condition` */

DROP TABLE IF EXISTS `jct_terms_and_condition`;

CREATE TABLE `jct_terms_and_condition` (
  `jct_tc_id` int(11) NOT NULL AUTO_INCREMENT,
  `jct_tc_description` varchar(20000) DEFAULT NULL,
  `jct_tc_created_by` varchar(255) DEFAULT NULL,
  `jct_tc_created_on` timestamp DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
  `jct_tc_modified_by` varchar(255) DEFAULT NULL,
  `jct_tc_modified_on` timestamp NULL DEFAULT '0000-00-00 00:00:00',
  `jct_user_profile` int(11) NOT NULL,
  `jct_tc_soft_delete` tinyint(1) DEFAULT '0',
  `jct_tc_user_type` tinyint(1) DEFAULT NULL,
  PRIMARY KEY (`jct_tc_id`)
) ENGINE=InnoDB AUTO_INCREMENT=6 DEFAULT CHARSET=utf8;

/*Data for the table `jct_terms_and_condition` */

/*Table structure for table `jct_user` */

DROP TABLE IF EXISTS `jct_user`;

CREATE TABLE `jct_user` (
  `jct_user_id` int(11) NOT NULL,
  `jct_user_name` varchar(50) DEFAULT NULL,
  `jct_password` varchar(40) NOT NULL,
  `jct_user_email` varchar(50) NOT NULL,
  `jct_active_yn` int(1) NOT NULL,
  `jct_version` int(11) DEFAULT NULL,
  `jct_created_ts` timestamp NULL DEFAULT NULL,
  `jct_lastmodified_by` varchar(50) DEFAULT NULL,
  `jct_created_by` varchar(50) DEFAULT NULL,
  `lastmodified_ts` timestamp NULL DEFAULT NULL,
  `jct_role_id` int(5) DEFAULT '1',
  `jct_user_soft_delete` tinyint(1) DEFAULT '0',
  `jct_account_expiration_date` timestamp NULL DEFAULT NULL,
  `jct_user_customer_id` varchar(200) DEFAULT NULL,
  `jct_user_disp_identifier` int(11) NULL DEFAULT '0',
  `jct_user_disabled` tinyint(1) DEFAULT '0',
  PRIMARY KEY (`jct_user_id`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8;

/*Data for the table `jct_user` */

insert  into `jct_user`(`jct_user_id`,`jct_user_name`,`jct_password`,`jct_user_email`,`jct_active_yn`,`jct_version`,`jct_created_ts`,`jct_lastmodified_by`,`jct_created_by`,`lastmodified_ts`,`jct_role_id`,`jct_user_soft_delete`,`jct_account_expiration_date`,`jct_user_customer_id`,`jct_user_disp_identifier`,`jct_user_disabled`) values 

(1,'Admin','y?5:2pb`ady?5:2pb`ad','jctadmin@interrait.com',2,1,'2014-10-10 00:00:00','Self','Self','2014-10-10 00:00:00',2,0,'2014-12-10 00:00:00','ADMINCUST001',0,0),

(2,'Ecommerce','NOTREQUIRED','NOTREQUIRED',2,1,'2014-10-10 00:00:00','Self','Self','2014-10-10 00:00:00',4,0,'2014-12-10 00:00:00','ECOMMERCE001',0,0);

/*Table structure for table `jct_user_backup` */

DROP TABLE IF EXISTS `jct_user_backup`;

CREATE TABLE `jct_user_backup` (
  `jct_user_id` int(11) NOT NULL,
  `jct_first_name` varchar(20) DEFAULT NULL,
  `jct_last_name` varchar(20) DEFAULT NULL,
  `jct_user_name` varchar(50) DEFAULT NULL,
  `jct_password` varchar(40) NOT NULL,
  `jct_user_email` varchar(50) NOT NULL,
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
  `jct_base_string` mediumtext,
  `jct_role_id` int(5) DEFAULT '1',
  `jct_user_region` varchar(80) DEFAULT NULL,
  `jct_user_levels` varchar(50) DEFAULT NULL,
  `jct_user_tenure` varchar(20) DEFAULT NULL,
  `jct_user_supervise_people` char(1) DEFAULT NULL,
  `jct_user_function_group` varchar(50) DEFAULT NULL,
  `jct_user_group_id` int(11) DEFAULT NULL,
  `jct_user_group_name` varchar(50) DEFAULT NULL,
  `jct_profile_id` int(11) DEFAULT NULL,
  `jct_profile_name` varchar(50) DEFAULT NULL,
  `jct_user_soft_delete` tinyint(1) DEFAULT '0',
  `jct_account_expiration_date` timestamp NULL DEFAULT NULL,
  `jct_facilitator_id` int(11) DEFAULT '0',
  `jct_admin_id` int(11) DEFAULT '0',
  PRIMARY KEY (`jct_user_id`),
  KEY `FK_jct_user_profile` (`jct_profile_id`),
  KEY `FK_jct_group` (`jct_user_group_id`),
  CONSTRAINT `FK_jct_group` FOREIGN KEY (`jct_user_group_id`) REFERENCES `jct_user_group` (`jct_user_group`),
  CONSTRAINT `FK_jct_user_profile` FOREIGN KEY (`jct_profile_id`) REFERENCES `jct_user_profile` (`jct_user_profile`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8;

/*Data for the table `jct_user_backup` */

/*Table structure for table `jct_user_details` */

DROP TABLE IF EXISTS `jct_user_details`;

CREATE TABLE `jct_user_details` (
  `jct_user_details_id` int(11) NOT NULL AUTO_INCREMENT,
  `jct_user_details_first_name` varchar(50) DEFAULT NULL,
  `jct_user_details_last_name` varchar(50) DEFAULT NULL,
  `jct_user_details_phone` varchar(50) DEFAULT NULL,
  `jct_user_details_gender` char(1) DEFAULT NULL,
  `jct_user_details_date_of_birth` date DEFAULT NULL,
  `jct_user_details_profile_picture` longblob,
  `jct_user_details_created_ts` timestamp NULL DEFAULT NULL,
  `jct_user_details_lastmodified_by` varchar(50) DEFAULT NULL,
  `jct_user_details_created_by` varchar(50) DEFAULT NULL,
  `jct_user_details_lastmodified_ts` timestamp NULL DEFAULT NULL,
  `jct_user_details_base_string` mediumtext,
  `jct_user_details_region` varchar(80) DEFAULT NULL,
  `jct_user_details_levels` varchar(50) DEFAULT 'N/A',
  `jct_user_details_tenure` varchar(20) DEFAULT NULL,
  `jct_user_details_supervise_people` char(1) DEFAULT NULL,
  `jct_user_details_function_group` varchar(50) DEFAULT 'N/A',
  `jct_user_details_group_id` int(11) DEFAULT NULL,
  `jct_user_details_group_name` varchar(50) DEFAULT NULL,
  `jct_user_details_profile_id` int(11) DEFAULT NULL,
  `jct_user_details_profile_name` varchar(50) DEFAULT NULL,
  `jct_user_details_facilitator_id` int(11) DEFAULT '0',
  `jct_user_details_admin_id` int(11) DEFAULT '0',
  `jct_user_id` int(11) DEFAULT '0',
  `jct_user_details_soft_delete` int(11) NULL DEFAULT '0',
  `jct_user_onet_occupation` varchar(100) NULL DEFAULT 'NOT-ACTIVATED',
  `jct_nos_of_years` varchar(15) DEFAULT '0 year',
  PRIMARY KEY (`jct_user_details_id`),
  KEY `FK_jct_user_details` (`jct_user_id`),
  KEY `FK_jct_user_details_profile` (`jct_user_details_profile_id`),
  KEY `FK_jct_user_details_group` (`jct_user_details_group_id`),
  CONSTRAINT `FK_jct_user_details` FOREIGN KEY (`jct_user_id`) REFERENCES `jct_user` (`jct_user_id`),
  CONSTRAINT `FK_jct_user_details_group` FOREIGN KEY (`jct_user_details_group_id`) REFERENCES `jct_user_group` (`jct_user_group`),
  CONSTRAINT `FK_jct_user_details_profile` FOREIGN KEY (`jct_user_details_profile_id`) REFERENCES `jct_user_profile` (`jct_user_profile`)
) ENGINE=InnoDB AUTO_INCREMENT=969 DEFAULT CHARSET=utf8;

/*Data for the table `jct_user_details` */

insert  into `jct_user_details`(`jct_user_details_id`,`jct_user_details_first_name`,`jct_user_details_last_name`,`jct_user_details_phone`,`jct_user_details_gender`,`jct_user_details_date_of_birth`,`jct_user_details_profile_picture`,`jct_user_details_created_ts`,`jct_user_details_lastmodified_by`,`jct_user_details_created_by`,`jct_user_details_lastmodified_ts`,`jct_user_details_base_string`,`jct_user_details_region`,`jct_user_details_levels`,`jct_user_details_tenure`,`jct_user_details_supervise_people`,`jct_user_details_function_group`,`jct_user_details_group_id`,`jct_user_details_group_name`,`jct_user_details_profile_id`,`jct_user_details_profile_name`,`jct_user_details_facilitator_id`,`jct_user_details_admin_id`,`jct_user_id`,`jct_user_details_soft_delete`,`jct_user_onet_occupation`,`jct_nos_of_years`) values 

(1,'Admin','Admin',NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,'',NULL,NULL,'',1,NULL,1,NULL,0,1,1,0,'NOT-ACTIVATED','0 year'),

(2,'ECOMMERCE','ECOMMERCE',NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,'N/A',NULL,NULL,'N/A',1,NULL,1,NULL,0,2,2,0,'NOT-ACTIVATED','0 year');

/*Table structure for table `jct_user_group` */

DROP TABLE IF EXISTS `jct_user_group`;

CREATE TABLE `jct_user_group` (
  `jct_user_group` int(11) NOT NULL,
  `jct_soft_delete` tinyint(1) NULL DEFAULT '0',
  `version` int(11) DEFAULT NULL,
  `jct_created_ts` timestamp NULL DEFAULT '0000-00-00 00:00:00',
  `jct_lastmodified_ts` timestamp NULL DEFAULT '0000-00-00 00:00:00',
  `jct_lastmodified_by` varchar(50) DEFAULT NULL,
  `jct_created_by` varchar(50) DEFAULT NULL,
  `jct_user_profile` int(11) NOT NULL,
  `jct_user_profile_desc` varchar(50) DEFAULT NULL,
  `jct_active_status` int(11) NULL DEFAULT '1',
  `jct_user_group_desc` varchar(50) DEFAULT NULL,
  `jct_user_customer_id` varchar(200) DEFAULT NULL,
  `jct_user_role_id` int(5) DEFAULT NULL,
  PRIMARY KEY (`jct_user_group`),
  KEY `FK_jct_user_group` (`jct_user_profile`),
  KEY `NewIndex1` (`jct_user_role_id`),
  CONSTRAINT `FK_jct_user_group` FOREIGN KEY (`jct_user_profile`) REFERENCES `jct_user_profile` (`jct_user_profile`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8;

/*Data for the table `jct_user_group` */

insert  into `jct_user_group`(`jct_user_group`,`jct_soft_delete`,`version`,`jct_created_ts`,`jct_lastmodified_ts`,`jct_lastmodified_by`,`jct_created_by`,`jct_user_profile`,`jct_user_profile_desc`,`jct_active_status`,`jct_user_group_desc`,`jct_user_customer_id`,`jct_user_role_id`) values 

(1,0,4,'2014-05-01 14:57:36','2014-06-26 18:08:42',NULL,NULL,2,'Default Profile',1,'Default User Group',NULL,2),

(2,0,4,'2014-08-06 14:57:36','2014-08-06 18:08:42',NULL,NULL,2,'Default Profile',1,'None','ALL',3);

/*Table structure for table `jct_user_info` */

DROP TABLE IF EXISTS `jct_user_info`;

CREATE TABLE `jct_user_info` (
  `jct_user_info_id` int(11) NOT NULL AUTO_INCREMENT,
  `jct_user_id` int(11) DEFAULT NULL,
  `jct_user_email` varchar(50) DEFAULT NULL,
  `jct_created_ts` timestamp NULL DEFAULT NULL,
  `jct_no_of_count` int(11) DEFAULT '1',
  `jct_soft_delete` int(1) DEFAULT '0',
  PRIMARY KEY (`jct_user_info_id`)
) ENGINE=InnoDB AUTO_INCREMENT=19 DEFAULT CHARSET=utf8;

/*Data for the table `jct_user_info` */

/*Table structure for table `jct_user_login_info` */

DROP TABLE IF EXISTS `jct_user_login_info`;

CREATE TABLE `jct_user_login_info` (
  `jct_user_login_info_id` int(11) NOT NULL AUTO_INCREMENT,
  `jct_jobref_no` varchar(50) DEFAULT NULL,
  `jct_user_id` int(11) NOT NULL,
  `jct_start_time` timestamp NULL DEFAULT '0000-00-00 00:00:00',
  `jct_end_time` timestamp NULL DEFAULT '0000-00-00 00:00:00',
  `jct_status_id` int(10) DEFAULT NULL,
  `jct_page_info` varchar(50) DEFAULT NULL,
  `version` int(11) DEFAULT NULL,
  `jct_as_created_ts` timestamp NULL DEFAULT '0000-00-00 00:00:00',
  `jct_as_lastmodified_ts` timestamp NULL DEFAULT '0000-00-00 00:00:00',
  `jct_as_lastmodified_by` varchar(50) DEFAULT NULL,
  `jct_as_created_by` varchar(50) DEFAULT NULL,
  `jct_completed` int(11) NULL DEFAULT '0',
  `jct_next_page` varchar(10) DEFAULT NULL,
  PRIMARY KEY (`jct_user_login_info_id`),
  KEY `FK_jct_user_login_info` (`jct_status_id`),
  CONSTRAINT `FK_jct_user_login_info` FOREIGN KEY (`jct_status_id`) REFERENCES `jct_status` (`jct_status_id`)
) ENGINE=InnoDB AUTO_INCREMENT=186 DEFAULT CHARSET=utf8;

/*Data for the table `jct_user_login_info` */

/*Table structure for table `jct_user_profile` */

DROP TABLE IF EXISTS `jct_user_profile`;

CREATE TABLE `jct_user_profile` (
  `jct_user_profile` int(11) NOT NULL,
  `jct_user_profile_desc` varchar(50) DEFAULT NULL,
  `jct_soft_delete` tinyint(1) NULL DEFAULT '0',
  `version` int(11) DEFAULT NULL,
  `jct_created_ts` timestamp NULL DEFAULT '0000-00-00 00:00:00',
  `jct_lastmodified_ts` timestamp NULL DEFAULT '0000-00-00 00:00:00',
  `jct_lastmodified_by` varchar(50) DEFAULT NULL,
  `jct_created_by` varchar(50) DEFAULT NULL,
  PRIMARY KEY (`jct_user_profile`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8;

/*Data for the table `jct_user_profile` */

insert  into `jct_user_profile`(`jct_user_profile`,`jct_user_profile_desc`,`jct_soft_delete`,`version`,`jct_created_ts`,`jct_lastmodified_ts`,`jct_lastmodified_by`,`jct_created_by`) values 

(1,'All',0,1,'2014-04-14 16:51:58','2014-04-14 16:51:58','admin@interrait.com','admin@interrait.com'),

(2,'Default Profile',0,1,'2014-04-22 16:03:52','2014-04-22 16:03:52',NULL,'admin@interrait.com');

/*Table structure for table `jct_user_role` */

DROP TABLE IF EXISTS `jct_user_role`;

CREATE TABLE `jct_user_role` (
  `jct_role_id` int(5) NOT NULL,
  `jct_role_code` varchar(15) NOT NULL,
  `jct_role_desc` varchar(40) DEFAULT NULL,
  PRIMARY KEY (`jct_role_id`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8;

/*Data for the table `jct_user_role` */

insert  into `jct_user_role`(`jct_role_id`,`jct_role_code`,`jct_role_desc`) values 

(1,'USER','GENERAL_USER'),

(2,'ADMIN','ADMIN_USER'),

(3,'FACILITATOR','FACILITATOR_USER'),

(4,'ECOMMERCE','ECOMMERCE');

/* Function  structure for function  `jct_before_sketch_to_after_sketch` */

/*!50003 DROP FUNCTION IF EXISTS `jct_before_sketch_to_after_sketch` */;
DELIMITER $$

/*!50003 CREATE DEFINER=`jobCraftingTool` FUNCTION `jct_before_sketch_to_after_sketch`(jobRefNo Varchar(30)) RETURNS int(11)
BEGIN
                
                DECLARE l_return INT DEFAULT  0;
    DECLARE l_count INT DEFAULT  0;
                DECLARE l_jobRefNo Varchar(30);
                DECLARE l_jct_bs_task_id Int;
                DECLARE l_jct_bs_energy Int;
                DECLARE l_jct_bs_task_desc  Varchar(200);
                                DECLARE l_createdBy Varchar(50);
                
                DECLARE l_jct_as_task_id Int;
                DECLARE l_jct_as_energy Int;
                DECLARE l_jct_as_task_desc Varchar(200);
                
                DECLARE  diff_id   Varchar(30); 
        DECLARE  diff_energy   Varchar(30); 
                DECLARE  diff_desc   Varchar(100); 
                DECLARE done          INT DEFAULT  0;
                DECLARE innerLoopCount         INT DEFAULT  0;
                DECLARE innerLoopCountTo         INT DEFAULT  0;
                DECLARE l_length         INT DEFAULT  0;
                DECLARE ll_jct_bs_task_id Int;
                DECLARE ll_jct_bs_energy Int;
                DECLARE ll_jct_bs_task_desc  Varchar(200);
                
                DECLARE ll_jct_as_task_id Int;
                DECLARE ll_jct_as_energy Int;
                DECLARE ll_jct_as_task_desc Varchar(200);
                
                DECLARE  ll_diff_id   Varchar(30); 
    DECLARE  ll_diff_energy   Varchar(30); 
                DECLARE  ll_diff_desc   Varchar(100); 
                DECLARE  l_status   Varchar(100) DEFAULT ''; 
    DECLARE lll_jct_as_task_desc Varchar(200);
                DECLARE beforeSketch CURSOR FOR 
                                SELECT jct_bs_task_id,jct_bs_energy,rtrim(ltrim(jct_bs_task_desc)),jct_bs_created_by FROM jct_before_sketch_details 
                where jct_bs_soft_delete=0 and  jct_bs_jobref_no= jobRefNo;
                
                DECLARE afterSketch CURSOR FOR 
                                SELECT jct_as_task_id,jct_as_task_energy,rtrim(ltrim(jct_as_task_desc)) FROM jct_after_sketch_finalpage_details 
                where  jct_as_jobref_no= jobRefNo and jct_as_soft_delete=0 and jct_as_task_id=l_jct_bs_task_id;
                DECLARE afterSketchTo CURSOR FOR 
                                SELECT jct_as_task_id,jct_as_task_energy,rtrim(ltrim(jct_as_task_desc)) FROM jct_after_sketch_finalpage_details 
                where  jct_as_jobref_no= jobRefNo and jct_as_soft_delete=0 and (jct_as_element_desc is null or jct_as_element_desc ='');
                DECLARE beforeSketchTo CURSOR FOR 
                                SELECT jct_bs_task_id,jct_bs_energy,rtrim(ltrim(jct_bs_task_desc)) FROM jct_before_sketch_details 
                where jct_bs_soft_delete=0 and  jct_bs_jobref_no= jobRefNo and jct_bs_task_id=ll_jct_as_task_id; 
                DECLARE CONTINUE HANDLER FOR NOT FOUND SET done = 1;
                
                DECLARE EXIT HANDLER FOR SQLEXCEPTION
                                BEGIN
                                                SET l_return=1;
                                                SET l_status= 'Error occurred';
                                END; 
                SET l_jobRefNo=jobRefNo;
    Select count(*) into l_count from jct_bs_to_as where jct_jobref_no=jobRefNo;
    if (l_count>0) then 
      update jct_bs_to_as set jct_soft_delete=1 where jct_jobref_no=jobRefNo;
    end if;
                OPEN beforeSketch;
                cursor_loop: LOOP
                                FETCH beforeSketch into l_jct_bs_task_id,l_jct_bs_energy,l_jct_bs_task_desc,l_createdBy;
                                IF done=1 THEN
                                                LEAVE cursor_loop;
                                END IF;
                                OPEN afterSketch;
                                inner_cursor_loop: LOOP
                                                
                                                FETCH afterSketch into l_jct_as_task_id,l_jct_as_energy,l_jct_as_task_desc;
                                                IF done=1 THEN
                                                                LEAVE inner_cursor_loop;
                                                END IF;
                                                set innerLoopCount=innerLoopCount+1;
                                if((l_jct_bs_task_id=l_jct_as_task_id) and (l_jct_bs_energy=l_jct_as_energy) and 
                                                                (l_jct_bs_task_desc=l_jct_as_task_desc)) then
                                                
                                              --  SET diff_energy=l_jct_bs_energy-l_jct_as_energy;
            SET diff_energy=l_jct_as_energy - l_jct_bs_energy;
                                                SET l_length= locate('(', l_jct_as_task_desc);
                                                SET diff_desc= substring(l_jct_as_task_desc, l_length);
                                                 insert  into jct_bs_to_as (`jct_bs_to_as_id`,`jct_jobref_no`,`jct_task_id`,`jct_bs_energy`,`jct_bs_time`,`jct_bs_task_desc`,`jct_as_energy`,`jct_as_time`,`jct_as_task_desc`,`jct_diff_status`,`jct_diff_energy`,`jct_diff_time`,`jct_diff_task_desc`,`version`,`jct_bs_to_as_created_by`,`jct_bs_to_as_task_desc_diff`) 
                                                                Values ( null,l_jobRefNo,l_jct_as_task_id,l_jct_bs_energy,'0',l_jct_bs_task_desc,l_jct_as_energy,'0',l_jct_as_task_desc,'No Change',diff_energy,'0',diff_desc,0,l_createdBy,1);
                                                LEAVE inner_cursor_loop;                                           
                                elseif( (l_jct_bs_task_id=l_jct_as_task_id) and  (l_jct_bs_task_desc=l_jct_as_task_desc)) then 
                                              --  SET diff_energy=l_jct_bs_energy-l_jct_as_energy;
            SET diff_energy=l_jct_as_energy - l_jct_bs_energy;
                                                SET l_length= locate('(', l_jct_as_task_desc);
                                                SET diff_desc= substring(l_jct_as_task_desc, l_length);
                                                 insert  into jct_bs_to_as (`jct_bs_to_as_id`,`jct_jobref_no`,`jct_task_id`,`jct_bs_energy`,`jct_bs_time`,`jct_bs_task_desc`,`jct_as_energy`,`jct_as_time`,`jct_as_task_desc`,`jct_diff_status`,`jct_diff_energy`,`jct_diff_time`,`jct_diff_task_desc`,`version`,`jct_bs_to_as_created_by`,`jct_bs_to_as_task_desc_diff`) 
                                                                Values (null,l_jobRefNo,l_jct_as_task_id,l_jct_bs_energy,'0',l_jct_bs_task_desc,l_jct_as_energy,'0',l_jct_as_task_desc,'Modified Task',diff_energy,'0',diff_desc,0,l_createdBy,1);
                                                LEAVE inner_cursor_loop;
                                elseif((l_jct_bs_task_id=l_jct_as_task_id) and (l_jct_bs_energy=l_jct_as_energy)) then 
                                                -- SET diff_energy=l_jct_bs_energy-l_jct_as_energy;
            SET diff_energy=l_jct_as_energy - l_jct_bs_energy;
                                                SET l_length= locate('(', l_jct_as_task_desc);
                                                SET diff_desc= substring(l_jct_as_task_desc, l_length);
                                                insert  into jct_bs_to_as (`jct_bs_to_as_id`,`jct_jobref_no`,`jct_task_id`,`jct_bs_energy`,`jct_bs_time`,`jct_bs_task_desc`,`jct_as_energy`,`jct_as_time`,`jct_as_task_desc`,`jct_diff_status`,`jct_diff_energy`,`jct_diff_time`,`jct_diff_task_desc`,`version`,`jct_bs_to_as_created_by`,`jct_bs_to_as_task_desc_diff`) 
                                                                Values (null,l_jobRefNo,l_jct_as_task_id,l_jct_bs_energy,'0',l_jct_bs_task_desc,l_jct_as_energy,'0',l_jct_as_task_desc,'Modified Task',diff_energy,'0',diff_desc,0,l_createdBy,0);
                                                LEAVE inner_cursor_loop;
          /*elseif((l_jct_bs_task_desc=l_jct_as_task_desc) and (l_jct_bs_energy=l_jct_as_energy)) then 
              SET diff_energy=l_jct_bs_energy-l_jct_as_energy;
              SET l_length= locate('(', l_jct_as_task_desc);
              SET diff_desc= substring(l_jct_as_task_desc, l_length);
              insert  into jct_bs_to_as (`jct_bs_to_as_id`,`jct_jobref_no`,`jct_task_id`,`jct_bs_energy`,`jct_bs_time`,`jct_bs_task_desc`,`jct_as_energy`,`jct_as_time`,`jct_as_task_desc`,`jct_diff_status`,`jct_diff_energy`,`jct_diff_time`,`jct_diff_task_desc`,`version`,`jct_bs_to_as_created_by`) 
                  Values ( jct_generate_id ('jct_bs_to_as'),l_jobRefNo,l_jct_as_task_id,l_jct_bs_energy,'0',l_jct_bs_task_desc,l_jct_as_energy,'0',l_jct_as_task_desc,'Modified Task',diff_energy,'0',diff_desc,0,l_createdBy);
              LEAVE inner_cursor_loop;
          else 
              SET diff_energy=l_jct_bs_energy-l_jct_as_energy;
              SET l_length= locate('(', l_jct_as_task_desc);
              SET diff_desc= substring(l_jct_as_task_desc, l_length);
              insert  into jct_bs_to_as (`jct_bs_to_as_id`,`jct_jobref_no`,`jct_task_id`,`jct_bs_energy`,`jct_bs_time`,`jct_bs_task_desc`,`jct_as_energy`,`jct_as_time`,`jct_as_task_desc`,`jct_diff_status`,`jct_diff_energy`,`jct_diff_time`,`jct_diff_task_desc`,`version`,`jct_bs_to_as_created_by`) 
                  Values ( jct_generate_id ('jct_bs_to_as'),l_jobRefNo,l_jct_as_task_id,l_jct_bs_energy,'0',l_jct_bs_task_desc,l_jct_as_energy,'0',l_jct_as_task_desc,'Modified Task',diff_energy,'0',diff_desc,0,l_createdBy);
              LEAVE inner_cursor_loop;  */   
         elseif((l_jct_bs_task_id=l_jct_as_task_id) ) then 
                                                -- SET diff_energy=l_jct_bs_energy-l_jct_as_energy;
            SET diff_energy=l_jct_as_energy - l_jct_bs_energy;
                                                SET l_length= locate('(', l_jct_as_task_desc);
            SET lll_jct_as_task_desc=trim(substring(l_jct_as_task_desc,1, l_length-1));
                                                SET diff_desc= substring(l_jct_as_task_desc, l_length);
            if(l_jct_bs_task_desc=lll_jct_as_task_desc) then
              insert  into jct_bs_to_as (`jct_bs_to_as_id`,`jct_jobref_no`,`jct_task_id`,`jct_bs_energy`,`jct_bs_time`,`jct_bs_task_desc`,`jct_as_energy`,`jct_as_time`,`jct_as_task_desc`,`jct_diff_status`,`jct_diff_energy`,`jct_diff_time`,`jct_diff_task_desc`,`version`,`jct_bs_to_as_created_by`,`jct_bs_to_as_task_desc_diff`) 
                                                                Values ( null,l_jobRefNo,l_jct_as_task_id,l_jct_bs_energy,'0',l_jct_bs_task_desc,l_jct_as_energy,'0',l_jct_as_task_desc,'Modified Task',diff_energy,'0',diff_desc,0,l_createdBy,1);
               LEAVE inner_cursor_loop;
            else
              insert  into jct_bs_to_as (`jct_bs_to_as_id`,`jct_jobref_no`,`jct_task_id`,`jct_bs_energy`,`jct_bs_time`,`jct_bs_task_desc`,`jct_as_energy`,`jct_as_time`,`jct_as_task_desc`,`jct_diff_status`,`jct_diff_energy`,`jct_diff_time`,`jct_diff_task_desc`,`version`,`jct_bs_to_as_created_by`,`jct_bs_to_as_task_desc_diff`) 
                                                                Values ( null,l_jobRefNo,l_jct_as_task_id,l_jct_bs_energy,'0',l_jct_bs_task_desc,l_jct_as_energy,'0',l_jct_as_task_desc,'Modified Task',diff_energy,'0',diff_desc,0,l_createdBy,0);
               LEAVE inner_cursor_loop;
            end if;
                                               
                                end if;
                                                
                                END LOOP inner_cursor_loop;
                                CLOSE afterSketch;
                                SET done=0;
                                
                                IF (innerLoopCount=0) THEN
        SET l_jct_as_task_desc='';
        SET l_jct_as_energy='';
        SET diff_energy='';
        SET diff_desc='';
        
                                insert  into jct_bs_to_as (`jct_bs_to_as_id`,`jct_jobref_no`,`jct_task_id`,`jct_bs_energy`,`jct_bs_time`,`jct_bs_task_desc`,`jct_as_energy`,`jct_as_time`,`jct_as_task_desc`,`jct_diff_status`,`jct_diff_energy`,`jct_diff_time`,`jct_diff_task_desc`,`version`,`jct_bs_to_as_created_by`,`jct_bs_to_as_task_desc_diff`) 
                                                                Values ( null,l_jobRefNo,l_jct_bs_task_id,l_jct_bs_energy,'0',l_jct_bs_task_desc,l_jct_as_energy,'0',l_jct_as_task_desc,'Deleted Task',diff_energy,'0',diff_desc,0,l_createdBy,0);
                                else
                                                set innerLoopCount=0;
                                END IF;
                                
                                
                END LOOP cursor_loop;
                CLOSE beforeSketch;
                SET done=0;
                OPEN afterSketchTo;
                outer_cursor_loop: LOOP
                FETCH afterSketchTo into ll_jct_as_task_id,ll_jct_as_energy,ll_jct_as_task_desc;
                                IF done=1 THEN
                                                LEAVE outer_cursor_loop;
                                END IF;
                                OPEN beforeSketchTo;
                                inner_cursor_loop: LOOP
                                FETCH beforeSketchTo into ll_jct_bs_task_id,ll_jct_bs_energy,ll_jct_bs_task_desc;
                                                IF done=1 THEN
                                                                LEAVE inner_cursor_loop;
                                                END IF;
                                                set innerLoopCountTo=innerLoopCountTo+1;
                                END LOOP inner_cursor_loop;
                                CLOSE beforeSketchTo;
                                SET done=0;
                                IF (innerLoopCountTo=0) THEN
                                insert  into jct_bs_to_as (`jct_bs_to_as_id`,`jct_jobref_no`,`jct_task_id`,`jct_bs_energy`,`jct_bs_time`,`jct_bs_task_desc`,`jct_as_energy`,`jct_as_time`,`jct_as_task_desc`,`jct_diff_status`,`jct_diff_energy`,`jct_diff_time`,`jct_diff_task_desc`,`version`,`jct_bs_to_as_created_by`,`jct_bs_to_as_task_desc_diff`) 
                                                                Values ( null,l_jobRefNo,ll_jct_as_task_id,'','0','',ll_jct_as_energy,'0',ll_jct_as_task_desc,'New Task','','0','',0,l_createdBy,0);
                                else
                                                set innerLoopCountTo=0;
                                END IF;
                
        END LOOP outer_cursor_loop;
                CLOSE afterSketchTo;
      RETURN l_return;  
    END */$$
DELIMITER ;

/* Function  structure for function  `jct_delete_users` */

/*!50003 DROP FUNCTION IF EXISTS `jct_delete_users` */;
DELIMITER $$

/*!50003 CREATE DEFINER=`jobCraftingTool` FUNCTION `jct_delete_users`(user_id int) RETURNS int(11)
BEGIN
  DECLARE l_return INT DEFAULT  0;
  delete from jct_before_sketch_question where jct_bs_jobref_no IN (select distinct(jct_jobref_no) from jct_user_login_info where jct_user_id = user_id);
  delete from jct_before_sketch_details where jct_bs_header_id IN (SELECT distinct(jct_bs_header_id) from jct_before_sketch_header where jct_bs_user_id = user_id);
  delete from jct_before_sketch_header where jct_bs_user_id = user_id;
  delete from jct_action_plan where jct_action_plan_user_id = user_id;
        delete from jct_before_sketch_details_temp where jct_bs_header_id IN (SELECT distinct(jct_bs_header_id) from jct_before_sketch_header_temp where jct_bs_temp_user_id = user_id);
        delete from jct_before_sketch_header_temp where jct_bs_temp_user_id = user_id;
                                                
  delete from jct_after_sketch_pageone_details_temp where jct_as_header_id IN (SELECT distinct(jct_as_header_id) from jct_after_sketch_header_temp where jct_as_temp_user_id = user_id);
        delete from jct_after_sketch_finalpage_details_temp where jct_as_header_id IN (SELECT distinct(jct_as_header_id) from jct_after_sketch_header_temp where jct_as_temp_user_id = user_id);
        delete from jct_after_sketch_header_temp where jct_as_temp_user_id = user_id;
  delete from jct_after_sketch_pageone_details where jct_as_header_id IN (SELECT distinct(jct_as_header_id) from jct_after_sketch_header where jct_as_user_id = user_id);
        delete from jct_after_sketch_finalpage_details where jct_as_header_id IN (SELECT distinct(jct_as_header_id) from jct_after_sketch_header where jct_as_user_id = user_id);
        delete from jct_after_sketch_header where jct_as_user_id = user_id;
  delete from jct_bs_to_as WHERE jct_jobref_no IN (select distinct(jct_jobref_no) from jct_user_login_info where jct_user_id = user_id);
        delete from jct_completion_status WHERE jct_job_reference_no IN (select distinct(jct_jobref_no) from jct_user_login_info where jct_user_id = user_id);
        delete from jct_status_search WHERE jct_jobref_no IN (select distinct(jct_jobref_no) from jct_user_login_info where jct_user_id = user_id);
        delete from jct_pdf_records where jct_user_id = user_id;
        delete from jct_facilitator_details WHERE jct_fac_user_id = user_id;
        delete from jct_payment_details where jct_pmt_hdr_id IN (SELECT distinct(jct_pmt_hdr_id) from jct_payment_header where jct_pmt_hdr_user_id = user_id);
        delete from jct_payment_header where jct_pmt_hdr_user_id = user_id;
  delete from jct_check_payment_user_details where jct_user_id = user_id;
        delete from jct_job_attribute_frozen WHERE jct_job_attribute_user_id = user_id;
        DELETE FROM jct_survey_questions WHERE jct_survey_taken_by_user_id = user_id;
        DELETE FROM jct_user_details WHERE jct_user_id = user_id;
        DELETE FROM jct_user_login_info WHERE jct_user_id = user_id;
        DELETE FROM jct_user WHERE jct_user_id = user_id;
        RETURN l_return; 
END */$$
DELIMITER ;

/* Function  structure for function  `jct_generate_id` */

/*!50003 DROP FUNCTION IF EXISTS `jct_generate_id` */;
DELIMITER $$

/*!50003 CREATE DEFINER=`jobCraftingTool` FUNCTION `jct_generate_id`(idName Varchar(50)) RETURNS int(11)
BEGIN
  DECLARE jobRefNo1 Varchar(30);
  DECLARE l_id_val INT DEFAULT  0;
  DECLARE l_return INT DEFAULT  0;
  DECLARE done          INT DEFAULT  0;
  DECLARE idGenerator CURSOR FOR 
    SELECT jct_id_val FROM jct_id_gen_table 
          where  jct_id_name= idName;
  DECLARE CONTINUE HANDLER FOR NOT FOUND SET done = 1;
  OPEN idGenerator;
  cursor_loop: LOOP
    FETCH idGenerator into l_id_val;
    IF done=1 THEN
      LEAVE cursor_loop;
    END IF;
  
  END LOOP cursor_loop;
  CLOSE idGenerator;
  SET l_return=l_id_val+1;
  update jct_id_gen_table set jct_id_val=l_return where jct_id_name= idName;
  
 RETURN l_return;
END */$$
DELIMITER ;

/* Function  structure for function  `jct_global_profile_change_mapping_attrs` */

/*!50003 DROP FUNCTION IF EXISTS `jct_global_profile_change_mapping_attrs` */;
DELIMITER $$

/*!50003 CREATE DEFINER=`jobCraftingTool` FUNCTION `jct_global_profile_change_mapping_attrs`(mappingType varchar(10),existingVal Varchar(9999), newVal Varchar(9999), existingDesc varchar(9999), newDesc varchar(9999)) RETURNS varchar(1100) CHARSET utf8
BEGIN
  DECLARE l_mappingType varchar(3);
  DECLARE l_existingVal varchar(9999);
  DECLARE l_newVal varchar(9999);
  DECLARE existing_concat varchar(9999);
  DECLARE modified_concat varchar(9999);
  DECLARE l_jct_as_finalpage_json varchar(9999999);
  DECLARE l_jct_as_finalpage_json_temp varchar(9999999);
  DECLARE l_jct_as_page_one_json varchar(9999999);
  DECLARE l_jct_as_page_one_json_temp varchar(9999999);
  DECLARE l_jct_as_finalpage_json_modified varchar(9999999);
  DECLARE l_return INT DEFAULT  0;
  DECLARE finished INTEGER DEFAULT 0;
  DECLARE done INTEGER DEFAULT 0;
  DECLARE existing_concat_final varchar(9999);
  DECLARE existing_concat_final_temp varchar(9999);
  DECLARE modified_concat_final varchar(9999);
  DECLARE modified_concat_final_temp varchar(9999);
  DECLARE existing_concat_page_one varchar(9999);
  DECLARE existing_concat_page_one_temp varchar(9999);
  DECLARE modified_concat_page_one varchar(9999);
  DECLARE modified_concat_page_one_temp varchar(9999);
  
  DECLARE after_sketch_page_one_cursor CURSOR FOR
    SELECT jct_as_pagene_json FROM jct_after_sketch_header WHERE jct_as_pagene_json LIKE (SELECT CONCAT('%"divValue":"', existingVal, '"%'));
  
  DECLARE after_sketch_diagram_cursor CURSOR FOR
    SELECT jct_as_finalpage_json FROM jct_after_sketch_header WHERE jct_as_finalpage_json LIKE (SELECT CONCAT('%"divValue":"', existingVal, '"%'));
  
  DECLARE after_sketch_page_one_cursor_temp CURSOR FOR
    SELECT jct_as_pagene_json FROM jct_after_sketch_header_temp WHERE jct_as_pagene_json LIKE (SELECT CONCAT('%"divValue":"', existingVal, '"%'));
  
  DECLARE after_sketch_diagram_cursor_temp CURSOR FOR
    SELECT jct_as_finalpage_json FROM jct_after_sketch_header_temp WHERE jct_as_finalpage_json LIKE (SELECT CONCAT('%"divValue":"', existingVal, '"%'));
  
  
  -- declare NOT FOUND handler
  DECLARE CONTINUE HANDLER 
  FOR NOT FOUND SET finished = 1;
  
  OPEN after_sketch_page_one_cursor;
  cursor_loop: LOOP
    FETCH after_sketch_page_one_cursor into l_jct_as_page_one_json;
    SELECT CONCAT('"divValue":"', existingVal, '"') INTO existing_concat_page_one;
    SELECT CONCAT('"divValue":"', newVal, '"') INTO modified_concat_page_one;
    UPDATE jct_after_sketch_header SET jct_as_pagene_json = (REPLACE(l_jct_as_page_one_json, existing_concat_page_one, modified_concat_page_one)) WHERE jct_as_pagene_json = l_jct_as_page_one_json;
  IF finished=1 THEN
    LEAVE cursor_loop;
  END IF;
  END LOOP cursor_loop;
  CLOSE after_sketch_page_one_cursor;
  SET finished=0;
  OPEN after_sketch_diagram_cursor;
  cursor_loop: LOOP
  FETCH after_sketch_diagram_cursor into l_jct_as_finalpage_json;
    SELECT CONCAT('"divValue":"', existingVal, '"') INTO existing_concat_final;
    SELECT CONCAT('"divValue":"', newVal, '"') INTO modified_concat_final;
    UPDATE jct_after_sketch_header SET jct_as_finalpage_json = (REPLACE(l_jct_as_finalpage_json, existing_concat_final, modified_concat_final)) WHERE jct_as_finalpage_json = l_jct_as_finalpage_json;
  IF finished=1 THEN
    LEAVE cursor_loop;
  END IF;
  END LOOP cursor_loop;
  CLOSE after_sketch_diagram_cursor;
  SET finished=0;
  OPEN after_sketch_page_one_cursor_temp;
  cursor_loop: LOOP
    FETCH after_sketch_page_one_cursor_temp into l_jct_as_page_one_json_temp;
    SELECT CONCAT('"divValue":"', existingVal, '"') INTO existing_concat_page_one_temp;
    SELECT CONCAT('"divValue":"', newVal, '"') INTO modified_concat_page_one_temp;
    UPDATE jct_after_sketch_header_temp SET jct_as_pagene_json = (REPLACE(l_jct_as_page_one_json_temp, existing_concat_page_one_temp, modified_concat_page_one_temp)) WHERE jct_as_pagene_json = l_jct_as_page_one_json_temp;
  IF finished=1 THEN
    LEAVE cursor_loop;
  END IF;
  END LOOP cursor_loop;
  CLOSE after_sketch_page_one_cursor_temp;
  SET finished=0;
  OPEN after_sketch_diagram_cursor_temp;
  cursor_loop: LOOP
  FETCH after_sketch_diagram_cursor_temp into l_jct_as_finalpage_json_temp;
    SELECT CONCAT('"divValue":"', existingVal, '"') INTO existing_concat_final_temp;
    SELECT CONCAT('"divValue":"', newVal, '"') INTO modified_concat_final_temp;
    UPDATE jct_after_sketch_header_temp SET jct_as_finalpage_json = (REPLACE(l_jct_as_finalpage_json_temp, existing_concat_final_temp, modified_concat_final_temp)) WHERE jct_as_finalpage_json = l_jct_as_finalpage_json_temp;
  IF finished=1 THEN
    LEAVE cursor_loop;
  END IF;
  END LOOP cursor_loop;
  CLOSE after_sketch_diagram_cursor_temp;
  SET finished=0;
  
  UPDATE jct_after_sketch_finalpage_details SET jct_as_element_desc = newVal WHERE jct_as_element_desc = existingVal AND jct_as_element_code = mappingType;
  UPDATE jct_after_sketch_finalpage_details_temp SET jct_as_element_desc = newVal WHERE jct_as_element_desc = existingVal AND jct_as_element_code = mappingType;
  UPDATE jct_job_attribute_frozen SET jct_job_attribute_name = newVal WHERE jct_job_attribute_name = existingVal AND jct_job_attribute_code = mappingType;
  UPDATE jct_job_attribute_frozen SET jct_job_attribute_desc = newDesc WHERE jct_job_attribute_desc = existingDesc AND jct_job_attribute_code = mappingType;
  
  -- INSERT A NEW ROW IN GLOBAL PROFILE CHANGE HISTORY TABLE
  INSERT INTO jct_global_profile_history (jct_global_profile_str_original, 
        jct_global_profile_str_changed, jct_global_profile_str_desc_original, 
        jct_global_profile_str_desc_changed, jct_global_profile_is_attribute, 
        jct_global_profile_soft_delete, jct_global_profile_attribute_type) VALUES 
        (existingVal, newVal, existingDesc, newDesc,'Y', 0, mappingType);
  RETURN "DONE";
    END */$$
DELIMITER ;

/* Function  structure for function  `jct_global_profile_change_question` */

/*!50003 DROP FUNCTION IF EXISTS `jct_global_profile_change_question` */;
DELIMITER $$

/*!50003 CREATE DEFINER=`jobCraftingTool` FUNCTION `jct_global_profile_change_question`(pageCode varchar(10),mainQuestion Varchar(200), existingSubQtn varchar(200), newSubQtn varchar(200) ) RETURNS varchar(100) CHARSET utf8
BEGIN   
                DECLARE done INTEGER DEFAULT 0;
    DECLARE ans_desc VARCHAR (20000);
    DECLARE bq_header_id INTEGER(11);
    DECLARE l_jct_bs_question_id INTEGER(11);
    DECLARE l_jct_bs_answar_desc VARCHAR (20000);
    DECLARE l_jct_bs_jobref_no VARCHAR (100);
    DECLARE l_jct_bs_question_desc VARCHAR (20000);
    DECLARE l_jct_bs_status_id INTEGER(11);
    DECLARE l_version INTEGER(11);
    DECLARE l_jct_ds_created_ts TIMESTAMP;
    DECLARE l_jct_bs_lastmodified_ts TIMESTAMP;
    DECLARE l_jct_bs_lastmodified_by VARCHAR (100);
    DECLARE l_jct_bs_created_by VARCHAR (100);
    DECLARE l_jct_bs_header_id INTEGER(11);
    DECLARE l_jct_bs_sub_question VARCHAR (20000);
    DECLARE l_jct_bs_user_id INTEGER(11);
    -- ACTION PLAN VARIABLES
    DECLARE l_ap_jct_action_plan_id INTEGER(11);
    DECLARE l_ap_jct_jobref_no VARCHAR (20000);
    DECLARE l_ap_jct_action_plan_flg VARCHAR (1);
    DECLARE l_ap_jct_action_plan_desc VARCHAR (20000);
    DECLARE l_ap_jct_as_answar_desc VARCHAR (20000);
    DECLARE l_ap_jct_as_question_desc VARCHAR (20000);
    DECLARE l_ap_jct_as_question_sub_desc VARCHAR (20000);
    DECLARE l_ap_jct_status_id INTEGER(11);
    DECLARE l_ap_version INTEGER(11);
    DECLARE l_ap_jct_as_created_ts TIMESTAMP;
    DECLARE l_ap_jct_as_lastmodified_ts TIMESTAMP;
    DECLARE l_ap_jct_as_lastmodified_by VARCHAR (100);
    DECLARE l_ap_jct_as_created_by VARCHAR (100);
    DECLARE l_ap_jct_as_header_id INTEGER(11);
    DECLARE l_ap_jct_action_plan_user_id INTEGER(11);
                
    DECLARE questionnaire_cursor CURSOR FOR
      SELECT jct_bs_question_id, jct_bs_answar_desc, jct_bs_jobref_no, jct_bs_question_desc, jct_bs_status_id, version, 
             jct_ds_created_ts, jct_bs_lastmodified_ts, jct_bs_lastmodified_by, jct_bs_created_by, 
             jct_bs_header_id, jct_bs_sub_question, jct_bs_user_id FROM jct_before_sketch_question WHERE 
             jct_bs_question_desc = mainQuestion AND jct_bs_sub_question = existingSubQtn AND jct_bs_soft_delete = 0;
      
    DECLARE action_plan_cursor CURSOR FOR
      SELECT jct_action_plan_id, jct_jobref_no, jct_action_plan_flg, jct_action_plan_desc, jct_as_answar_desc, jct_as_question_desc, 
             jct_as_question_sub_desc, jct_status_id, version, jct_as_created_ts, jct_as_lastmodified_ts, jct_as_lastmodified_by, 
             jct_as_created_by, jct_as_header_id, jct_action_plan_user_id FROM jct_action_plan  
             WHERE jct_as_question_desc = mainQuestion AND jct_as_question_sub_desc = existingSubQtn AND jct_action_plan_soft_delete = 0;
      
    -- declare NOT FOUND handler
    DECLARE CONTINUE HANDLER 
    FOR NOT FOUND SET done = 1;
    -- QUESTIONNAIRE --
    IF pageCode = "BS" THEN 
      OPEN questionnaire_cursor;
      cursor_loop: LOOP
        FETCH questionnaire_cursor into l_jct_bs_question_id, l_jct_bs_answar_desc, l_jct_bs_jobref_no, 
        l_jct_bs_question_desc, l_jct_bs_status_id, l_version, l_jct_ds_created_ts, l_jct_bs_lastmodified_ts, 
        l_jct_bs_lastmodified_by, l_jct_bs_created_by, l_jct_bs_header_id, l_jct_bs_sub_question, l_jct_bs_user_id;
        IF done=1 THEN
          LEAVE cursor_loop;
        END IF;
        -- UPDATE THE EXISTNG SOFT DELETE TO 1
        UPDATE jct_before_sketch_question SET jct_bs_soft_delete = 1 WHERE jct_bs_question_id = l_jct_bs_question_id;
        -- INSERT A CORRESPONDING ROW WITH NEW VALUE AND SOFT DELETE AS 0
        INSERT INTO jct_before_sketch_question (jct_bs_answar_desc, jct_bs_jobref_no, jct_bs_question_desc, jct_bs_status_id, version, 
              jct_ds_created_ts, jct_bs_lastmodified_ts, jct_bs_lastmodified_by, jct_bs_created_by, 
              jct_bs_header_id, jct_bs_sub_question, jct_bs_user_id, jct_bs_soft_delete) VALUES (l_jct_bs_answar_desc, 
        l_jct_bs_jobref_no, l_jct_bs_question_desc, l_jct_bs_status_id, l_version, l_jct_ds_created_ts, 
        l_jct_bs_lastmodified_ts, l_jct_bs_lastmodified_by, l_jct_bs_created_by, l_jct_bs_header_id, newSubQtn, l_jct_bs_user_id, 0);
      END LOOP cursor_loop;
      CLOSE questionnaire_cursor;
      SET done=0;
      -- UPDATE THE EXISTING ROWS IF ANY
      UPDATE jct_global_profile_history SET jct_global_profile_soft_delete = 1 WHERE jct_global_profile_bs_main_qtn = mainQuestion AND jct_global_profile_bs_sub_qtn_changed = existingSubQtn;
      -- INSERT A NEW ROW
      INSERT INTO jct_global_profile_history (jct_global_profile_bs_sub_qtn_original, 
          jct_global_profile_bs_sub_qtn_changed, 
          jct_global_profile_bs_main_qtn, jct_global_profile_soft_delete) VALUES (existingSubQtn, 
          newSubQtn, mainQuestion, 0);
          END IF; 
    
    -- ACTION PLAN --
    IF pageCode = "AS" THEN 
      OPEN action_plan_cursor;
      cursor_loop: LOOP
        FETCH action_plan_cursor into l_ap_jct_action_plan_id, l_ap_jct_jobref_no, l_ap_jct_action_plan_flg, l_ap_jct_action_plan_desc, 
        l_ap_jct_as_answar_desc, l_ap_jct_as_question_desc, l_ap_jct_as_question_sub_desc, l_ap_jct_status_id, l_ap_version, 
        l_ap_jct_as_created_ts, l_ap_jct_as_lastmodified_ts, l_ap_jct_as_lastmodified_by, l_ap_jct_as_created_by, l_ap_jct_as_header_id, 
        l_ap_jct_action_plan_user_id;
        IF done=1 THEN
          LEAVE cursor_loop;
        END IF;
        -- UPDATE THE EXISTNG SOFT DELETE TO 1
        UPDATE jct_action_plan SET jct_action_plan_soft_delete = 1 WHERE jct_action_plan_id = l_ap_jct_action_plan_id;
        -- INSERT A CORRESPONDING ROW WITH NEW VALUE AND SOFT DELETE AS 0
        INSERT INTO jct_action_plan (jct_jobref_no, jct_action_plan_flg, jct_action_plan_desc, jct_as_answar_desc, jct_as_question_desc, 
        jct_as_question_sub_desc, jct_status_id, version, jct_as_created_ts, jct_as_lastmodified_ts, jct_as_lastmodified_by, 
        jct_as_created_by, jct_as_header_id, jct_action_plan_user_id, jct_action_plan_soft_delete) VALUES (l_ap_jct_jobref_no, l_ap_jct_action_plan_flg, l_ap_jct_action_plan_desc, 
        l_ap_jct_as_answar_desc, l_ap_jct_as_question_desc, newSubQtn, l_ap_jct_status_id, l_ap_version, 
        l_ap_jct_as_created_ts, l_ap_jct_as_lastmodified_ts, l_ap_jct_as_lastmodified_by, l_ap_jct_as_created_by, l_ap_jct_as_header_id, 
        l_ap_jct_action_plan_user_id, 0);
      END LOOP cursor_loop;
      CLOSE action_plan_cursor;
      SET done=0;
      -- UPDATE THE EXISTING ROWS IF ANY
      UPDATE jct_global_profile_history SET jct_global_profile_soft_delete = 1 WHERE jct_global_profile_ap_main_qtn = mainQuestion AND jct_global_profile_ap_sub_qtn_changed = existingSubQtn;
      -- INSERT A NEW ROW
      INSERT INTO jct_global_profile_history (jct_global_profile_ap_sub_qtn_original, 
          jct_global_profile_ap_sub_qtn_changed, 
          jct_global_profile_ap_main_qtn, jct_global_profile_soft_delete) VALUES (existingSubQtn, 
          newSubQtn, mainQuestion, 0);
          END IF;
 
RETURN "DONE";
    END */$$
DELIMITER ;

/* Function  structure for function  `jct_new_diagram` */

/*!50003 DROP FUNCTION IF EXISTS `jct_new_diagram` */;
DELIMITER $$

/*!50003 CREATE DEFINER=`jobCraftingTool` FUNCTION `jct_new_diagram`(jobFerNo Varchar(50),decision Varchar(1)) RETURNS int(11)
BEGIN                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                   
                DECLARE l_jobRefNo Varchar(50);  
                                DECLARE l_return INT DEFAULT  0; 
                DECLARE l_status INT DEFAULT  5; 
                DECLARE l_decision CHAR;
                DECLARE l_status_org INT DEFAULT  4; 
                DECLARE l_jct_bs_header_id int(11);
                                DECLARE l_jct_as_header_id int(11);
                                DECLARE l_version int(11);
                                DECLARE l_jct_as_insertion_sequence int(11);
                DECLARE EXIT HANDLER FOR SQLEXCEPTION
                BEGIN
                  --  l_return=1;   
                END; 
                
                                SET l_jobRefNo=jobFerNo;   
                                                                SET l_decision=decision;
                IF(l_decision='N') THEN
                     -- select before sketch header id based on job referance number
                     select    max(jct_bs_header_id) into l_jct_bs_header_id FROM jct_before_sketch_header
                                Where jct_bs_soft_delete=0 and jct_bs_status_id=l_status and jct_bs_jobref_no=l_jobRefNo;
                     -- delete before sketch date  based on job referance number
                     delete from jct_before_sketch_details where jct_bs_header_id=l_jct_bs_header_id;
                                     delete from jct_before_sketch_header where jct_bs_header_id=l_jct_bs_header_id;
                     delete from jct_before_sketch_question where jct_bs_header_id=l_jct_bs_header_id;
                     -- select after sketch header id based on job referance number
                     select    max(jct_as_header_id) into l_jct_as_header_id FROM jct_after_sketch_header
                                Where jct_as_soft_delete=0 and jct_as_status_id=l_status and jct_as_jobref_no=l_jobRefNo;
                     select    version,jct_as_insertion_sequence into l_version,l_jct_as_insertion_sequence FROM jct_after_sketch_header
                                Where jct_as_header_id=l_jct_as_header_id and jct_as_jobref_no=l_jobRefNo;
                       -- delete before sketch date  based on job referance number
                     delete from jct_after_sketch_finalpage_details where jct_as_header_id=l_jct_as_header_id;
                                     update  jct_after_sketch_header set jct_as_status_id=4,jct_as_finalpage_json =null,jct_as_finalpage_time_spent=0,version=l_version+1,jct_as_finalpage_snapshot=null,jct_as_finalpage_initial_json_value=null,jct_as_finalpage_snapshot_string=null,
                                jct_final_page_div_height=null,jct_final_page_div_width=null,jct_as_total_count=0,jct_as_insertion_sequence=l_jct_as_insertion_sequence+1 
                                where jct_as_header_id=l_jct_as_header_id ;
                    update  jct_after_sketch_pageone_details set jct_as_soft_delete=0
                                where jct_as_header_id=l_jct_as_header_id ;
                                    delete from jct_before_sketch_question where jct_bs_header_id=l_jct_bs_header_id;
                                    delete from jct_action_plan where jct_as_header_id=l_jct_as_header_id;
      -- disable the bs to as reports
        update jct_bs_to_as set jct_soft_delete=1 where jct_jobref_no = l_jobRefNo;
                  ELSE 
                                 -- select before sketch header id based on job referance number
                     select    max(jct_bs_header_id) into l_jct_bs_header_id FROM jct_before_sketch_header
                                Where jct_bs_soft_delete=0 and jct_bs_status_id=l_status_org and jct_bs_jobref_no=l_jobRefNo;
                     -- delete before sketch date  based on job referance number
                     delete from jct_before_sketch_details where jct_bs_header_id=l_jct_bs_header_id;
                                     delete from jct_before_sketch_header where jct_bs_header_id=l_jct_bs_header_id;
                     delete from jct_before_sketch_question where jct_bs_header_id=l_jct_bs_header_id;
                     -- select after sketch header id based on job referance number
                     select    max(jct_as_header_id) into l_jct_as_header_id FROM jct_after_sketch_header
                                Where jct_as_soft_delete=0 and jct_as_status_id=l_status_org and jct_as_jobref_no=l_jobRefNo;
                     select    version,jct_as_insertion_sequence into l_version,l_jct_as_insertion_sequence FROM jct_after_sketch_header
                                Where jct_as_header_id=l_jct_as_header_id and jct_as_jobref_no=l_jobRefNo;
                       -- delete before sketch date  based on job referance number
                     delete from jct_after_sketch_finalpage_details where jct_as_header_id=l_jct_as_header_id;
                                     update  jct_after_sketch_header set jct_as_status_id=4,jct_as_finalpage_json =null,jct_as_finalpage_time_spent=0,version=l_version+1,jct_as_finalpage_snapshot=null,jct_as_finalpage_initial_json_value=null,jct_as_finalpage_snapshot_string=null,
                                jct_final_page_div_height=null,jct_final_page_div_width=null,jct_as_total_count=0,jct_as_insertion_sequence=l_jct_as_insertion_sequence+1 
                                where jct_as_header_id=l_jct_as_header_id ;
                    update  jct_after_sketch_pageone_details set jct_as_soft_delete=0
                                where jct_as_header_id=l_jct_as_header_id ;
                                    delete from jct_before_sketch_question where jct_bs_header_id=l_jct_bs_header_id;
                                    delete from jct_action_plan where jct_as_header_id=l_jct_as_header_id;
      -- disable the bs to as reports
        update jct_bs_to_as set jct_soft_delete=1 where jct_jobref_no = l_jobRefNo;
                  END IF;
                               
                
      RETURN l_return;                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                       
END */$$
DELIMITER ;

/* Function  structure for function  `jct_populate_after_sketch` */

/*!50003 DROP FUNCTION IF EXISTS `jct_populate_after_sketch` */;
DELIMITER $$

/*!50003 CREATE DEFINER=`jobCraftingTool` FUNCTION `jct_populate_after_sketch`(jobFerNo Varchar(50)) RETURNS int(11)
BEGIN                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                   
  DECLARE l_jobRefNo Varchar(50);                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                         
  DECLARE l_return INT DEFAULT  0; 
  DECLARE l_status INT DEFAULT  5; 
  DECLARE l_max INT(11);
/* After sketch header*/
  DECLARE l_jct_as_header_id int(11);                                                                                
  DECLARE l_jct_as_jobref_no varchar(50);                                                                        
  DECLARE l_jct_as_status_id int(10);                                                                            
  DECLARE l_jct_as_pagene_json mediumtext;                                                                                    
  DECLARE l_jct_as_finalpage_json mediumtext;                                                                                 
  DECLARE l_jct_as_pageone_time_spent int(11);                                                                   
  DECLARE l_jct_as_finalpage_time_spent int(11);                                                                 
  DECLARE l_version int(11);                                                                                     
  DECLARE l_jct_as_created_ts timestamp;                       
  DECLARE l_jct_as_lastmodified_ts timestamp  DEFAULT '0000-00-00 00:00:00';                                          
  DECLARE l_jct_as_lastmodified_by varchar(50);                                                                  
  DECLARE l_jct_as_created_by varchar(50);                                                                       
  DECLARE l_jct_as_user_job_title varchar(30);                                                                   
  DECLARE l_jct_as_pageone_snapshot longblob;                                                                                 
  DECLARE l_jct_as_finalpage_snapshot longblob;                                                                               
  DECLARE l_jct_as_finalpage_initial_json_value mediumtext;                                                                   
  DECLARE l_jct_as_pageone_checked_value mediumtext;                                                                          
  DECLARE l_jct_as_pageone_snapshot_string longtext;                                                                          
  DECLARE l_jct_as_finalpage_snapshot_string longtext;                                                                        
  DECLARE l_jct_final_page_div_height varchar(25);                                                               
  DECLARE l_jct_final_page_div_width varchar(25);                                                                
  DECLARE l_jct_checked_strength varchar(7);                                                                     
  DECLARE l_jct_checked_passion varchar(7);                                                                      
  DECLARE l_jct_checked_value varchar(7);                                                                        
  DECLARE l_jct_as_soft_delete int(1);
  DECLARE l_jct_as_total_count int(11);
  DECLARE l_as_insertion_sequence int(11);
  DECLARE l_hdr_jct_as_user_id int(11);
/* End After sketch header*/
/*jct_after_sketch_finalpage_details*/
  DECLARE l_jct_as_finalpage_details_id int(11);                                                                                         
  /*DECLARE l_jct_as_jobref_no varchar(50);      */                                                                                      
  DECLARE l_jct_as_element_code varchar(5);                                                                                          
  DECLARE l_jct_as_element_desc varchar(50);                                                                                         
  DECLARE l_jct_as_position varchar(50);   
  DECLARE l_jct_as_role_desc varchar(130);                                                                                           
  DECLARE l_jct_as_task_energy int(11);                                                                                              
  DECLARE l_jct_as_task_id int(11);                                                                                                  
  DECLARE l_jct_as_task_desc varchar(200);                                                                                           
  DECLARE ll_jct_as_status_id int(10);                                                                                                
  /*DECLARE l_jct_as_soft_delete  tinyint(1) DEFAULT 0;
  DECLARE l_version int(11); */                                                                                                         
  DECLARE ll_jct_as_created_ts timestamp DEFAULT '0000-00-00 00:00:00';                                                                   
  DECLARE ll_jct_as_lastmodified_ts timestamp  DEFAULT '0000-00-00 00:00:00';                                                              
  DECLARE ll_jct_as_lastmodified_by varchar(50);                                                                                     
  DECLARE ll_jct_as_created_by varchar(50);                                                                                           
  /*DECLARE l_jct_as_header_id int(11);   */                                                                                                  
  DECLARE l_jct_as_additional_desc varchar(200);
  /* end jct_after_sketch_finalpage_details*/
  /*  jct_after_sketch_pageone_details*/
  DECLARE l_jct_as_pageone_details_id int(11);
  DECLARE ll_jct_as_element_code varchar(5);                                                                                        
  DECLARE l_jct_as_element_id int(11);                                                                                             
  DECLARE ll_jct_as_position varchar(50);
  DECLARE l_jct_version int(11);                                                                                                       
  DECLARE lll_jct_as_created_ts timestamp ;                                         
  DECLARE lll_jct_as_lastmodified_ts timestamp DEFAULT '0000-00-00 00:00:00';                                                            
  DECLARE lll_jct_as_lastmodified_by varchar(50);                                                                                    
  DECLARE lll_jct_as_created_by varchar(50);                                                                                             
  DECLARE l_jct_as_value_count int(11);                                                                                            
  DECLARE l_jct_as_passion_count int(11);                                                                                          
  DECLARE l_jct_as_strength_count int(11);  
  /* end jct_after_sketch_pageone_details*/
/*jct_action_plan*/
  DECLARE ll_max INT(11);
  DECLARE  l_jct_action_plan_id int(11);                                                                   
  DECLARE  l_jct_jobref_no varchar(50);                                                                
  DECLARE  l_jct_action_plan_flg varchar(1);                                                           
  DECLARE  l_jct_action_plan_desc varchar(100);                                                        
  DECLARE  l_jct_as_answar_desc varchar(1500);                                                         
  DECLARE  l_jct_as_question_desc varchar(200);                                                        
  DECLARE  l_jct_as_question_sub_desc varchar(200);
  DECLARE  l_ap_jct_action_plan_user_id int(11);
  /*DECLARE  l_version int(11);     */                                                                     
  DECLARE  llll_jct_as_created_ts timestamp ;            
  DECLARE  llll_jct_as_lastmodified_ts timestamp DEFAULT '0000-00-00 00:00:00';                               
  DECLARE  llll_jct_as_lastmodified_by varchar(50);                                                       
  DECLARE  llll_jct_as_created_by varchar(50);
  DECLARE l_jct_action_plan_soft_delete int(10);
/*jct_action_plan*/
  DECLARE done          INT DEFAULT  0;
  DECLARE aftersketchFinalPage CURSOR FOR 
  select jct_as_jobref_no,jct_as_element_code,jct_as_element_desc,jct_as_position,jct_as_role_desc,jct_as_task_energy,jct_as_task_id,jct_as_task_desc,jct_as_status_id,                                                                                    
  jct_as_created_ts,jct_as_lastmodified_ts,jct_as_lastmodified_by,jct_as_created_by,jct_as_additional_desc  
  from jct_after_sketch_finalpage_details WHERE jct_as_soft_delete=1 and jct_as_jobref_no=l_jobRefNo and jct_as_header_id=l_max;
  DECLARE aftersketchPageOne CURSOR FOR 
  select jct_as_element_code, jct_as_element_id,jct_as_position,version,jct_as_created_ts,jct_as_lastmodified_ts,jct_as_lastmodified_by,jct_as_created_by,                                                                                             
  jct_as_value_count,jct_as_passion_count,jct_as_strength_count  
  from jct_after_sketch_pageone_details WHERE jct_as_soft_delete=1 and jct_as_jobref_no=l_jobRefNo and jct_as_header_id=l_max;
  
  DECLARE actionPlan CURSOR FOR 
  select  jct_jobref_no,jct_action_plan_flg, jct_action_plan_desc, jct_as_answar_desc, jct_as_question_desc,jct_as_question_sub_desc,                                                    
                version,jct_as_created_ts,jct_as_lastmodified_ts,jct_as_lastmodified_by,jct_as_created_by,jct_action_plan_soft_delete,jct_action_plan_user_id 
  from jct_action_plan WHERE jct_action_plan_soft_delete=1  and jct_jobref_no=l_jobRefNo and jct_as_header_id=ll_max;
  DECLARE CONTINUE HANDLER FOR NOT FOUND SET done = 1;
  DECLARE EXIT HANDLER FOR SQLEXCEPTION
  BEGIN
    SET l_return= 1;
  END; 
  set l_jobRefNo=jobFerNo;
  -- set l_jct_as_header_id=jct_generate_id ('jct_after_sketch_header_id');
  Select max(jct_as_header_id) into l_max from jct_after_sketch_header where jct_as_soft_delete=1 and jct_as_jobref_no=l_jobRefNo and jct_as_status_id=l_status;
  Select max(jct_as_header_id) into ll_max from jct_action_plan where jct_action_plan_soft_delete=1 and jct_jobref_no=l_jobRefNo and jct_status_id=l_status;
  Select max(jct_as_insertion_sequence) into l_as_insertion_sequence from jct_after_sketch_header where jct_as_soft_delete=1 and jct_as_jobref_no=l_jobRefNo;
  
  select jct_as_jobref_no,jct_as_status_id,jct_as_pagene_json,jct_as_finalpage_json,jct_as_pageone_time_spent,jct_as_finalpage_time_spent,                                                                 
  version,jct_as_created_ts,jct_as_lastmodified_ts,jct_as_lastmodified_by,jct_as_created_by,jct_as_user_job_title,jct_as_pageone_snapshot,jct_as_finalpage_snapshot,                                                                               
  jct_as_finalpage_initial_json_value,jct_as_pageone_checked_value,jct_as_pageone_snapshot_string,jct_as_finalpage_snapshot_string,jct_final_page_div_height,                                                               
  jct_final_page_div_width,jct_checked_strength,jct_checked_passion,jct_checked_value,jct_as_soft_delete,jct_as_total_count,jct_as_insertion_sequence,jct_as_user_id into
  l_jct_as_jobref_no,l_jct_as_status_id,l_jct_as_pagene_json,l_jct_as_finalpage_json,l_jct_as_pageone_time_spent,l_jct_as_finalpage_time_spent,                                                                 
  l_version,l_jct_as_created_ts,l_jct_as_lastmodified_ts,l_jct_as_lastmodified_by,l_jct_as_created_by,l_jct_as_user_job_title,l_jct_as_pageone_snapshot,l_jct_as_finalpage_snapshot,l_jct_as_finalpage_initial_json_value,l_jct_as_pageone_checked_value,l_jct_as_pageone_snapshot_string,l_jct_as_finalpage_snapshot_string,l_jct_final_page_div_height,                                                             
  l_jct_final_page_div_width,l_jct_checked_strength,l_jct_checked_passion,l_jct_checked_value,l_jct_as_soft_delete,l_jct_as_total_count,l_as_insertion_sequence,l_hdr_jct_as_user_id 
  from jct_after_sketch_header where jct_as_soft_delete=1 and jct_as_jobref_no=l_jobRefNo and jct_as_header_id=l_max;
  
  insert into jct_after_sketch_header (jct_as_header_id,jct_as_jobref_no,jct_as_status_id,jct_as_pagene_json,jct_as_finalpage_json,jct_as_pageone_time_spent,jct_as_finalpage_time_spent,                                                                 
  version,jct_as_created_ts,jct_as_lastmodified_ts,jct_as_lastmodified_by,jct_as_created_by,jct_as_user_job_title,jct_as_pageone_snapshot,jct_as_finalpage_snapshot,                                                                               
  jct_as_finalpage_initial_json_value,jct_as_pageone_checked_value,jct_as_pageone_snapshot_string,jct_as_finalpage_snapshot_string,jct_final_page_div_height,                                                               
  jct_final_page_div_width,jct_checked_strength,jct_checked_passion,jct_checked_value,jct_as_soft_delete, jct_as_total_count,jct_as_insertion_sequence, jct_as_user_id) values
  (null,l_jct_as_jobref_no,l_status,l_jct_as_pagene_json,l_jct_as_finalpage_json,l_jct_as_pageone_time_spent,l_jct_as_finalpage_time_spent,                                                                 
  '0',l_jct_as_created_ts,l_jct_as_lastmodified_ts,l_jct_as_lastmodified_by,l_jct_as_created_by,l_jct_as_user_job_title,l_jct_as_pageone_snapshot,l_jct_as_finalpage_snapshot,l_jct_as_finalpage_initial_json_value,l_jct_as_pageone_checked_value,l_jct_as_pageone_snapshot_string,l_jct_as_finalpage_snapshot_string,l_jct_final_page_div_height,                                                             
  l_jct_final_page_div_width,l_jct_checked_strength,l_jct_checked_passion,l_jct_checked_value,'0',l_jct_as_total_count,l_as_insertion_sequence+1,l_hdr_jct_as_user_id);
  
  Select max(jct_as_header_id) into l_jct_as_header_id from jct_after_sketch_header where jct_as_soft_delete=0 and jct_as_jobref_no=l_jobRefNo and jct_as_status_id=l_status;
  
  OPEN aftersketchFinalPage;
  cursor_loop: LOOP
  FETCH aftersketchFinalPage into l_jct_as_jobref_no,l_jct_as_element_code,l_jct_as_element_desc,l_jct_as_position,l_jct_as_role_desc,l_jct_as_task_energy,                                                                                              
  l_jct_as_task_id,l_jct_as_task_desc,ll_jct_as_status_id,ll_jct_as_created_ts,ll_jct_as_lastmodified_ts,ll_jct_as_lastmodified_by,ll_jct_as_created_by,l_jct_as_additional_desc;
  
  IF done=1 THEN
    LEAVE cursor_loop;
  END IF;
  insert into jct_after_sketch_finalpage_details  (`jct_as_finalpage_details_id`,`jct_as_jobref_no`,`jct_as_element_code`,`jct_as_element_desc`,`jct_as_position`,`jct_as_role_desc`,`jct_as_task_energy`,`jct_as_task_id`,`jct_as_task_desc`,               
        `jct_as_status_id`,`jct_as_soft_delete`,`version`,`jct_as_created_ts`,`jct_as_lastmodified_ts`,`jct_as_lastmodified_by`,`jct_as_created_by`,`jct_as_header_id`,`jct_as_additional_desc`) 
  values (null,l_jct_as_jobref_no,l_jct_as_element_code,l_jct_as_element_desc,l_jct_as_position,l_jct_as_role_desc,l_jct_as_task_energy,l_jct_as_task_id,l_jct_as_task_desc,ll_jct_as_status_id,'0',                                                                                   
    '0',ll_jct_as_created_ts,ll_jct_as_lastmodified_ts,ll_jct_as_lastmodified_by,ll_jct_as_created_by,l_jct_as_header_id,l_jct_as_additional_desc);
  END LOOP cursor_loop;
  CLOSE aftersketchFinalPage;
SET done=0;
  OPEN aftersketchPageOne;
  cursor_loop1: LOOP
  FETCH aftersketchPageOne into ll_jct_as_element_code,l_jct_as_element_id,ll_jct_as_position,l_jct_version,lll_jct_as_created_ts,lll_jct_as_lastmodified_ts,lll_jct_as_lastmodified_by,lll_jct_as_created_by,
  l_jct_as_value_count,l_jct_as_passion_count,l_jct_as_strength_count;
  
  IF done=1 THEN
    LEAVE cursor_loop1;
  END IF;
  insert into jct_after_sketch_pageone_details(`jct_as_pageone_details_id`,`jct_as_jobref_no`,`jct_as_element_code`,`jct_as_element_id`,`jct_as_position`,`jct_as_soft_delete`,`version`,
        `jct_as_created_ts`,`jct_as_lastmodified_ts`,`jct_as_lastmodified_by`,`jct_as_created_by`,`jct_as_header_id`,`jct_as_value_count`,`jct_as_passion_count`,`jct_as_strength_count`)
    values(null,l_jobRefNo,ll_jct_as_element_code,l_jct_as_element_id,ll_jct_as_position,'0','0',
      lll_jct_as_created_ts,lll_jct_as_lastmodified_ts,lll_jct_as_lastmodified_by,lll_jct_as_created_by,l_jct_as_header_id,l_jct_as_value_count,l_jct_as_passion_count,l_jct_as_strength_count);
  END LOOP cursor_loop1;
  CLOSE aftersketchPageOne;
  SET done=0;
  OPEN actionPlan;
  cursor_loop: LOOP
  FETCH actionPlan into l_jct_jobref_no,l_jct_action_plan_flg,l_jct_action_plan_desc,l_jct_as_answar_desc,l_jct_as_question_desc,l_jct_as_question_sub_desc,
     l_version,llll_jct_as_created_ts,llll_jct_as_lastmodified_ts,llll_jct_as_lastmodified_by,llll_jct_as_created_bY,l_jct_action_plan_soft_delete,l_ap_jct_action_plan_user_id;
  
  IF done=1 THEN
    LEAVE cursor_loop;
  END IF;
  insert into jct_action_plan (`jct_action_plan_id`,                                                                   
                   `jct_jobref_no`,                                                                
                   `jct_action_plan_flg`,                                                           
                   `jct_action_plan_desc`,                                                        
                   `jct_as_answar_desc`,                                                         
                   `jct_as_question_desc`,                                                        
                   `jct_as_question_sub_desc`,                                                    
                   `jct_status_id`,                                                                    
                   `version`,                                                                          
                   `jct_as_created_ts`,            
                   `jct_as_lastmodified_ts`,                               
                   `jct_as_lastmodified_by`,                                                       
                   `jct_as_created_by`,
       `jct_action_plan_soft_delete`,
       `jct_as_header_id`,
       `jct_action_plan_user_id`) 
      values (
        null,
        l_jobRefNo,                                                                
        l_jct_action_plan_flg,                                                           
        l_jct_action_plan_desc,                                                        
        l_jct_as_answar_desc,                                                         
        l_jct_as_question_desc,                                                        
        l_jct_as_question_sub_desc,                                                    
        l_status,                                                                   
        '0',                                                                          
        llll_jct_as_created_ts,            
        llll_jct_as_lastmodified_ts,                               
        llll_jct_as_lastmodified_by,                                                       
        llll_jct_as_created_by,
        '0',
        l_jct_as_header_id,
        l_ap_jct_action_plan_user_id);
  END LOOP cursor_loop;
  CLOSE actionPlan;
  
      RETURN l_return;                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                       
END */$$
DELIMITER ;

/* Function  structure for function  `jct_populate_before_sketch` */

/*!50003 DROP FUNCTION IF EXISTS `jct_populate_before_sketch` */;
DELIMITER $$

/*!50003 CREATE DEFINER=`jobCraftingTool` FUNCTION `jct_populate_before_sketch`(jobFerNo Varchar(50)) RETURNS int(11)
BEGIN                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                   
                DECLARE l_jobRefNo Varchar(50);                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                         
                DECLARE l_return INT DEFAULT  0; 
                DECLARE l_status INT DEFAULT  5; 
                DECLARE l_jct_bs_header_id int(11);
                DECLARE l_max INT(11);
                /*jct_before_sketch_header*/
                DECLARE l_jct_bs_jobref_no varchar(50);                                                                         
                DECLARE l_jct_bs_status_id int(10);                                                                             
                DECLARE l_jct_bs_json mediumtext;                                                                                            
                DECLARE l_jct_bs_time_spent int(11); 
                DECLARE l_jct_ds_created_ts timestamp ;                      
                DECLARE l_jct_bs_lastmodified_ts timestamp DEFAULT '0000-00-00 00:00:00';                                           
                DECLARE l_jct_bs_lastmodified_by varchar(50);                                                                   
                DECLARE l_jct_bs_created_by varchar(50);                                                                        
                DECLARE l_jct_bs_user_job_title varchar(30);                                                                    
                DECLARE l_jct_bs_snapshot longblob ;                                                                                 
                DECLARE l_jct_bs_initial_json_value mediumtext ;                                                                     
                DECLARE l_jct_bs_snapshot_string longtext ; 
                DECLARE l_bs_insertion_sequence int(11);
                /*jct_before_sketch_header*/
                /*jct_before_sketch_details*/
                DECLARE l_jct_bs_task_id int(11) ;             
                /*DECLARE l_jct_bs_jobref_no varchar(50) ; */      
                DECLARE l_jct_bs_energy int(11) ;              
                DECLARE l_jct_bs_time int(11) ;                
                DECLARE l_jct_bs_task_desc varchar(200) ;      
                DECLARE ll_jct_bs_status_id int(10) ;           
                DECLARE l_jct_bs_position varchar(50) ; 
                DECLARE ll_jct_ds_created_ts timestamp ;                
                DECLARE ll_jct_bs_lastmodified_ts timestamp  DEFAULT '0000-00-00 00:00:00';      
                DECLARE ll_jct_bs_lastmodified_by varchar(50) ; 
                DECLARE ll_jct_bs_created_by varchar(50) ;  
                DECLARE done          INT DEFAULT  0;
                /*jct_before_sketch_details*/
                DECLARE ll_max INT(11);
                /*jct_before_sketch_question*/
                DECLARE l_jct_bs_answar_desc varchar(2000);
                DECLARE l_jct_bs_question_desc varchar(500);                                                                      
                DECLARE l_version int(11);                                                                                        
                DECLARE lll_jct_ds_created_ts timestamp  DEFAULT '0000-00-00 00:00:00';                                                  
                DECLARE lll_jct_bs_lastmodified_ts timestamp  DEFAULT '0000-00-00 00:00:00';                                            
                DECLARE lll_jct_bs_lastmodified_by varchar(50);                                                                     
                DECLARE lll_jct_bs_created_by varchar(50); 
                DECLARE l_jct_bs_sub_question varchar(2000);
    DECLARE l_jct_bs_user_id int(11);
    DECLARE l_jct_bsq_user_id int(11);
                
                /*jct_before_sketch_question*/
                DECLARE reflectionQuestion CURSOR FOR 
                                select jct_bs_answar_desc,jct_bs_question_desc,jct_ds_created_ts,jct_bs_lastmodified_ts,jct_bs_lastmodified_by,jct_bs_created_by,jct_bs_sub_question, jct_bs_user_id 
                FROM jct_before_sketch_question where  jct_bs_jobref_no= l_jobRefNo and jct_bs_soft_delete=1 and jct_bs_header_id=ll_max;
                DECLARE beforeSketch CURSOR FOR 
                                select jct_bs_task_id,jct_bs_energy,jct_bs_time,jct_bs_task_desc,jct_bs_status_id,           
                                jct_bs_position,jct_ds_created_ts,jct_bs_lastmodified_ts,jct_bs_lastmodified_by,jct_bs_created_by  
                FROM jct_before_sketch_details             where  jct_bs_jobref_no= l_jobRefNo and jct_bs_soft_delete=1 and jct_bs_header_id=l_max;
                DECLARE CONTINUE HANDLER FOR NOT FOUND SET done = 1;
                DECLARE EXIT HANDLER FOR SQLEXCEPTION
                BEGIN
                                SET l_return= 1;
                END; 
                set l_jobRefNo=jobFerNo;
                -- set l_jct_bs_header_id=jct_generate_id ('jct_before_sketch_header_id');
                Select max(jct_bs_header_id) into ll_max from jct_before_sketch_question where jct_bs_soft_delete=1 and jct_bs_jobref_no=l_jobRefNo and jct_bs_status_id=l_status;
                Select max(jct_bs_header_id) into l_max from jct_before_sketch_header where jct_bs_soft_delete=1 and jct_bs_jobref_no=l_jobRefNo  and jct_bs_status_id=l_status;
                Select max(jct_bs_insertion_sequence) into l_bs_insertion_sequence from jct_before_sketch_header where jct_bs_soft_delete=1 and jct_bs_jobref_no=l_jobRefNo;
                
                
                Select jct_bs_jobref_no ,jct_bs_status_id,jct_bs_json,jct_bs_time_spent,jct_ds_created_ts,jct_bs_lastmodified_ts ,jct_bs_lastmodified_by,                                                                   
                                jct_bs_created_by,jct_bs_user_job_title,jct_bs_snapshot,jct_bs_initial_json_value,jct_bs_snapshot_string,jct_bs_insertion_sequence,jct_bs_user_id into  
                l_jct_bs_jobref_no ,l_jct_bs_status_id,l_jct_bs_json,l_jct_bs_time_spent,l_jct_ds_created_ts,l_jct_bs_lastmodified_ts ,l_jct_bs_lastmodified_by,                                                                   
                                l_jct_bs_created_by,l_jct_bs_user_job_title,l_jct_bs_snapshot,l_jct_bs_initial_json_value,l_jct_bs_snapshot_string,l_bs_insertion_sequence, l_jct_bs_user_id
                from  jct_before_sketch_header where jct_bs_soft_delete=1 and jct_bs_jobref_no=l_jobRefNo and jct_bs_header_id=l_max;
                
                insert into jct_before_sketch_header(`jct_bs_header_id`,`jct_bs_jobref_no`,`jct_bs_status_id`,`jct_bs_json`,`jct_bs_time_spent`,`version`,`jct_ds_created_ts`,`jct_bs_lastmodified_ts`,
                `jct_bs_lastmodified_by`,`jct_bs_created_by`,`jct_bs_user_job_title`,`jct_bs_snapshot`,`jct_bs_initial_json_value`,`jct_bs_snapshot_string`,                  
                `jct_bs_soft_delete`,`jct_bs_insertion_sequence`, `jct_bs_user_id`) Values
                ( null,l_jobRefNo,l_status,l_jct_bs_json,l_jct_bs_time_spent,'0',l_jct_ds_created_ts,l_jct_bs_lastmodified_ts,
                l_jct_bs_lastmodified_by,l_jct_bs_created_by,l_jct_bs_user_job_title,l_jct_bs_snapshot,l_jct_bs_initial_json_value,l_jct_bs_snapshot_string,'0',l_bs_insertion_sequence+1, l_jct_bs_user_id);
                
                -- Select jct_bs_header_id into l_jct_bs_header_id from jct_before_sketch_header where jct_bs_soft_delete=1 and jct_bs_jobref_no=l_jobRefNo and jct_bs_header_id=l_max;
                Select jct_bs_header_id into l_jct_bs_header_id from jct_before_sketch_header where jct_bs_soft_delete=0 and jct_bs_jobref_no=l_jobRefNo and jct_bs_status_id=l_status;
                OPEN beforeSketch;
                cursor_loop: LOOP
                FETCH beforeSketch into l_jct_bs_task_id,l_jct_bs_energy,l_jct_bs_time,l_jct_bs_task_desc,ll_jct_bs_status_id,l_jct_bs_position,
                ll_jct_ds_created_ts,ll_jct_bs_lastmodified_ts,ll_jct_bs_lastmodified_by,ll_jct_bs_created_by;
                
                IF done=1 THEN
                                LEAVE cursor_loop;
                END IF;
                Insert into jct_before_sketch_details(`jct_bs_detail_id`,`jct_bs_task_id`,`jct_bs_jobref_no`,`jct_bs_energy`,`jct_bs_time`,`jct_bs_task_desc`,`jct_bs_status_id`,
                `jct_bs_position`,`jct_bs_soft_delete`,`version`,`jct_ds_created_ts`,`jct_bs_lastmodified_ts`,`jct_bs_lastmodified_by`,`jct_bs_created_by`,`jct_bs_header_id`) Values 
                (null,l_jct_bs_task_id,l_jobRefNo,l_jct_bs_energy,l_jct_bs_time,l_jct_bs_task_desc,ll_jct_bs_status_id,
                l_jct_bs_position,'0','0',ll_jct_ds_created_ts,ll_jct_bs_lastmodified_ts,ll_jct_bs_lastmodified_by,ll_jct_bs_created_by,l_jct_bs_header_id);
                END LOOP cursor_loop;
                CLOSE beforeSketch;
                SET done=0;
                
                OPEN reflectionQuestion;
                cursor_loop: LOOP
                FETCH reflectionQuestion into l_jct_bs_answar_desc,l_jct_bs_question_desc,lll_jct_ds_created_ts,lll_jct_bs_lastmodified_ts,                                             
                lll_jct_bs_lastmodified_by,lll_jct_bs_created_by,l_jct_bs_sub_question, l_jct_bsq_user_id;
                
                IF done=1 THEN
                                LEAVE cursor_loop;
                END IF;
                Insert into jct_before_sketch_question (`jct_bs_answar_desc`,`jct_bs_question_id`,`jct_bs_jobref_no`,`jct_bs_question_desc`,`jct_bs_status_id`,`version`,`jct_ds_created_ts`,
                                    `jct_bs_lastmodified_ts`,`jct_bs_lastmodified_by`,`jct_bs_created_by`,`jct_bs_soft_delete`,`jct_bs_header_id`, `jct_bs_sub_question`, `jct_bs_user_id`) values
                                    (l_jct_bs_answar_desc,null,l_jobRefNo,l_jct_bs_question_desc,l_status,'0',lll_jct_ds_created_ts,lll_jct_bs_lastmodified_ts,                                             
                                    lll_jct_bs_lastmodified_by,lll_jct_bs_created_by,'0',l_jct_bs_header_id,l_jct_bs_sub_question, l_jct_bsq_user_id);
                END LOOP cursor_loop;
                CLOSE reflectionQuestion;
      RETURN l_return;                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                       
END */$$
DELIMITER ;

/* Function  structure for function  `jct_update_status` */

/*!50003 DROP FUNCTION IF EXISTS `jct_update_status` */;
DELIMITER $$

/*!50003 CREATE DEFINER=`jobCraftingTool` FUNCTION `jct_update_status`(jobFerNo Varchar(50),decision char(1)) RETURNS int(11)
BEGIN                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                   
                DECLARE l_jobRefNo Varchar(50);                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                         
                DECLARE l_return INT DEFAULT  0; 
                DECLARE l_status INT DEFAULT  5; 
                DECLARE l_decision CHAR;
                DECLARE l_status_org INT DEFAULT  4; 
                DECLARE l_jct_bs_snapshot longblob;
                DECLARE l_jct_bs_snapshot_string longtext;
                DECLARE l_jct_as_finalpage_snapshot longblob;
                DECLARE l_jct_as_finalpage_snapshot_string longtext;
                DECLARE l_jct_bs_user_job_title varchar(50);
                DECLARE l_as_header_count INT DEFAULT  0;
                DECLARE l_bs_header_count INT DEFAULT  0;
    DECLARE l_jct_completion_id int(11);
    DECLARE l_max_status_id int(11);
    DECLARE l_count INT(11); 
    DECLARE l_pdf_id INT(11); 
    -- OCCUPATION --
    DECLARE l_occupation_code varchar(50);
    DECLARE l_user_id INT DEFAULT  0;
                DECLARE EXIT HANDLER FOR SQLEXCEPTION
                BEGIN
                               -- update jct_action_plan set jct_status_id=l_status where jct_jobref_no= l_jobRefNo;     
                             --   update jct_after_sketch_header set jct_as_status_id=l_status where jct_as_jobref_no= l_jobRefNo; 
                               -- update jct_before_sketch_details set jct_bs_status_id=l_status where jct_bs_jobref_no= l_jobRefNo;  
               --   update jct_before_sketch_header set jct_bs_status_id=l_status where jct_bs_jobref_no= l_jobRefNo;  
                              --  update jct_before_sketch_question set jct_bs_status_id=l_status where jct_bs_jobref_no= l_jobRefNo; 
                              --  SET l_return= 1;
                END; 
                
                                SET l_jobRefNo=jobFerNo;      
                                SET l_decision=decision;   
                                SELECT jct_bs_snapshot,jct_bs_snapshot_string,jct_bs_user_job_title into l_jct_bs_snapshot,l_jct_bs_snapshot_string,l_jct_bs_user_job_title FROM jct_before_sketch_header 
                                where jct_bs_soft_delete=0 and  jct_bs_jobref_no= l_jobRefNo;
                
        -- SELECT THE USER ID FOR THE JOB REFERENCE NUMBER --
        SELECT max(jct_user_id) INTO l_user_id from jct_user_login_info where jct_jobref_no = l_jobRefNo;
        SELECT max(jct_user_onet_occupation) INTO l_occupation_code from jct_user_details where jct_user_id = l_user_id AND jct_user_details_soft_delete = 0;
        
                                SELECT jct_as_finalpage_snapshot,jct_as_finalpage_snapshot_string into l_jct_as_finalpage_snapshot,l_jct_as_finalpage_snapshot_string
                                FROM jct_after_sketch_header where jct_as_soft_delete=0 and  jct_as_jobref_no= l_jobRefNo;
        SELECT count(jct_completion_id) into l_jct_completion_id FROM jct_completion_status
        Where jct_job_reference_no=l_jobRefNo;
        Select jct_file_id into l_pdf_id from jct_pdf_records where jct_job_reference_no=l_jobRefNo and jct_show_pdf=0 
                          and jct_created_timestamp = (select max(jct_created_timestamp) from jct_pdf_records where jct_job_reference_no=l_jobRefNo);                     
        
                                IF (l_jct_completion_id>0) THEN
          SELECT count(jct_bs_header_id) into l_bs_header_count FROM jct_before_sketch_header 
          where jct_bs_soft_delete=1  and jct_bs_status_id=l_status and  jct_bs_jobref_no= l_jobRefNo;
      
          SELECT count(jct_as_header_id) into l_as_header_count
          FROM jct_after_sketch_header where jct_as_soft_delete=1 and jct_as_status_id=l_status and  jct_as_jobref_no= l_jobRefNo;
        ELSE
          SELECT count(jct_bs_header_id) into l_bs_header_count FROM jct_before_sketch_header 
          where jct_bs_soft_delete=1  and jct_bs_status_id=l_status_org and  jct_bs_jobref_no= l_jobRefNo;
      
          SELECT count(jct_as_header_id) into l_as_header_count
          FROM jct_after_sketch_header where jct_as_soft_delete=1 and jct_as_status_id=l_status_org and  jct_as_jobref_no= l_jobRefNo;
        END IF;
        -- 07-082014
        SELECT count(jct_status_search_id) into l_count FROM jct_status_search where jct_jobref_no = jobFerNo;
                                
                                IF (l_decision='C') THEN  
                                                update jct_status_search set jct_soft_delete=1 where jct_jobref_no= l_jobRefNo and jct_soft_delete=0;                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                
                                                -- update jct_action_plan set jct_status_id=l_status where jct_jobref_no= l_jobRefNo;  
                                                -- update jct_after_sketch_header set jct_as_status_id=l_status where jct_as_jobref_no= l_jobRefNo; 
                                                -- update jct_before_sketch_header set jct_bs_status_id=l_status where jct_bs_jobref_no= l_jobRefNo;  
                                                -- update jct_before_sketch_question set jct_bs_status_id=l_status where jct_bs_jobref_no= l_jobRefNo; 
                         
                                              --  Insert into jct_status_search (`jct_status_search_id`,`jct_jobref_no`,`jct_as_snapshot`,`jct_bs_snapshot`,`jct_bs_snapshot_string`,`jct_as_snapshot_string`,`jct_status_id`,`jct_soft_delete`,`version`,`jct_user_levels`)
                                                  --              values ( jct_generate_id ('jct_status_id'),l_jobRefNo,l_jct_as_finalpage_snapshot,l_jct_bs_snapshot,l_jct_bs_snapshot_string,l_jct_as_finalpage_snapshot_string,'5','0','0',l_jct_bs_user_job_title);
              Insert into jct_status_search (`jct_status_search_id`,`jct_jobref_no`,`jct_as_snapshot`,`jct_bs_snapshot`,`jct_bs_snapshot_string`,`jct_as_snapshot_string`,`jct_status_id`,`jct_soft_delete`,`version`,`jct_user_levels`, `jct_created_ts`, occupation_code)
                                                                values ( null,l_jobRefNo,l_jct_as_finalpage_snapshot,l_jct_bs_snapshot,l_jct_bs_snapshot_string,l_jct_as_finalpage_snapshot_string,'5','0','0',l_jct_bs_user_job_title, (select CURRENT_TIMESTAMP()), l_occupation_code);
                                
                        update jct_action_plan set jct_action_plan_soft_delete=1,jct_status_id=l_status where jct_jobref_no= l_jobRefNo and jct_action_plan_soft_delete=0;  
                        update jct_after_sketch_header set jct_as_soft_delete=1,jct_as_status_id=l_status where jct_as_jobref_no= l_jobRefNo and jct_as_soft_delete=0; 
                        update jct_after_sketch_finalpage_details set jct_as_soft_delete=1 where jct_as_jobref_no= l_jobRefNo and jct_as_soft_delete=0;
                        update jct_after_sketch_pageone_details set jct_as_soft_delete=1 where jct_as_jobref_no= l_jobRefNo and jct_as_soft_delete=0;
                        update jct_before_sketch_header set jct_bs_soft_delete=1,jct_bs_status_id=l_status where jct_bs_jobref_no= l_jobRefNo and jct_bs_soft_delete=0;  
                        update jct_before_sketch_details set jct_bs_soft_delete=1 where jct_bs_jobref_no= l_jobRefNo and jct_bs_soft_delete=0; 
                        update jct_before_sketch_question set jct_bs_soft_delete=1,jct_bs_status_id=l_status  where jct_bs_jobref_no= l_jobRefNo and jct_bs_soft_delete=0;        
            Update jct_pdf_records SET jct_show_pdf = 1 WHERE jct_file_id = l_pdf_id;
                                ELSEIF(l_decision='P') THEN
                                        IF(l_bs_header_count> 0 and l_as_header_count>0) THEN 
                                                update jct_status_search set jct_soft_delete=1 where jct_jobref_no= l_jobRefNo and jct_soft_delete=0 and jct_status_decision=0;
                                        END IF;                               
                                                                     
                      
                        update jct_action_plan set jct_action_plan_soft_delete=1,jct_status_id=l_status where jct_jobref_no= l_jobRefNo and jct_action_plan_soft_delete=0;  
                        update jct_after_sketch_header set jct_as_soft_delete=1,jct_as_status_id=l_status where jct_as_jobref_no= l_jobRefNo and jct_as_soft_delete=0; 
                        update jct_after_sketch_finalpage_details set jct_as_soft_delete=1 where jct_as_jobref_no= l_jobRefNo and jct_as_soft_delete=0;
                        update jct_after_sketch_pageone_details set jct_as_soft_delete=1 where jct_as_jobref_no= l_jobRefNo and jct_as_soft_delete=0;
                        update jct_before_sketch_header set jct_bs_soft_delete=1,jct_bs_status_id=l_status where jct_bs_jobref_no= l_jobRefNo and jct_bs_soft_delete=0;  
                        update jct_before_sketch_details set jct_bs_soft_delete=1 where jct_bs_jobref_no= l_jobRefNo and jct_bs_soft_delete=0; 
                        update jct_before_sketch_question set jct_bs_soft_delete=1,jct_bs_status_id=l_status  where jct_bs_jobref_no= l_jobRefNo and jct_bs_soft_delete=0;
              ELSEIF(l_decision='K') THEN 
                        -- update jct_action_plan set jct_action_plan_soft_delete=1,jct_status_id=l_status where jct_jobref_no= l_jobRefNo and jct_action_plan_soft_delete=0;  
                        -- update jct_after_sketch_header set jct_as_soft_delete=1 where jct_as_jobref_no= l_jobRefNo and jct_as_soft_delete=0 and jct_as_status_id=l_status_org; 
                        -- update jct_after_sketch_finalpage_details set jct_as_soft_delete=1 where jct_as_jobref_no= l_jobRefNo and jct_as_soft_delete=0 ;
                        -- update jct_after_sketch_pageone_details set jct_as_soft_delete=1 where jct_as_jobref_no= l_jobRefNo and jct_as_soft_delete=0;
                        -- update jct_before_sketch_header set jct_bs_soft_delete=1 where jct_bs_jobref_no= l_jobRefNo and jct_bs_soft_delete=0  and jct_bs_status_id =l_status_org;  
                        -- update jct_before_sketch_details set jct_bs_soft_delete=1 where jct_bs_jobref_no= l_jobRefNo and jct_bs_soft_delete=0 ;  
                        -- update jct_before_sketch_question set jct_bs_soft_delete=1  where jct_bs_jobref_no= l_jobRefNo and jct_bs_soft_delete=0  and jct_bs_status_id=l_status_org;
                        update jct_action_plan set jct_action_plan_soft_delete=1,jct_status_id=l_status where jct_jobref_no= l_jobRefNo and jct_action_plan_soft_delete=0;  
                        update jct_after_sketch_header set jct_as_soft_delete=1,jct_as_status_id=l_status where jct_as_jobref_no= l_jobRefNo and jct_as_soft_delete=0; 
                        update jct_after_sketch_finalpage_details set jct_as_soft_delete=1 where jct_as_jobref_no= l_jobRefNo and jct_as_soft_delete=0;
                        update jct_after_sketch_pageone_details set jct_as_soft_delete=1 where jct_as_jobref_no= l_jobRefNo and jct_as_soft_delete=0;
                        update jct_before_sketch_header set jct_bs_soft_delete=1,jct_bs_status_id=l_status where jct_bs_jobref_no= l_jobRefNo and jct_bs_soft_delete=0;  
                        update jct_before_sketch_details set jct_bs_soft_delete=1 where jct_bs_jobref_no= l_jobRefNo and jct_bs_soft_delete=0; 
                        update jct_before_sketch_question set jct_bs_soft_delete=1,jct_bs_status_id=l_status  where jct_bs_jobref_no= l_jobRefNo and jct_bs_soft_delete=0;        
                      IF(l_count>1) THEN
                        -- NEW IMPL - 
                        update jct_status_search set jct_soft_delete=1 where jct_jobref_no= l_jobRefNo and jct_soft_delete=0 and jct_status_decision=1;
                        update jct_status_search set jct_soft_delete=1 where jct_jobref_no= l_jobRefNo and jct_soft_delete=0 and jct_status_decision=0;
                        update jct_status_search set jct_soft_delete=1 where jct_jobref_no= l_jobRefNo and jct_soft_delete=2 and jct_status_decision=1;                   
                        update jct_status_search set jct_soft_delete=1 where jct_jobref_no= l_jobRefNo and jct_soft_delete=2 and jct_status_decision=0;
                        -- NEW IMPL
                        SELECT max(jct_status_search_id) into l_max_status_id
                          FROM jct_status_search where jct_soft_delete=1 and jct_status_id=l_status and  jct_jobref_no= l_jobRefNo and jct_status_decision=0;                       
                        update jct_status_search set jct_soft_delete=0 where jct_status_search_id=l_max_status_id;
                      ELSE IF(l_count=1) THEN
                        update jct_status_search set jct_soft_delete=1 where jct_jobref_no= l_jobRefNo and jct_soft_delete=0 and jct_status_decision=1;
                        update jct_status_search set jct_soft_delete=1 where jct_jobref_no= l_jobRefNo and jct_soft_delete=0 and jct_status_decision=0;
                        update jct_status_search set jct_soft_delete=1 where jct_jobref_no= l_jobRefNo and jct_soft_delete=2 and jct_status_decision=1;                   
                        update jct_status_search set jct_soft_delete=1 where jct_jobref_no= l_jobRefNo and jct_soft_delete=2 and jct_status_decision=0;
                        -- NEW IMPL
                        SELECT max(jct_status_search_id) into l_max_status_id
                          FROM jct_status_search where jct_soft_delete=1 and jct_status_id=l_status and  jct_jobref_no= l_jobRefNo and jct_status_decision=0;                       
                        update jct_status_search set jct_soft_delete=0 where jct_status_search_id=l_max_status_id;
                        END IF;
                            END IF;
              
              
              ELSE  
                        update jct_action_plan set jct_status_id=l_status_org where jct_jobref_no= l_jobRefNo and jct_action_plan_soft_delete=0;  
                        update jct_after_sketch_header set jct_as_status_id=l_status_org where jct_as_jobref_no= l_jobRefNo and jct_as_soft_delete=0; 
                        update jct_before_sketch_header set jct_bs_status_id=l_status_org where jct_bs_jobref_no= l_jobRefNo and jct_bs_soft_delete=0;  
                        update jct_before_sketch_question set jct_bs_status_id=l_status_org  where jct_bs_jobref_no= l_jobRefNo and jct_bs_soft_delete=0;
                
                            END IF;
                
                                
                
                              
                
                               
                
      RETURN l_return;                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                       
END */$$
DELIMITER ;

/* Function  structure for function  `jct_update_facilitator_type_after_chk_realization` */

/*!50003 DROP FUNCTION IF EXISTS `jct_update_facilitator_type_after_chk_realization` */;
DELIMITER $$

/*!50003 CREATE DEFINER=`jobCraftingTool` FUNCTION `jct_update_facilitator_type_after_chk_realization`(headerId int(11), customerId Varchar(100), createdBy Varchar(100)) RETURNS int(11)
BEGIN
                DECLARE l_return INT DEFAULT  0;
                DECLARE i_headerId INT DEFAULT  0;
    DECLARE l_customerId Varchar(100);
                DECLARE i_createdBy Varchar(100);
    DECLARE i_temp_type Varchar(100);
    DECLARE i_temp_replace_type Varchar(100);
    DECLARE done INT DEFAULT  0;
                
    DECLARE jct_facilitator_details_cursor CURSOR FOR
    select dt.jct_fac_type from jct_facilitator_details dt where dt.jct_pmt_hdr_id = headerId and 
    dt.jct_user_customer_id = customerId and dt.jct_fac_created_by = createdBy;
    -- declare NOT FOUND handler
    DECLARE CONTINUE HANDLER FOR NOT FOUND SET done = 1;
    
    OPEN jct_facilitator_details_cursor;
    cursor_loop: LOOP
      FETCH jct_facilitator_details_cursor into i_temp_type;
    IF (i_temp_type='RW_CHECK') THEN
      set i_temp_replace_type = 'RW';
    END IF;
    IF (i_temp_type='AD_CHECK') THEN
      set i_temp_replace_type = 'AD';
    END IF;
    Update jct_facilitator_details set jct_fac_type = i_temp_replace_type where 
    jct_pmt_hdr_id = headerId and jct_user_customer_id = customerId and jct_fac_created_by = createdBy and 
    jct_fac_type = i_temp_type;
    IF done=1 THEN
      LEAVE cursor_loop;
    END IF; 
      
    END LOOP cursor_loop;
    CLOSE jct_facilitator_details_cursor;
    RETURN l_return;  
    END */$$
DELIMITER ;

/* Procedure structure for procedure `jct_before_sketch_to_after_sketch` */

/*!50003 DROP PROCEDURE IF EXISTS  `jct_before_sketch_to_after_sketch` */;

DELIMITER $$

/*!50003 CREATE DEFINER=`jobCraftingTool` PROCEDURE `jct_before_sketch_to_after_sketch`(IN jobRefNo Varchar(30))
BEGIN
  DECLARE l_jobRefNo Varchar(30);
  DECLARE l_jct_bs_task_id Int;
  DECLARE l_jct_bs_energy Int;
  DECLARE l_jct_bs_task_desc  Varchar(200);
  
  DECLARE l_jct_as_task_id Int;
  DECLARE l_jct_as_energy Int;
  DECLARE l_jct_as_task_desc Varchar(200);
  
  DECLARE  diff_id   Varchar(30); 
        DECLARE  diff_energy   Varchar(30); 
  DECLARE  diff_desc   Varchar(100); 
  DECLARE done          INT DEFAULT  0;
  DECLARE innerLoopCount         INT DEFAULT  0;
  DECLARE innerLoopCountTo         INT DEFAULT  0;
  DECLARE l_length         INT DEFAULT  0;
  DECLARE ll_jct_bs_task_id Int;
  DECLARE ll_jct_bs_energy Int;
  DECLARE ll_jct_bs_task_desc  Varchar(200);
  
  DECLARE ll_jct_as_task_id Int;
  DECLARE ll_jct_as_energy Int;
  DECLARE ll_jct_as_task_desc Varchar(200);
  
  DECLARE  ll_diff_id   Varchar(30); 
        DECLARE  ll_diff_energy   Varchar(30); 
  DECLARE  ll_diff_desc   Varchar(100); 
  DECLARE  l_status   Varchar(100) DEFAULT ''; 
  DECLARE beforeSketch CURSOR FOR 
    SELECT jct_bs_task_id,jct_bs_energy,rtrim(ltrim(jct_bs_task_desc)) FROM jct_before_sketch_details 
          where jct_bs_soft_delete=0 and  jct_bs_jobref_no= jobRefNo;
  
  DECLARE afterSketch CURSOR FOR 
    SELECT jct_as_task_id,jct_as_task_energy,rtrim(ltrim(jct_as_task_desc)) FROM jct_after_sketch_finalpage_details 
          where  jct_as_jobref_no= jobRefNo and jct_as_soft_delete=0 and jct_as_task_id=l_jct_bs_task_id;
  DECLARE afterSketchTo CURSOR FOR 
    SELECT jct_as_task_id,jct_as_task_energy,rtrim(ltrim(jct_as_task_desc)) FROM jct_after_sketch_finalpage_details 
          where  jct_as_jobref_no= jobRefNo and jct_as_soft_delete=0 and (jct_as_element_id is null or jct_as_element_id ='');
  DECLARE beforeSketchTo CURSOR FOR 
    SELECT jct_bs_task_id,jct_bs_energy,rtrim(ltrim(jct_bs_task_desc)) FROM jct_before_sketch_details 
          where jct_bs_soft_delete=0 and  jct_bs_jobref_no= jobRefNo and jct_bs_task_id=ll_jct_as_task_id; 
  DECLARE CONTINUE HANDLER FOR NOT FOUND SET done = 1;
  
  DECLARE EXIT HANDLER FOR SQLEXCEPTION
    BEGIN
      ROLLBACK;
      SET l_status= 'Error occurred';
    END; 
  SET l_jobRefNo=jobRefNo;
  OPEN beforeSketch;
  cursor_loop: LOOP
    FETCH beforeSketch into l_jct_bs_task_id,l_jct_bs_energy,l_jct_bs_task_desc;
    IF done=1 THEN
      LEAVE cursor_loop;
    END IF;
    OPEN afterSketch;
    inner_cursor_loop: LOOP
      
      FETCH afterSketch into l_jct_as_task_id,l_jct_as_energy,l_jct_as_task_desc;
      IF done=1 THEN
        LEAVE inner_cursor_loop;
      END IF;
      set innerLoopCount=innerLoopCount+1;
    if((l_jct_bs_task_id=l_jct_as_task_id) and (l_jct_bs_energy=l_jct_as_energy) and 
        (l_jct_bs_task_desc=l_jct_as_task_desc)) then
      
      SET diff_energy=l_jct_bs_energy-l_jct_as_energy;
      SET l_length= locate('(', l_jct_as_task_desc);
      SET diff_desc= substring(l_jct_as_task_desc, l_length);
       insert  into jct_bs_to_as (`jct_bs_to_as_id`,`jct_jobref_no`,`jct_task_id`,`jct_bs_energy`,`jct_bs_time`,`jct_bs_task_desc`,`jct_as_energy`,`jct_as_time`,`jct_as_task_desc`,`jct_diff_status`,`jct_diff_energy`,`jct_diff_time`,`jct_diff_task_desc`,`version`) 
        Values ( jct_generate_id ('jct_bs_to_as'),l_jobRefNo,l_jct_as_task_id,l_jct_bs_energy,'0',l_jct_bs_task_desc,l_jct_as_energy,'0',l_jct_as_task_desc,'No Change',diff_energy,'0',diff_desc,0);
      LEAVE inner_cursor_loop;      
    elseif( (l_jct_bs_task_id=l_jct_as_task_id) and  (l_jct_bs_task_desc=l_jct_as_task_desc)) then 
      SET diff_energy=l_jct_bs_energy-l_jct_as_energy;
      SET l_length= locate('(', l_jct_as_task_desc);
      SET diff_desc= substring(l_jct_as_task_desc, l_length);
       insert  into jct_bs_to_as (`jct_bs_to_as_id`,`jct_jobref_no`,`jct_task_id`,`jct_bs_energy`,`jct_bs_time`,`jct_bs_task_desc`,`jct_as_energy`,`jct_as_time`,`jct_as_task_desc`,`jct_diff_status`,`jct_diff_energy`,`jct_diff_time`,`jct_diff_task_desc`,`version`) 
        Values ( jct_generate_id ('jct_bs_to_as'),l_jobRefNo,l_jct_as_task_id,l_jct_bs_energy,'0',l_jct_bs_task_desc,l_jct_as_energy,'0',l_jct_as_task_desc,'Modifide Task',diff_energy,'0',diff_desc,0);
      LEAVE inner_cursor_loop;
    elseif((l_jct_bs_task_id=l_jct_as_task_id) and (l_jct_bs_energy=l_jct_as_energy)) then 
      SET diff_energy=l_jct_bs_energy-l_jct_as_energy;
      SET l_length= locate('(', l_jct_as_task_desc);
      SET diff_desc= substring(l_jct_as_task_desc, l_length);
      insert  into jct_bs_to_as (`jct_bs_to_as_id`,`jct_jobref_no`,`jct_task_id`,`jct_bs_energy`,`jct_bs_time`,`jct_bs_task_desc`,`jct_as_energy`,`jct_as_time`,`jct_as_task_desc`,`jct_diff_status`,`jct_diff_energy`,`jct_diff_time`,`jct_diff_task_desc`,`version`) 
        Values ( jct_generate_id ('jct_bs_to_as'),l_jobRefNo,l_jct_as_task_id,l_jct_bs_energy,'0',l_jct_bs_task_desc,l_jct_as_energy,'0',l_jct_as_task_desc,'Modifide Task',diff_energy,'0',diff_desc,0);
      LEAVE inner_cursor_loop;
    elseif((l_jct_bs_task_desc=l_jct_as_task_desc) and (l_jct_bs_energy=l_jct_as_energy)) then 
      SET diff_energy=l_jct_bs_energy-l_jct_as_energy;
      SET l_length= locate('(', l_jct_as_task_desc);
      SET diff_desc= substring(l_jct_as_task_desc, l_length);
      insert  into jct_bs_to_as (`jct_bs_to_as_id`,`jct_jobref_no`,`jct_task_id`,`jct_bs_energy`,`jct_bs_time`,`jct_bs_task_desc`,`jct_as_energy`,`jct_as_time`,`jct_as_task_desc`,`jct_diff_status`,`jct_diff_energy`,`jct_diff_time`,`jct_diff_task_desc`,`version`) 
        Values ( jct_generate_id ('jct_bs_to_as'),l_jobRefNo,l_jct_as_task_id,l_jct_bs_energy,'0',l_jct_bs_task_desc,l_jct_as_energy,'0',l_jct_as_task_desc,'Modifide Task',diff_energy,'0',diff_desc,0);
      LEAVE inner_cursor_loop;
    else 
      SET diff_energy=l_jct_bs_energy-l_jct_as_energy;
      SET l_length= locate('(', l_jct_as_task_desc);
      SET diff_desc= substring(l_jct_as_task_desc, l_length);
      insert  into jct_bs_to_as (`jct_bs_to_as_id`,`jct_jobref_no`,`jct_task_id`,`jct_bs_energy`,`jct_bs_time`,`jct_bs_task_desc`,`jct_as_energy`,`jct_as_time`,`jct_as_task_desc`,`jct_diff_status`,`jct_diff_energy`,`jct_diff_time`,`jct_diff_task_desc`,`version`) 
        Values ( jct_generate_id ('jct_bs_to_as'),l_jobRefNo,l_jct_as_task_id,l_jct_bs_energy,'0',l_jct_bs_task_desc,l_jct_as_energy,'0',l_jct_as_task_desc,'Modifide Task',diff_energy,'0',diff_desc,0);
      LEAVE inner_cursor_loop;      
    end if;
      
    END LOOP inner_cursor_loop;
    CLOSE afterSketch;
    SET done=0;
    select 'aaaaabbc';
    IF (innerLoopCount=0) THEN
    insert  into jct_bs_to_as (`jct_bs_to_as_id`,`jct_jobref_no`,`jct_task_id`,`jct_bs_energy`,`jct_bs_time`,`jct_bs_task_desc`,`jct_as_energy`,`jct_as_time`,`jct_as_task_desc`,`jct_diff_status`,`jct_diff_energy`,`jct_diff_time`,`jct_diff_task_desc`,`version`) 
        Values ( jct_generate_id ('jct_bs_to_as'),l_jobRefNo,l_jct_bs_task_id,l_jct_bs_energy,'0',l_jct_bs_task_desc,l_jct_as_energy,'0',l_jct_as_task_desc,'Deleted Task',diff_energy,'0',diff_desc,0);
    else
      set innerLoopCount=0;
    END IF;
    
    
  END LOOP cursor_loop;
  CLOSE beforeSketch;
  SET done=0;
  OPEN afterSketchTo;
  outer_cursor_loop: LOOP
  FETCH afterSketchTo into ll_jct_as_task_id,ll_jct_as_energy,ll_jct_as_task_desc;
    IF done=1 THEN
      LEAVE outer_cursor_loop;
    END IF;
    OPEN beforeSketchTo;
    inner_cursor_loop: LOOP
    FETCH beforeSketchTo into ll_jct_bs_task_id,ll_jct_bs_energy,ll_jct_bs_task_desc;
      IF done=1 THEN
        LEAVE inner_cursor_loop;
      END IF;
      set innerLoopCountTo=innerLoopCountTo+1;
    END LOOP inner_cursor_loop;
    CLOSE beforeSketchTo;
    SET done=0;
select 'aaaaa';
    IF (innerLoopCountTo=0) THEN
    insert  into jct_bs_to_as (`jct_bs_to_as_id`,`jct_jobref_no`,`jct_task_id`,`jct_bs_energy`,`jct_bs_time`,`jct_bs_task_desc`,`jct_as_energy`,`jct_as_time`,`jct_as_task_desc`,`jct_diff_status`,`jct_diff_energy`,`jct_diff_time`,`jct_diff_task_desc`,`version`) 
        Values ( jct_generate_id ('jct_bs_to_as'),l_jobRefNo,ll_jct_as_task_id,'','0','',ll_jct_as_energy,'0',ll_jct_as_task_desc,'New Task','','0','',0);
    else
      set innerLoopCountTo=0;
    END IF;
  
        END LOOP outer_cursor_loop;
  CLOSE afterSketchTo;
  
END */$$
DELIMITER ;

/* Procedure structure for procedure `jct_create_user` */

/*!50003 DROP PROCEDURE IF EXISTS  `jct_create_user` */;

DELIMITER $$

/*!50003 CREATE DEFINER=`jobCraftingTool` PROCEDURE `jct_create_user`(userNo INT(10), userType INT(1))
BEGIN
    DECLARE userNo_1 INT(10);
  DECLARE uName varchar(50);
  DECLARE i INT(10);
  DECLARE custId varchar(100);
  DECLARE custId_DB varchar(100);
  DECLARE userId INT(10);
  DECLARE userDtlsId INT(10); 
  
  SET i = 1;
  SET userId = 500;
  SET userDtlsId = 200;
  IF userType = 1 THEN          # 1 for user, 3 for facilitator
    SET custId = '991505201500000000';  # User starts with 99
  ELSE
    SET custId = '981505201500000000';  # User starts with 98
  END IF;
  
  while i <= userNo do
    SET uName = CONCAT('user',i,'@user.com');
    SET custId_DB = CONCAT(custId,i);
    
    insert into jct_user (jct_user_id,jct_user_name,jct_password,jct_user_email,jct_active_yn,jct_version,jct_created_ts,jct_lastmodified_by,jct_created_by,lastmodified_ts,jct_role_id,jct_user_soft_delete,jct_account_expiration_date,jct_user_customer_id) 
      values (userId,uName,'y?5:2pb`ady?5:2pb`ad',uName,3,0,'2015-05-15 10:11:34','jctadmin@interrait.com','jctadmin@interrait.com','2015-05-15 10:11:34','1','0','2015-06-15 10:11:34',custId_DB);
  
    insert into jct_user_details (jct_user_details_id,jct_user_details_group_id,jct_user_details_group_name,jct_user_details_profile_id,jct_user_details_profile_name,jct_user_details_admin_id,jct_user_id,jct_user_details_soft_delete,jct_user_onet_occupation) 
      values (userDtlsId,1,'Default Profile',2,'Default Profile','1',userId,0,'13-2071.00');    
    
    SET userId = userId + 1;
    SET userDtlsId = userDtlsId + 1;
    SET i = i + 1;
  end while;
  commit; 
END */$$
DELIMITER ;

/*Table structure for table `jct_action_plan_view` */

DROP TABLE IF EXISTS `jct_action_plan_view`;

/*!50001 DROP VIEW IF EXISTS `jct_action_plan_view` */;
/*!50001 DROP TABLE IF EXISTS `jct_action_plan_view` */;

/*!50001 CREATE TABLE  `jct_action_plan_view`(
 `jct_as_created_by` varchar(50) ,
 `jct_user_occupation` varchar(100) ,
 `jct_user_function_group` varchar(50) ,
 `jct_user_tenure` varchar(20) ,
 `jct_user_levels` varchar(50) ,
 `jct_user_region` varchar(80) ,
 `jct_first_name` varchar(50) ,
 `jct_last_name` varchar(50) ,
 `jct_as_answar_desc` varchar(2000) ,
 `jct_as_question_desc` varchar(200) ,
 `jct_as_question_sub_desc` varchar(200) ,
 `jct_action_plan_soft_delete` int(1) ,
 `jct_status_id` int(11) ,
 `jct_as_header_id` int(11) 
)*/;

/*Table structure for table `jct_after_sketch_view` */

DROP TABLE IF EXISTS `jct_after_sketch_view`;

/*!50001 DROP VIEW IF EXISTS `jct_after_sketch_view` */;
/*!50001 DROP TABLE IF EXISTS `jct_after_sketch_view` */;

/*!50001 CREATE TABLE  `jct_after_sketch_view`(
 `jct_user_function_group` varchar(50) ,
 `jct_user_occupation` varchar(100) ,
 `jct_user_tenure` varchar(20) ,
 `jct_user_levels` varchar(50) ,
 `jct_user_region` varchar(80) ,
 `jct_first_name` varchar(50) ,
 `jct_last_name` varchar(50) ,
 `jct_as_created_by` varchar(50) ,
 `jct_as_finalpage_time_spent` int(11) ,
 `jct_as_task_energy` int(11) ,
 `jct_as_task_desc` varchar(200) ,
 `jct_as_element_code` varchar(5) ,
 `jct_as_element_desc` varchar(50) ,
 `jct_as_role_desc` varchar(130) ,
 `jct_as_soft_delete` tinyint(1) ,
 `jct_as_status_id` int(10) ,
 `jct_as_header_id` int(11) 
)*/;

/*Table structure for table `jct_before_sketch_to_after_sketch_view` */

DROP TABLE IF EXISTS `jct_before_sketch_to_after_sketch_view`;

/*!50001 DROP VIEW IF EXISTS `jct_before_sketch_to_after_sketch_view` */;
/*!50001 DROP TABLE IF EXISTS `jct_before_sketch_to_after_sketch_view` */;

/*!50001 CREATE TABLE  `jct_before_sketch_to_after_sketch_view`(
 `jct_bs_to_as_created_by` varchar(50) ,
 `jct_user_occupation` varchar(100) ,
 `jct_user_function_group` varchar(50) ,
 `jct_user_tenure` varchar(20) ,
 `jct_user_levels` varchar(50) ,
 `jct_user_region` varchar(80) ,
 `jct_first_name` varchar(50) ,
 `jct_last_name` varchar(50) ,
 `jct_soft_delete` int(1) ,
 `jct_bs_task_desc` varchar(200) ,
 `jct_as_task_desc` varchar(200) ,
 `jct_diff_task_desc` varchar(200) ,
 `jct_bs_energy` int(11) ,
 `jct_as_energy` int(11) ,
 `jct_diff_energy` varchar(11) ,
 `jct_diff_status` varchar(20) 
)*/;

/*Table structure for table `jct_before_sketch_view` */

DROP TABLE IF EXISTS `jct_before_sketch_view`;

/*!50001 DROP VIEW IF EXISTS `jct_before_sketch_view` */;
/*!50001 DROP TABLE IF EXISTS `jct_before_sketch_view` */;

/*!50001 CREATE TABLE  `jct_before_sketch_view`(
 `jct_bs_created_by` varchar(50) ,
 `jct_user_occupation` varchar(100) ,
 `jct_user_function_group` varchar(50) ,
 `jct_user_tenure` varchar(20) ,
 `jct_user_levels` varchar(50) ,
 `jct_user_region` varchar(80) ,
 `jct_first_name` varchar(50) ,
 `jct_last_name` varchar(50) ,
 `jct_bs_energy` int(11) ,
 `jct_bs_task_desc` varchar(200) ,
 `jct_bs_time_spent` int(11) ,
 `jct_bs_soft_delete` tinyint(1) ,
 `jct_bs_status_id` int(10) ,
 `jct_bs_header_id` int(11) 
)*/;

/*View structure for view jct_action_plan_view */

/*!50001 DROP TABLE IF EXISTS `jct_action_plan_view` */;
/*!50001 DROP VIEW IF EXISTS `jct_action_plan_view` */;

/*!50001 CREATE ALGORITHM=UNDEFINED DEFINER=`jobCraftingTool` SQL SECURITY DEFINER VIEW `jct_action_plan_view` AS select `actionplan`.`jct_as_created_by` AS `jct_as_created_by`,`userdetails`.`jct_user_onet_occupation` AS `jct_user_occupation`,`userdetails`.`jct_user_details_function_group` AS `jct_user_function_group`,`userdetails`.`jct_user_details_tenure` AS `jct_user_tenure`,`userdetails`.`jct_user_details_levels` AS `jct_user_levels`,`userdetails`.`jct_user_details_region` AS `jct_user_region`,`userdetails`.`jct_user_details_first_name` AS `jct_first_name`,`userdetails`.`jct_user_details_last_name` AS `jct_last_name`,`actionplan`.`jct_as_answar_desc` AS `jct_as_answar_desc`,`actionplan`.`jct_as_question_desc` AS `jct_as_question_desc`,`actionplan`.`jct_as_question_sub_desc` AS `jct_as_question_sub_desc`,`actionplan`.`jct_action_plan_soft_delete` AS `jct_action_plan_soft_delete`,`actionplan`.`jct_status_id` AS `jct_status_id`,`actionplan`.`jct_as_header_id` AS `jct_as_header_id` from ((`jct_user` `jctuser` join `jct_action_plan` `actionplan`) join `jct_user_details` `userdetails`) where ((`userdetails`.`jct_user_onet_occupation` <> 'NOT-ACTIVATED') and (`jctuser`.`jct_user_id` = `userdetails`.`jct_user_id`) and (((`actionplan`.`jct_action_plan_soft_delete` = 0) and (`jctuser`.`jct_user_id` = `userdetails`.`jct_user_id`) and (`actionplan`.`jct_status_id` = 5) and (`jctuser`.`jct_role_id` = 1)) or ((`actionplan`.`jct_action_plan_soft_delete` = 1) and (`jctuser`.`jct_user_id` = `userdetails`.`jct_user_id`) and (`actionplan`.`jct_status_id` = 5) and (`jctuser`.`jct_role_id` = 1))) and `jctuser`.`jct_user_email` in (select distinct `myhdr`.`jct_as_created_by` from `jct_action_plan` `myhdr` where (((`myhdr`.`jct_as_created_by` = `actionplan`.`jct_as_created_by`) and (`actionplan`.`jct_action_plan_soft_delete` = 0) and (`actionplan`.`jct_status_id` = 5) and (`jctuser`.`jct_role_id` = 1)) or ((`myhdr`.`jct_as_created_by` = `actionplan`.`jct_as_created_by`) and (`actionplan`.`jct_action_plan_soft_delete` = 1) and (`actionplan`.`jct_status_id` = 5) and (`jctuser`.`jct_role_id` = 1)))) and (not((`actionplan`.`jct_jobref_no` like '%A%')))) */;

/*View structure for view jct_after_sketch_view */

/*!50001 DROP TABLE IF EXISTS `jct_after_sketch_view` */;
/*!50001 DROP VIEW IF EXISTS `jct_after_sketch_view` */;

/*!50001 CREATE ALGORITHM=UNDEFINED DEFINER=`jobCraftingTool` SQL SECURITY DEFINER VIEW `jct_after_sketch_view` AS select `userdetails`.`jct_user_details_function_group` AS `jct_user_function_group`,`userdetails`.`jct_user_onet_occupation` AS `jct_user_occupation`,`userdetails`.`jct_user_details_tenure` AS `jct_user_tenure`,`userdetails`.`jct_user_details_levels` AS `jct_user_levels`,`userdetails`.`jct_user_details_region` AS `jct_user_region`,`userdetails`.`jct_user_details_first_name` AS `jct_first_name`,`userdetails`.`jct_user_details_last_name` AS `jct_last_name`,`header`.`jct_as_created_by` AS `jct_as_created_by`,`header`.`jct_as_finalpage_time_spent` AS `jct_as_finalpage_time_spent`,`dtl`.`jct_as_task_energy` AS `jct_as_task_energy`,`dtl`.`jct_as_task_desc` AS `jct_as_task_desc`,`dtl`.`jct_as_element_code` AS `jct_as_element_code`,`dtl`.`jct_as_element_desc` AS `jct_as_element_desc`,`dtl`.`jct_as_role_desc` AS `jct_as_role_desc`,`header`.`jct_as_soft_delete` AS `jct_as_soft_delete`,`header`.`jct_as_status_id` AS `jct_as_status_id`,`header`.`jct_as_header_id` AS `jct_as_header_id` from (((`jct_user` `jctuser` join `jct_after_sketch_header` `header`) join `jct_after_sketch_finalpage_details` `dtl`) join `jct_user_details` `userdetails`) where ((`userdetails`.`jct_user_onet_occupation` <> 'NOT-ACTIVATED') and (`header`.`jct_as_header_id` = `dtl`.`jct_as_header_id`) and (`jctuser`.`jct_user_id` = `userdetails`.`jct_user_id`) and (`header`.`jct_as_jobref_no` = `dtl`.`jct_as_jobref_no`) and (not((`header`.`jct_as_jobref_no` like '00000000ADPRV%'))) and `jctuser`.`jct_user_email` in (select distinct `jct_after_sketch_header`.`jct_as_created_by` from `jct_after_sketch_header` where ((`jct_after_sketch_header`.`jct_as_created_by` = `header`.`jct_as_created_by`) and (((`header`.`jct_as_soft_delete` = 0) and (`header`.`jct_as_status_id` = 4)) or ((`header`.`jct_as_soft_delete` = 0) and (`header`.`jct_as_status_id` = 5)) or ((`header`.`jct_as_soft_delete` = 1) and (`header`.`jct_as_status_id` = 5)))) order by `dtl`.`jct_as_role_desc`)) */;

/*View structure for view jct_before_sketch_to_after_sketch_view */

/*!50001 DROP TABLE IF EXISTS `jct_before_sketch_to_after_sketch_view` */;
/*!50001 DROP VIEW IF EXISTS `jct_before_sketch_to_after_sketch_view` */;

/*!50001 CREATE ALGORITHM=UNDEFINED DEFINER=`jobCraftingTool` SQL SECURITY DEFINER VIEW `jct_before_sketch_to_after_sketch_view` AS select `bstoas`.`jct_bs_to_as_created_by` AS `jct_bs_to_as_created_by`,`userdetails`.`jct_user_onet_occupation` AS `jct_user_occupation`,`userdetails`.`jct_user_details_function_group` AS `jct_user_function_group`,`userdetails`.`jct_user_details_tenure` AS `jct_user_tenure`,`userdetails`.`jct_user_details_levels` AS `jct_user_levels`,`userdetails`.`jct_user_details_region` AS `jct_user_region`,`userdetails`.`jct_user_details_first_name` AS `jct_first_name`,`userdetails`.`jct_user_details_last_name` AS `jct_last_name`,`bstoas`.`jct_soft_delete` AS `jct_soft_delete`,`bstoas`.`jct_bs_task_desc` AS `jct_bs_task_desc`,`bstoas`.`jct_as_task_desc` AS `jct_as_task_desc`,`bstoas`.`jct_diff_task_desc` AS `jct_diff_task_desc`,`bstoas`.`jct_bs_energy` AS `jct_bs_energy`,`bstoas`.`jct_as_energy` AS `jct_as_energy`,`bstoas`.`jct_diff_energy` AS `jct_diff_energy`,`bstoas`.`jct_diff_status` AS `jct_diff_status` from ((`jct_user` `jctuser` join `jct_bs_to_as` `bstoas`) join `jct_user_details` `userdetails`) where ((`userdetails`.`jct_user_onet_occupation` <> 'NOT-ACTIVATED') and (not((`bstoas`.`jct_jobref_no` like '00000000ADPRV%'))) and (`jctuser`.`jct_user_id` = `userdetails`.`jct_user_id`) and `jctuser`.`jct_user_email` in (select distinct `myhdr`.`jct_bs_to_as_created_by` from `jct_bs_to_as` `myhdr` where (`myhdr`.`jct_bs_to_as_created_by` = `bstoas`.`jct_bs_to_as_created_by`))) */;

/*View structure for view jct_before_sketch_view */

/*!50001 DROP TABLE IF EXISTS `jct_before_sketch_view` */;
/*!50001 DROP VIEW IF EXISTS `jct_before_sketch_view` */;

/*!50001 CREATE ALGORITHM=UNDEFINED DEFINER=`jobCraftingTool` SQL SECURITY DEFINER VIEW `jct_before_sketch_view` AS select distinct `header`.`jct_bs_created_by` AS `jct_bs_created_by`,`userdetails`.`jct_user_onet_occupation` AS `jct_user_occupation`,`userdetails`.`jct_user_details_function_group` AS `jct_user_function_group`,`userdetails`.`jct_user_details_tenure` AS `jct_user_tenure`,`userdetails`.`jct_user_details_levels` AS `jct_user_levels`,`userdetails`.`jct_user_details_region` AS `jct_user_region`,`userdetails`.`jct_user_details_first_name` AS `jct_first_name`,`userdetails`.`jct_user_details_last_name` AS `jct_last_name`,`dtl`.`jct_bs_energy` AS `jct_bs_energy`,`dtl`.`jct_bs_task_desc` AS `jct_bs_task_desc`,`header`.`jct_bs_time_spent` AS `jct_bs_time_spent`,`header`.`jct_bs_soft_delete` AS `jct_bs_soft_delete`,`header`.`jct_bs_status_id` AS `jct_bs_status_id`,`header`.`jct_bs_header_id` AS `jct_bs_header_id` from (((`jct_user` `jctuser` join `jct_before_sketch_header` `header`) join `jct_before_sketch_details` `dtl`) join `jct_user_details` `userdetails`) where ((`userdetails`.`jct_user_onet_occupation` <> 'NOT-ACTIVATED') and (`header`.`jct_bs_header_id` = `dtl`.`jct_bs_header_id`) and (`jctuser`.`jct_user_id` = `userdetails`.`jct_user_id`) and (`header`.`jct_bs_jobref_no` = `dtl`.`jct_bs_jobref_no`) and (not((`dtl`.`jct_bs_jobref_no` like '00000000ADPRV%'))) and (((`header`.`jct_bs_soft_delete` = 0) and (`header`.`jct_bs_status_id` = 4)) or ((`header`.`jct_bs_soft_delete` = 0) and (`header`.`jct_bs_status_id` = 5)) or ((`header`.`jct_bs_soft_delete` = 1) and (`header`.`jct_bs_status_id` = 5))) and ((`dtl`.`jct_bs_soft_delete` = 0) or (`dtl`.`jct_bs_soft_delete` = 1)) and `jctuser`.`jct_user_email` in (select distinct `myhdr`.`jct_bs_created_by` from `jct_before_sketch_header` `myhdr` where ((`myhdr`.`jct_bs_created_by` = `header`.`jct_bs_created_by`) and (((`myhdr`.`jct_bs_soft_delete` = 0) and (`myhdr`.`jct_bs_status_id` = 4)) or ((`myhdr`.`jct_bs_soft_delete` = 0) and (`myhdr`.`jct_bs_status_id` = 5)) or ((`myhdr`.`jct_bs_soft_delete` = 1) and (`myhdr`.`jct_bs_status_id` = 5)))))) */;

/*!40101 SET SQL_MODE=@OLD_SQL_MODE */;
/*!40014 SET FOREIGN_KEY_CHECKS=@OLD_FOREIGN_KEY_CHECKS */;
/*!40014 SET UNIQUE_CHECKS=@OLD_UNIQUE_CHECKS */;
/*!40111 SET SQL_NOTES=@OLD_SQL_NOTES */;
