package com.vmware.jct.common.utility;
/**
 * 
 * <p><b>Class name:</b> StatusConstants.java</p>
 * <p><b>Author:</b> InterraIT</p>
 * <p><b>Purpose:</b> This class acts as a constant class for application status messages in integer..
 * <p><b>Copyrights:</b>All rights reserved by Interra IT and should only be used for its internal application development.</p>
 * <p><b>Created Date:</b> 10/Apr/2014</p>
 * <p><b>Revision History:</b>
 * 	<li></li>
 * </p>
 */
public class StatusConstants {
	public static final Integer STATUS_SUCCESS_WITH_RESULT 				= new Integer(200);
	public static final Integer STATUS_SUCCESS_WITH_NO_RESULT 			= new Integer(201);
	public static final Integer STATUS_FAILURE 							= new Integer(500);
	public static final Integer STATUS_ACTIVE 							= new Integer(1);
	public static final Integer STATUS_INACTIVE 						= new Integer(2);
	public static final Integer USER_ONLY_REGISTERED 					= new Integer(200);
	public static final Integer USER_REGISTERED_MAIL_SENT 				= new Integer(201);
	public static final Integer USER_REGISTERED_MAIL_BY_CRON 			= new Integer(202);
	public static final Integer USER_REGISTERED_MAIL_SEND_ERROR 		= new Integer(203);
	public static final Integer USER_NOT_REGISTERED 					= new Integer(204);
	public static final Integer USER_EMAIL_NOT_SENT 					= new Integer(0);
	public static final Integer USER_EMAIL_SENT 						= new Integer(1);
	public static final Integer USER_REGISTERED_MAIL_SEND_ERROR_ADMIN 	= new Integer(210);
	public static final Integer USER_REGISTERED_MAIL_SEND_ERROR_USERS 	= new Integer(211);
	public static final Integer INVALID_EMAIL_ID_FOUND_CSV_UPLD		 	= new Integer(212);
	public static final Integer ALREADY_REG_EMAIL_CSV_UPLD		 		= new Integer(213);
	public static final Integer TOTAL_CSV_UPLD_FAILURE			 		= new Integer(214);
	public static final Integer ALL_CORRECT						 		= new Integer(215);
	public static final Integer ACC_DEACTIVATION_MAIL_SEND_ERROR 		= new Integer(216);
	public static final Integer USER_DELETION_SUCCESS 					= new Integer(400);
	public static final Integer USER_DELETION_FAILURE 					= new Integer(401);
	public static final Integer USER_ACTIVATION_SUCCESS 				= new Integer(400);
	public static final Integer USER_FAILURE_FAILURE 					= new Integer(401);
	public static final Integer USER_UPDATION_SUCCESS 					= new Integer(400);
	public static final Integer USER_UPDATION_FAILURE 					= new Integer(401);
	public static final Integer USER_PASSWORD_RESET_SUCCESS 			= new Integer(200);
	public static final Integer USER_PASSWORD_RESET_FAILURE 			= new Integer(201);
	public static final Integer USER_PASSWORD_RESET_CANNOT_PROCEED 		= new Integer(222);
	public static final Integer ALREADY_EXISTS 							= new Integer(600);
	public static final Integer DOES_NOT_EXIST 							= new Integer(601);
	public static final Integer MAX_LIMIT 								= new Integer(602);
	public static final Integer USER_PASSWORD_RESET_AFTER_COMPETION 	= new Integer(5);
	public static final Integer USER_PASSWORD_RESET_BEFORE_COMPETION	= new Integer(4);
	public static final Integer STATUS_SUCCESS 							= new Integer(200);
	public static final Integer STATUS_USER_EXIST						= new Integer(204);
	public static final Integer STATUS_GEN_USER_EXIST					= new Integer(205);
	public static final Integer STATUS_FAILED 							= new Integer(201);
	public static final Integer TABLE_OR_DATA_RELATED_ERROR			    = new Integer(999);
	public static final Integer NO_FACILITATOR_FOR_ADD			    	= new Integer(701);
	public static final Integer NO_FACILITATOR_FOR_RENEW			    = new Integer(702);
	
	public static final Integer ICON_STATUS_SUCCESS 					= new Integer(1);
	public static final Integer ICON_STATUS_FAILURE					    = new Integer(2);
	public static final Integer ICON_STATUS_WARNING				    	= new Integer(3);
}
