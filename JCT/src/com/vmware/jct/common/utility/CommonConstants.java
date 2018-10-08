package com.vmware.jct.common.utility;
/**
 * 
 * <p><b>Class name:</b> CommonConstants.java</p>
 * <p><b>Author:</b> InterraIT</p>
 * <p><b>Purpose:</b> This class acts as a constant class for application status messages in string..
 * <p><b>Copyrights:</b>All rights reserved by Interra IT and should only be used for its internal application development.</p>
 * <p><b>Created Date:</b> 10/Apr/2014</p>
 * <p><b>Revision History:</b>
 * 	<li></li>
 * </p>
 */
public class CommonConstants {

	/************************* USER STATUS ************************/
	public static final int CREATED 								= 1;					// WHEN ADMIN CREATES THE USER
	public static final int ACTIVATED_ONLY_EMAIL 					= 2;		// WHEN USER VALIDATES HIMSELF AND CHANGES HIS PASSWORD
	public static final int ACTIVATION_COMPLETE 					= 3;		// WHEN USER ENTERS ALL REQUIRED DATA
	public static final int DEACTIVATED 							= 4;
	public static final int USER_PASSWORD_RESET_AFTER_COMPETION 	= new Integer(5);
	public static final int CHECK_PAYMENT_NOT_CLEARED			    = 6;		//	PAYMENT_CHEQUE NOT CLEARED BY ADMIN
	public static final int USER_EXIST_BUT_DISABLED			    	= 7;		//	USER_EXIST_BUT_DISABLED	
	
	
	/************************* APPLICATION STATUS **************************/
	public static final String SUCCESSFUL 							= "200";
	public static final String PASSWORD_RESET_NOT_COMPLETE 			= "874";
	public static final String CAPTCHA_VALIDATION_FAILED 			= "555";
	public static final String USER_REGISTRATION_FAILED 			= "201";
	public static final String USER_NOT_ACTIVATED 					= "201";
	public static final String USER_DISABLED 						= "202";
	public static final String USER_INACTIVATED_BY_ADMIN 			= "987";
	public static final String USER_PASSWORD_RESET_STATUS 			= "222";
	public static final String USER_NOT_FOUND 						= "404";
	public static final String USER_VALIDATION_FAILED 				= "224";
	public static final String PAYMENT_CHEQUE_NOT_CLEARED		 	= "990";
	public static final String FORGOT_PASSWORD_LOCKED 				= "989";
	public static final String PASSWORD_RESET_FAILED 				= "400";
	public static final String UNKNOWN_ERROR 						= "000";
	public static final String DATABASE_NOT_FOUND 					= "999";
	public static final String MAIL_SEND_ERROR 						= "888";
	public static final String DATA_NOT_FOUND 						= "100";
	public static final String FAILURE 								= "500";
	public static final String USER_ACCOUNT_EXPIRED 				= "776";
	public static final String ONET_NOT_FOUND		 				= "865";
	
	
	/************************ TASK STATUS *************************/
	public static final int TASK_PENDING 							= 4;
	public static final int TASK_COMPLETED 							= 5;
	
	/************************ PAGE SEQUENCE *************************/
	public static final int PAGE_SEQUENCE_BEFORE_SKETCH 			= 1;
	public static final int PAGE_SEQUENCE_QUESTIONAIR 				= 2;
	public static final int PAGE_SEQUENCE_AFTER_SKETCH_ONE 			= 3;
	public static final int PAGE_SEQUENCE_AFTER_SKETCH_TWO 			= 4;
	public static final int PAGE_SEQUENCE_ACTION_PLAN 				= 5;
	
	/************************ SOFT DELETE *************************/
	public static final int ENTRY_SOFT_DELETED 						= 1;
	public static final int ENTRY_NOT_DELETED 						= 0;
	public static final int ENTRY_FOR_PREVIEW_TOOL 					= 2;
	
	/************************ Job attribute  *************************/
	
	public static final String JOB_ATTRIBUTE_STRENGTH 				= "STR";
	public static final String JOB_ATTRIBUTE_PASSION 				= "PAS";
	public static final String JOB_ATTRIBUTE_VALUE 					= "VAL";
	
/************************ Before Sketch and After Sketch Defination  *************************/
	
	public static final String BEFORE_SKETCH 						= "BS";
	public static final String AFTER_SKETCH 						= "AS";
	
		/************************ Job Level  *************************/
	public static final String JOB_LEVEL_ALL 						= "All";
	
	public static final int FREE_TEXT								= new Integer(1);
	public static final int RADIO_BTN								= new Integer(2);
	public static final int DROP_DOWN								= new Integer(3);
	public static final int CHECK_BOX								= new Integer(4);
	public static final int GENERAL_USER							= new Integer(1);
	public static final int ADMIN_USER								= new Integer(2);
	public static final int FACILITATOR_USER						= new Integer(3);
	
}
