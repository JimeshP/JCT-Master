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

	public static final int CREATED 							= 1;
	public static final int ACTIVATED 							= 2;
	public static final int CHECK_PAYMENT_NOT_ACTIVATED			= 6;
	
	public static final int FACI_ACCOUNT_JUST_CREATED			= 10;
	public static final int FACI_PASSWORD_RESET			    	= 11;
	public static final String FACI_ACCOUNT_JUST_CREATED_STRING	= "10";
	public static final String FACI_PASSWORD_RESET_STRING		= "11";
	public static final String PASSWORD_RESET_FAILED 			= "400";
	
	public static final String SUCCESSFUL 						= "200";
	public static final String USER_FOUND 						= "402";
	public static final String USER_NOT_FOUND 					= "404";
	public static final String USUER_CHEQUE_NOT_CLEARED		 	= "990";
	
	
	
	public static final String DATABASE_NOT_FOUND 				= "999";
	public static final String FACILITATOR_TRANSFERRED			= "350";
	public static final String DATA_NOT_FOUND 					= "100";
	public static final int ENTRY_SOFT_DELETED 					= 1;
	public static final int ENTRY_NOT_DELETED 					= 0;
	public static final String JOB_ATTRIBUTE_STRENGTH 			= "STR";
	public static final String JOB_ATTRIBUTE_PASSION 			= "PAS";
	public static final String JOB_ATTRIBUTE_VALUE 				= "VAL";
	public static final String EMAIL_PATTERN 					= 
			"^[_A-Za-z0-9-\\+]+(\\.[_A-Za-z0-9-]+)*@"
			+ "[A-Za-z0-9-]+(\\.[A-Za-z0-9]+)*(\\.[A-Za-z]{2,})$";
	public static final String NEW_FACILITATOR_CREATION			= "AD";
	public static final String RENEW_SUBSCRIBE_FACILITATOR		= "RN";
	public static final String FACILITATOR_RENEWAL				= "SD";
	
	public static final String INDIVIDUAL_USER_CUST_CODE		= "98";
	public static final String FACILITATOR_USER_CUST_CODE		= "99";
	
	public static final int FREE_TEXT							= new Integer(1);
	public static final int RADIO_BTN							= new Integer(2);
	public static final int DROP_DOWN							= new Integer(3);
	public static final int CHECK_BOX							= new Integer(4);
	
	public static final int GENERAL_USER						= new Integer(1);
	public static final int ADMIN_USER							= new Integer(2);
	public static final int FACILITATOR_USER					= new Integer(3);
	
	public static final String SHOPIFY_INDIVIDUAL_CUSTOMER		= "Individual";
	public static final String SHOPIFY_FACILITATOR_CUSTOMER		= "Facilitator";
	public static final String SHOPIFY_NEW_SUBSCRIPTION			= "New";
	public static final String SHOPIFY_RENEW_SUBSCRIPTION		= "Renew";
	public static final String SHOPIFY_ECOMMERCE				= "ECOMMERCE";
	
}
