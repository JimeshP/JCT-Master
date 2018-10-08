package com.vmware.jct.controller;
/**
 * 
 * <p><b>Class name:</b> BasicController.java</p>
 * <p><b>Author:</b> InterraIT</p>
 * <p><b>Purpose:</b> This class acts as constant file for all Request Mapping.
 * <p><b>Description:</b> This class acts as constant file for all Request Mapping.. </p>
 * <p><b>Copyrights:</b> 	All rights reserved by InterraIT and should only be used for its internal application development.</p>
 * <p><b>Created Date:</b> 10/Apr/2014</p>
 * <p><b>Revision History:</b>
 * 	<li></li>
 * </p>
 */
public class BasicController {
	public static final String ACTION_AUTHENTICATE 				= "/authenticate";
	public static final String ACTION_FACILITATOR_AUTHENTICATE  = "/facilitatorAuthenticate";
	public static final String ACTION_CHECK_USER 				= "/checkUser";
	public static final String ACTION_REGISTER 					= "/register";
	public static final String ACTION_SAVE_SURVEY				= "/saveSurveyQuestions";
	public static final String ACTION_SEARCH_OCCUPTN_LIST		= "/searchOccupationList";
	public static final String ACTION_RESET_PASSWORD 			= "/resetPassword";
	public static final String ACTION_FORGOT_PASSWORD 			= "/forgotPassword";
	public static final String ACTION_LOGOUT 					= "/logout";
	public static final String ACTION_UPDATE_LOGIN_INFO_ROW 	=  "/updateLoginInfoRowData";
	public static final String ACTION_SIGNUP_MASTER 			= "/populateSignupMasterData";
	public static final String ACTION_POPULATE_USER_DETAILS 	= "/populateUserDetails";
	public static final String ACTION_UPDATE_DEMO_INFO 			= "/updateDemographicalInfo";
	public static final String ACTION_DECIDE_PAGE 				= "/decidePage";
	public static final String ACTION_FETCH_QUESTIONNAIRE 		= "/fetchQuestionnaire";
	public static final String ACTION_SAVE_ANSWERS 				= "/saveAnswers";
	public static final String ACTION_GO_PREVIOUS 				= "/goPrevious";
	public static final String ACTION_SAVE_AFTER_SKETCH_1 		= "/saveAfterSketchBasic";
	public static final String ACTION_SAVE_AFTER_SKETCH_2 		= "/saveAfterSketchDiagram";
	public static final String ACTION_SEARCH_BEFORE_SKETCH 		= "/searchBeforeSketch";
	public static final String ACTION_SEARCH_ALL_DIAGRAMS 		= "/searchAllDiagrams";
	public static final String ACTION_SEARCH_AFTER_SKETCH 		= "/searchAfterSketch";
	public static final String ACTION_POPULATE_MY_ACCOUNT 		= "/populateMyAccount";
	public static final String ACTION_IDENTIFY_SEARCH_TABS 		= "/getSearchTabsTobeViewed";
	public static final String ACTION_GET_SELECTED_DIAGRAMS 	= "/getSelectedDiagrams";
	public static final String ACTION_GET_PREV_DIAGRAM 			= "/getPrevDiag";
	public static final String ACTION_FETCH_ACTION_PLAN 		= "/fetchActionPlan";
	public static final String ACTION_SAVE_ACTION_PLAN 			= "/saveActionPlan";
	public static final String ACTION_ADD_DIAG_WITHOUT_EDITS 	= "/addDiagWithoutEdits";
	public static final String ACTION_ADD_DIAG_WITH_EDITS 		= "/addEdits";
	public static final String ACTION_DISABLE_STATUS 			= "/disableStatus";
	public static final String ACTION_KEEP_PREVIOUS 			= "/keepPrevious";
	public static final String ACTION_STORE_BEFORE_SKETCH 		= "/storeBeforeSketch";
	public static final String ACTION_POPULATE_INSTRUCTION 		= "/populateInstruction";
	public static final String ACTION_RESTART_EXCERSIZE 		= "/restartExcersize";
	public static final String ACTION_POPULATE_POPUP			= "/populatePopUp";
	public static final String ACTION_GET_SHOPIFY_LINK			= "/getShopifyLink";
	/***************	Terms & Conditions	 ***********/
	public static final String ACTION_GET_TERMS_AND_CONDITIONS	= "/fetchTnC";
	/***************	PDF search	 ***********/
	public static final String ACTION_GET_OLD_PDF				= "/fetchOldPdf";
	/***************	Update Time	 ***********/
	public static final String ACTION_SILENT_TIME_UPDATE		= "/silentTimeUpdate";
}
