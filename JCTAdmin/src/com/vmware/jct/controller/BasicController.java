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
	public static final String ACTION_AUTHENTICATE 									= "/authenticateAdmin";
	public static final String ACTION_CHECK_USER 									= "/checkUser";
	public static final String ACTION_REGISTER 										= "/register";
	public static final String ACTION_FORGOT_PASSWORD 								= "/forgotPassword";
	public static final String ACTION_LOGOUT 										= "/logout";
	public static final String ACTION_BEFORE_SKETCH_REPORT 							= "/populateBeforeSketchReport";
	public static final String ACTION_BEFORE_SKETCH_SELECTED_REPORT 				= "/populateBeforeSketchSelectedReport";
	public static final String ACTION_ACTION_PLAN_REPORT 							= "/populateActionPlanReport";
	public static final String ACTION_ACTION_PLAN_SELECTED_REPORT 					= "/populateActionPlanSelectedReport";
	public static final String ACTION_AFTER_SKETCH_REPORT 							= "/populateAfterSketchReport";
	public static final String ACTION_AFTER_SKETCH_SELECTED_REPORT 					= "/populateAfterSketchSelectedReport";
	public static final String ACTION_AFTER_SKETCH_MAPPING_DETAILS 					= "/populateAfterSketchMappingDetails";
	public static final String ACTION_BEFORE_SKETCH_AFTER_SKETCH_REPORT 			= "/populateBeforeSketchToAfterSketchReport";
	public static final String ACTION_BEFORE_SKETCH_AFTER_SKETCH_SELECTED_REPORT 	= "/populateBeforeSketchToAfterSketchSelectedReport";
	public static final String ACTION_MISC_REPORT 									= "/populateMiscReport";
	public static final String ACTION_MISC_DETAILED_REPORT 							= "/populateMiscDetailedReport";
	public static final String ACTION_USER_PROFILE_DROPDOWN 						= "/populateProfileDropDown";
	public static final String ACTION_POPULATE_EXTNG_USR_RPFL 						= "/populateExistingUserProfile";
	public static final String ACTION_SAVE_USER_PROFILE 							= "/saveUserProfile";
	public static final String VALIDATE_USER_PROFILE 								= "/validateExistence";
	public static final String VALIDATE_USER_GROUP 									= "/validateExistenceUserGrp";
	public static final String ACTION_POPULATE_EXTNG_USR_GRP 						= "/populateExistingUserGroup";
	public static final String ACTION_POPULATE_EXTNG_VIDEOS 						= "/populateExistingVideos";
	public static final String ACTION_POPULATE_VIDEOS_SELECTN 						= "/populateVideosBySelection";
	public static final String ACTION_POPULATE_TEXT_INSTRUCTN 						= "/populateTextInsBySelection";
	public static final String ACTION_SAVE_POPUP_INSTRUCTION						= "/savePopupInstruction";
	public static final String ACTION_SAVE_USER_GROUP 								= "/saveUserGroup";
	public static final String ACTION_UPDATE_USER_GROUP_STATUS 						= "/toogleGroupStatus";
	public static final String ACTION_UPDATE_USER_GROUP 							= "/updateUserGroup";
	public static final String ACTION_POPULATE_USER_PROFILE 						= "/populateUserProfile";
	public static final String ACTION_POPULATE_USER_GROUP 							= "/populateUserGroup";
	public static final String ACTION_SAVE_USER 									= "/saveUser";
	public static final String ACTION_SAVE_FACILITATOR_PAYMENT						= "/saveFacilitatorPayment";
	public static final String ACTION_SAVE_GEN_USER_MANUAL_PAYMENT					= "/saveGeneralUserPaymentManual";
	public static final String ACTION_SAVE_GEN_USER_CSV_PAYMENT						= "/saveGeneralUserPaymentCSV";
	public static final String ACTION_SAVE_MANUAL_USER 								= "/saveManualUser";	
	public static final String ACTION_POPULATE_EXISTING_USERS 						= "/populateExistingUsers";
	public static final String ACTION_POPULATE_EXISTING_USERS_GR_TYPE				= "/populateExistingUsersByUserTypeAndGroup";
	public static final String ACTION_POPULATE_ACTIVE_INACTIVE_USER 				= "/specificUserList";
	public static final String ACTION_POPULATE_ACTIVE_INACTIVE_USER_ADMIN			= "/specificUserListForAdmin";
	public static final String ACTION_DELETE_USER 									= "/deleteExistingUsers";
	public static final String ACTION_CHANGE_BATCH_USER_STATUS 						= "/changeUserStatusInBatch";
	public static final String ACTION_USER_STATUS 									= "/updateStatus";
	public static final String ACTION_RESET_PASSWORD 								= "/resetPassword";
	public static final String ACTION_RESET_PASSWORD_IN_BATCH 						= "/resetPasswordInBatch";
	public static final String ACTION_FETCH_EXISTING_USERNAME 						= "/fetchExistingUserName";
	public static final String ACTION_FETCH_INDIVIDUAL_RENEW_DETAILS				= "/fetchIndividualDetails";
	public static final String ACTION_FETCH_FACILITATOR_RENEW_DETAILS				= "/fetchFacilitatorDetails";
	public static final String ACTION_FETCH_EXISTING_USERNAME_ADMIN					= "/fetchExistingUserNameForAdmin";
	public static final String ACTION_POPULATE_EXTNG_REF_QTN 						= "/populateExistingRefQtn";
	public static final String ACTION_SAVE_REF_QTN 									= "/saveRefQtn";
	public static final String ACTION_UPDATE_REF_QTN 								= "/updateRefQtn";
	public static final String ACTION_POPULATE_USER_PROFILE_REF_QTN 				= "/populateUserProfileRefQtn";
	public static final String ACTION_DELETE_REF_QTN 								= "/deleteRefQtn";
	public static final String ACTION_SAVE_ACTION_PLAN 								= "/saveActionPlan";
	public static final String ACTION_UPDATE_ACTION_PLAN 							= "/updateActionPlan";
	public static final String ACTION_DELETE_ACTION_PLAN 							= "/deleteActionPlan";
	public static final String ACTION_POPULATE_EXTNG_FUNC_GRP 						= "/populateExistingFunctionGroup";
	public static final String ACTION_POPULATE_EXTNG_JOB_LEVEL 						= "/populateExistingJobLevel";
	public static final String ACTION_SAVE_FUNC_GRP 								= "/saveFunctionGroup";
	public static final String ACTION_SAVE_JOB_LEVEL 								= "/saveJobLevel";
	public static final String ACTION_POPULATE_EXTNG_REGION 						= "/populateExistingRegion";
	public static final String ACTION_SAVE_REGION 									= "/saveRegion";
	public static final String ACTION_UPDATE_REGION 								= "/updateRegion";
	public static final String ACTION_DELETE_REGION 								= "/deleteRegion";
	public static final String ACTION_POPULATE_EXTNG_MAPPING 						= "/populateExistingMapping";
	public static final String ACTION_SAVE_STRENGTH 								= "/saveStrength";
	public static final String ACTION_SAVE_VALUE 									= "/saveValue";
	public static final String ACTION_SAVE_PASSION 									= "/savePassion";
	public static final String ACTION_UPDATE_STRENGTH 								= "/updateStrength";
	public static final String ACTION_UPDATE_VALUE 									= "/updateValue";
	public static final String ACTION_UPDATE_PASSION 								= "/updatePassion";
	public static final String ACTION_DELETE_STRENGTH 								= "/deleteStrength";
	public static final String ACTION_DELETE_VALUE 									= "/deleteValue";
	public static final String ACTION_DELETE_PASSION 								= "/deletePassion";
	public static final String ACTION_POPULATE_EXTNG_USER_DATA 						= "/populateExistingUserData";
	public static final String ACTION_UPDATE_FIRST_NAME 							= "/updateFirstName";
	public static final String ACTION_UPDATE_LAST_NAME 								= "/updateLastName";
	public static final String ACTION_UPDATE_EMAIL_ID 								= "/updateEmailId";
	public static final String ACTION_UPDATE_PASSWORD 								= "/updatePassword";
	public static final String ACTION_POPULATE_USER_PROFILE_GENERAL 				= "/populateUserProfileGeneral";
	public static final String ACTION_SAVE_INSTRUCTION 								= "/saveInstruction";
	public static final String ACTION_POPULATE_EXTNG_INSTRUCTION 					= "/populateExistingInstruction";
	public static final String ACTION_UPDATE_INSTRUCTION 							= "/updateInstruction";
	public static final String ACTION_SAVE_APP_SETTINGS 							= "saveAppSettings";
	public static final String ACTION_FETCH_COLOR 									= "/fetchColor";
	public static final String ACTION_POPULATE_MAIN_QTN 							= "/populateMainQtn";
	public static final String ACTION_POPULATE_SUB_QTN 								= "/populateSubQtn";
	public static final String ACTION_SAVE_MAIN_QTN_ORDER 							= "/saveMainQtnOrder";
	public static final String ACTION_SAVE_SUB_QTN_ORDER 							= "/saveSubQtnOrder";
	public static final String ACTION_SAVE_FUNCTION_GROUP_ORDER 					= "/saveFuncGrpOrder";
	public static final String ACTION_POPULATE_EXTNG_FUNC_GRP_ORDER 				= "/populateFunctionGroupByOrder";
	public static final String ACTION_SAVE_JOB_LEVEL_ORDER 							= "/saveJobLevelOrder";
	public static final String ACTION_POPULATE_EXTNG_JOB_LEVEL_ORDER 				= "/populateJobLevelByOrder";
	public static final String ACTION_POPULATE_ATTRIBUTE_ORDER 						= "/populateMappingList";
	public static final String ACTION_SAVE_ATTRIBUTE_ORDER 							= "/saveAttributeOrder";
	public static final String ACTION_SAVE_REGION_ORDER 							= "/saveRegionOrder";
	/**
	 * THIS IS FOR JCT PUBLIC VERSION
	 */
	public static final String ACTION_SAVE_OCCUPATION_DATA 							= "/uploadFile";
	public static final String ACTION_SAVE_VIDEO_DATA 								= "/uploadVideo";
	public static final String ACTION_POPULATE_ONET_OCCUPATION 						= "/populateOnetOccupationList";
	public static final String ACTION_POPULATE_EXTNG_FACILITATOR_DATA 				= "/populateExistingFacilitatorData";
	public static final String ACTION_UPDATE_USER_NAME_FACILITATOR					= "/updateUserNameFacilitator";
	public static final String ACTION_UPDATE_EMAIL_ID_FACILITATOR					= "/updateEmailIdFacilitator";
	public static final String ACTION_UPDATE_PASSWORD_FACILITATOR					= "/updatePasswordFacilitator";
	public static final String ACTION_POPULATE_SUBSCRIBED_USER						= "/populateSubscribedUser";
	public static final String ACTION_SAVE_MANUAL_USER_FACILITATOR					= "/saveManualUserFacilitator";
	public static final String ACTION_SAVE_USER_CSV_FACILITATOR					    = "/saveCSVFacilitator";
	public static final String ACTION_VALIDATE_CSV_FACILITATOR					    = "/validateUserDataFacilitator";
	public static final String ACTION_POPULATE_USERS_BY_FACILITATOR					= "/populateUsersByFacilitatorId";
	public static final String ACTION_POPULATE_ACTIVE_INACTIVE_USER_FACILITATOR		= "/fetchActiveInactiveUserListFacilitator";	
	public static final String ACTION_FETCH_USER_RESET_PASSWORD_FACILITATOR			= "/fetchUserListForResetPassword";
	public static final String ACTION_CHANGE_BATCH_USER_STATUS_FACILITATOR			= "/changeUserStatusInBatchFacilitator";
	public static final String ACTION_RESET_PASSWORD_IN_BATCH_FACILITATOR			= "/resetPasswordInBatchFacilitator";
	public static final String ACTION_GENERATE_FAC_IND_UI_RPRT						= "/generateFacIndUIReport";
	public static final String ACTION_RENEW_USER_FACILITATOR						= "/renewUserByFacilitator";
	public static final String ACTION_POPULATE_USER_DROP_DOWN						= "/populateUserDropDown";
	public static final String ACTION_POPULATE_EXISTING_FACILITATOR					= "/populateExistingFacilitator";
	public static final String ACTION_UPDATE_USER_TO_FACILITATOR					= "/updateUserToFacilitator";
	public static final String ACTION_ADD_FACILITATOR_TO_CHANGE_ROLE				= "/addNewFacilitator";
	
	/********** SURVEY QUESTIONS **************/
	public static final String ACTION_SAVE_FREE_QTN_SURVEY_QTN						= "/saveFreeText";
	public static final String ACTION_SAVE_RADIO_ANSWERS							= "/saveRadio";
	public static final String ACTION_SAVE_DROP_DOWN								= "/saveDropDown";
	public static final String ACTION_SAVE_CHECK_BOX								= "/saveCheckBox";
	public static final String ACTION_FETCH_ALL_EXISTING_SURVEY_QTNS				= "/fetchAllExistingSurveyQtns";
	public static final String ACTION_FETCH_ALL_EXISTING_SURVEY_QTNS_BY_USR_TYP  	= "/fetchAllExistingSurveyQtnsByUserType";
	public static final String ACTION_FETCH_ALL_SURV_QTNS_BY_USR_TYPE_QTN_TYPE		= "/fetchAllExistingSurveyQtnsByUserTypeAndQtnType";
	public static final String ACTION_FETCH_ALL_SURV_QTNS_ALL_SELECTED				= "/fetchAllExistingSurveyQtnsAllSelected";
	public static final String ACTION_FETCH_ALL_SURV_QTNS_USER_PROFILE				= "/fetchAllExistingSurveyQtnsOnlyUserProfile";
	public static final String ACTION_FETCH_ALL_SURV_QTNS_USER_PROFILE_TYPE				= "/fetchSurveyQtnsByProfileAndUserType";
	public static final String ACTION_FETCH_ALL_SURV_QTNS_BY_QTN_TYPE				= "/fetchAllExistingSurveyQtnsByQtnType";
	public static final String ACTION_DELETE_SURVEY_QUESTION						= "/deleteSurveyQtn";
	public static final String ACTION_UPDATE_TEXT_SURVEY_QUESTION					= "/updateTextSurveyQtn";
	public static final String ACTION_UPDATE_RADIO_SURVEY_QUESTION					= "/updateRadioSurveyQtn";
	public static final String ACTION_UPDATE_DROP_DOWN_ENTRY						= "/updateDropdownSurveyQtn";
	public static final String ACTION_UPDATE_CHECK_BOX_ENTRY						= "/updateCheckboxSurveyQtn";
	public static final String ACTION_SAVE_SURVEY_QTN_ORDER							= "/saveSurvryQtnOrder";
	
	
	/*********** REMOVE DIAGRAMS **************/
	public static final String ACTION_SEARCH_SHARED_DIAGRAM							= "/searchUserDiagram";
	public static final String ACTION_REMOVE_SHARED_DIAGRAM							= "/removeDiagram";
	public static final String ACTION_POPULATE_ALL_DIAGRAM							= "/populateAllDiagram";
	public static final String ACTION_POPULATE_DIAGRAM_COUNT						= "/fetchAllDiagramCount";
	/*****************************************/
	public static final String ACTION_PROGRESS_REPORT								= "/populateNewProgressReport";
	public static final String ACTION_POPULATE_ALL_REPORT							= "/populateAllProgressReports";
	public static final String ACTION_RENEW_SUBSCRIBE_FACILITATOR					= "/renewSubscribe";
	public static final String ACTION_GET_FACILITATOR_EMAIL							= "/getFacilitatorEmail";
	
	/*********** SURVEY REPORT **************/
	public static final String SURVEY_REPORT							            = "/populateSurveyReport";

	public static final String ACTION_GENERATE_FAC_IND_UI_SURVEY_RPRT				= "/generateSurveyReport";
	public static final String ACTION_CHEQUE_DETAILS								= "/fetchChequeDetails";
	public static final String ACTION_CHEQUE_EXISTING_DETAILS						= "/fetchExistingChequeUsers";

	
	public static final String ACTION_SIGNUP_MASTER 								= "/populateSignupMasterData";
	public static final String ACTION_SAVE_SURVEY_ANSWER							= "/saveSurveyQuestions";
	public static final String ACTION_UPDATE_ACTIVE_STATUS_FACILITATOR				= "/updateActiveStatusFacilitator";
	
	public static final String ACTION_HONORED_CHECK 								= "/honoredCheck";
	public static final String ACTION_DISHONORED_CHECK 								= "/dishonoredCheck";
	
	public static final String ACTION_POPULATE_DEFAULT_PROFILE_ID 					= "/fetchDefaultProfileId";
	public static final String ACTION_UPDATE_USER_GROUP_FACILITATOR 				= "/updateUserGroupForFacilitator";
	public static final String ACTION_POPULATE_USER_GROUP_FACILITATOR 				= "/populateUserGroupDroDwnFacilitator";
	public static final String ACTION_POPULATE_PROGRESS_REPORT						= "/populateProgressReports";
	
	public static final String ACTION_SEARCH_USER_REFUND_REQUEST					= "/searchUserForRefundRequest";
	public static final String ACTION_DISABLE_USER_REFUND_REQUEST					= "/disableUserForRefundRequest";
	public static final String ACTION_VALIDATE_BULK_RENEW_FACILITATOR				= "/validateBulkRenewCSV";
	public static final String ACTION_BULK_RENEW_BY_FACILITATOR						= "/renewBulkUserByFacilitator";
	public static final String ACTION_POPULATE_RENEWED_USERS_BY_FACILITATOR			= "/populateRenewedUsersByFacilitator";
	
	public static final String ACTION_RENEW_GEN_USER_MANUAL_PAYMENT 				= "/renewGeneralUserPaymentManual";
	public static final String ACTION_RENEW_GEN_USER_CSV_PAYMENT 					= "/renewGeneralUserCSV";
	public static final String ACTION_VALIDATE_CSV_ENTRY							= "/validateCSVEntry";
	public static final String ACTION_VALIDATE_CSV_ENTRY_ADD_USER					= "/validateCSVEntryOnAddUser";
	public static final String ACTION_GET_SHOPIFY_LINK								= "/getShopifyLink";
	
	/*********** TERMS & CONDITIONS **************/
	
	public static final String ACTION_FETCH_TERMS_AND_CONDITIONS					="/fetchTermsAndConditions";
	public static final String ACTION_SAVE_TERMS_AND_CONDITIONS						="/saveAndUpdateTermAndCondition";	
	public static final String ACTION_BULK_ASSIGN_GROUP_BY_FACILITATOR				= "/assignGroupInBatchFacilitator";
	public static final String ACTION_GET_TERMS_AND_CONDITIONS						= "/fetchTnC";
}