package com.vmware.jct.controller;

import java.text.SimpleDateFormat;
import java.util.Calendar;
import java.util.Date;
import java.util.HashMap;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.Random;
import java.util.StringTokenizer;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpSession;

import org.codehaus.jackson.JsonNode;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.context.MessageSource;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.RequestBody;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestMethod;
import org.springframework.web.bind.annotation.ResponseBody;

import com.vmware.jct.common.utility.CommonConstants;
import com.vmware.jct.common.utility.CommonUtility;
import com.vmware.jct.common.utility.MailNotificationHelper;
import com.vmware.jct.dao.dto.FunctionDTO;
import com.vmware.jct.dao.dto.JobAttributeDTO;
import com.vmware.jct.dao.dto.LevelDTO;
import com.vmware.jct.dao.dto.OccupationListDTO;
import com.vmware.jct.dao.dto.RegionDTO;
import com.vmware.jct.exception.JCTException;
import com.vmware.jct.exception.MailingException;
import com.vmware.jct.model.JctUser;
import com.vmware.jct.model.JctUserDetails;
import com.vmware.jct.model.JctUserLoginInfo;
import com.vmware.jct.service.IActionPlanService;
import com.vmware.jct.service.IAfterSketchService;
import com.vmware.jct.service.IAuthenticatorService;
import com.vmware.jct.service.IBeforeSketchService;
import com.vmware.jct.service.IConfirmationService;
import com.vmware.jct.service.IJobAttributeFrozenService;
import com.vmware.jct.service.IQuestionnaireService;
import com.vmware.jct.service.ISketchSearchService;
import com.vmware.jct.service.ISurveyQuestionService;
import com.vmware.jct.service.vo.ForgotPasswordVO;
import com.vmware.jct.service.vo.ReturnVO;
import com.vmware.jct.service.vo.ShopifyVO;
import com.vmware.jct.service.vo.TermsAndConditionsVO;
import com.vmware.jct.service.vo.UserLoginInfoVO;
import com.vmware.jct.service.vo.UserVO;

/**
 * 
 * <p><b>Class name:</b> AuthenticatingContoller.java</p>
 * <p><b>Author:</b> InterraIT</p>
 * <p><b>Purpose:</b> This class acts as a controller for all the Authenticating and registering requests..
 * AuthenticatingContoller extends BasicController  and has following  methods.
 * -authenticate()
 * -checkUser()
 * -register()
 * -resetPassword()
 * -forgotPassword()
 * -logout()
 * <p><b>Description:</b> This class basically intercept all the authentication related service calls and 
 * then delegated the it to the service layer for further actions. </p>
 * <p><b>Copyrights:</b> 	All rights reserved by Interra IT and should only be used for its internal application development.</p>
 * <p><b>Created Date:</b> 10/Jan/2014</p>
 * <p><b>Revision History:</b>
 * 	<li>InterraIT - 04/Feb/2014 - Implemented comments</li>
 * </p>
 */
@Controller
@RequestMapping(value="/auth")
public class AuthenticatingContoller extends BasicController {
	
	@Autowired
	private IAuthenticatorService authenticatorService;
	
	@Autowired
	private IBeforeSketchService beforeSketchService;
	
	@Autowired
	private IAfterSketchService afterSketchService;
	
	@Autowired
	private IConfirmationService confirmationService;
	
	@Autowired
	private IQuestionnaireService iQuestionnaireService;
	
	@Autowired
	private IJobAttributeFrozenService attributeFrozenService;
	
	@Autowired
	private ISurveyQuestionService surveyQuestionService;
		
	@Autowired
	private ISketchSearchService searchService;
	
	@Autowired
	private IActionPlanService actionPlanService;
	
	@Autowired
	private MailNotificationHelper mailer;
	
	@Autowired  
	private MessageSource messageSource;
	
	private static final Logger LOGGER = LoggerFactory.getLogger(AuthenticatingContoller.class);
	
	/**
	 * Method authenticates the user and returns to the specified 
	 * page.
	 * @return UserVO
	 */
	@RequestMapping(value = ACTION_AUTHENTICATE, method = RequestMethod.POST)
	@ResponseBody()
	public UserVO authenticate(@RequestBody String body, HttpServletRequest request) {
		LOGGER.info(">>>> AuthenticatingContoller.authenticate");
		HttpSession httpSession = request.getSession();
		
		UserVO userVO = new UserVO();
		//JsonNode node = getNode(body);
		JsonNode node = CommonUtility.getNode(body);
		userVO.setEmail(node.get("emailId").toString().replaceAll("\"", ""));
		userVO.setPassword(node.get("password").toString().replaceAll("\"", ""));
		String screenResolution = node.get("userScreenResolution").asText();
		String operatingSystems = node.get("userOS").asText();
		SimpleDateFormat sdf = new SimpleDateFormat("dd-MM-yy:HH:mm:SS");
		String accessDate = sdf.format(new Date());
		
		LOGGER.info("********************************** LOGIN EXECUTING **************************************");
		LOGGER.info("<<======== DATE: "+accessDate+", USER ID: "+node.get("emailId").asText()+
				", SCREEN RESOLUTION: "+screenResolution+
				", OPERATING SYS: "+operatingSystems+"=========>>");
		LOGGER.info("********************************** LOGIN EXECUTED ***************************************");
		
		LOGGER.info("----> AUTHENTICATION STARTED FOR: "+userVO.getEmail());
		try {
			int userId = 0;
			int isCompleted = 0;
			String lastPage = "";
			String pageInfo = "";
			List<JctUser> userList = authenticatorService.authenticateUser(userVO);
			if (userList.size() > 0) {
				JctUser user = (JctUser) userList.get(0);
				int activeStatus = user.getJctActiveYn();
				int jctUserDisabled = user.getJctUserDisabled();				
				
				/** THIS IS DONE FOR THE PUBLIC VERSION **/
				userVO.setJctUserId(user.getJctUserId());
				/*****************************************/
				
				if ((activeStatus == CommonConstants.ACTIVATION_COMPLETE) && (jctUserDisabled == 0) && (user.getJctUserSoftDelete() == CommonConstants.ENTRY_NOT_DELETED)) {
					//Check for the account expiration date
					//int accountExpirationCount = getExpirationDetails(user.getJctAccountExpirationDate());
					int accountExpirationCount = authenticatorService.getDaysToExpire(user.getJctUserId());
					if (accountExpirationCount < 0) { // Account Expired
						LOGGER.info("ACCOUNT EXPIRED FOR THE USER: "+userVO.getEmail());
						userVO.setStatusCode(CommonConstants.USER_ACCOUNT_EXPIRED);
						userVO.setStatusMsg("accountExpired");
						userVO.setStatusDesc(this.messageSource.getMessage("error.account.expired",null, null));
						userVO.setProfileId(user.getJctUserDetails().getJctUserDetailsProfileId());
						// Convert the date in mm-dd-yyyy format
						Date expiredDate = user.getJctAccountExpirationDate();
						SimpleDateFormat DATE_FORMAT = new SimpleDateFormat("MMM-dd-yyyy");
				        String date = DATE_FORMAT.format(expiredDate);
						// Add entry in userVO to set the expitred date
				        userVO.setExpirationDate(date);
						
					}  else {
						if ((accountExpirationCount == 0)) {
							String message = this.messageSource.getMessage("warning.account.expiration.today",null, null);
							userVO.setAccountExpWarning(message);
						} else if ((accountExpirationCount > 0) && (accountExpirationCount <=7)) {
							String message = this.messageSource.getMessage("warning.account.expiration1",null, null)+" "+accountExpirationCount+" "+this.messageSource.getMessage("warning.account.expiration2",null, null);
							userVO.setAccountExpWarning(message);
						} else {
							userVO.setAccountExpWarning("NA");
						}
						//Check the total Failed Counts
						int totalFailedCount = authenticatorService.getTotalLoginFailedCounts(user.getJctUserId(), user.getJctUserEmail());
						if (totalFailedCount < 5) {
							try {
								authenticatorService.changeInfoStatus(user.getJctUserId(), user.getJctUserEmail());
							} catch (JCTException exception) {
								LOGGER.error("UNABLE TO CHANGE THE SOFT DELETE TO 1");
							}
							
							httpSession.setAttribute("isLoggedIn", true);
							httpSession.setAttribute("isLoggedPdfIn", true);
							LOGGER.info("----> AUTHENTICATED "+userVO.getEmail()+". STATUS: ACTIVE");
							userVO.setStatusCode(CommonConstants.SUCCESSFUL);
							userVO.setStatusMsg("loginSuccess");
							userVO.setStatusDesc("Successfully Logged in");
							/** Added user profile and user group **/
							userVO.setProfileId(user.getJctUserDetails().getJctUserDetailsProfileId());
							userVO.setGroupId(user.getJctUserDetails().getJctUserDetailsGroupId());
							userVO.setJobTitle(user.getJctUserDetails().getJctUserDetailsLevels());
							userVO.setFirstName(user.getJctUserDetails().getJctUserDetailsFirstName());
							String jobRefNo = "";
							userId = user.getJctUserId();
							//Check if user is already in mid of any task. jct_user_login_info and status is pending
							List<JctUserLoginInfo> userInfoList = authenticatorService.getTaskPendingDetails(user.getJctUserId());
							if(userInfoList.size() == 0) { // No pending or no entry at all
								//Generate a job reference number
								jobRefNo = generateJobReferenceNo(user.getJctUserId());
								userVO.setBsJson("none");
								userVO.setInitialJson("none");
								userVO.setTotalTime("00:00:00");
								userVO.setUrl("/user/view/beforeSketch.jsp");
								userVO.setIsCompleted(0);
								userVO.setLastPage("none");
								pageInfo = "/user/view/beforeSketch.jsp";
								
								/** THIS IS DONE FOR THE PUBLIC VERSION **/
								freezeContent(jobRefNo, user.getJctUserEmail(), userId, user.getJctUserDetails().getJctUserDetailsProfileId());
								/*****************************************/
								
							} else { 
								long totalMillis = 0L;
								for(int index=0; index<userInfoList.size(); index++){
									JctUserLoginInfo object = (JctUserLoginInfo)userInfoList.get(index);
									Calendar startCal = Calendar.getInstance();
									startCal.setTime(object.getJctStartTime());
									
									Calendar endCal = Calendar.getInstance();
									endCal.setTime(object.getJctEndTime());
									totalMillis += (endCal.getTimeInMillis() - startCal.getTimeInMillis());
									if((userInfoList.size() - index) == 1){
										userVO.setUrl(object.getJctPageInfo());
										jobRefNo = object.getJctJobrefNo();
										pageInfo = userVO.getUrl();
										lastPage = object.getJctNextPage();
									}
									
									if (object.getJctCompleted() == 1) {
										isCompleted = 1;
									} else {
										isCompleted = 0;
									}
								}
								int pageSequence = getPageSequence(pageInfo);
								if(pageSequence == CommonConstants.PAGE_SEQUENCE_BEFORE_SKETCH) {
									userVO = populateBeforeSketch(userVO, jobRefNo);
								} else if(pageSequence == CommonConstants.PAGE_SEQUENCE_QUESTIONAIR) {
									userVO = populateQuestions(userVO, jobRefNo ,user.getJctUserDetails().getJctUserDetailsProfileId());
								} else if(pageSequence == CommonConstants.PAGE_SEQUENCE_AFTER_SKETCH_ONE){
									userVO = populateMappingYourself(userVO, jobRefNo, userId);
								} else if(pageSequence == CommonConstants.PAGE_SEQUENCE_AFTER_SKETCH_TWO){
									userVO = populateAfterDiagram(userVO, jobRefNo);
								} else if(pageSequence == CommonConstants.PAGE_SEQUENCE_ACTION_PLAN){
									userVO = populateActionPlan(userVO, jobRefNo);
								}
								userVO.setTotalTime(CommonUtility.convertMillis(totalMillis));
								userVO.setLastPage(lastPage);
							}
							userVO.setJobRefNo(jobRefNo);
							userVO.setPageSequence(getPageSequence(pageInfo));
							userVO.setIsCompleted(isCompleted);
							userVO = insertLoginInfo(userVO, jobRefNo, userId, pageInfo, isCompleted);
							userVO.setExcCompletionCount(confirmationService.getExcersizeCompletionCount(jobRefNo));
							userVO.setSource("Tool");
						} else {
							userVO.setStatusCode(CommonConstants.USER_VALIDATION_FAILED);
							userVO.setStatusMsg("loginFailure");
							userVO.setStatusDesc(this.messageSource.getMessage("error.uservalidation.failed6",null, null));
						}
					}
						
				} else if (jctUserDisabled == 1) { // User disabled by Admin
					LOGGER.info("----> AUTHENTICATED "+userVO.getEmail()+". STATUS: USER DISABLED [USER EXIST]");
					userVO.setStatusCode(CommonConstants.USER_DISABLED);
					userVO.setStatusMsg("loginFailure");
					userVO.setStatusDesc(this.messageSource.getMessage("error.exist.butDisabled",null, null));
					userVO.setProfileId(user.getJctUserDetails().getJctUserDetailsProfileId());
				} else if ((activeStatus == CommonConstants.CREATED)) { // Admin created it
					LOGGER.info("----> AUTHENTICATED "+userVO.getEmail()+". STATUS: CREATED [NOT ACTIVATED]");
					userVO.setStatusCode(CommonConstants.USER_NOT_ACTIVATED);
					userVO.setStatusMsg("loginFailure");
					userVO.setStatusDesc(this.messageSource.getMessage("error.notActivated",null, null));
					userVO.setProfileId(user.getJctUserDetails().getJctUserDetailsProfileId());
					userVO.setUrl("/user/view/activateAccount.jsp?emailId="+user.getJctUserEmail()+"&showFlg=1");
				} else if (activeStatus == CommonConstants.ACTIVATED_ONLY_EMAIL) { //user gave only email
					LOGGER.info("----> AUTHENTICATED "+userVO.getEmail()+". STATUS: CREATED [NOT ACTIVATED]");
					userVO.setStatusCode(CommonConstants.USER_NOT_ACTIVATED);
					userVO.setStatusMsg("loginFailure");
					userVO.setStatusDesc(this.messageSource.getMessage("error.notActivated",null, null));
					userVO.setUrl("/user/view/activateAccount.jsp?emailId="+user.getJctUserEmail()+"&showFlg=2");
					userVO.setProfileId(user.getJctUserDetails().getJctUserDetailsProfileId());
				} else if (user.getJctUserSoftDelete() == CommonConstants.ENTRY_SOFT_DELETED) { //user gave only email
					LOGGER.info("----> AUTHENTICATED "+userVO.getEmail()+". STATUS: INACTIVATED BY ADMIN.");
					userVO.setStatusCode(CommonConstants.USER_INACTIVATED_BY_ADMIN);
					userVO.setStatusMsg("loginFailure");
					userVO.setStatusDesc(this.messageSource.getMessage("error.admin.deactivated",null, null));
					userVO.setProfileId(user.getJctUserDetails().getJctUserDetailsProfileId());
				} else if ((activeStatus == CommonConstants.CHECK_PAYMENT_NOT_CLEARED)) { // PAYMENT_CHEQUE NOT CLEARED BY ADMIN
					LOGGER.info("----> AUTHENTICATED "+userVO.getEmail()+". STATUS: PAYMENT_CHEQUE NOT CLEARED BY ADMIN");
					userVO.setStatusCode(CommonConstants.PAYMENT_CHEQUE_NOT_CLEARED);
					userVO.setStatusMsg("loginFailure");
					userVO.setStatusDesc(this.messageSource.getMessage("error.chequeNotCleared",null, null));
					userVO.setProfileId(user.getJctUserDetails().getJctUserDetailsProfileId());					
				} else {
					LOGGER.error("----> AUTHENTICATION FAILED FOR: "+userVO.getEmail()+". UNKNOWN ERROR");
					userVO.setStatusCode(CommonConstants.UNKNOWN_ERROR);
					userVO.setStatusMsg("loginFailure");
					userVO.setStatusDesc("erroe.loggedFailed");
					userVO.setProfileId(user.getJctUserDetails().getJctUserDetailsProfileId());
				}
				// Fetch the demographic information
				String message = this.messageSource.getMessage("label.demo.details.confirmation",null, null);
				String demoInfo = authenticatorService.getDemographicInformation(user.getJctUserEmail());
				//userVO.setDemographicPopupMsg(message + demoInfo);
				userVO.setDemographicPopupMsg(demoInfo);
				userVO.setCorrect(this.messageSource.getMessage("lable.demo.correct",null, null));
				userVO.setWrong(this.messageSource.getMessage("lable.demo.change",null, null));
			} else { //LOGIN FAILURE
				LOGGER.error("----> AUTHENTICATION FAILED FOR: "+userVO.getEmail()+". USER NOT FOUND");
				//Check if email is available. If yes then put it in user_login_info as failed login 
				List<JctUser> userListByEmail = authenticatorService.checkUser(userVO);
				if (userListByEmail.size() > 0) { //This means user name is present. Hence login failure
					JctUser user = (JctUser) userListByEmail.get(0);
					userVO.setStatusCode(CommonConstants.USER_VALIDATION_FAILED);
					userVO.setStatusMsg("loginFailure");
					if (user.getJctUserRole().getJctRoleId() == 1) {
						//save it to user_info table
						int totalFailedCount = authenticatorService.saveUserValidationFails(user.getJctUserId(), user.getJctUserEmail());
						if (totalFailedCount == 1) {
							userVO.setStatusDesc(this.messageSource.getMessage("error.uservalidation.failed1",null, null));
						} else if (totalFailedCount == 2) {
							userVO.setStatusDesc(this.messageSource.getMessage("error.uservalidation.failed2",null, null));
						} else if (totalFailedCount == 3) {
							userVO.setStatusDesc(this.messageSource.getMessage("error.uservalidation.failed3",null, null));
						} else if (totalFailedCount == 4) {
							userVO.setStatusDesc(this.messageSource.getMessage("error.uservalidation.failed4",null, null));
						} else if (totalFailedCount == 99) { //Started again
							userVO.setStatusDesc(this.messageSource.getMessage("error.uservalidation.failed1",null, null));
						}  else { // 5 and above
							userVO.setStatusDesc(this.messageSource.getMessage("error.uservalidation.failed5",null, null));
						}
					} else {
						userVO.setStatusDesc(this.messageSource.getMessage("error.userPasswordNotMatch",null, null));
					}
				} else {
					userVO.setStatusCode(CommonConstants.USER_NOT_FOUND);
					userVO.setStatusMsg("loginFailure");
					userVO.setStatusDesc(this.messageSource.getMessage("error.userPasswordNotMatch",null, null));
				}
			}
			
		} catch (Exception connExe) {
			LOGGER.error("----> AUTHENTICATION FAILED FOR: "+userVO.getEmail()+". PROBLEM WITH DATABASE: "+connExe.getMessage());
			userVO.setStatusCode(CommonConstants.DATABASE_NOT_FOUND);
			userVO.setStatusMsg("loginFailure");
			userVO.setStatusDesc(this.messageSource.getMessage("error.remortResourceUnreachable",null, null));
		}
		LOGGER.info("<<<< AuthenticatingContoller.authenticate");
		return userVO;
	}
	
	private UserVO populateActionPlan(UserVO userVO, String jobRefNo) throws JCTException{
		LOGGER.info(">>>> AuthenticatingContoller.populateActionPlan");
		//Fetch the Action Plan
		/*List dummyList = afterSketchService.fetchActionPlans(jobRefNo, userVO.getProfileId());
		List returnList = new ArrayList();
		int loopCounter = 1;
		for(int index = 0; index<dummyList.size(); index++){
			if(loopCounter < dummyList.size()){
				returnList.add(dummyList.get(index));
			} else {
				userVO.setExisting(dummyList.get(index).toString());
			}
			loopCounter = loopCounter + 1;
		}
		userVO.setList(returnList);*/
		
		//Fetch the snapshot url of the aftersketch diagram
		userVO.setSnapShot(afterSketchService.getFinalPageSnapShot(jobRefNo, "N"));
		LOGGER.info("<<<< AuthenticatingContoller.populateActionPlan");
		return userVO;
	}
	private UserVO insertLoginInfo(UserVO userVO, String jobRefNo, int userId, String pageInfo, int isCompleted) throws JCTException{
		LOGGER.info(">>>> AuthenticatingContoller.insertLoginInfo");
		//Insert the record in user_login_info
		UserLoginInfoVO loginInfoVO = new UserLoginInfoVO();
		loginInfoVO.setJobReferenceNo(jobRefNo);
		loginInfoVO.setUserId(userId);
		loginInfoVO.setPageInfo(pageInfo);
		loginInfoVO.setIsCompleted(isCompleted);
		String status = authenticatorService.saveLoginInfo(loginInfoVO);
		LOGGER.info("---------> LOGIN INFO INSERTED: "+status);
		if(status.contains("success")){
			userVO.setIdentifier(status.substring(status.indexOf("#")+1, status.length()));
		} else{
			userVO.setIdentifier("0");
		}
		LOGGER.info("<<<< AuthenticatingContoller.insertLoginInfo");
		return userVO;
	}
	/**
	 * Method checks if the user already exists or not
	 * @param selectedUserName
	 * @param request
	 * @return UserVO
	 */
	@RequestMapping(value = ACTION_CHECK_USER, method = RequestMethod.POST)
	@ResponseBody()
	public UserVO checkUser(@RequestBody String body, HttpServletRequest request) {
		LOGGER.info(">>>> AuthenticatingContoller.checkUser");
		UserVO userVO = new UserVO();
		JsonNode node = CommonUtility.getNode(body);
		userVO.setEmail(node.get("emailId").toString().replaceAll("\"", ""));
		LOGGER.info("----> CHECKING USER: "+userVO.getEmail());
		try {
			List<JctUser> userList = authenticatorService.checkUser(userVO);
			if (userList.size() > 0) {
				LOGGER.info("----> USER: "+userVO.getEmail()+" EXISTS...");
				userVO.setStatusMsg("userExists");
				userVO.setStatusDesc(this.messageSource.getMessage("error.alreadyRegistered",null, null));
			} else {
				userVO.setStatusMsg("userDoesntExist");
				userVO.setStatusDesc("");
			}
		} catch (Exception connExe) {
			LOGGER.error(connExe.getMessage());
			userVO.setStatusCode(CommonConstants.DATABASE_NOT_FOUND);
			userVO.setStatusMsg("userDoesntExist");
			userVO.setStatusDesc(this.messageSource.getMessage("error.remortResourceUnreachable",null, null));
		}
		LOGGER.info("<<<< AuthenticatingContoller.checkUser");
		return userVO;
	}
	/**
	 * Method to register new user
	 * @param body
	 * @param request
	 * @return
	 */
	@RequestMapping(value = ACTION_REGISTER, method = RequestMethod.POST)
	@ResponseBody()
	public UserVO register(@RequestBody String body, HttpServletRequest request) {
		LOGGER.info(">>>> AuthenticatingContoller.register");
		HttpSession httpSession = request.getSession();
		UserVO userVO = new UserVO();
		JsonNode node = CommonUtility.getNode(body);
		String status = "";
		String statusCode = "";
		try {
			userVO.setFirstName(node.get("firstNames").asText());
			userVO.setLastName(node.get("lastNames").asText());
			userVO.setLocation(node.get("location").asText());
			userVO.setGender(node.get("genders").asText());
			userVO.setSupervise(node.get("supervisePeoples").asText());
			userVO.setTenure(node.get("tenures").asText());
			userVO.setFunctionGrpSelected(node.get("functionGrps").asText());
			userVO.setJobLevelSelected(node.get("jobLevels").asText());
			userVO.setEmail(node.get("emailIds").asText());
			/** THIS IS DONE FOR THE PUBLIC VERSION **/
			userVO.setOccupationVal(node.get("occupationVal").asText());
			userVO.setNosOfYrs(node.get("nosOfYrs").asText());
			/*****************************************/
			
			try {
				// Get the user Id
				JctUser user = authenticatorService.checkInactiveUser(userVO).get(0);
				int userId = user.getJctUserId();
				httpSession.setAttribute("isLoggedIn", true);
				httpSession.setAttribute("isLoggedPdfIn", true);
				status = authenticatorService.register(userVO);
				
				if (!status.equals("onet-not-found")) {
					String jobRefNo = "";
					String pageInfo = "";
					int accountExpirationCount = authenticatorService.getDaysToExpire(user.getJctUserId());
					if (accountExpirationCount < 0) { // Account Expired
						LOGGER.info("ACCOUNT EXPIRED FOR THE USER: "+userVO.getEmail());
						userVO.setStatusCode(CommonConstants.USER_ACCOUNT_EXPIRED);
						userVO.setStatusMsg("accountExpired");
						userVO.setStatusDesc(this.messageSource.getMessage("error.account.expired",null, null));
						userVO.setProfileId(user.getJctUserDetails().getJctUserDetailsProfileId());
					}  else {
						if ((accountExpirationCount == 0)) {
							String message = this.messageSource.getMessage("warning.account.expiration.today",null, null);
							userVO.setAccountExpWarning(message);
						} else if ((accountExpirationCount > 0) && (accountExpirationCount <=3)) {
							String message = this.messageSource.getMessage("warning.account.expiration1",null, null)+" "+accountExpirationCount+" "+this.messageSource.getMessage("warning.account.expiration2",null, null);
							userVO.setAccountExpWarning(message);
						} else {
							userVO.setAccountExpWarning("NA");
						}
						
						statusCode = CommonConstants.SUCCESSFUL;
						userVO.setStatusDesc(this.messageSource.getMessage("error.confirmation",null, null));
						userVO.setUserName(userVO.getUserName());
						userVO.setEmail(userVO.getEmail());
						
						/***************Survey Questions**************/
						// Survey questions
						String surveyQtnTextString = node.get("surveyQtnTextString").asText();
						String surveyQtnRadioString = node.get("surveyQtnRadioString").asText();
						String surveyQtnCheckBoxString = node.get("surveyQtnCheckBoxString").asText();
						String surveyQtnDropdownString = node.get("surveyQtnDropdownString").asText();
						processSurveyQuestions(surveyQtnTextString, surveyQtnRadioString, surveyQtnCheckBoxString, surveyQtnDropdownString, node.get("emailIds").asText(), userId);
						/************** STORE SURVEY QTN ************/
						
						List<JctUserLoginInfo> userInfoList = authenticatorService.getTaskPendingDetails(userId);
						if(userInfoList.size() == 0) { // No pending or no entry at all
							//Generate a job reference number
							jobRefNo = generateJobReferenceNo(userId);
							userVO.setBsJson("none");
							userVO.setInitialJson("none");
							userVO.setTotalTime("00:00:00");
							userVO.setUrl("/user/view/beforeSketch.jsp");
							userVO.setIsCompleted(0);
							userVO.setLastPage("none");
							pageInfo = "/user/view/beforeSketch.jsp";
							userVO.setJobTitle(userVO.getJobLevelSelected());
						} else {
							LOGGER.error("AFTER REGISTER THERE SHOULD NOT BE ANY PENDING TASK... DELETE THE LIST FIRST!");
						}
					}
					/** THIS IS DONE FOR THE PUBLIC VERSION **/
					userVO.setProfileId(user.getJctUserDetails().getJctUserDetailsProfileId());
					freezeContent(jobRefNo, user.getJctUserEmail(), userId, user.getJctUserDetails().getJctUserDetailsProfileId());
					userVO.setJctUserId(userId);
					userVO.setJobRefNo(jobRefNo);
					userVO = insertLoginInfo(userVO, jobRefNo, userId, pageInfo, 0);
					userVO.setJobTitle(userVO.getJobLevelSelected());
					/*****************************************/
					
				} else {
					LOGGER.error("ONET DATA IS NOT FOUND IN THE DATABASE FOR: "+userVO.getOccupationVal());
				}
			} 
			catch (JCTException e) {
				statusCode = CommonConstants.USER_REGISTRATION_FAILED;
				LOGGER.error("Inisde outer try"+e.getMessage());
			}
			if (!status.equals("onet-not-found")) {
				userVO.setStatusMsg(status);
				userVO.setStatusCode(statusCode);
			} else {
				userVO.setStatusCode(CommonConstants.ONET_NOT_FOUND);
				userVO.setStatusMsg("onet-not-found");
				userVO.setStatusDesc(this.messageSource.getMessage("error.onet.not.found",null, null));
			}
		}  catch (Exception connExe) {
			LOGGER.error(connExe.getMessage());
			userVO.setStatusCode(CommonConstants.DATABASE_NOT_FOUND);
			userVO.setStatusMsg("userDoesntExist");
			userVO.setStatusDesc(this.messageSource.getMessage("error.remortResourceUnreachable",null, null));
	}
		LOGGER.info("<<<< AuthenticatingContoller.register");
		return userVO;
	}
	/**
	 * Method saves survey questions. This will be saved when only for the first time 
	 * when the user logs in for the very first time. From next login the user will not 
	 * get the survey questions to answers.
	 * @param surveyQtnTextString
	 * @param surveyQtnRadioString
	 * @param surveyQtnCheckBoxString
	 * @param surveyQtnDropdownString
	 */
	private void processSurveyQuestions(String surveyQtnTextString,
			String surveyQtnRadioString, String surveyQtnCheckBoxString,
			String surveyQtnDropdownString, String emailId, int userId) {
		LOGGER.info(">>>>>>>> AuthenticatingContoller.processSurveyQuestions");
		try {
			surveyQuestionService.saveSurveyQuestions(surveyQtnTextString, surveyQtnRadioString, 
					surveyQtnDropdownString, surveyQtnCheckBoxString, emailId, userId);
		} catch (JCTException ex) {
			LOGGER.error("UNABLE TO SAVE SURVEY QUESTIONS: "+ex.getLocalizedMessage());
		}
		LOGGER.info("<<<<<<<< AuthenticatingContoller.processSurveyQuestions");
	}
	/**
	 * Method handles reset password functionality
	 * @param body
	 * @param request
	 * @return
	 */
	@RequestMapping(value = ACTION_RESET_PASSWORD, method = RequestMethod.POST)
	@ResponseBody()
	public ReturnVO resetPassword(@RequestBody String body,
			HttpServletRequest request) {
		LOGGER.info(">>>> AuthenticatingContoller.resetPassword");
		HttpSession httpSession = request.getSession();
		UserVO userVO = new UserVO();
		ReturnVO returnVO = new ReturnVO();
		JsonNode node = CommonUtility.getNode(body);
		userVO.setEmail(node.get("forUser").toString().replaceAll("\"", "")); // This is the username
		userVO.setInitialPassword(node.get("initialPassword").toString()
				.replaceAll("\"", ""));
		userVO.setPassword(node.get("userPassword").toString()
				.replaceAll("\"", ""));
		userVO.setClickStatus(node.get("clickStatus").toString().replaceAll("\"", ""));
		String mailedPassword = node.get("mailedPwd").toString().replaceAll("\"", "");
		userVO.setMailedPassword(mailedPassword);
				
		int dist = node.get("dist").asInt();
		String status = "failure";
		String statusPwd = "failure"; 
		
		
		
		//Check for status
		//userVO.setActiveYn(Integer.parseInt(node.get("status").toString().replaceAll("\"", "")));
		// Go got updation
		try {
			// If already activated once
			if (authenticatorService.alreadyRegisteredOnce(userVO.getEmail())) {
				returnVO.setStatusCode("128");
				userVO.setActiveYn(3);
			} else {
				//Check for status
				userVO.setActiveYn(Integer.parseInt(node.get("status").toString().replaceAll("\"", "")));
				
			}
			
			
			httpSession.setAttribute("isLoggedIn", true);
			httpSession.setAttribute("isLoggedPdfIn", true);
			if (!mailedPassword.equals("NA")) {
				statusPwd = authenticatorService.getMailedPassword(userVO);
			} else {
				// Check if initial password are correct
				statusPwd = authenticatorService.getInitialPassword(userVO);
			}
			if(statusPwd.equalsIgnoreCase("success")){
				status = authenticatorService.resetPassword(userVO, dist);
			if (!status.contains("not-found")) {
				returnVO.setStatusCode(CommonConstants.SUCCESSFUL);
				returnVO.setStatusMsg(status);
				returnVO.setStatusDesc(this.messageSource.getMessage("error.passwordChangedSuccessfully",null, null));
				List<JctUser> userListByEmail = authenticatorService.checkUser(userVO);
				JctUser user = (JctUser) userListByEmail.get(0);
				returnVO.setProfileId(user.getJctUserDetails().getJctUserDetailsProfileId());
				if (authenticatorService.alreadyRegisteredOnce(userVO.getEmail())) {
					returnVO.setStatusCode("128");
				}
			}
			else if (!status.contains("not-complete")) {
				returnVO.setStatusCode(CommonConstants.PASSWORD_RESET_NOT_COMPLETE);
				returnVO.setStatusMsg(status);
				returnVO.setStatusDesc(this.messageSource.getMessage("error.passwordResetNotComplete",null, null));
			} else {
				returnVO.setStatusCode(CommonConstants.USER_NOT_FOUND);
				returnVO.setStatusMsg("failure");
				returnVO.setStatusDesc(this.messageSource.getMessage("error.sendPasswordMissMatch",null, null));
			}
			}
			else{
				returnVO.setStatusCode(CommonConstants.FAILURE);
				returnVO.setStatusMsg("failure");
				returnVO.setStatusDesc(this.messageSource.getMessage("error.generated.password.incorrect",null, null));
			}
			
			
			
		} catch (JCTException e) {
			LOGGER.error(e.getMessage());
			returnVO.setStatusCode(CommonConstants.PASSWORD_RESET_FAILED);
			returnVO.setStatusDesc(this.messageSource.getMessage("error.changePasswordOnceAgain",null, null));
		} catch (Exception connExe) {
			LOGGER.error(connExe.getMessage());
			returnVO.setStatusCode(CommonConstants.DATABASE_NOT_FOUND);
			returnVO.setStatusMsg("userDoesntExist");
			returnVO.setStatusDesc(this.messageSource.getMessage("error.remortResourceUnreachable",null, null));
		}
		LOGGER.info("<<<< AuthenticatingContoller.resetPassword");
		return returnVO;
	}
	/**
	 * Method handles the forgot password functionality
	 * @param body
	 * @param request
	 * @return
	 */
	@RequestMapping(value = ACTION_FORGOT_PASSWORD, method = RequestMethod.POST)
	@ResponseBody()
	public ForgotPasswordVO forgotPassword(@RequestBody String body, HttpServletRequest request) {
		LOGGER.info(">>>> AuthenticatingContoller.forgotPassword");
		HttpSession httpSession = request.getSession();
		UserVO userVO = new UserVO();
		ReturnVO returnVO = null;
		ForgotPasswordVO vo = new ForgotPasswordVO();
		
		JsonNode node = CommonUtility.getNode(body);
		userVO.setEmail(node.get("myEmailId").toString().replaceAll("\"", ""));
		String generatedPwd = generatePassword();
		String encryptedPwd = encriptString(generatedPwd); // stored in db
		userVO.setPassword(encryptedPwd);
		try {
			httpSession.setAttribute("isLoggedIn", true);
			httpSession.setAttribute("isLoggedPdfIn", true);
			
			//If his account is locked then do not allow
			boolean isAllowed = authenticatorService.isForgotPwdAllowed(node.get("myEmailId").toString().replaceAll("\"", ""));
			if (isAllowed) {
				returnVO = authenticatorService.forgotPassword(userVO);
				if (returnVO.getStatusCode().equals(CommonConstants.SUCCESSFUL)) {
					mailer.sendPasswordByMail(userVO.getEmail(), generatedPwd, returnVO.getUserName(), returnVO.getfName(), returnVO.getlName());
					//returnVO.setEmail(userVO.getEmail());
					vo.setEmail(userVO.getEmail());
					vo.setStatusCode(returnVO.getStatusCode());
					vo.setStatusMsg(returnVO.getStatusMsg());
					vo.setStatusDesc(returnVO.getStatusDesc());
				} else if (returnVO.getStatusCode().equals(CommonConstants.USER_NOT_FOUND)) {
					vo.setEmail(userVO.getEmail());
					vo.setStatusCode(returnVO.getStatusCode());
					vo.setStatusMsg(returnVO.getStatusMsg());
					vo.setStatusDesc(returnVO.getStatusDesc());
				}
			} else {
				//returnVO.setStatusCode(CommonConstants.FORGOT_PASSWORD_LOCKED);
				//returnVO.setStatusMsg("FORGOT_PASSWORD_LOCKED");
				//returnVO.setStatusDesc(this.messageSource.getMessage("error.forgot.password.locked",null, null));
				vo.setStatusCode(CommonConstants.FORGOT_PASSWORD_LOCKED);
				vo.setStatusMsg("FORGOT_PASSWORD_LOCKED");
				vo.setStatusDesc(this.messageSource.getMessage("error.forgot.password.locked",null, null));
			}
			
			
		} catch (MailingException e) {
			//e.printStackTrace();
			//returnVO.setStatusCode(CommonConstants.MAIL_SEND_ERROR);
			//returnVO.setStatusMsg("mailingExep");
			//returnVO.setStatusDesc(this.messageSource.getMessage("error.sendMailFailure",null, null));
			vo.setStatusCode(CommonConstants.MAIL_SEND_ERROR);
			vo.setStatusMsg("mailingExep");
			vo.setStatusDesc(this.messageSource.getMessage("error.sendMailFailure",null, null));
			LOGGER.error(e.getMessage());
		} catch (JCTException e) {
			LOGGER.error(e.getMessage());
			//returnVO.setStatusCode(CommonConstants.DATABASE_NOT_FOUND);
			//returnVO.setStatusMsg("userDoesntExist");
			//returnVO.setStatusDesc(this.messageSource.getMessage("error.remortResourceUnreachable",null, null));
			vo.setStatusCode(CommonConstants.DATABASE_NOT_FOUND);
			vo.setStatusMsg("userDoesntExist");
			vo.setStatusDesc(this.messageSource.getMessage("error.remortResourceUnreachable",null, null));
		} catch (Exception connExe) {
			LOGGER.error(connExe.getMessage());
			//returnVO.setStatusCode(CommonConstants.DATABASE_NOT_FOUND);
			//returnVO.setStatusMsg("userDoesntExist");
			//returnVO.setStatusDesc(this.messageSource.getMessage("error.remortResourceUnreachable",null, null));
			vo.setStatusCode(CommonConstants.DATABASE_NOT_FOUND);
			vo.setStatusMsg("userDoesntExist");
			vo.setStatusDesc(this.messageSource.getMessage("error.remortResourceUnreachable",null, null));
		}
		LOGGER.info("<<<< AuthenticatingContoller.forgotPassword");
		return vo;
	}
	/**
	 * Method takes care of logout.
	 * @param body
	 * @param request
	 * @return null
	 */
	@RequestMapping(value = ACTION_LOGOUT, method = RequestMethod.POST)
	@ResponseBody()
	public ReturnVO logout(@RequestBody String body,
			HttpServletRequest request) {
		LOGGER.info(">>>> AuthenticatingContoller.logout");
		
		HttpSession httpSession = request.getSession();
		//JsonNode node = getNode(body);
		JsonNode node = CommonUtility.getNode(body);
		String jobReferenceNumber = node.get("jobReferenceString").toString().replaceAll("\"", "");
		int rowId = Integer.parseInt(node.get("rowId").toString().replaceAll("\"", ""));
		String lastActivePage = node.get("lastActivePage").toString().replaceAll("\"", "");
		try {
			String result = authenticatorService.updateLoginInfo(jobReferenceNumber, rowId, lastActivePage);
			httpSession.removeAttribute("isLoggedIn");
			httpSession.removeAttribute("isLoggedPdfIn");
			LOGGER.info("---------> LOGOUT: "+result);
		} catch (JCTException e) {
			LOGGER.error(e.getMessage());
		}
		LOGGER.info("<<<< AuthenticatingContoller.logout");
		return null;
	}
	/**
	 * Method encripts the string [Custom encryption]
	 * @param string
	 * @return
	 */
	private String encriptString(String string){
		LOGGER.info(">>>> AuthenticatingContoller.encriptString");
		//String ref = "1&(9%2^*37@86^54&&^%&^0CKj=mo(%np(^qs$YTuQ%^rxyd=ahg)&ei=lBfWzC@*KJM*&%ioON=PQajpSTUVR~XYDA*&^*HGE=ILBFWZ";
		//String ref = "Q%1wS~(RxGzPA$By&N!Oa_2#^CM>D)<bFWLE@3c4*Kd5J+IVe-6Hf7g8=h9iT0jU";
		//String ref = "P!QR#S$T%U&V'W(X)Y*Z+[,-].^/_0`1a2b3c4d5e6f7g8h9i:j;k<l=m>n?o@pAqBrCsDtEuFvGwHxIyJzK{L|M}N~O";
		String ref = "P!QR#S$T%U&VW(X)Y*Z+[,-].^/_0`1a2b3c4d5e6f7g8h9i:j;k<l=m>n?o@pAqBrCsDtEuFvGwHxIyJzK{L|M}N~O";
		StringBuilder result = new StringBuilder("");
			for (int count=0; count<string.length(); count++) {
				String charStr = string.substring (count, count+1); 
				int num = ref.indexOf(charStr);
				String encodeChar = ref.substring(num+1, num+2);
				result.append(encodeChar);
			}
			result.append(result.toString());
		LOGGER.info("<<<< AuthenticatingContoller.encriptString");
		return result.toString();
	}
	/**
	 * Method generates random password
	 * @return generated password
	 */
	private String generatePassword() {
		LOGGER.info(">>>> AuthenticatingContoller.generatePassword");
		Random rand = new Random();
		// begin method with return string and you may or may not give input
		char[] values1 = { 'w', 'e', 'l', 'c', 's', 'p', 'a', 'm', 'S', 'M',
				'A', 'H', 'E', 'V', 'X', 'q' };
		// char[] values2 = {'_','&','$','#','%','*'};
		char[] values2 = { '_' };
		char[] values3 = { '1', '2', '3', '4', '5', '6', '7', '8', '9', '0' };
		StringBuilder sb = new StringBuilder("");
		
		for (int i = 0; i < 6; i++) {
			int idx = rand.nextInt(values1.length);
			sb.append(values1[idx]);
		}

		for (int i = 0; i < 1; i++) {
			int idx = rand.nextInt(values2.length);
			sb.append(values2[idx]);
		}
		
		for (int i = 0; i < 3; i++) {
			int idx = rand.nextInt(values3.length);
			sb.append(values3[idx]);
		}
		LOGGER.info("<<<< AuthenticatingContoller.generatePassword");
		return sb.toString();
	}
	/**
	 * Method generates the job reference number
	 * @return yyyyMMddXXXXXXXX
	 */
	private String generateJobReferenceNo(int userId) {
		LOGGER.info(">>>> AuthenticatingContoller.generateJobReferenceNo");
		StringBuilder jobRefBuilder = new StringBuilder("");
		SimpleDateFormat fmt = new SimpleDateFormat("yyyyMMdd");
		String dateInString = fmt.format(new Date());
		jobRefBuilder.append(dateInString);
		try {
			int id = authenticatorService.getGeneratedJRId();
			jobRefBuilder.append(generateJobRefNumber(id, String.valueOf(id).length()));
			jobRefBuilder.append(userId);
		} catch (JCTException e) {
			LOGGER.error(e.getMessage());
		}
		LOGGER.info("<<<< AuthenticatingContoller.generateJobReferenceNo");
		return jobRefBuilder.toString();
	}
	/**
	 * Method generates Job reference number if the user is logging
	 * for the first time.
	 * @param id
	 * @param length
	 * @return reference number (String)
	 */
	private String generateJobRefNumber(int id, int length){
		LOGGER.info(">>>> AuthenticatingContoller.generateJobRefNumber");
		
		String idString = "";
		switch (length) {
		case 1:
			idString = "0000000"+id;
			break;
		case 2:
			idString = "000000"+id;
			break;
		case 3:
			idString = "00000"+id;
			break;
		case 4:
			idString = "0000"+id;
			break;
		case 5:
			idString = "000"+id;
			break;
		case 6:
			idString = "00"+id;
			break;
		case 7:
			idString = "0"+id;
			break;

		default:
			idString = ""+id;
			break;
		}
		LOGGER.info("<<<< AuthenticatingContoller.generateJobRefNumber");
		return idString;
	}
	/**
	 * Method returns the page sequence
	 * @param pageInfo
	 * @return
	 */
	private int getPageSequence(String pageInfo) {
		LOGGER.info(">>>> AuthenticatingContoller.getPageSequence");
		
		int pageSeq = 0;
		if(pageInfo.contains("beforeSketch.jsp")){
			pageSeq = 1;
		} else if(pageInfo.contains("Questionaire.jsp")){
			pageSeq = 2;
		} else if(pageInfo.contains("afterSketch1.jsp")){
			pageSeq = 3;
		} else if(pageInfo.contains("afterSketch2.jsp")){
			pageSeq = 4;
		} else if(pageInfo.contains("actionPlan.jsp")){
			pageSeq = 5;
		} else {
			pageSeq = 6;
		}
		LOGGER.info("<<<< AuthenticatingContoller.getPageSequence");
		return pageSeq;
	}
	/**
	 * Method populates the before sketch screen (1st screen)
	 * @param userVO
	 * @param jobRefNo
	 * @return
	 * @throws JCTException
	 */
	private UserVO populateBeforeSketch(UserVO userVO, String jobRefNo) throws JCTException {
		LOGGER.info(">>>> AuthenticatingContoller.populateBeforeSketch");
		
		userVO.setBsJson(beforeSketchService.getJson(jobRefNo, "N"));
		userVO.setInitialJson(beforeSketchService.getInitialJson(jobRefNo, "N"));
		LOGGER.info("<<<< AuthenticatingContoller.populateBeforeSketch");
		return userVO;
	}
	/**
	 * Method populates the populate questions screen (2nd screen)
	 * @param userVO
	 * @param jobRefNo
	 * @return
	 * @throws JCTException
	 */
	private UserVO populateQuestions(UserVO userVO, String jobRefNo, int profileId) throws JCTException{
		LOGGER.info(">>>> AuthenticatingContoller.populateQuestions");
		//userVO.setList(beforeSketchService.fetchQuestionAnswer(jobRefNo,profileId));
		userVO.setJctBaseString(beforeSketchService.fetchBase64String(jobRefNo));
		LOGGER.info("<<<< AuthenticatingContoller.populateQuestions");
		return userVO;
	}
	/**
	 * Method populates the Mapping yourself screen (3rd screen)
	 * @param userVO
	 * @param jobRefNo
	 * @return
	 * @throws JCTException
	 */
	private UserVO populateMappingYourself(UserVO userVO, String jobRefNo, int userId) throws JCTException {
		LOGGER.info(">>>> AuthenticatingContoller.populateMappingYourself");
		
		//Populate the check box
		List<JobAttributeDTO> jctJobAttributeList=null;
		Map<Integer, String> strengthMap=new HashMap<Integer, String>();
		Map<Integer, String> passionMap=new HashMap<Integer, String>();
		Map<Integer, String> valueMap=new HashMap<Integer, String>();
		
		/** THIS IS FOR PUBLIC VERSION **/
		//jctJobAttributeList=iQuestionnaireService.fetchJobAttribute(profileId);
		jctJobAttributeList = attributeFrozenService.fetchJobAttribute(userVO.getProfileId(), userVO.getEmail(), userId);
		/********************************/
		
		LOGGER.debug("**********"+jctJobAttributeList.size());
		int j=0;
		for(int i=1;i<=jctJobAttributeList.size();i++){
			if(jctJobAttributeList.get(j).getJctJobAttributeCode().equals(CommonConstants.JOB_ATTRIBUTE_STRENGTH)){
				strengthMap.put(jctJobAttributeList.get(j).jctJobAttributeId, jctJobAttributeList.get(j).getJctJobAttributeName()+"+++"+jctJobAttributeList.get(j).getJctJobAttributeDesc().replace(",", "~~"));
			}else if(jctJobAttributeList.get(j).getJctJobAttributeCode().equals(CommonConstants.JOB_ATTRIBUTE_PASSION)){
				passionMap.put(jctJobAttributeList.get(j).jctJobAttributeId, jctJobAttributeList.get(j).getJctJobAttributeName()+"+++"+jctJobAttributeList.get(j).getJctJobAttributeDesc().replace(",", "~~"));
			}else{
				valueMap.put(jctJobAttributeList.get(j).jctJobAttributeId, jctJobAttributeList.get(j).getJctJobAttributeName()+"+++"+jctJobAttributeList.get(j).getJctJobAttributeDesc().replace(",", "~~"));
			}
			j++;
		}
		userVO.setStrengthMap(strengthMap);
		userVO.setPassionMap(passionMap);
		userVO.setValueMap(valueMap);
		userVO.setAfterSketchCheckedEle(afterSketchService.fetchCheckedElements(jobRefNo, "N")); //CheckedElements
		userVO.setAfterSkPageOneTotalJson(afterSketchService.fetchASOneJson(jobRefNo, "N"));// json
		userVO.setPassionCount(afterSketchService.getPassionCount(jobRefNo));
		userVO.setValueCount(afterSketchService.getValueCount(jobRefNo));
		userVO.setStrengthCount(afterSketchService.getStrengthCount(jobRefNo));
		
		String str = afterSketchService.getCheckedItems(jobRefNo, "N");
		StringTokenizer strToken = new StringTokenizer(str, "#");
		while(strToken.hasMoreTokens()){
			userVO.setCheckedStrength(strToken.nextToken().toString());
			userVO.setCheckedPassion(strToken.nextToken().toString());
			userVO.setCheckedValue(strToken.nextToken().toString());
		}
		LOGGER.info("<<<< AuthenticatingContoller.populateMappingYourself");
		return userVO;
	}
	
	private UserVO populateAfterDiagram(UserVO userVO, String jobRefNo) throws JCTException {
		LOGGER.info(">>>> AuthenticatingContoller.populateAfterDiagram");
		
		//Fetch total_json_object [string / value / strength json]
		userVO.setPageOneJson(afterSketchService.fetchASOneJson(jobRefNo, "N"));
		//Fetch total_json_add_task
		userVO.setTotalJsonAddTask(beforeSketchService.getJson(jobRefNo, "N"));
		//Set div_initial_value
		//userVO.setDivInitialValue(beforeSketchService.getInitialJson(jobRefNo));
		userVO.setDivInitialValue(afterSketchService.getFinalPageInitialJson(jobRefNo, "N"));
		//fetch the total json of after task
		userVO.setAfterSketch2TotalJsonObj(afterSketchService.fetchASFinalPageJson(jobRefNo, "N"));
		String str = afterSketchService.getDivHeightAndWidth(jobRefNo, "N");
		StringTokenizer token = new StringTokenizer(str, "#");
		String height = "";
		String width = "";
		try{
			while(token.hasMoreTokens()){
				height = (String)token.nextToken();
				width = (String)token.nextToken();
			}
			}	catch(Exception ex){
				LOGGER.error("---------------------------->WIDTH AND HEIGHT ISSUE AFTER DRAGGING");
		}
		
		userVO.setDivHeight(height);
		userVO.setDivWidth(width);
		userVO.setAsTotalCount(afterSketchService.getAsTotalCount(jobRefNo, "N"));
		LOGGER.info("<<<< AuthenticatingContoller.populateAfterDiagram");
		return userVO;
	}
	
	/**
	 * Method populates the Region master
	 * @param userVO
	 * @param Null
	 * @return
	 * @throws JCTException
	 */
	private UserVO populateRegion(UserVO userVO) throws JCTException {
		LOGGER.info(">>>> AuthenticatingContoller.populateRegion");
		
		List<RegionDTO> jctRegionList=null;
		Map<String, String> regionMap=new LinkedHashMap<String, String>();
		jctRegionList=authenticatorService.jctRegionList();
			int j=0;
			for(int i=1;i<=jctRegionList.size();i++){
				//regionMap.put(jctRegionList.get(j).getJctRegionId(), jctRegionList.get(j).getJctRegionName());
				regionMap.put(jctRegionList.get(j).getJctRegionName(), jctRegionList.get(j).getJctRegionName());
				j++;
		
			userVO.setRegion(regionMap);
		}
		LOGGER.info("<<<< AuthenticatingContoller.populateRegion");
		return userVO;
	}
	
	/**
	 * Method populates the Function Group master
	 * @param userVO
	 * @param Null
	 * @return
	 * @throws JCTException
	 */
	private UserVO populateFunction(UserVO userVO) throws JCTException {
		LOGGER.info(">>>> AuthenticatingContoller.populateFunction");
		List<FunctionDTO> jctFunctionList=null;
		Map<String, String> functionGroupMap=new LinkedHashMap<String, String>();
			jctFunctionList=authenticatorService.jctFunctionList();
			int j=0;
			for(int i=1;i<=jctFunctionList.size();i++){
				//functionGroupMap.put(jctFunctionList.get(j).getJctFunctionId(), jctFunctionList.get(j).getJctFunctionName());
				functionGroupMap.put(jctFunctionList.get(j).getJctFunctionName(), jctFunctionList.get(j).getJctFunctionName());
				j++;
			}
			userVO.setFunctionGroup(functionGroupMap);
			LOGGER.info("<<<< AuthenticatingContoller.populateFunction");
		return userVO;
	}
	
	/**
	 * Method populates the Level master
	 * @param userVO
	 * @param Null
	 * @return
	 * @throws JCTException
	 */
	private UserVO populateLevel(UserVO userVO) throws JCTException {
		LOGGER.info(">>>> AuthenticatingContoller.populateLevel");
		List<LevelDTO> jctLevelList=null;
		Map<String, String> jobLevelMap=new LinkedHashMap<String, String>();
			jctLevelList=authenticatorService.jctLevelList();
			int j=0;
			for(int i=1;i<=jctLevelList.size();i++){
				//jobLevelMap.put(jctLevelList.get(j).getJctLevelId(), jctLevelList.get(j).getJctLevelName());
				jobLevelMap.put(jctLevelList.get(j).getJctLevelName(), jctLevelList.get(j).getJctLevelName());
				j++;
			}
			userVO.setJobLevel(jobLevelMap);
			LOGGER.info("<<<< AuthenticatingContoller.populateLevel");
		return userVO;
	}
	/**
	 * Method takes care of logout.
	 * @param body
	 * @param request
	 * @return null
	 */
	@RequestMapping(value = ACTION_SIGNUP_MASTER, method = RequestMethod.POST)
	@ResponseBody()
	public UserVO populateSignupMasterData(@RequestBody String body,
			HttpServletRequest request) {
		LOGGER.info(">>>> AuthenticatingContoller.populateSignupMasterData");
		UserVO userVO = new UserVO();
		try {
			userVO = populateRegion(userVO);
			userVO = populateFunction(userVO);
			userVO = populateLevel(userVO);
			
			//Fetch survey questions for general users if configured by admin
			userVO = authenticatorService.getSurveyQuestions(userVO);
			
			userVO.setStatusCode(CommonConstants.SUCCESSFUL);
			userVO.setStatusDesc("Success");
		} catch (JCTException e) {
			LOGGER.error(e.getMessage());
			userVO.setStatusCode(CommonConstants.DATA_NOT_FOUND);
			userVO.setStatusDesc(this.messageSource.getMessage(
					"error.noDataFound", null, null));
		}
		LOGGER.info("<<<< AuthenticatingContoller.populateSignupMasterData");
		return userVO;
	}
	
	@RequestMapping(value = ACTION_DECIDE_PAGE, method = RequestMethod.POST)
	@ResponseBody()
	public UserVO decidePage(@RequestBody String body,
			HttpServletRequest request) {
		LOGGER.info(">>>> AuthenticatingContoller.decidePage");
		UserVO userVO = new UserVO();
		JsonNode node = CommonUtility.getNode(body);
		userVO.setEmail(node.get("emailId").toString().replaceAll("\"", ""));
		//emailId
		try {
			List<JctUser> userList = authenticatorService.checkInactiveUser(userVO);
			if (userList.size() > 0) {
				JctUser user = (JctUser) userList.get(0);
				JctUserDetails dtls = user.getJctUserDetails();
				userVO.setFirstName(dtls.getJctUserDetailsFirstName());
				userVO.setLastName(dtls.getJctUserDetailsLastName());
				int accountExpirationCount = authenticatorService.getDaysToExpire(user.getJctUserId());
				userVO.setProfileId(user.getJctUserDetails().getJctUserDetailsProfileId());
				if (accountExpirationCount < 0) { // Account Expired
					LOGGER.info("ACCOUNT EXPIRED FOR THE USER: "+userVO.getEmail());
					userVO.setStatusCode(CommonConstants.USER_ACCOUNT_EXPIRED);
					userVO.setStatusMsg("accountExpired");
					userVO.setStatusDesc(this.messageSource.getMessage("error.account.expired",null, null));
					//userVO.setProfileId(user.getJctUserDetails().getJctUserDetailsProfileId());
				}  else {
					if ((accountExpirationCount == 0)) {
						String message = this.messageSource.getMessage("warning.account.expiration.today",null, null);
						userVO.setAccountExpWarning(message);
					} else if ((accountExpirationCount > 0) && (accountExpirationCount <=3)) {
						String message = this.messageSource.getMessage("warning.account.expiration1",null, null)+" "+accountExpirationCount+" "+this.messageSource.getMessage("warning.account.expiration2",null, null);
						userVO.setAccountExpWarning(message);
					} else {
						userVO.setAccountExpWarning("NA");
					}
					
					// Mandatory data available
					if ((null != user.getJctUserDetails().getJctUserDetailsLevels()) && (user.getJctUserDetails().getJctUserDetailsLevels().trim().length() > 0)) {
						userVO.setStatusCode(CommonConstants.USER_NOT_ACTIVATED);
						userVO.setStatusMsg("loginFailure");
						userVO.setStatusDesc(this.messageSource.getMessage("error.notActivated",null, null));
						userVO.setIsCompleted(9);
					}
					else {
						int activeStatus = user.getJctActiveYn();
						if (activeStatus == CommonConstants.CREATED) { // Admin created it
							userVO.setStatusCode(CommonConstants.USER_NOT_ACTIVATED);
							userVO.setStatusMsg("loginFailure");
							userVO.setStatusDesc(this.messageSource.getMessage("error.notActivated",null, null));
							userVO.setIsCompleted(1);
						} else if (activeStatus == CommonConstants.ACTIVATED_ONLY_EMAIL) { //user gave only email
							LOGGER.info("----> AUTHENTICATED "+userVO.getEmail()+". STATUS: CREATED [NOT ACTIVATED]");
							userVO.setStatusCode(CommonConstants.USER_NOT_ACTIVATED);
							userVO.setStatusMsg("loginFailure");
							userVO.setStatusDesc(this.messageSource.getMessage("error.notActivated",null, null));
							userVO.setIsCompleted(2);
						}
					}
				}
			} else {
				List<JctUser> lst = authenticatorService.checkUser(userVO);
				if (lst.size() == 0) {
					userVO.setIsCompleted(99); // no user exist
				}
			}
			
		} catch (JCTException e) {
			LOGGER.info(e.getLocalizedMessage());
		}
		LOGGER.info("<<<< AuthenticatingContoller.decidePage");
		return userVO;
	}
	/**
	 * Method populates the user details.
	 * @param body
	 * @param request
	 * @return UserVO
	 */
	@RequestMapping(value = ACTION_POPULATE_USER_DETAILS, method = RequestMethod.POST)
	@ResponseBody()
	public UserVO populateUserDetails(@RequestBody String body,
			HttpServletRequest request) {
		LOGGER.info(">>>> AuthenticatingContoller.populateUserDetails");
		UserVO userVO = new UserVO();
		JsonNode node = CommonUtility.getNode(body);
		String emailId = node.get("emailId").asText();
		try {
			//userVO = populateRegion(userVO);
			//userVO = populateFunction(userVO);
			//userVO = populateLevel(userVO);
			//userVO = populateOccupationData(emailId);						
			userVO.setOccupationVal(authenticatorService.getOccupationData(emailId));
			userVO.setAccountVO(searchService.getAccountDetails(emailId));
			userVO.setStatusCode(CommonConstants.SUCCESSFUL);
			userVO.setStatusDesc(this.messageSource.getMessage("label.your.demographic.details", null, null));
		} catch (JCTException e) {
			LOGGER.error(e.getMessage());
			userVO.setStatusCode(CommonConstants.DATA_NOT_FOUND);
			userVO.setStatusDesc(this.messageSource.getMessage(
					"error.noDataFound", null, null));
		}
		LOGGER.info("<<<< AuthenticatingContoller.populateUserDetails");
		return userVO;
	}
	/**
	 * Method populates the user details.
	 * @param body
	 * @param request
	 * @return UserVO
	 */
	@RequestMapping(value = ACTION_UPDATE_DEMO_INFO, method = RequestMethod.POST)
	@ResponseBody()
	public UserVO updateDemographicalInfo(@RequestBody String body,
			HttpServletRequest request) {
		LOGGER.info(">>>> AuthenticatingContoller.updateDemographicalInfo");
		UserVO userVO = new UserVO();
		JsonNode node = CommonUtility.getNode(body);
		userVO.setOccupationVal(node.get("occupationData").asText());
		//userVO.setFunctionGrpSelected(node.get("functionGrps").asText());
		//userVO.setJobLevelSelected(node.get("jobLevels").asText());
		userVO.setEmail(node.get("emailIds").asText());
		try {
			int status = authenticatorService.updateDemographicInfo(userVO);
			if (status == 200) {
				userVO.setStatusCode(CommonConstants.SUCCESSFUL);
			} else if (status == 500) {
				userVO.setStatusCode(CommonConstants.FAILURE);
				userVO.setStatusDesc(this.messageSource.getMessage("label.unable.to.save", null, null));
			} else {
				userVO.setStatusCode(CommonConstants.FAILURE);
				userVO.setStatusDesc(this.messageSource.getMessage("label.unable.to.save.data.not.found", null, null));
			}
		} catch (JCTException e) {
			LOGGER.error(e.getMessage());
			userVO.setStatusCode(CommonConstants.DATA_NOT_FOUND);
			userVO.setStatusDesc(this.messageSource.getMessage("label.unable.to.save", null, null));
		}
		LOGGER.info("<<<< AuthenticatingContoller.updateDemographicalInfo");
		return userVO;
	}
	/**
	 * Method updates the login info last active page details when the user clicks on any link on final page.
	 * @param body
	 * @param request
	 * @return null
	 */
	@RequestMapping(value = ACTION_UPDATE_LOGIN_INFO_ROW, method = RequestMethod.POST)
	@ResponseBody()
	public ReturnVO updateLoginInfoRowData(@RequestBody String body,
			HttpServletRequest request) {
		LOGGER.info(">>>> AuthenticatingContoller.updateLoginInfoRowData");
		JsonNode node = CommonUtility.getNode(body);
		String jobReferenceNumber = node.get("jobReferenceString").toString().replaceAll("\"", "");
		int rowId = Integer.parseInt(node.get("rowId").toString().replaceAll("\"", ""));
		String newActivePage = node.get("newActivePage").toString().replaceAll("\"", "");
		try {
			String result = authenticatorService.updateLoginInfo(jobReferenceNumber, rowId, newActivePage);
			LOGGER.info("---------> UPDATION RESULT: "+result);
		} catch (JCTException e) {
			LOGGER.error(e.getMessage());
		}
		LOGGER.info("<<<< AuthenticatingContoller.updateLoginInfoRowData");
		return null;
	}
	
	
	/**
	 * Method authenticates the user and returns to before sketch page.
	 * This method only works when trying to access the tool from admin 
	 * panel (facilitator).
	 * @return UserVO
	 */
	@RequestMapping(value = ACTION_FACILITATOR_AUTHENTICATE, method = RequestMethod.POST)
	@ResponseBody()
	public UserVO authenticate2(@RequestBody String body, HttpServletRequest request) {
		LOGGER.info(">>>> AuthenticatingContoller.authenticate2");
		HttpSession httpSession = request.getSession();
		httpSession.setAttribute("isLoggedIn", true);
		httpSession.setAttribute("isLoggedPdfIn", true);
		UserVO userVO = new UserVO();
		JsonNode node = CommonUtility.getNode(body);
		userVO.setEmail(node.get("jctEmail").asText());
		try {
			int isCompleted = 0;
			String pageInfo = "";
			//List<JctUser> userList = authenticatorService.checkUser(userVO);
			List<JctUser> userList = authenticatorService.checkFacilitatorUser(userVO);
			if (userList.size() > 0) {
				JctUser user = (JctUser) userList.get(0);
				LOGGER.info("----> AUTHENTICATED "+userVO.getEmail()+". STATUS: ACTIVE");
				userVO.setStatusCode(CommonConstants.SUCCESSFUL);
				userVO.setStatusMsg("loginSuccess");
				userVO.setStatusDesc("Successfully Logged in");
				userVO.setJctUserId(user.getJctUserId());
				/** Added user profile and user group **/
				userVO.setProfileId(user.getJctUserDetails().getJctUserDetailsProfileId());
				userVO.setGroupId(user.getJctUserDetails().getJctUserDetailsGroupId());
				userVO.setJobTitle("Facilitator");
				userVO.setFirstName("Facilitator");
				String jobRefNo = "";
				jobRefNo = "00000000ADPRV"+user.getJctUserId();
				userVO.setBsJson("none");
				userVO.setInitialJson("none");
				userVO.setTotalTime("00:00:00");
				userVO.setUrl("/user/view/beforeSketch.jsp");
				//userVO.setUrl("/user/view/landing-page.jsp");
				userVO.setIsCompleted(0);
				userVO.setLastPage("none");
				//pageInfo = "/user/view/beforeSketch.jsp";
				pageInfo = "/user/view/landing-page.jsp";
				userVO.setJobRefNo(jobRefNo);
				userVO.setPageSequence(getPageSequence(pageInfo));
				userVO.setIsCompleted(isCompleted);
				userVO.setIdentifier("0");
				userVO.setExcCompletionCount(0);
				userVO.setSource("Admin");
				
				// Delete all the existing records with the generated job Reference number
				authenticatorService.deleteDummyData(jobRefNo);
				
				/** THIS IS DONE FOR THE PUBLIC VERSION **/
				freezeContent(jobRefNo, user.getJctUserEmail(), user.getJctUserId(), user.getJctUserDetails().getJctUserDetailsProfileId());
				/*****************************************/
				
				userVO = insertLoginInfo(userVO, jobRefNo, user.getJctUserId(), pageInfo, isCompleted);
			} 
		} catch (Exception connExe) {
			LOGGER.error("----> AUTHENTICATION FAILED FOR: "+userVO.getEmail()+". PROBLEM WITH DATABASE: "+connExe.getMessage());
			userVO.setStatusCode(CommonConstants.DATABASE_NOT_FOUND);
			userVO.setStatusMsg("loginFailure");
			userVO.setStatusDesc(this.messageSource.getMessage("error.remortResourceUnreachable",null, null));
		}
		LOGGER.info("<<<< AuthenticatingContoller.authenticate");
		return userVO;
	}
	
	private void freezeContent(String jrNo, String user, int userId,
			int profileId) {
		LOGGER.info(">>>> AuthenticatingContoller.freezeContent");
		try {
			// Before Sketch questionnaire...
			// Check if empty only then
			int countQuestionnaire = iQuestionnaireService.getQuestCount(jrNo, user, userId, profileId);
			if (countQuestionnaire == 0) {
				iQuestionnaireService.freezeQuestionnaireContent(jrNo, user, userId, profileId);
			}
			
			//Action Plan
			int countActionPlan = actionPlanService.getActionPlanCount(jrNo, user, userId, profileId);
			if (countActionPlan == 0) {
				actionPlanService.freezeActionPlanContent(jrNo, user, userId, profileId);
			}
			
			//Value, Passion and Strength
			attributeFrozenService.freezeContent(jrNo, user, profileId, userId);
		} catch (JCTException ex) {
			LOGGER.error("UNABLE TO FREEZE THE CONTENT. REASON: "+ex.getLocalizedMessage());
		}
		LOGGER.info("<<<< AuthenticatingContoller.freezeContent");
	}
	
	/**
	 * Method handles reset password functionality
	 * @param body
	 * @param request
	 * @return
	 */
	@RequestMapping(value = ACTION_SAVE_SURVEY, method = RequestMethod.POST)
	@ResponseBody()
	public ReturnVO saveSurveyQuestions(@RequestBody String body,
			HttpServletRequest request) {
		ReturnVO vo = new ReturnVO();
		vo.setStatusCode("500");
		JsonNode node = CommonUtility.getNode(body);
		int userId = node.get("userId").asInt();
		String emailId = node.get("emailIds").asText();
		/** THIS IS DONE FOR THE PUBLIC VERSION **/
		// Survey questions
		String surveyQtnTextString = node.get("surveyQtnTextString").asText();
		String surveyQtnRadioString = node.get("surveyQtnRadioString").asText();
		String surveyQtnCheckBoxString = node.get("surveyQtnCheckBoxString").asText();
		String surveyQtnDropdownString = node.get("surveyQtnDropdownString").asText();
		/*****************************************/
		processSurveyQuestions(surveyQtnTextString, surveyQtnRadioString, surveyQtnCheckBoxString, surveyQtnDropdownString, emailId, userId);
		vo.setStatusCode("200");
		return vo;
	}
	/**
	 * Method searches the occupation list
	 * @param body
	 * @param request
	 * @return
	 */
	@RequestMapping(value = ACTION_SEARCH_OCCUPTN_LIST, method = RequestMethod.POST)
	@ResponseBody()
	public ReturnVO searchOccupationList(@RequestBody String body,
			HttpServletRequest request) {
		LOGGER.info(">>>>>>>> AuthenticatingContoller.searchOccupationList");
		ReturnVO vo = new ReturnVO();
		vo.setStatusCode(CommonConstants.FAILURE);
		JsonNode node = CommonUtility.getNode(body);
		String searchString = node.get("searchString").asText();
		try {
			List<OccupationListDTO> list = authenticatorService.searchOccupationList(searchString);
			if (list.size() > 0) {
				vo.setStatusCode(CommonConstants.SUCCESSFUL);
			} else {
				vo.setStatusCode(CommonConstants.FAILURE);
				vo.setStatusMsg(this.messageSource.getMessage("error.onet.searched.not.found",null, null));
			}
			vo.setDtoList(list);
		} catch (JCTException e) {
			LOGGER.error("UNABLE TO FETCH OCCUPATION LIST");
		}
		LOGGER.info("<<<<<<<< AuthenticatingContoller.searchOccupationList");
		return vo;
	}
	/**
	 * Method returns shopify link
	 * @param body
	 * @param request
	 * @return
	 */
	@RequestMapping(value = ACTION_GET_SHOPIFY_LINK, method = RequestMethod.POST)
	@ResponseBody()
	public ShopifyVO getShopifyLink(@RequestBody String body,
			HttpServletRequest request) {
		LOGGER.info(">>>>>>>> AuthenticatingContoller.getShopifyLink");
		ShopifyVO vo = new ShopifyVO();
		StringBuffer sb = new StringBuffer("");
		sb.append(this.messageSource.getMessage("shopify.api.key",null, null)); //API KEY
		sb.append(":"); //DELIMETER
		sb.append(this.messageSource.getMessage("shopify.api.pwd",null, null)); //API PASSWORD
		sb.append("@");
		sb.append(this.messageSource.getMessage("shopify.ind.renew.url",null, null)); //SERVIE URL
		vo.setShopifyFormUrl(sb.toString());
		vo.setProductVarient(this.messageSource.getMessage("shopify.ind-renew.product.varient",null, null)); // VARIENT
		LOGGER.info("<<<<<<<< AuthenticatingContoller.getShopifyLink");
		return vo;
	}

	/**
	 * Method to fetch TnC by user profile
	 * @param body
	 * @param request
	 * @return
	 */
	@RequestMapping(value = ACTION_GET_TERMS_AND_CONDITIONS, method = RequestMethod.POST)
	@ResponseBody()
	public TermsAndConditionsVO fetchTnC(@RequestBody String body,
			HttpServletRequest request) {
		LOGGER.info(">>>>>>>> AuthenticatingContoller.fetchTnC");
		JsonNode node = CommonUtility.getNode(body);
		int userProfile = node.get("userProfile").asInt();
		int userType = node.get("userType").asInt();
		TermsAndConditionsVO termsAndConditionsVO = new TermsAndConditionsVO();
		String tnC = null;
		try {
			tnC = authenticatorService.getTermsAndConditions(userProfile, userType);		
			if(null != tnC){
				termsAndConditionsVO.setJctTcDescription(tnC);
				termsAndConditionsVO.setStatusCode(Integer.parseInt(CommonConstants.SUCCESSFUL));
				termsAndConditionsVO.setStatusDesc(this.messageSource.getMessage("label.tnC.found",null, null));
			} else {
				termsAndConditionsVO.setStatusCode(Integer.parseInt(CommonConstants.DATA_NOT_FOUND));
				termsAndConditionsVO.setStatusDesc(this.messageSource.getMessage("label.tnC.noTC",null, null));				
			}
		} catch (JCTException e) {
			LOGGER.error("UNABLE TO FETCH TnC for USER PROFILE ID: "+userProfile+", ---- "+e.getLocalizedMessage());
			e.printStackTrace();
		}
		LOGGER.info("<<<<<<<< AuthenticatingContoller.fetchTnC");
		return termsAndConditionsVO;
	}
	
	@RequestMapping(value = ACTION_SILENT_TIME_UPDATE, method = RequestMethod.POST)
	@ResponseBody()
	public void silentTimeUpdate (@RequestBody String body,
			HttpServletRequest request) {
		LOGGER.info(">>>>>>>> AuthenticatingContoller.silentTimeUpdate");
		JsonNode node = CommonUtility.getNode(body);
		try {
			authenticatorService.silentUpdateTime(node.get("rowId").asInt());
		} catch (JCTException e) {
			
		}
		LOGGER.info("<<<<<<<< AuthenticatingContoller.silentTimeUpdate");
	}
}