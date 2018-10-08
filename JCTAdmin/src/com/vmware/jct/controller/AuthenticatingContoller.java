package com.vmware.jct.controller;

import java.util.List;

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
import com.vmware.jct.dao.dto.OccupationListDTO;
import com.vmware.jct.exception.JCTException;
import com.vmware.jct.model.JctUser;
import com.vmware.jct.service.IAdminMenuService;
import com.vmware.jct.service.IAuthenticatorService;
import com.vmware.jct.service.ISurveyQuestionService;
import com.vmware.jct.service.vo.ReturnVO;
import com.vmware.jct.service.vo.ShopifyVO;
import com.vmware.jct.service.vo.TermsAndConditionsVO;
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
@RequestMapping(value="/authAdmin")
public class AuthenticatingContoller extends BasicController {
	
	@Autowired
	private IAuthenticatorService authenticatorService;
	
	@Autowired
	private IAdminMenuService adminMenuService;
	
	@Autowired
	private ISurveyQuestionService surveyQuestionService;
	
	@Autowired
	private MailNotificationHelper mailer;
	
	@Autowired  
	private MessageSource messageSource;
	
	private static final Logger logger = LoggerFactory.getLogger(AuthenticatingContoller.class);
	
	/**
	 * Method authenticates the admin user and returns to the specified 
	 * page.
	 * @return UserVO
	 */
	@RequestMapping(value = ACTION_AUTHENTICATE, method = RequestMethod.POST)
	@ResponseBody()
	public UserVO authenticateAdmin(@RequestBody String body, HttpServletRequest request) {
		logger.info(">>>> AuthenticatingContoller.authenticateAdmin");
		HttpSession httpSession = request.getSession();
		UserVO userVO = new UserVO();
		JsonNode node = CommonUtility.getNode(body);
		userVO.setEmail(node.get("emailId").asText());
		userVO.setPassword(node.get("password").asText());
		logger.info("----> AUTHENTICATION STARTED FOR: "+userVO.getEmail());
		try {
			List<JctUser> userList = authenticatorService.authenticateUser(userVO);
			if (userList.size() > 0) {
				JctUser user = (JctUser) userList.get(0);
				int activeStatus = user.getJctActiveYn();								
				if (activeStatus == CommonConstants.ACTIVATED) {
					httpSession.setAttribute("isLoggedIn", true);
					logger.info("----> AUTHENTICATED "+userVO.getEmail()+". STATUS: ACTIVE");
					userVO.setStatusCode(CommonConstants.SUCCESSFUL);
					userVO.setStatusMsg("loginSuccess");
					userVO.setStatusDesc("Successfully Logged in");
					
					
					// CHANGE
					/*userVO.setJobTitle(user.getJctUserLevels());
					userVO.setFirstName(user.getJctFirstName());*/
					userVO.setJobTitle(user.getJctUserDetails().getJctUserDetailsLevels());
					userVO.setFirstName(user.getJctUserDetails().getJctUserDetailsFirstName());
					userVO.setProfileId(user.getJctUserDetails().getJctUserDetailsProfileId());
					//Set the menu items
					userVO.setMenu(adminMenuService.getMenuString(user.getJctUserRole().getJctRoleId()));
					//set the user type
					userVO.setRoleId(user.getJctUserRole().getJctRoleId());
					//Set the un-authorization message
					userVO.setUnauthMsg(this.messageSource.getMessage("jct.not.auth",null, null));
					//Set the interface title
					if (userVO.getRoleId() == 2) {
						userVO.setInterfaceTitleStr(this.messageSource.getMessage("label.job.crafting.panel",null, null));
					} else {
						userVO.setInterfaceTitleStr(this.messageSource.getMessage("label.job.crafting.facilitator.panel",null, null));
					}
					//set customer Id
					userVO.setCustomerId(user.getJctUserCustomerId());
					userVO.setActiveYn(2);
					userVO.setJctUserId(user.getJctUserId());
				} else if (activeStatus == CommonConstants.CREATED){  // For Facilitator First Log in
					/**
					 * CHANGE FOR JCT PUBLIC VERSION
					 */
					httpSession.setAttribute("isLoggedIn", true);
					userVO.setStatusCode(CommonConstants.SUCCESSFUL);
					userVO.setStatusMsg("loginSuccess");
					userVO.setStatusDesc("Successfully Logged in");
										
					userVO.setJobTitle(user.getJctUserDetails().getJctUserDetailsLevels());
					userVO.setFirstName(user.getJctUserDetails().getJctUserDetailsFirstName());
					userVO.setProfileId(user.getJctUserDetails().getJctUserDetailsProfileId());
					//Set the menu items
					userVO.setMenu(adminMenuService.getMenuString(user.getJctUserRole().getJctRoleId()));
					//set the user type
					userVO.setRoleId(user.getJctUserRole().getJctRoleId());
					//Set the un-authorization message
					userVO.setUnauthMsg(this.messageSource.getMessage("jct.not.auth",null, null));
					//Set the interface title
					if (userVO.getRoleId() == 2) {
						userVO.setInterfaceTitleStr(this.messageSource.getMessage("label.job.crafting.panel",null, null));
					} else {
						userVO.setInterfaceTitleStr(this.messageSource.getMessage("label.job.crafting.facilitator.panel",null, null));
					}
					//set customer Id
					userVO.setCustomerId(user.getJctUserCustomerId());
					userVO.setActiveYn(1);			
					userVO.setJctUserId(user.getJctUserId());
				} else if (activeStatus == CommonConstants.CHECK_PAYMENT_NOT_ACTIVATED){ // CHEQUE_PAYMENT_NOT_PROCESSED
					//set the user type
					userVO.setRoleId(user.getJctUserRole().getJctRoleId());
					userVO.setStatusCode(CommonConstants.USUER_CHEQUE_NOT_CLEARED);
					userVO.setStatusDesc(this.messageSource.getMessage("error.userChequeProblem",null, null));
					userVO.setActiveYn(1);
				} else if (activeStatus == CommonConstants.FACI_ACCOUNT_JUST_CREATED){ // FACILITATOR ACCOUNT JUST CREATED
					httpSession.setAttribute("isLoggedIn", true);
					userVO.setRoleId(user.getJctUserRole().getJctRoleId());
					userVO.setStatusCode(CommonConstants.FACI_ACCOUNT_JUST_CREATED_STRING);
					userVO.setStatusDesc(this.messageSource.getMessage("error.userChequeProblem",null, null));
					userVO.setActiveYn(10);
				} else if (activeStatus == CommonConstants.FACI_PASSWORD_RESET){ // FACILITATOR PASSWORD RESET
					httpSession.setAttribute("isLoggedIn", true);
					userVO.setRoleId(user.getJctUserRole().getJctRoleId());
					userVO.setStatusCode(CommonConstants.FACI_PASSWORD_RESET_STRING);
					userVO.setStatusDesc("Password Reset");
					userVO.setActiveYn(11);
				}
			} else {
				logger.error("----> AUTHENTICATION FAILED FOR: "+userVO.getEmail()+". USER NOT FOUND");
				userVO.setStatusCode(CommonConstants.USER_NOT_FOUND);
				userVO.setStatusMsg("loginFailure");
				userVO.setStatusDesc(this.messageSource.getMessage("error.userPasswordNotMatch",null, null));
			}
		} catch (Exception connExe) {
			logger.error("----> AUTHENTICATION FAILED FOR: "+userVO.getEmail()+". PROBLEM WITH DATABASE: "+connExe.getMessage());
			userVO.setStatusCode(CommonConstants.DATABASE_NOT_FOUND);
			userVO.setStatusMsg("loginFailure");
			userVO.setStatusDesc(this.messageSource.getMessage("error.remortResourceUnreachable",null, null));
		}
		logger.info("<<<< AuthenticatingContoller.authenticateAdmin");
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
		logger.info(">>>> AuthenticatingContoller.checkUser");
		UserVO userVO = new UserVO();
		//JsonNode node = getNode(body);
		JsonNode node = CommonUtility.getNode(body);
		userVO.setEmail(node.get("emailId").toString().replaceAll("\"", ""));
		logger.info("----> CHECKING USER: "+userVO.getEmail());
		try {
			List<JctUser> userList = authenticatorService.checkUser(userVO);
			if (userList.size() > 0) {
				logger.info("----> USER: "+userVO.getEmail()+" EXISTS...");
				userVO.setStatusMsg("userExists");
				userVO.setStatusDesc(this.messageSource.getMessage("error.alreadyRegistered",null, null));
			} else {
				userVO.setStatusMsg("userDoesntExist");
				userVO.setStatusDesc("");
			}
		} catch (JCTException jctExe) {
			logger.error(jctExe.getMessage());
			userVO.setStatusCode(CommonConstants.DATABASE_NOT_FOUND);
			userVO.setStatusMsg("userDoesntExist");
			userVO.setStatusDesc(this.messageSource.getMessage("error.remortResourceUnreachable",null, null));
		} catch (Exception connExe) {
			logger.error(connExe.getMessage());
			userVO.setStatusCode(CommonConstants.DATABASE_NOT_FOUND);
			userVO.setStatusMsg("userDoesntExist");
			userVO.setStatusDesc(this.messageSource.getMessage("error.remortResourceUnreachable",null, null));
		}
		logger.info("<<<< AuthenticatingContoller.checkUser");
		return userVO;
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
		logger.info(">>>> AuthenticatingContoller.logout");
		HttpSession httpSession = request.getSession();
		//JsonNode node = getNode(body);
		//JsonNode node = CommonUtility.getNode(body);
		//String jobReferenceNumber = node.get("jobReferenceString").toString().replaceAll("\"", "");
		//int rowId = Integer.parseInt(node.get("rowId").toString().replaceAll("\"", ""));
		//String lastActivePage = node.get("lastActivePage").toString().replaceAll("\"", "");
		//try {
		//	String result = authenticatorService.updateLoginInfo(jobReferenceNumber, rowId, lastActivePage);
		httpSession.removeAttribute("isLoggedIn");
		//	logger.info("---------> LOGOUT: "+result);
		//} catch (JCTException e) {
		//	logger.error(e.getMessage());
		//}
		logger.info("<<<< AuthenticatingContoller.logout");
		return null;
	}
	
	/**
	 * Method populates master data on user's first login.
	 * @param body
	 * @param request
	 * @return null
	 */
	@RequestMapping(value = ACTION_SIGNUP_MASTER, method = RequestMethod.POST)
	@ResponseBody()
	public UserVO populateSignupMasterData(@RequestBody String body,
			HttpServletRequest request) {
		logger.info(">>>> AuthenticatingContoller.populateSignupMasterData");
		UserVO userVO = new UserVO();
		try {				
			//Fetch survey questions for general users if configured by admin
			userVO = authenticatorService.getSurveyQuestions(userVO);		
			userVO.setStatusCode(CommonConstants.SUCCESSFUL);
			userVO.setStatusDesc("Success");
		} catch (JCTException e) {
			logger.error(e.getMessage());
			userVO.setStatusCode(CommonConstants.DATA_NOT_FOUND);
			userVO.setStatusDesc(this.messageSource.getMessage(
					"error.noDataFound", null, null));
		}
		logger.info("<<<< AuthenticatingContoller.populateSignupMasterData");
		return userVO;
	}
	
	/**
	 * Method save survey answers for facilitator
	 * @param body
	 * @param request
	 * @return
	 */
	@RequestMapping(value = ACTION_SAVE_SURVEY_ANSWER, method = RequestMethod.POST)
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
		logger.info(">>>>>>>> AuthenticatingContoller.processSurveyQuestions");
		try {
			surveyQuestionService.saveSurveyQuestions(surveyQtnTextString, surveyQtnRadioString, 
					surveyQtnDropdownString, surveyQtnCheckBoxString, emailId, userId);
		} catch (JCTException ex) {
			logger.error("UNABLE TO SAVE SURVEY QUESTIONS: "+ex.getLocalizedMessage());
		}
		logger.info("<<<<<<<< AuthenticatingContoller.processSurveyQuestions");
	}
	
	/**
	 * Method update active status for facilitator
	 * @param body
	 * @param request
	 * @return
	 */
	@RequestMapping(value = ACTION_UPDATE_ACTIVE_STATUS_FACILITATOR, method = RequestMethod.POST)
	@ResponseBody()
	public ReturnVO updateActiveStatusFacilitator(@RequestBody String body,
			HttpServletRequest request) {
		ReturnVO vo = new ReturnVO();
		JsonNode node = CommonUtility.getNode(body);
		int userId = node.get("userId").asInt();
		String emailId = node.get("emailIds").asText();	
		try {
			surveyQuestionService.updateActiveStatusFacilitator(emailId, userId);
			vo.setStatusCode(CommonConstants.SUCCESSFUL);
			vo.setStatusDesc("Success");
		} catch (JCTException ex) {
			vo.setStatusCode(CommonConstants.USER_NOT_FOUND);
			logger.error("UNABLE TO SAVE SURVEY QUESTIONS: "+ex.getLocalizedMessage());
		}
		
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
		logger.info(">>>>>>>> AuthenticatingContoller.getShopifyLink");
		ShopifyVO vo = new ShopifyVO();
		StringBuffer sb = new StringBuffer("");
		sb.append(this.messageSource.getMessage("shopify.api.key",null, null)); //API KEY
		sb.append(":"); //DELIMETER
		sb.append(this.messageSource.getMessage("shopify.api.pwd",null, null)); //API PASSWORD
		sb.append("@");
		sb.append(this.messageSource.getMessage("shopify.ind.renew.url",null, null)); //SERVIE URL
		vo.setShopifyFormUrl(sb.toString());
		vo.setProductVarient(this.messageSource.getMessage("shopify.ind-renew.product.varient",null, null)); // VARIENT
		logger.info("<<<<<<<< AuthenticatingContoller.getShopifyLink");
		return vo;
	}
	
	
	/**
	 * Method searches the occupation list
	 * @param body
	 * @param request
	 * @return
	 */
	@RequestMapping(value = "/searchOccupationList", method = RequestMethod.POST)
	@ResponseBody()
	public ReturnVO searchOccupationList(@RequestBody String body,
			HttpServletRequest request) {
		logger.info(">>>>>>>> AuthenticatingContoller.searchOccupationList");
		ReturnVO vo = new ReturnVO();
		vo.setStatusCode("500");
		JsonNode node = CommonUtility.getNode(body);
		String searchString = node.get("searchString").asText();
		try {
			List<OccupationListDTO> list = authenticatorService.searchOccupationList(searchString);
			if (list.size() > 0) {
				vo.setStatusCode(CommonConstants.SUCCESSFUL);
			} else {
				vo.setStatusCode("500");
				vo.setStatusMsg(this.messageSource.getMessage("error.onet.searched.not.found",null, null));
			}
			vo.setDtoList(list);
		} catch (JCTException e) {
			logger.error("UNABLE TO FETCH OCCUPATION LIST");
		}
		logger.info("<<<<<<<< AuthenticatingContoller.searchOccupationList");
		return vo;
	}
	
	@RequestMapping(value = "/resetFacilitatorInitialPwd", method = RequestMethod.POST)
	@ResponseBody()
	public UserVO resetFacilitatorInitialPwd(@RequestBody String body, HttpServletRequest request) {
		logger.info(">>>> AuthenticatingContoller.authenticateAdmin");
		HttpSession httpSession = request.getSession();
		UserVO userVO = new UserVO();
		JsonNode node = CommonUtility.getNode(body);
		String emailId = node.get("forUser").asText();
		String initialPassword = node.get("initialPassword").asText();
		String userPassword = node.get("userPassword").asText();
		//String mailedPwd = node.get("mailedPwd").asText();
		// Check if username and password exists
		userVO.setEmail(emailId);
		userVO.setPassword(initialPassword);
		logger.info("----> AUTHENTICATION STARTED FOR: "+emailId);
		try {
			List<JctUser> userList2 = authenticatorService.authenticateUser(userVO);
			if (userList2.size() > 0) {
				if(userList2.get(0).getJctActiveYn() == 10) {
					userVO.setStatusCodeInt(10);
				} else {
					userVO.setStatusCodeInt(11);
				}
				String status = authenticatorService.updateUser(userPassword, 2, emailId);
				if (status.equals("success")) {
					userVO.setEmail(emailId);
					userVO.setPassword(userPassword);
					try {
						List<JctUser> userList = authenticatorService.authenticateUser(userVO);
						if (userList.size() > 0) {
							JctUser user = (JctUser) userList.get(0);
							int activeStatus = user.getJctActiveYn();								
							if (activeStatus == CommonConstants.ACTIVATED) {
								httpSession.setAttribute("isLoggedIn", true);
								logger.info("----> AUTHENTICATED "+userVO.getEmail()+". STATUS: ACTIVE");
								userVO.setStatusCode(CommonConstants.SUCCESSFUL);
								userVO.setStatusMsg("loginSuccess");
								userVO.setStatusDesc("Successfully Logged in");
								userVO.setJobTitle(user.getJctUserDetails().getJctUserDetailsLevels());
								userVO.setFirstName(user.getJctUserDetails().getJctUserDetailsFirstName());
								userVO.setProfileId(user.getJctUserDetails().getJctUserDetailsProfileId());
								
								//Set the menu items
								userVO.setMenu(adminMenuService.getMenuString(user.getJctUserRole().getJctRoleId()));
								//set the user type
								userVO.setRoleId(user.getJctUserRole().getJctRoleId());
								//Set the un-authorization message
								userVO.setUnauthMsg(this.messageSource.getMessage("jct.not.auth",null, null));
								//Set the interface title
								if (userVO.getRoleId() == 2) {
									userVO.setInterfaceTitleStr(this.messageSource.getMessage("label.job.crafting.panel",null, null));
								} else {
									userVO.setInterfaceTitleStr(this.messageSource.getMessage("label.job.crafting.facilitator.panel",null, null));
								}
								//set customer Id
								userVO.setCustomerId(user.getJctUserCustomerId());
								userVO.setActiveYn(2);
								userVO.setJctUserId(user.getJctUserId());
							} else if (activeStatus == CommonConstants.CHECK_PAYMENT_NOT_ACTIVATED){ // CHEQUE_PAYMENT_NOT_PROCESSED
								//set the user type
								userVO.setRoleId(user.getJctUserRole().getJctRoleId());
								userVO.setStatusCode(CommonConstants.USUER_CHEQUE_NOT_CLEARED);
								userVO.setStatusDesc(this.messageSource.getMessage("error.userChequeProblem",null, null));
								userVO.setActiveYn(1);
							} 
						} else {
							logger.error("----> AUTHENTICATION FAILED FOR: "+userVO.getEmail()+". USER NOT FOUND");
							userVO.setStatusCode(CommonConstants.USER_NOT_FOUND);
							userVO.setStatusMsg("loginFailure");
							userVO.setStatusDesc(this.messageSource.getMessage("error.userPasswordNotMatch",null, null));
						}
					} catch (Exception connExe) {
						logger.error("----> AUTHENTICATION FAILED FOR: "+userVO.getEmail()+". PROBLEM WITH DATABASE: "+connExe.getMessage());
						userVO.setStatusCode(CommonConstants.DATABASE_NOT_FOUND);
						userVO.setStatusMsg("loginFailure");
						userVO.setStatusDesc(this.messageSource.getMessage("error.remortResourceUnreachable",null, null));
					}
				}
			} else {
				logger.error("----> AUTHENTICATION FAILED FOR: "+userVO.getEmail()+". USER NOT FOUND");
				userVO.setStatusCode(CommonConstants.USER_NOT_FOUND);
				userVO.setStatusMsg("loginFailure");
				userVO.setStatusDesc(this.messageSource.getMessage("error.faci.system.pass",null, null));
			}
			// Merge the user object
			
		} catch (Exception we) {
			
		}
		return userVO;
	}
	
	@RequestMapping(value = "/decideDiv", method = RequestMethod.POST)
	@ResponseBody()
	public UserVO decideDiv(@RequestBody String body, HttpServletRequest request) {
		logger.info(">>>> AuthenticatingContoller.decideDiv");
		UserVO userVO = new UserVO();
		JsonNode node = CommonUtility.getNode(body);
		String emailId = node.get("email").asText();
		try {
			List<JctUser> list = authenticatorService.getUser(emailId);
			if (list.size() > 0) {
				userVO.setStatusCode("200");
			} else {
				userVO.setStatusCode("300");
			}
		} catch (Exception ex) {
			userVO.setStatusCode("300");
		}
		return userVO;
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
		logger.info(">>>>>>>> AuthenticatingContoller.fetchTnC");
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
			logger.error("UNABLE TO FETCH TnC for USER PROFILE ID: "+userProfile+", ---- "+e.getLocalizedMessage());
			e.printStackTrace();
		}
		logger.info("<<<<<<<< AuthenticatingContoller.fetchTnC");
		return termsAndConditionsVO;
	}
}
