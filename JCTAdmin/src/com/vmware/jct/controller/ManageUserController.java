package com.vmware.jct.controller;

import java.io.BufferedReader;
import java.io.File;
import java.io.FileReader;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Date;
import java.util.Iterator;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.Random;
import java.util.StringTokenizer;

import javax.naming.NamingException;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import javax.servlet.http.HttpSession;

import org.codehaus.jackson.JsonNode;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.context.MessageSource;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.PathVariable;
import org.springframework.web.bind.annotation.RequestBody;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestMethod;
import org.springframework.web.bind.annotation.ResponseBody;

import com.vmware.jct.common.utility.CommonConstants;
import com.vmware.jct.common.utility.CommonUtility;
import com.vmware.jct.common.utility.MailNotificationHelper;
import com.vmware.jct.common.utility.StatusConstants;
import com.vmware.jct.dao.dto.ExistingUserDTO;
import com.vmware.jct.dao.dto.FacilitatorDetailDTO;
import com.vmware.jct.dao.dto.NewUserDTO;
import com.vmware.jct.dao.dto.PaymentDetailsDTO;
import com.vmware.jct.dao.dto.UserAccountDTO;
import com.vmware.jct.dao.dto.UserGroupDTO;
import com.vmware.jct.exception.JCTException;
import com.vmware.jct.exception.MailingException;
import com.vmware.jct.model.JctUser;
import com.vmware.jct.service.IManageUserService;
import com.vmware.jct.service.vo.ConfirmationMessageVO;
import com.vmware.jct.service.vo.CrossBrowserUserQueryVO;
import com.vmware.jct.service.vo.ExistingFacilitatorVO;
import com.vmware.jct.service.vo.ExistingUsersVO;
import com.vmware.jct.service.vo.FacilitatorAccountVO;
import com.vmware.jct.service.vo.GeneralUserAccountVO;
import com.vmware.jct.service.vo.ManageUserVO;
import com.vmware.jct.service.vo.PaymentVO;
import com.vmware.jct.service.vo.ReportsVO;
import com.vmware.jct.service.vo.ReturnVO;
import com.vmware.jct.service.vo.UserGroupVO;
import com.vmware.jct.service.vo.UserVO;

/**
 * 
 * <p><b>Class name:</b> ManageUserControllers.java</p>
 * <p><b>Author:</b> InterraIT</p>
 * <p><b>Purpose:</b> This class acts as a controller for all User Management on Admin Side..
 * ManageUserController extends BasicController  and has following  methods.
 * -populateExistingUserProfile()
 * <p><b>Description:</b> This class acts as a controller for all User Management on Admin Side. </p>
 * <p><b>Copyrights:</b> All rights reserved by InterraIT and should only be used for its internal application development.</p>
 * <p><b>Created Date:</b> 26/Mar/2014</p>
 * <p><b>Revision History:</b>
 * 	<li></li>
 * </p>
 */
@Controller
@RequestMapping(value = "/manageuser")
public class ManageUserController extends BasicController {
	
	
	@Autowired
	private IManageUserService service;
	
	@Autowired  
	private MessageSource messageSource;
	
	@Autowired
	private MailNotificationHelper mailer;
	
	private static final Logger LOGGER = LoggerFactory.getLogger(ManageUserController.class);
	
	/**
	 * Method populates the User Profile List.
	 * @param body
	 * @param request
	 * @return ManageUserVO containing user profiles.
	 */
	@RequestMapping(value = ACTION_POPULATE_EXTNG_USR_RPFL, method = RequestMethod.POST)
	@ResponseBody()
	public ManageUserVO populateExistingUserProfile(@RequestBody String body, HttpServletRequest request) {
		LOGGER.info(">>>>>> ManageUserController.populateExistingUserProfile");
		ManageUserVO manageVO = null;
		try {
			manageVO = service.populateExistingUserProfile();
		} catch (JCTException e) {
			manageVO = new ManageUserVO();
			manageVO.setStatusCode(StatusConstants.STATUS_FAILURE);
			manageVO.setStatusDesc(this.messageSource.getMessage("exception.dao.error",null, null));
			LOGGER.error(e.getLocalizedMessage());
		}
		LOGGER.info("<<<<<< ManageUserController.populateExistingUserProfile");
		return manageVO;
	}
	/**
	 * Method saves new user profile
	 * @param body
	 * @param request
	 * @return return status code only.
	 */
	@RequestMapping(value = ACTION_SAVE_USER_PROFILE, method = RequestMethod.POST)
	@ResponseBody()
	public ManageUserVO saveUserProfile(@RequestBody String body, HttpServletRequest request) throws JCTException {
		LOGGER.info(">>>>>> ManageUserController.saveUserProfile");
		ManageUserVO manageVO = new ManageUserVO();
		JsonNode node = CommonUtility.getNode(body);
		String userProfile = node.get("userProfileVal").toString().replaceAll("\"" , "").toUpperCase().trim();
		String createdBy = node.get("createdBy").toString().replaceAll("\"" , "");		
		String updationSuccess = service.validateExistence(userProfile);
		if (updationSuccess.equals("success")) {
			manageVO.setStatusCode(StatusConstants.ALREADY_EXISTS);
		} else {
			manageVO.setStatusCode(StatusConstants.DOES_NOT_EXIST);
			try {
				manageVO = service.saveUserProfile(node.get("userProfileVal").toString().replaceAll("\"" , "").trim(), createdBy);
				ManageUserVO unused = service.populateExistingUserProfile();
				manageVO.setExistingUserProfileList(unused.getExistingUserProfileList());
				unused = null;
			} catch (JCTException e) {
				manageVO = new ManageUserVO();
				manageVO.setStatusCode(StatusConstants.STATUS_FAILURE);
				manageVO.setStatusDesc(this.messageSource.getMessage("exception.dao.user.profile.save",null, null));
				LOGGER.error(e.getLocalizedMessage());
			} 
		}		
		LOGGER.info("<<<<<< ManageUserController.saveUserProfile");
		return manageVO;
	}
	/**
	 * Method populates the User Group List as well as User profile list.
	 * @param body
	 * @param request
	 * @return ManageUserVO containing user profiles.
	 */
	@RequestMapping(value = ACTION_POPULATE_EXTNG_USR_GRP, method = RequestMethod.POST)
	@ResponseBody()
	public ManageUserVO populateExistingUserGroup(@RequestBody String body, HttpServletRequest request) {
		LOGGER.info(">>>>>> ManageUserController.populateExistingUserGroup");
		JsonNode node = CommonUtility.getNode(body);
		int profileId = Integer.parseInt(node.get("profileId").toString().replaceAll("\"" , ""));
		String customerId = node.get("customerId").toString().replaceAll("\"" , "");
		int roleId = Integer.parseInt(node.get("roleId").toString().replaceAll("\"" , ""));
		ManageUserVO manageVO = new ManageUserVO();
		try {
			manageVO.setUserProfileMap(this.service.populateUserProfile());			
			ManageUserVO unused = this.service.populateExistingUserGroup(profileId, customerId, roleId);
			manageVO.setExistingUserGroupList(unused.getExistingUserGroupList());
			unused = null;
			manageVO.setStatusCode(StatusConstants.STATUS_SUCCESS_WITH_RESULT);
		} catch (JCTException e) {
			manageVO = new ManageUserVO();
			manageVO.setStatusCode(StatusConstants.STATUS_FAILURE);
			manageVO.setStatusDesc(this.messageSource.getMessage("exception.dao.error",null, null));
			LOGGER.error(e.getLocalizedMessage());
		}
		LOGGER.info("<<<<<< ManageUserController.populateExistingUserGroup");
		return manageVO;
	}
	/**
	 * Method populates the User Group List.
	 * @param body
	 * @param request
	 * @return ManageUserVO containing user profiles.
	 */
	@RequestMapping(value = ACTION_POPULATE_USER_PROFILE, method = RequestMethod.POST)
	@ResponseBody()
	public ManageUserVO populateUserProfile(@RequestBody String body, HttpServletRequest request) {
		LOGGER.info(">>>>>> ManageUserController.populateUserProfile");
		ManageUserVO manageVO = new ManageUserVO();
		try {
			manageVO.setUserProfileMap(service.populateUserProfile());
			manageVO.setStatusCode(StatusConstants.STATUS_SUCCESS_WITH_RESULT);
		} catch (JCTException e) {
			manageVO = new ManageUserVO();
			manageVO.setStatusCode(StatusConstants.STATUS_FAILURE);
			manageVO.setStatusDesc(this.messageSource.getMessage("exception.dao.error",null, null));
			LOGGER.error(e.getLocalizedMessage());
		}
		LOGGER.info("<<<<<< ManageUserController.populateUserProfile");
		return manageVO;
	}
	/**
	 * Method saves new user profile
	 * @param body
	 * @param request
	 * @return return status code only.
	 */
	@RequestMapping(value = ACTION_SAVE_USER_GROUP, method = RequestMethod.POST)
	@ResponseBody()
	public ManageUserVO saveUserGroup(@RequestBody String body, HttpServletRequest request) throws JCTException{
		LOGGER.info(">>>>>> ManageUserController.saveUserGroup");
		ManageUserVO manageVO = new ManageUserVO();
		UserGroupVO userGroupVO = null;
		JsonNode node = CommonUtility.getNode(body);
		String userGroup = node.get("userGroupVal").toString().replaceAll("\"" , "").toUpperCase().trim();
		String userGroup1 = node.get("userGroupVal").toString().replaceAll("\"" , "").trim();
		String userProfileValue = node.get("userProfileVal").toString().replaceAll("\"" , "");
		Integer userProfileId = Integer.parseInt(node.get("userProfileId").toString().replaceAll("\"" , ""));
		String createdBy = node.get("createdBy").toString().replaceAll("\"" , "");
		String customerId = node.get("customerId").toString().replaceAll("\"" , "");
		Integer roleId = Integer.parseInt(node.get("roleId").toString().replaceAll("\"" , ""));
		//int profileId = Integer.parseInt(node.get("profileId").toString().replaceAll("\"" , ""));
		
		userGroupVO = new UserGroupVO();
		userGroupVO.setUserGroupName(userGroup1);
		userGroupVO.setUserProfileName(userProfileValue);
		userGroupVO.setUserProfileId(userProfileId);
		userGroupVO.setCreatedBy(createdBy);	
		userGroupVO.setCustomerId(customerId);
		userGroupVO.setRoleId(roleId);
		
		String updationSuccess = service.validateExistenceUserGrp(userGroup, userProfileId, roleId, customerId);
		if (updationSuccess.equals("success")) {
			manageVO.setStatusCode(StatusConstants.ALREADY_EXISTS);
		} else {
			manageVO.setStatusCode(StatusConstants.DOES_NOT_EXIST);
			try {
				manageVO = service.saveUserGroup(userGroupVO);
				ManageUserVO unused = service.populateExistingUserGroup(userProfileId, customerId, roleId);
				manageVO.setExistingUserGroupList(unused.getExistingUserGroupList());
				unused = null;
			} catch (JCTException e) {
				manageVO = new ManageUserVO();
				manageVO.setStatusCode(StatusConstants.STATUS_FAILURE);
				manageVO.setStatusDesc(this.messageSource.getMessage("exception.dao.user.group.save",null, null));
				LOGGER.error(e.getLocalizedMessage());
			} 
		}
		LOGGER.info("<<<<<< ManageUserController.saveUserGroup");
		return manageVO;
	}
	/**
	 * Method updates user group status
	 * @param body
	 * @param request
	 * @return return status code only.
	 */
	@RequestMapping(value = ACTION_UPDATE_USER_GROUP_STATUS, method = RequestMethod.POST)
	@ResponseBody()
	public ManageUserVO toogleGroupStatus(@RequestBody String body, HttpServletRequest request) {
		LOGGER.info(">>>>>> ManageUserController.saveUserGroup");
		ManageUserVO manageVO = null;
		UserGroupVO userGroupVO = null;
		JsonNode node = CommonUtility.getNode(body);
		Integer tablePkId = Integer.parseInt(node.get("tablePkId").toString().replaceAll("\"" , ""));
		Integer status = Integer.parseInt(node.get("status").toString().replaceAll("\"" , ""));
		int profileId = Integer.parseInt(node.get("profileId").toString().replaceAll("\"" , ""));
		String customerId = node.get("customerId").toString().replaceAll("\"" , "");
		Integer roleId = Integer.parseInt(node.get("roleId").toString().replaceAll("\"" , ""));
		userGroupVO = new UserGroupVO();
		userGroupVO.setPrimaryKeyVal(tablePkId);
		userGroupVO.setActiveStatus(status);
		userGroupVO.setUserGroupName(node.get("userGroup").toString().replaceAll("\"" , ""));
		try {
			manageVO = service.updateUserGroup(userGroupVO, "A");
			ManageUserVO unused = service.populateExistingUserGroup(profileId, customerId, roleId);
			manageVO.setExistingUserGroupList(unused.getExistingUserGroupList());
			unused = null;
		} catch (JCTException e) {
			manageVO = new ManageUserVO();
			manageVO.setStatusCode(StatusConstants.STATUS_FAILURE);
			manageVO.setStatusDesc(this.messageSource.getMessage("exception.dao.user.group.save",null, null));
			LOGGER.error(e.getLocalizedMessage());
		}
		LOGGER.info("<<<<<< ManageUserController.saveUserGroup");
		return manageVO;
	}
	/**
	 * Method update existing user group
	 * @param body
	 * @param request
	 * @return return status code only.
	 */
	@RequestMapping(value = ACTION_UPDATE_USER_GROUP, method = RequestMethod.POST)
	@ResponseBody()
	public ManageUserVO updateUserGroup(@RequestBody String body, HttpServletRequest request) throws JCTException{
		LOGGER.info(">>>>>> ManageUserController.updateUserGroup");
		ManageUserVO manageVO = new ManageUserVO();
		UserGroupVO userGroupVO = null;
		JsonNode node = CommonUtility.getNode(body);
		String userGroup = node.get("userGroupVal").toString().replaceAll("\"" , "").toUpperCase().trim();
		String userProfileValue = node.get("userProfileVal").toString().replaceAll("\"" , "");
		Integer userProfileId = Integer.parseInt(node.get("userProfileId").toString().replaceAll("\"" , ""));
		Integer tablePkId = Integer.parseInt(node.get("tablePkId").toString().replaceAll("\"" , ""));
		String customerId = node.get("customerId").toString().replaceAll("\"" , "");
		Integer roleId = Integer.parseInt(node.get("roleId").toString().replaceAll("\"" , ""));
		//int profileId = Integer.parseInt(node.get("profileId").toString().replaceAll("\"" , ""));
		userGroupVO = new UserGroupVO();
		userGroupVO.setUserGroupName(node.get("userGroupVal").toString().replaceAll("\"" , "").trim());
		userGroupVO.setUserProfileName(userProfileValue);
		userGroupVO.setUserProfileId(userProfileId);
		userGroupVO.setPrimaryKeyVal(tablePkId);		
		String updationSuccess = service.validateExistenceUserGrp(userGroup, userProfileId, roleId, customerId);
		if (updationSuccess.equals("success")) {
			manageVO.setStatusCode(StatusConstants.ALREADY_EXISTS);
		} else {
			manageVO.setStatusCode(StatusConstants.DOES_NOT_EXIST);
			try {
				manageVO = service.updateUserGroup(userGroupVO, null);
				ManageUserVO unused = service.populateExistingUserGroup(userProfileId, customerId, roleId);
				manageVO.setExistingUserGroupList(unused.getExistingUserGroupList());
				unused = null;
			} catch (JCTException e) {
				manageVO = new ManageUserVO();
				manageVO.setStatusCode(StatusConstants.STATUS_FAILURE);
				manageVO.setStatusDesc(this.messageSource.getMessage("exception.dao.user.group.update",null, null));
				LOGGER.error(e.getLocalizedMessage());
			} 
		}
		LOGGER.info("<<<<<< ManageUserController.updateUserGroup");
		return manageVO;
	}	
	/**
	 * Method populate existing user group
	 * @param body
	 * @param request
	 * @return return Map.
	 */
	@RequestMapping(value = ACTION_POPULATE_USER_GROUP, method = RequestMethod.POST)
	@ResponseBody()
	public ReportsVO populateUserGroup(@RequestBody String body, HttpServletRequest request) throws JCTException {
		LOGGER.info(">>>>>> ManageUserController.populateUserGroup");
		ReportsVO reportVO = new ReportsVO();
		List<UserGroupDTO> userGroupList = null;
		Map<Integer, String> userGroupMap = new LinkedHashMap<Integer, String>();
		userGroupList = service.getUserGroupList();
			int j = 0;
			for (int i=1; i<=userGroupList.size(); i++) {
				userGroupMap.put(userGroupList.get(j).getJctUserGroupId(), userGroupList.get(j).getUserGroupName());
				j++;
			}
			reportVO.setStatusCode(StatusConstants.STATUS_SUCCESS_WITH_RESULT);
			reportVO.setUserGroupMap(userGroupMap);
			LOGGER.info("<<<<<< ManageUserController.populateUserGroup");
		return reportVO;
	}
	/**
	 * Method populate existing users
	 * @param body
	 * @param request
	 * @return return Map.
	 */
	@RequestMapping(value = ACTION_POPULATE_EXISTING_USERS, method = RequestMethod.POST)
	@ResponseBody()
	public ExistingUsersVO populateExistingUsers(@RequestBody String body, HttpServletRequest request) throws JCTException {
		LOGGER.info(">>>>>> ManageUserController.populateExistingUsers");
		ExistingUsersVO existingUserVO = new ExistingUsersVO();
		List<ExistingUserDTO> userList = service.getUserList(null); // we do not need user group hence null
		if ( null == userList ) {
			userList = new ArrayList<ExistingUserDTO>();
		}
		existingUserVO.setStatusCode(StatusConstants.STATUS_SUCCESS_WITH_RESULT);
		existingUserVO.setExistingUsers(userList);
		LOGGER.info("<<<<<< ManageUserController.populateExistingUsers");
		return existingUserVO;
	}
/*	*//**
	 * Method populate existing facilitator
	 * @param body
	 * @param request
	 * @return return Map.
	 *//*
	@RequestMapping(value = ACTION_POPULATE_EXISTING_FACILITATOR, method = RequestMethod.POST)
	@ResponseBody()
	public ExistingUsersVO populateExistingFacilitator(@RequestBody String body, HttpServletRequest request) throws JCTException {
		LOGGER.info(">>>>>> ManageUserController.populateExistingUsers");
		ExistingUsersVO existingUserVO = new ExistingUsersVO();
		List<ExistingUserDTO> userList = service.getUserList(null); // we do not need user group hence null
		if ( null == userList ) {
			userList = new ArrayList<ExistingUserDTO>();
		}
		existingUserVO.setStatusCode(StatusConstants.STATUS_SUCCESS_WITH_RESULT);
		existingUserVO.setExistingUsers(userList);
		LOGGER.info("<<<<<< ManageUserController.populateExistingUsers");
		return existingUserVO;
	}*/
	/**
	 * Method populate existing users by user type and group
	 * @param body
	 * @param request
	 * @return return Map.
	 */
	@RequestMapping(value = ACTION_POPULATE_EXISTING_USERS_GR_TYPE, method = RequestMethod.POST)
	@ResponseBody()
	public ExistingUsersVO populateExistingUsersByUserTypeAndGroup(@RequestBody String body, HttpServletRequest request) throws JCTException {
		LOGGER.info(">>>>>> ManageUserController.populateExistingUsersByUserTypeAndGroup");
		ExistingUsersVO existingUserVO = new ExistingUsersVO();
		JsonNode node = CommonUtility.getNode(body);
		String userGrpVal = node.get("userGrpVal").asText();
		String usrType = node.get("usrType").asText();
		
		List<ExistingUserDTO> userList = service.getUserListByUserTypeAndGroup(userGrpVal, usrType);
		if ( null == userList ) {
			userList = new ArrayList<ExistingUserDTO>();
		}
		existingUserVO.setStatusCode(StatusConstants.STATUS_SUCCESS_WITH_RESULT);
		existingUserVO.setExistingUsers(userList);
		LOGGER.info("<<<<<< ManageUserController.populateExistingUsersByUserTypeAndGroup");
		return existingUserVO;
	}
	
	/**
	 * Method populates Active Inactive user list
	 * @param body
	 * @param request
	 * @return return Map.
	 */
	private ExistingUsersVO populateActiveInactiveUserList(String userGroup, String type, int userType) throws JCTException {
		LOGGER.info(">>>>>> ManageUserController.populateActiveInactiveUserList");
		ExistingUsersVO existingUserVO = new ExistingUsersVO();
		List<ExistingUserDTO> userList = service.populateActiveInactiveUserList(userGroup, type, userType);
		if ( null == userList ) {
			userList = new ArrayList<ExistingUserDTO>();
		}
		existingUserVO.setStatusCode(StatusConstants.STATUS_SUCCESS_WITH_RESULT);
		existingUserVO.setExistingUsers(userList);
		LOGGER.info("<<<<<< ManageUserController.populateActiveInactiveUserList");
		return existingUserVO;
	}

	/**
	 * Method deletes existing users
	 * @param body
	 * @param request
	 * @return return Map.
	 */
	@RequestMapping(value = ACTION_DELETE_USER, method = RequestMethod.POST)
	@ResponseBody()
	public ExistingUsersVO deleteExistingUsers(@RequestBody String body, HttpServletRequest request) 
			throws JCTException {
		LOGGER.info(">>>>>> ManageUserController.deleteExistingUsers");
		ExistingUsersVO existingUserVO = new ExistingUsersVO();
		JsonNode node = CommonUtility.getNode(body);
		String emailId = node.get("emailId").toString().replaceAll("\"" , "");
		int userType = node.get("userType").asInt();
		String deletionSuccess = service.deleteUser(emailId, userType);
		if (deletionSuccess.equals("success")) {
			existingUserVO.setStatusCode(StatusConstants.USER_DELETION_SUCCESS);
		} else {
			existingUserVO.setStatusCode(StatusConstants.USER_DELETION_FAILURE);
		}
		existingUserVO = populateExistingUsers("", null);
		LOGGER.info("<<<<<< ManageUserController.deleteExistingUsers");
		return existingUserVO;
	}
	/**
	 * Method updates the soft delete of existing users
	 * @param body
	 * @param request
	 * @return return Map.
	 */
	@RequestMapping(value = ACTION_USER_STATUS, method = RequestMethod.POST)
	@ResponseBody()
	public ExistingUsersVO updateStatus(@RequestBody String body, HttpServletRequest request) throws JCTException {
		LOGGER.info(">>>>>> ManageUserController.updateStatus");
		ExistingUsersVO existingUserVO = new ExistingUsersVO();
		JsonNode node = CommonUtility.getNode(body);
		String emailId = node.get("emailId").toString().replaceAll("\"" , "");
		Integer softDelete = Integer.parseInt(node.get("updatedStt").toString().replaceAll("\"" , ""));
		String updationSuccess = service.updateUserActivationStatus(emailId , softDelete);
		String fName = "";
		if (updationSuccess.equals("success")) {
			if(softDelete == 1){
				try {
					fName = service.getFirstName(emailId, 1);
					mailer.mailAccDeactivate(emailId,fName);
				} catch (MailingException e) {
					existingUserVO.setStatusCode(StatusConstants.ACC_DEACTIVATION_MAIL_SEND_ERROR);
					LOGGER.error(e.getLocalizedMessage());
				}
			}
			existingUserVO.setStatusCode(StatusConstants.USER_UPDATION_SUCCESS);
		} else {
			existingUserVO.setStatusCode(StatusConstants.USER_UPDATION_FAILURE);
		}
		existingUserVO = populateExistingUsers("", null);
		LOGGER.info("<<<<<< ManageUserController.updateStatus");
		return existingUserVO;
	}
	/**
	 * Method resets the user password
	 * @param body
	 * @param request
	 * @return return Map.
	 */
	@RequestMapping(value = ACTION_RESET_PASSWORD, method = RequestMethod.POST)
	@ResponseBody()
	public ExistingUsersVO resetPassword(@RequestBody String body, HttpServletRequest request) 
			throws JCTException {
		ExistingUsersVO existingUserVO = new ExistingUsersVO();
		JsonNode node = CommonUtility.getNode(body);
		String emailId = node.get("emailId").asText();
		int roleId = node.get("roleId").asInt();
		String generatedPassword = this.generatePassword();
		String encryptedPassword = this.encriptString(generatedPassword);
		String passwordResetStatus = service.resetPassword(emailId, encryptedPassword, roleId);
		String fname = "";
		if (passwordResetStatus.equals("success")) {
			existingUserVO.setStatusCode(StatusConstants.USER_PASSWORD_RESET_SUCCESS);
			try {
				fname = service.getFirstName(emailId, roleId);
				if(roleId == 1){										
					mailer.mailResetIndiFaci(emailId, generatedPassword, fname, "reset");					
				} else {
					mailer.mailResetIndiFaci(emailId, generatedPassword, fname, "resetFaci");					
				}
				
			} catch (MailingException e) {
				existingUserVO.setStatusCode(StatusConstants.USER_REGISTERED_MAIL_SEND_ERROR_USERS);
				LOGGER.error(e.getLocalizedMessage());
			}
			
		} else if (passwordResetStatus.equals("changed")) {
			existingUserVO.setStatusCode(StatusConstants.USER_PASSWORD_RESET_CANNOT_PROCEED);
			existingUserVO.setStatusDesc(this.messageSource.getMessage("error.password.cannot.be.changed",null, null));
		} else {
			existingUserVO.setStatusCode(StatusConstants.USER_PASSWORD_RESET_FAILURE);
		}
		return existingUserVO;
	}
	
	
	/**
	 * Method adds new user
	 * @param body
	 * @param request
	 * @return ManageUserVO.
	 */
	@RequestMapping(value = ACTION_SAVE_MANUAL_USER, method = RequestMethod.POST)
	@ResponseBody()
	public ManageUserVO saveManualUser(@RequestBody String body, HttpServletRequest request) {
		LOGGER.info(">>>>>> ManageUserController.saveManualUser");
		ManageUserVO manageVO = new ManageUserVO();
		ArrayList<String> validEmailList = new ArrayList<String>();
		ArrayList<String> inValidEmailList = new ArrayList<String>();
		JsonNode node = CommonUtility.getNode(body);
		String userGroup = node.get("userGroup").toString().replaceAll("\"" , "");
		String emailIdString = node.get("emailIdString").toString().replaceAll("\"" , "");
		String createdBy = node.get("createdBy").toString().replaceAll("\"" , "");
		String paymentType = node.get("paymentType").toString().replaceAll("\"" , "");
		String[] grpSplit = userGroup.split("!"); //0 is the id and 1 is the value
		StringTokenizer emailToken = new StringTokenizer(emailIdString, "~");
		while (emailToken.hasMoreTokens()) {
			String email = (String) emailToken.nextToken().replaceAll("\"" , "");
			//Boolean isValid = this.isEmailValid(email);
			Boolean isValid = CommonUtility.isEmailValid(email);
			if (isValid) {
				validEmailList.add(email);
			} else {
				inValidEmailList.add(email);
			}
		}
		try {
			// Check if the users are present
			// if a single user exists with user role as individual user, return the admin saying to correct it
			int registeredUserCount = service.alreadyRegUserCount(validEmailList);
			if (registeredUserCount == 0) {
				NewUserDTO saveUserDTO = service.saveUsers(validEmailList, Integer.parseInt(grpSplit[0]), grpSplit[1], createdBy, paymentType);
				manageVO.setStatusCode(StatusConstants.USER_ONLY_REGISTERED);
				manageVO.setStatusDesc(this.messageSource.getMessage("manage.user.registered.no.email",null, null));
				try{
					LOGGER.info("INTIMATING THE ADMIN...");
					mailer.intimateAdmin(saveUserDTO, createdBy, inValidEmailList);
				} catch (MailingException e) {
					LOGGER.error("UNABLE TO SEND MAIL TO ADMIN: "+e.getLocalizedMessage());
					manageVO.setStatusCode(StatusConstants.USER_REGISTERED_MAIL_SEND_ERROR_ADMIN);
					manageVO.setStatusDesc(this.messageSource.getMessage("manage.user.registered.no.email.error",null, null));
				}
				//postpone the mail send and save it to database. Cron Job will send when triggered
				if (saveUserDTO.getNewRegistration().size() > 10) {
					LOGGER.info("AS NUMBER OF VALID REGISTERED USER IS MORE THAN 10, CRON WILL BE TRIGGERED TO SEND MAIL TO REGISTERED USER");
					service.saveCronUsers(saveUserDTO.getNewRegistration(), createdBy);
					manageVO.setStatusCode(StatusConstants.USER_REGISTERED_MAIL_BY_CRON);
					// if all the user email ids provided are valid and registered
					if ((inValidEmailList.size() == 0) && (validEmailList.size() == saveUserDTO.getNewRegistration().size())) {
						manageVO.setStatusDesc(this.messageSource.getMessage("manage.user.registration.admin.cron.confirmation.record.matched",null, null));
					}  else {
						int userExistsCount = validEmailList.size() - saveUserDTO.getNewRegistration().size();
						manageVO.setStatusDesc(saveUserDTO.getNewRegistration().size()+" "+
											   this.messageSource.getMessage("manage.user.registration.admin.cron.confirmation.record.mismatched",null, null) + 
											   "\n"+inValidEmailList.size()+" "+this.messageSource.getMessage("manage.user.registration.admin.cron.confirmation.record.invalid",null, null) + 
											   "\n"+userExistsCount+" "+this.messageSource.getMessage("manage.user.registration.admin.cron.confirmation.record.exists",null, null) +
											   "\n\n"+this.messageSource.getMessage("manage.user.registration.admin.cron.confirmation.footer",null, null));
					}
				} else {
					Iterator<NewUserDTO> itr = saveUserDTO.getNewRegistration().iterator();
					int userMailSentFailedCounter = 0;
					while (itr.hasNext()) {
						NewUserDTO dto = (NewUserDTO) itr.next();
						try {
							mailer.intimateUser(dto.getEmailId(), dto.getPassword(), dto.getFirstName(), "register",null);
						} catch (MailingException e) {
							service.saveCronUsers(saveUserDTO.getNewRegistration(), createdBy);
							userMailSentFailedCounter = userMailSentFailedCounter + 1;
						}
					}
					if ((inValidEmailList.size() == 0) && (validEmailList.size() == saveUserDTO.getNewRegistration().size())) {
						manageVO.setStatusDesc(this.messageSource.getMessage("manage.user.registration.admin.manual.success",null, null));
					} else {
						int userExistsCount = validEmailList.size() - saveUserDTO.getNewRegistration().size();
						manageVO.setStatusDesc(saveUserDTO.getNewRegistration().size()+" "+
								   this.messageSource.getMessage("manage.user.registration.admin.cron.confirmation.record.mismatched",null, null) + 
								   "\n"+inValidEmailList.size()+" "+this.messageSource.getMessage("manage.user.registration.admin.cron.confirmation.record.invalid",null, null) + 
								   "\n"+userExistsCount+" "+this.messageSource.getMessage("manage.user.registration.admin.cron.confirmation.record.exists",null, null) +
								   "\n"+userMailSentFailedCounter+" "+this.messageSource.getMessage("manage.user.registration.admin.cron.confirmation.email.failed",null, null) +
								   "\n\n"+this.messageSource.getMessage("manage.user.registration.admin.cron.confirmation.footer",null, null));
					}
				}
			} else {
				manageVO.setStatusCode(StatusConstants.USER_NOT_REGISTERED);
				manageVO.setStatusDesc(this.messageSource.getMessage("manage.user.registration.failed.manual.already.reg",null, null));
			}
		}   catch (JCTException e) {
			LOGGER.error("-----> "+e.getLocalizedMessage());
			manageVO.setStatusCode(StatusConstants.USER_NOT_REGISTERED);
			manageVO.setStatusDesc(this.messageSource.getMessage("manage.user.registration.failed",null, null));
		}
		LOGGER.info("<<<<<< ManageUserController.saveManualUser");
		return manageVO;
	}
	
	/**
	 * Method adds new user
	 * @param body
	 * @param request
	 * @return ManageUserVO.
	 */
	@RequestMapping(value = ACTION_SAVE_USER, method = RequestMethod.POST)
	@ResponseBody()
	public ManageUserVO saveUser(@RequestBody String body, HttpServletRequest request) {
		LOGGER.info(">>>>>> ManageUserController.saveUser");
		ManageUserVO manageVO = new ManageUserVO();
		ArrayList<String> validEmailList = new ArrayList<String>();
		ArrayList<String> inValidEmailList = new ArrayList<String>();
		JsonNode node = CommonUtility.getNode(body);
		String userGroup = node.get("userGroup").asText();
		String emailIdString = node.get("emailIdString").toString().replaceAll("\"" , "");
		String createdBy = node.get("createdBy").toString().replaceAll("\"" , "");
		String paymentType = node.get("paymentType").toString().replaceAll("\"" , "");
		String[] grpSplit = userGroup.split("!"); //0 is the id and 1 is the value
		StringTokenizer emailToken = new StringTokenizer(emailIdString, "~");
		while (emailToken.hasMoreTokens()) {
			String email = (String) emailToken.nextToken().replaceAll("\"" , "");
			//Boolean isValid = this.isEmailValid(email);
			Boolean isValid = CommonUtility.isEmailValid(email);
			if (isValid) {
				validEmailList.add(email);
			} else {
				inValidEmailList.add(email);
			}
		}
		try {
			NewUserDTO saveUserDTO = service.saveUsers(validEmailList, Integer.parseInt(grpSplit[0]), grpSplit[1], createdBy, paymentType);
			manageVO.setStatusCode(StatusConstants.USER_ONLY_REGISTERED);
			manageVO.setStatusDesc(this.messageSource.getMessage("manage.user.registered.no.email",null, null));
			try{
				LOGGER.info("INTIMATING THE ADMIN...");
				mailer.intimateAdmin(saveUserDTO, createdBy, inValidEmailList);
			} catch (MailingException e) {
				LOGGER.error("UNABLE TO SEND MAIL TO ADMIN: "+e.getLocalizedMessage());
				manageVO.setStatusCode(StatusConstants.USER_REGISTERED_MAIL_SEND_ERROR_ADMIN);
				manageVO.setStatusDesc(this.messageSource.getMessage("manage.user.registered.no.email.error",null, null));
			}
			//postpone the mail send and save it to database. Cron Job will send when triggered
			if (saveUserDTO.getNewRegistration().size() > 10) {
				LOGGER.info("AS NUMBER OF VALID REGISTERED USER IS MORE THAN 10, CRON WILL BE TRIGGERED TO SEND MAIL TO REGISTERED USER");
				service.saveCronUsers(saveUserDTO.getNewRegistration(), createdBy);
				manageVO.setStatusCode(StatusConstants.USER_REGISTERED_MAIL_BY_CRON);
				// if all the user email ids provided are valid and registered
				if ((inValidEmailList.size() == 0) && (validEmailList.size() == saveUserDTO.getNewRegistration().size())) {
					manageVO.setStatusDesc(this.messageSource.getMessage("manage.user.registration.admin.cron.confirmation.record.matched",null, null));
				}  else {
					int userExistsCount = validEmailList.size() - saveUserDTO.getNewRegistration().size();
					manageVO.setStatusDesc(saveUserDTO.getNewRegistration().size()+" "+
										   this.messageSource.getMessage("manage.user.registration.admin.cron.confirmation.record.mismatched",null, null) + 
										   "\n"+inValidEmailList.size()+" "+this.messageSource.getMessage("manage.user.registration.admin.cron.confirmation.record.invalid",null, null) + 
										   "\n"+userExistsCount+" "+this.messageSource.getMessage("manage.user.registration.admin.cron.confirmation.record.exists",null, null) +
										   "\n\n"+this.messageSource.getMessage("manage.user.registration.admin.cron.confirmation.footer",null, null));
				}
			} else {
				Iterator<NewUserDTO> itr = saveUserDTO.getNewRegistration().iterator();
				int userMailSentFailedCounter = 0;
				while (itr.hasNext()) {
					NewUserDTO dto = (NewUserDTO) itr.next();
					try {
						mailer.intimateUser(dto.getEmailId(), dto.getPassword(), dto.getFirstName(), "register",null);
					} catch (MailingException e) {
						service.saveCronUsers(saveUserDTO.getNewRegistration(), createdBy);
						//manageVO.setStatusCode(StatusConstants.USER_REGISTERED_MAIL_SEND_ERROR_USERS);
						userMailSentFailedCounter = userMailSentFailedCounter + 1;
					}
				}
				if ((inValidEmailList.size() == 0) && (validEmailList.size() == saveUserDTO.getNewRegistration().size())) {
					manageVO.setStatusDesc(this.messageSource.getMessage("manage.user.registration.admin.manual.success",null, null));
				} else {
					int userExistsCount = validEmailList.size() - saveUserDTO.getNewRegistration().size();
					manageVO.setStatusDesc(saveUserDTO.getNewRegistration().size()+" "+
							   this.messageSource.getMessage("manage.user.registration.admin.cron.confirmation.record.mismatched",null, null) + 
							   "\n"+inValidEmailList.size()+" "+this.messageSource.getMessage("manage.user.registration.admin.cron.confirmation.record.invalid",null, null) + 
							   "\n"+userExistsCount+" "+this.messageSource.getMessage("manage.user.registration.admin.cron.confirmation.record.exists",null, null) +
							   "\n"+userMailSentFailedCounter+" "+this.messageSource.getMessage("manage.user.registration.admin.cron.confirmation.email.failed",null, null) +
							   "\n\n"+this.messageSource.getMessage("manage.user.registration.admin.cron.confirmation.footer",null, null));
				}
				//manageVO.setStatusCode(StatusConstants.USER_REGISTERED_MAIL_SENT);
			}
		}   catch (JCTException e) {
			LOGGER.error("-----> "+e.getLocalizedMessage());
			manageVO.setStatusCode(StatusConstants.USER_NOT_REGISTERED);
			manageVO.setStatusDesc(this.messageSource.getMessage("manage.user.registration.failed",null, null));
		}
		LOGGER.info("<<<<<< ManageUserController.saveUser");
		return manageVO;
	}
	/**
	 * Method resets password in batch
	 * @param body
	 * @param request
	 * @return ManageUserVO.
	 */
	@RequestMapping(value = ACTION_RESET_PASSWORD_IN_BATCH, method = RequestMethod.POST)
	@ResponseBody()
	public ManageUserVO resetPasswordInBatch(@RequestBody String body, HttpServletRequest request) {
		LOGGER.info(">>>>>> ManageUserController.resetPasswordInBatch");
		ManageUserVO manageVO = new ManageUserVO();
		JsonNode node = CommonUtility.getNode(body);
		String emailIdString = node.get("emailIdString").toString().replaceAll("\"" , "");
		String updatedBy = node.get("updatedBy").toString().replaceAll("\"" , "");
		StringTokenizer token = new StringTokenizer(emailIdString, "~");
		int totalEmailCount = token.countTokens();
		LOGGER.info("TOTAL NUMBER OF EMAILS: "+totalEmailCount);
		while (token.hasMoreTokens()) {
			try {
				String emailString = (String) token.nextToken();
				manageVO = resetMassPassword(emailString.split("`!!`")[0] , totalEmailCount, 
						manageVO, updatedBy, Integer.parseInt(emailString.split("`!!`")[1]));
			} catch (JCTException e) {
				LOGGER.error("PASSWORD RESET FAILURE: "+e.getLocalizedMessage());
				manageVO.setStatusCode(StatusConstants.USER_PASSWORD_RESET_FAILURE);
			}
		}
		LOGGER.info("<<<<<< ManageUserController.resetPasswordInBatch");
		return manageVO;
	}
	
	/**
	 * Method resets password 
	 * @param body
	 * @param request
	 * @return ManageUserVO.
	 */	
	private ManageUserVO resetMassPassword(String emailId, int totalEmailCount, ManageUserVO manageVO, String updatedBy, int profileId) throws JCTException {
		LOGGER.info(">>>>>> ManageUserController.resetMassPassword");
		String generatedPassword = this.generatePassword();
		String encryptedPassword = this.encriptString(generatedPassword);
		String passwordResetStatus = service.resetPassword(emailId, encryptedPassword, profileId);
		String fName = "";
		if (passwordResetStatus.equals("changed")) {
			manageVO.setStatusCode(StatusConstants.USER_PASSWORD_RESET_CANNOT_PROCEED);
			manageVO.setStatusDesc(this.messageSource.getMessage("error.password.cannot.be.changed",null, null));
		} else {
			manageVO.setStatusCode(StatusConstants.USER_PASSWORD_RESET_SUCCESS);
		}
		if (passwordResetStatus.equals("success")) {
			try {
				if (totalEmailCount <= 10) {
					fName = service.getFirstName(emailId, profileId);
					if(profileId == 1){						
						mailer.mailResetIndiFaci(emailId, generatedPassword, fName, "reset");						
					} else {
						mailer.mailResetIndiFaci(emailId, generatedPassword, fName, "resetFaci");						
					}
					
				} else {
					//Store in the db
					String status = service.updateCronUsers(emailId, generatedPassword, updatedBy);
					LOGGER.info("====== STATUS OF UPDATE: "+status+" =======");
				}
			} catch (MailingException e) {
				manageVO.setStatusCode(StatusConstants.USER_REGISTERED_MAIL_SEND_ERROR_USERS);
				LOGGER.error(e.getLocalizedMessage());
			}
		} else {
			manageVO.setStatusCode(StatusConstants.USER_PASSWORD_RESET_FAILURE);
		}
		LOGGER.info("<<<<<< ManageUserController.resetMassPassword");
		return manageVO;
	}
	/**
	 * Method encripts the string [Custom encryption]
	 * @param string
	 * @return
	 */
	private String encriptString(String string){
		LOGGER.info(">>>>>> ManageUserController.encriptString");
		String ref = "P!QR#S$T%U&VW(X)Y*Z+[,-].^/_0`1a2b3c4d5e6f7g8h9i:j;k<l=m>n?o@pAqBrCsDtEuFvGwHxIyJzK{L|M}N~O";
		StringBuilder result = new StringBuilder("");
			for (int count=0; count<string.length(); count++) {
				String charStr = string.substring (count, count+1); 
				int num = ref.indexOf(charStr);
				String encodeChar = ref.substring(num+1, num+2);
				result.append(encodeChar);
			}
			result.append(result.toString());
			LOGGER.info("<<<<<< ManageUserController.encriptString");
		return result.toString();
	}
	/**
	 * Method generates random password
	 * @return generated password
	 */
	private String generatePassword() {
		LOGGER.info(">>>>>> ManageUserController.generatePassword");
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
		LOGGER.info("<<<<<< ManageUserController.generatePassword");
		return sb.toString();
	}
	/**
	 * Method fetches username
	 * @param body
	 * @param request
	 * @return return Map.
	 */
	@RequestMapping(value = ACTION_FETCH_EXISTING_USERNAME, method = RequestMethod.POST)
	@ResponseBody()
	public ExistingUsersVO fetchExistingUserName(@RequestBody String body, HttpServletRequest request) 
			throws JCTException {
		LOGGER.info(">>>>>> ManageUserController.fetchExistingUserName");
		JsonNode node = CommonUtility.getNode(body);
		ExistingUsersVO vo = new ExistingUsersVO();
		String userGroup = node.get("userGroup").toString().replaceAll("\"" , "");
		List<ExistingUserDTO> userList = service.getUserList(userGroup.split("!")[1]);
		if ( null == userList ) {
			userList = new ArrayList<ExistingUserDTO>();
		}
		vo.setStatusCode(StatusConstants.STATUS_SUCCESS_WITH_RESULT);
		vo.setExistingUsers(userList);
		LOGGER.info("<<<<<< ManageUserController.fetchExistingUserName");
		return vo;
	}
	/**
	 * Method deletes users in batch
	 * @param body
	 * @param request
	 * @return return Map.
	 */
	@RequestMapping(value = ACTION_CHANGE_BATCH_USER_STATUS, method = RequestMethod.POST)
	@ResponseBody()
	public ExistingUsersVO changeUserStatusInBatch(@RequestBody String body, HttpServletRequest request) 
			throws JCTException {
		LOGGER.info(">>>>>> ManageUserController.changeUserStatusInBatch");
		ExistingUsersVO existingUserVO = new ExistingUsersVO();
		JsonNode node = CommonUtility.getNode(body);
		String emailIdString = node.get("emailIdString").toString().replaceAll("\"" , "");
		int status = Integer.parseInt(node.get("statusToBe").toString().replaceAll("\"" , ""));
		int userType = node.get("userType").asInt();
		String userGrpVal = node.get("userGrpVal").toString().replaceAll("\"" , "");
		String updationSuccess = service.updateUserActivationStatusInBatch(emailIdString , status, userType);
		if (updationSuccess.equals("success")) {
			existingUserVO.setStatusCode(StatusConstants.USER_UPDATION_SUCCESS);
		} else {
			existingUserVO.setStatusCode(StatusConstants.USER_UPDATION_FAILURE);
		}
		String type = "inactive";
		if(status == 0) {
			type = "active";
		} 
		existingUserVO = populateActiveInactiveUserList(userGrpVal.split("!")[1], type, userType);
		LOGGER.info("<<<<<< ManageUserController.changeUserStatusInBatch");
		return existingUserVO;
	}
	/**
	 * Method is used to fetch the sample PDF
	 * @param body
	 * @param request
	 * @return
	 */
	@RequestMapping(value = "/downloadCSVFile/{fileName}", method = RequestMethod.GET)
	//@ResponseBody()
	public ReturnVO downloadCSVFile(@PathVariable("fileName") String fileName, HttpServletRequest request, HttpServletResponse response) {
        try {
        	LOGGER.info(">>>> ManageUserController.downloadCSVFile");
        	String csvPath = this.messageSource.getMessage("csv.file.location",null, null);
        	final File file = new File(csvPath);
            response.setContentType("text/csv");
            response.setHeader("Content-Disposition",
                    "attachment; filename=" + file.getName());

            final BufferedReader br = new BufferedReader(new FileReader(file));
            try {
                String line;
                while ((line = br.readLine()) != null) {
                    response.getWriter().write(line + "\n");
                }
            } finally {
                br.close();
            }	
        
        } catch (Exception e) {
        	LOGGER.error(e.getLocalizedMessage());
			//e.printStackTrace();
		}
        LOGGER.info("<<<<<< ManageUserController.downloadCSVFile");
        return null;
	}
	/**
	 * Method populates active or inactive user list
	 * @param body
	 * @param request
	 * @return return Map.
	 */
	@RequestMapping(value = ACTION_POPULATE_ACTIVE_INACTIVE_USER_ADMIN, method = RequestMethod.POST)
	@ResponseBody()
	public ExistingUsersVO specificUserListForAdmin(@RequestBody String body, HttpServletRequest request) throws JCTException {
		LOGGER.info(">>>>>> ManageUserController.specificUserList");
		JsonNode node = CommonUtility.getNode(body);
		int status = Integer.parseInt(node.get("statusToBe").toString().replaceAll("\"" , ""));
		String userGrpVal = node.get("userGroup").toString().replaceAll("\"" , "");
		String emailId = node.get("emailId").asText();
		String userType = node.get("userType").asText();
		ExistingUsersVO existingUserVO = new ExistingUsersVO();
		String type = "active";
		if(status == 0) {
			type = "inactive";
		} 
		int userTypeInt = 1;
		if (userType.equals("F")) {
			userTypeInt = 3;
		}
		List<ExistingUserDTO> userList = service.populateActiveInactiveUserListForAdmin(userGrpVal.split("!")[1], type, emailId, userTypeInt);
		if ( null == userList ) {
			userList = new ArrayList<ExistingUserDTO>();
		}
		existingUserVO.setStatusCode(StatusConstants.STATUS_SUCCESS_WITH_RESULT);
		existingUserVO.setExistingUsers(userList);
		existingUserVO.setUserType(userTypeInt);
		LOGGER.info("<<<<<< ManageUserController.specificUserList");
		return existingUserVO;
	}
	/**
	 * Method fetches existing user names
	 * @param body
	 * @param request
	 * @return return Map.
	 */
	@RequestMapping(value = ACTION_FETCH_EXISTING_USERNAME_ADMIN, method = RequestMethod.POST)
	@ResponseBody()
	public ExistingUsersVO fetchExistingUserNameForAdmin(@RequestBody String body, HttpServletRequest request) 
			throws JCTException {
		LOGGER.info(">>>>>> ManageUserController.fetchExistingUserNameForAdmin");
		JsonNode node = CommonUtility.getNode(body);
		ExistingUsersVO vo = new ExistingUsersVO();
		String userGroup = node.get("userGroup").toString().replaceAll("\"" , "");
		String emailId = node.get("emailId").asText();
		String userType = node.get("userType").asText();
		int userTypeInt = 1;
		if (userType.equals("F")) {
			userTypeInt = 3;
		}
		List<ExistingUserDTO> userList = service.getUserListForAdmin(userGroup.split("!")[1], emailId, userTypeInt);
		if ( null == userList ) {
			userList = new ArrayList<ExistingUserDTO>();
		}
		vo.setStatusCode(StatusConstants.STATUS_SUCCESS_WITH_RESULT);
		vo.setExistingUsers(userList);
		vo.setUserType(userTypeInt);
		LOGGER.info("<<<<<< ManageUserController.fetchExistingUserNameForAdmin");
		return vo;
	}
	/**
	 * Method adds new facilitator.
	 * @param body
	 * @param request
	 * @return ConfirmationMessageVO.
	 */
	@RequestMapping(value = ACTION_SAVE_FACILITATOR_PAYMENT, method = RequestMethod.POST)
	@ResponseBody()
	public ConfirmationMessageVO saveFacilitatorPayment(@RequestBody String body, HttpServletRequest request) {
		LOGGER.info(">>>>>> ManageUserController.saveFacilitatorPayment");
		JsonNode node = CommonUtility.getNode(body);
		String userGroup = node.get("userGroup").asText();
		//String bankDetails = node.get("bankDetails").asText();
		String facilitatorEmail = node.get("facilitatorEmail").asText();
		String paymentDate = node.get("paymentDate").asText();
		String numberOfUsers = node.get("numberOfUsers").asText();
		String chequeNos = node.get("chequeNos").asText();
		String totalAmountToBePaid = node.get("totalAmountToBePaid").asText();
		String facToHaveAccount = node.get("facToHaveAccount").asText();
		String createdBy = node.get("createdBy").asText();
		String paymentType = node.get("paymentType").asText();
		String facFirstName = node.get("facFirstName").asText();
		String facLastName = node.get("facLastName").asText();
		//Boolean isValid = this.isEmailValid();
		Boolean isValid = CommonUtility.isEmailValid(facilitatorEmail);
		int status = 0;
		ConfirmationMessageVO vo = new ConfirmationMessageVO();
		try {
			if (isValid) {
				FacilitatorAccountVO facilitatorChequeVO = new FacilitatorAccountVO(userGroup, facilitatorEmail, paymentDate, 
						numberOfUsers, chequeNos, totalAmountToBePaid, facToHaveAccount, createdBy, paymentType, facFirstName, facLastName);
				status = service.saveFacilitatorViaChequePayment(facilitatorChequeVO,paymentType);
				if (status == StatusConstants.STATUS_SUCCESS) {
					StringBuilder msgBuilder = new StringBuilder(this.messageSource.getMessage("msg.registered1",null, null));
					msgBuilder.append(facilitatorEmail);
					msgBuilder.append(this.messageSource.getMessage("msg.registered2",null, null));
					msgBuilder.append(this.messageSource.getMessage("msg.registered3",null, null));
					vo.setMessage(msgBuilder.toString());
				} else if (status == StatusConstants.USER_REGISTERED_MAIL_SEND_ERROR) {
					StringBuilder msgBuilder = new StringBuilder(this.messageSource.getMessage("msg.registered1",null, null));
					msgBuilder.append(facilitatorEmail);
					msgBuilder.append(this.messageSource.getMessage("msg.registered2",null, null));
					msgBuilder.append(this.messageSource.getMessage("msg.registered4",null, null));
					vo.setMessage(msgBuilder.toString());
				} else if (status == StatusConstants.TABLE_OR_DATA_RELATED_ERROR) {
					StringBuilder msgBuilder = new StringBuilder(this.messageSource.getMessage("msg.registered5",null, null));
					vo.setMessage(msgBuilder.toString());
				} else if (status == StatusConstants.STATUS_USER_EXIST) {
					StringBuilder msgBuilder = new StringBuilder(this.messageSource.getMessage("msg.registered7",null, null));
					vo.setMessage(msgBuilder.toString());
				} else if (status == StatusConstants.STATUS_GEN_USER_EXIST) {
					StringBuilder msgBuilder = new StringBuilder(this.messageSource.getMessage("msg.registered8",null, null));
					vo.setMessage(msgBuilder.toString());
				} else { //FAILED
					StringBuilder msgBuilder = new StringBuilder(this.messageSource.getMessage("msg.registered5",null, null));
					vo.setMessage(msgBuilder.toString());
					vo.setStatusCode(123);
				}
			} else {
				StringBuilder msgBuilder = new StringBuilder(this.messageSource.getMessage("msg.registered6",null, null));
				vo.setMessage(msgBuilder.toString());
				vo.setStatusCode(888);
			}
		} catch (JCTException e) {
			LOGGER.error("UNABLE TO SAVE FACILITATOR: "+e.getLocalizedMessage());
		}
		LOGGER.info(">>>>>> ManageUserController.saveFacilitatorPayment");
		return vo;
	}
	
	/**
	 * Method adds new general user. This method takes care for Manual entry.
	 * @param body
	 * @param request
	 * @return ConfirmationMessageVO.
	 */
	@RequestMapping(value = ACTION_SAVE_GEN_USER_MANUAL_PAYMENT, method = RequestMethod.POST)
	@ResponseBody()
	public ConfirmationMessageVO saveGeneralUserPaymentManual(@RequestBody String body, HttpServletRequest request) {
		LOGGER.info(">>>>>> ManageUserController.saveGeneralUserPaymentManual");
		
		ConfirmationMessageVO vo = new ConfirmationMessageVO();
		ArrayList<String> validEmailList = new ArrayList<String>();
		ArrayList<String> inValidEmailList = new ArrayList<String>();
		JsonNode node = CommonUtility.getNode(body);
		String userGroup = node.get("userGroup").asText();
		//String bankDetails = node.get("bankDetails").asText();
		String emailIdString = node.get("emailIdString").asText();
		String paymentDate = node.get("paymentDate").asText();
		String numberOfUsers = node.get("numberOfUsers").asText();
		String chequeNos = node.get("chequeNos").asText();
		String totalAmountToBePaid = node.get("totalAmountToBePaid").asText();
		String expiryDuration = node.get("expiryDuration").asText();
		String createdBy = node.get("createdBy").asText();
		String paymentType = node.get("paymentType").asText();
		StringTokenizer emailToken = new StringTokenizer(emailIdString, "~");
		while (emailToken.hasMoreTokens()) {			
			validEmailList.add(emailToken.nextToken().toString());
						
			//Boolean isValid = this.isEmailValid(email);
			/*Boolean isValid = CommonUtility.isEmailValid(email);
			if (isValid) {
				validEmailList.add(email);
			} else {
				inValidEmailList.add(email);
			}*/
		}
		GeneralUserAccountVO accountVO = new GeneralUserAccountVO(createdBy,
				paymentType, expiryDuration, userGroup,
				validEmailList, inValidEmailList, paymentDate, numberOfUsers,
				chequeNos, totalAmountToBePaid);
		try {
			vo.setMessage(service.saveGeneralUserViaChequePayment(accountVO, validEmailList.size(),paymentType));
		} catch (JCTException ex) {
			LOGGER.error("UNABLE TO SAVE GENERAL USER BY MANUAL: "+ex.getLocalizedMessage());
		}
		LOGGER.info(">>>>>> ManageUserController.saveGeneralUserPaymentManual");
		return vo;
	}
	/**
	 * Method adds new general user. This method takes care for CSV entry.
	 * Payment Mode CHECK
	 * @param body
	 * @param request
	 * @return ConfirmationMessageVO.
	 */
	@RequestMapping(value = ACTION_SAVE_GEN_USER_CSV_PAYMENT, method = RequestMethod.POST)
	@ResponseBody()
	public ConfirmationMessageVO saveGeneralUserPaymentCSV(@RequestBody String body, HttpServletRequest request) {
		LOGGER.info(">>>>>> ManageUserController.saveGeneralUserPaymentCSV");
		ConfirmationMessageVO vo = new ConfirmationMessageVO();
		ArrayList<String> validEmailList = new ArrayList<String>();
		ArrayList<String> inValidEmailList = new ArrayList<String>();
		JsonNode node = CommonUtility.getNode(body);
		String userGroup = node.get("userGroup").asText();
		//String bankDetails = node.get("bankDetails").asText();
		String emailIdString = node.get("emailIdString").asText();
		String paymentDate = node.get("paymentDate").asText();
		String numberOfUsers = node.get("numberOfUsers").asText();
		String chequeNos = node.get("chequeNos").asText();
		String totalAmountToBePaid = node.get("totalAmountToBePaid").asText();
		String expiryDuration = node.get("expiryDuration").asText();
		String createdBy = node.get("createdBy").asText();
		String paymentType = node.get("paymentType").asText();
		StringTokenizer emailToken = new StringTokenizer(emailIdString, "~");
		
		while (emailToken.hasMoreTokens()) {		
			String email = (String) emailToken.nextToken().replaceAll("\"" , "");
			if (email.trim().length() > 5) {
			validEmailList.add(email);
			}
			/*if (email.trim().length() > 5) {
				//Boolean isValid = this.isEmailValid(email);
				Boolean isValid = CommonUtility.isEmailValid(email);
				if (isValid) {
					validEmailList.add(email);
				} else {
					inValidEmailList.add(email);
				}
			}*/
		}
		
		GeneralUserAccountVO accountVO = new GeneralUserAccountVO(createdBy,
				paymentType, expiryDuration, userGroup,
				validEmailList, inValidEmailList, paymentDate, numberOfUsers,
				chequeNos, totalAmountToBePaid);
		try {
			vo.setMessage(service.saveGeneralUserViaChequePayment(accountVO, validEmailList.size(),paymentType));
		} catch (JCTException ex) {
			LOGGER.error("UNABLE TO SAVE GENERAL USER BY CSV: "+ex.getLocalizedMessage());
		}
		LOGGER.info(">>>>>> ManageUserController.saveGeneralUserPaymentCSV");
		return vo;
	}	
	
	
	/**
	 * Method renew / subscribe facilitator for CHK / CASH / FREE payment.
	 * @param body
	 * @param request
	 * @return ConfirmationMessageVO.
	 */
	@RequestMapping(value = ACTION_RENEW_SUBSCRIBE_FACILITATOR, method = RequestMethod.POST)
	@ResponseBody()
	public ConfirmationMessageVO renewSubscribe(@RequestBody String body, HttpServletRequest request) {
		LOGGER.info(">>>>>> ManageUserController.renewSubscribe");
		JsonNode node = CommonUtility.getNode(body);
		
		String createdBy = node.get("createdBy").asText();
		String facilitatorID = node.get("facilitatorID").asText();
		String paymentType = node.get("paymentType").asText();
		//String bankDetails = node.get("bankDetails").asText();		
		String paymentDate = node.get("paymentDate").asText();
		String numberOfUsers = node.get("numberOfUsers").asText();
		String chequeNos = "";
		if(paymentType.equals("CHECK")){	// check no is available only for CHECK payment
			chequeNos = node.get("chequeNos").asText();
		}
		String totalAmountToBePaid = node.get("totalAmountToBePaid").asText();		
		String inputRenewSub = node.get("inputRenewSub").asText();
				
		int status = 0;
		ConfirmationMessageVO vo = new ConfirmationMessageVO();
		try {
			    //String facilitatorEmail = service.getFacilitatorEmailByID(facilitatorID);
			    if( facilitatorID == null ){	//	no such user
			    	StringBuilder msgBuilder = new StringBuilder(this.messageSource.getMessage("email.renewSubscribe.userNotExist",null, null));
					vo.setMessage(msgBuilder.toString());
					vo.setStatusCode(StatusConstants.STATUS_SUCCESS_WITH_NO_RESULT);
			    } else {
					FacilitatorAccountVO facilitatorChequeVO = new FacilitatorAccountVO(createdBy,facilitatorID,facilitatorID, paymentDate, numberOfUsers,
							chequeNos, totalAmountToBePaid, paymentType);					
					status = service.renewSubscribeViaChequePayment(facilitatorChequeVO,inputRenewSub);
					if (status == StatusConstants.STATUS_SUCCESS) {
						StringBuilder msgBuilder = new StringBuilder(this.messageSource.getMessage("msg.registered1",null, null));
						msgBuilder.append(facilitatorID);
						msgBuilder.append(this.messageSource.getMessage("email.renewSubscribe.registered2",null, null));
						vo.setMessage(msgBuilder.toString());
						vo.setStatusCode(status);
					} else if (status == StatusConstants.USER_REGISTERED_MAIL_SEND_ERROR) {
						StringBuilder msgBuilder = new StringBuilder(this.messageSource.getMessage("msg.registered1",null, null));
						msgBuilder.append(facilitatorID);
						msgBuilder.append(this.messageSource.getMessage("msg.registered2",null, null));
						msgBuilder.append(this.messageSource.getMessage("msg.registered4",null, null));
						vo.setMessage(msgBuilder.toString());
					} else if (status == StatusConstants.TABLE_OR_DATA_RELATED_ERROR) {
						StringBuilder msgBuilder = new StringBuilder(this.messageSource.getMessage("msg.registered5",null, null));
						vo.setMessage(msgBuilder.toString());
					} else { //FAILED
						StringBuilder msgBuilder = new StringBuilder(this.messageSource.getMessage("msg.registered5",null, null));
						vo.setMessage(msgBuilder.toString());
						vo.setStatusCode(123);
					}
			    }
			
		} catch (JCTException e) {
			LOGGER.error("UNABLE TO SAVE FACILITATOR: "+e.getLocalizedMessage());
			StringBuilder msgBuilder = new StringBuilder(this.messageSource.getMessage("msg.registered5",null, null));
			vo.setMessage(msgBuilder.toString()); 
			vo.setStatusCode(123);
		}
		LOGGER.info(">>>>>> ManageUserController.renewSubscribe");
		return vo;
	}
	
	
	/**
	 * Method to get facilitator email
	 * @param body
	 * @param request
	 * @return ConfirmationMessageVO.
	 */
	@RequestMapping(value = ACTION_GET_FACILITATOR_EMAIL, method = RequestMethod.POST)
	@ResponseBody()
	public UserVO getFacilitatorEmail(@RequestBody String body, HttpServletRequest request) {
		LOGGER.info(">>>>>> ManageUserController.getFacilitatorEmail");
		JsonNode node = CommonUtility.getNode(body);		
		String facilitatorID = node.get("facilitatorID").asText();				
		UserVO vo = new UserVO();
		try {
				String facilitatorEmail = service.getFacilitatorEmailByID(facilitatorID);
			    if( facilitatorEmail == null ) {	//	no such user			    	
			    	vo.setStatusCode(StatusConstants.STATUS_FAILURE.toString());
					vo.setStatusDesc(this.messageSource.getMessage("email.renewSubscribe.userNotExist",null, null));			    	
			    } else {
			    	vo.setStatusCode(StatusConstants.STATUS_SUCCESS_WITH_RESULT.toString());
			    	vo.setEmail(facilitatorEmail);			    	
			    }
		} catch (JCTException e) {
			LOGGER.error("UNABLE TO SAVE FACILITATOR: "+e.getLocalizedMessage());
			StringBuilder msgBuilder = new StringBuilder(this.messageSource.getMessage("msg.renewSub.fail",null, null));
			vo.setStatusDesc(msgBuilder.toString()); 
			vo.setStatusCode("123");
		}
		LOGGER.info(">>>>>> ManageUserController.getFacilitatorEmail");
		return vo;
	}
	/**
	 * Method fetches all existing users having cheque payment	
	 * @param body
	 * @param request
	 * @return
	 */
	@RequestMapping(value = ACTION_CHEQUE_EXISTING_DETAILS, method = RequestMethod.POST)
	@ResponseBody()
	public PaymentVO fetchExistingChequeUsers(@RequestBody String body, HttpServletRequest request) {
		LOGGER.info(">>>>>> ManageUserController.fetchExistingChequeUsers");
		JsonNode node = CommonUtility.getNode(body);
		String fetchType = node.get("fetchType").asText();
		String chequeNum = null;
		String userName = null;
		if(!fetchType.equals("")){
			chequeNum = node.get("chequeNum").asText();
			userName = node.get("userName").asText();
		}
		PaymentVO paymentVO=new PaymentVO();
		String userList = null;
		try{
			userList = service.fetchExistingChequeUsers(fetchType, chequeNum, userName);			
						 
			paymentVO.setChkDetailsString(userList);
			paymentVO.setStatusCode(StatusConstants.STATUS_SUCCESS);
			paymentVO.setStatusDesc("SUCCESS");
		}catch(JCTException e){
			paymentVO.setStatusCode(StatusConstants.STATUS_FAILURE);
			paymentVO.setStatusDesc(e.getLocalizedMessage());
			LOGGER.error(e.getLocalizedMessage());	
		}
		return paymentVO;
	}
		
	/**
	 * Method resets the user password, check register / honored for Individual and Facilitator
	 * @param body
	 * @param request
	 * @return return Map.
	 */
	@RequestMapping(value = ACTION_HONORED_CHECK, method = RequestMethod.POST)
	@ResponseBody()
	public PaymentVO honoredCheck(@RequestBody String body, HttpServletRequest request) 
			throws JCTException {
		PaymentVO vo = new PaymentVO();
		JsonNode node = CommonUtility.getNode(body);
		String tranId = node.get("tranId").asText();
		String createdBy = node.get("createdBy").asText();
		int honoredResultStatus = 0;	//	failed
		JctUser userObj = null;
		String fName = "";
		try{			
			honoredResultStatus = service.honoredCheck(tranId,createdBy);			
			if (honoredResultStatus != 0) {	//	success, user found
				List<PaymentDetailsDTO> userList =  service.getExistingUserByTranId(tranId);
				if(userList.size() > 0) {					
					try {		//	mail area starts from here
						for(int i = 0 ; i < userList.size(); i++){
							PaymentDetailsDTO obj = userList.get(i);
						
							if(obj.getRoleId() == 1) {	//	individual	
								fName = service.getFirstName(obj.getUserName(), 1);								
								if (!obj.getTypeDesc().equals("NEW_USER")) {	//	CHECK registration mail
									if(!obj.getTypeDesc().equals("RENEW") ){
										mailer.intimateUser(obj.getUserName(), decriptString(obj.getPassword()), fName,
												"register","CHK-CLEARED-INDIVIDUAL");	//	here paymentType =  CHK-CLEARED-INDIVIDUAL
									} 									
								} else {	//	//	registration check cleared
									mailer.intimateUser(obj.getUserName(), decriptString(obj.getPassword()), fName,
											"ChkCleared","CHK-CLEARED-INDIVIDUAL");	//	here paymentType =  CHK-CLEARED-INDIVIDUAL									
								}
							} else if (obj.getRoleId() == 3) {		//	3 - facilitator
								fName = service.getFirstName(obj.getUserName(), 3);
								if (obj.getTypeDesc().equals("NEW_FACILITATOR")) {		//	registration check cleared
									mailer.intimateFaci(obj.getUserName(), decriptString(obj.getPassword()), fName,
											"register","CHK-CLEARED-FACILITATOR");	//	here paymentType =  CHK-CLEARED-FACILITATOR	
								} else { // Sub / Renew check cleared									
									mailer.intimateFaci(obj.getUserName(), decriptString(obj.getPassword()), fName,
											"subRenew","CHK-CLEARED-FACILITATOR");	//	here paymentType =  CHK-CLEARED-FACILITATOR								
								}
							} else if (obj.getRoleId() == 5) {		//	5 - facilitator	+ individual			//	registration check cleared
								userObj = service.getPassword(obj.getUserName(), 1);
								fName = service.getFirstName(obj.getUserName(), 1);
								mailer.intimateUser(obj.getUserName(), decriptString(userObj.getJctPassword()), fName,
										"register","CHK-CLEARED-INDIVIDUAL");	//	here paymentType =  CHK-CLEARED-INDIVIDUAL	
								fName = service.getFirstName(obj.getUserName(), 3);
								userObj = service.getPassword(obj.getUserName(), 3);
								mailer.intimateFaci(obj.getUserName(), decriptString(obj.getPassword()),fName,
											"register","CHK-CLEARED-FACILITATOR");	//	here paymentType =  CHK-CLEARED-FACILITATOR	
							}
						}	//	for
					} catch (MailingException e) {
						LOGGER.error("FACILITATOR REGISTERED....COULD NOT SEND EMAIL...  REASON: "+e.getLocalizedMessage());
						vo.setStatusCode(StatusConstants.USER_REGISTERED_MAIL_SEND_ERROR);
					}
				}
				
				vo.setStatusCode(StatusConstants.STATUS_SUCCESS);
				vo.setStatusDesc("SUCCESS");
			} else {
				vo.setStatusCode(StatusConstants.STATUS_FAILURE);
				vo.setStatusDesc("FAILURE");
			}
		} catch(JCTException e) {
			vo.setStatusCode(StatusConstants.STATUS_FAILURE);
			vo.setStatusDesc(e.getLocalizedMessage());
			LOGGER.error(e.getLocalizedMessage());	
		}	
		return vo;
	}
	
/**
 * Method to dishoner payment check
 * @param body
 * @param request
 * @return return Map.
 */
@RequestMapping(value = ACTION_DISHONORED_CHECK, method = RequestMethod.POST)
@ResponseBody()
public PaymentVO dishonoredCheck(@RequestBody String body, HttpServletRequest request) 
		throws JCTException {
	LOGGER.info(">>>>>> ManageUserController.dishonoredCheck");
	PaymentVO vo = new PaymentVO();
	JsonNode node = CommonUtility.getNode(body);
	String tranId = node.get("tranId").asText();
	String createdBy = node.get("createdBy").asText();
	String fName = "";
	try{
		String dishonoredResult = service.dishonoredCheck(tranId,createdBy);
		if (dishonoredResult.length() != 0) {
				try {
					String[] singleUserStr = dishonoredResult.split("#");
					for (int i = 0 ; i < singleUserStr.length ; i++){
						String[] eachInfo = singleUserStr[i].split("~");
						fName = eachInfo[5];					
						mailer.intimateDishonoredUser(eachInfo[0], fName, eachInfo[1], Integer.parseInt(eachInfo[2]), eachInfo[3], eachInfo[4]);//	to,chkNo,role,paymntType,payDate
					}
					
				} catch (MailingException e) {
					LOGGER.error("FACILITATOR REGISTERED....COULD NOT SEND EMAIL...  REASON: "+e.getLocalizedMessage());
					vo.setStatusCode(StatusConstants.USER_REGISTERED_MAIL_SEND_ERROR);
				}			
			vo.setStatusCode(StatusConstants.STATUS_SUCCESS);
			vo.setStatusDesc("SUCCESS");
		} else {
			vo.setStatusCode(StatusConstants.STATUS_FAILURE);
			vo.setStatusDesc("FAILURE");
		}
	} catch(JCTException e) {
		vo.setStatusCode(StatusConstants.STATUS_FAILURE);
		vo.setStatusDesc(e.getLocalizedMessage());
		LOGGER.error(e.getLocalizedMessage());	
	}
	LOGGER.info("<<<<<< ManageUserController.dishonoredCheck");
	return vo;
}
/**
 * Method decript the password string 
 * @param string
 * @return string
 */
private String decriptString(String string){	  
	String ref = "P!QR#S$T%U&VW(X)Y*Z+[,-].^/_0`1a2b3c4d5e6f7g8h9i:j;k<l=m>n?o@pAqBrCsDtEuFvGwHxIyJzK{L|M}N~O";
    StringBuilder result = new StringBuilder("");
    for(int i=0;i<string.length()/2;i++){
                    String charStr     = string.substring (i, i+1);
                    int num = ref.indexOf(charStr);
                    String original = ref.substring(num-1, num);
                    result.append(original);
    }
    return result.toString();
}	
/**
 * Method update existing user group
 * @param body
 * @param request
 * @return return status code only.
 */
@RequestMapping(value = ACTION_UPDATE_USER_GROUP_FACILITATOR, method = RequestMethod.POST)
@ResponseBody()
public ManageUserVO updateUserGroupForFacilitator(@RequestBody String body, HttpServletRequest request) throws JCTException{
	LOGGER.info(">>>>>> ManageUserController.updateUserGroupForFacilitator");
	ManageUserVO manageVO = new ManageUserVO();
	UserGroupVO userGroupVO = null;
	JsonNode node = CommonUtility.getNode(body);
	String userGroup = node.get("userGroupVal").toString().replaceAll("\"" , "").toUpperCase().trim();
	String userProfileValue = node.get("userProfileVal").toString().replaceAll("\"" , "");
	Integer userProfileId = Integer.parseInt(node.get("userProfileId").toString().replaceAll("\"" , ""));
	Integer tablePkId = Integer.parseInt(node.get("tablePkId").toString().replaceAll("\"" , ""));
	String customerId = node.get("customerId").toString().replaceAll("\"" , "");
	Integer roleId = Integer.parseInt(node.get("roleId").toString().replaceAll("\"" , ""));
	//int profileId = Integer.parseInt(node.get("profileId").toString().replaceAll("\"" , ""));
	userGroupVO = new UserGroupVO();
	userGroupVO.setUserGroupName(node.get("userGroupVal").toString().replaceAll("\"" , "").trim());
	userGroupVO.setUserProfileName(userProfileValue);
	userGroupVO.setUserProfileId(userProfileId);
	userGroupVO.setPrimaryKeyVal(tablePkId);		
	String updationSuccess = service.validateExistenceUserGrp(userGroup, userProfileId, roleId, customerId);
	if (updationSuccess.equals("success")) {
		manageVO.setStatusCode(StatusConstants.ALREADY_EXISTS);
	} else {
		manageVO.setStatusCode(StatusConstants.DOES_NOT_EXIST);
		try {
			manageVO = service.updateUserGroup(userGroupVO, null);
			ManageUserVO unused = service.populateExistingUserGroup(userProfileId, customerId, roleId);
			manageVO.setExistingUserGroupList(unused.getExistingUserGroupList());
			unused = null;
		} catch (JCTException e) {
			manageVO = new ManageUserVO();
			manageVO.setStatusCode(StatusConstants.STATUS_FAILURE);
			manageVO.setStatusDesc(this.messageSource.getMessage("exception.dao.user.group.update",null, null));
			LOGGER.error(e.getLocalizedMessage());
		} 
	}
	LOGGER.info("<<<<<< ManageUserController.updateUserGroupForFacilitator");
		return manageVO;
}

@RequestMapping(value = ACTION_RENEW_GEN_USER_CSV_PAYMENT, method = RequestMethod.POST)
@ResponseBody()
public ManageUserVO renewGeneralUserCSV(@RequestBody String body, HttpServletRequest request) {
	LOGGER.info(">>>>>> ManageUserController.renewGeneralUserPaymentManual");	
	ManageUserVO manageVO = new ManageUserVO();
	JsonNode node = CommonUtility.getNode(body);	
	String emailIdString = node.get("emailIdString").asText();
	String paymentDate = node.get("paymentDate").asText();
	String numberOfUsers = node.get("numberOfUsers").asText();
	String chequeNos = node.get("chequeNos").asText();
	String totalAmountToBePaid = node.get("totalAmountToBePaid").asText();
	String createdBy = node.get("createdBy").asText();
	String paymentType = node.get("paymentType").asText();
	
	List<String> validEmailList = new ArrayList<String>();
	List<String> emailWithExpDateList = new ArrayList<String>();
	StringTokenizer emailToken = new StringTokenizer(emailIdString, "~");
	System.out.println(emailToken.countTokens());
	while (emailToken.hasMoreTokens()) {
		String tok = (String) emailToken.nextToken().replaceAll("\"" , "");
		emailWithExpDateList.add(tok);
		StringTokenizer emailIds = new StringTokenizer(tok, "#`#");
		while (emailIds.hasMoreTokens()) {
			validEmailList.add((String) emailIds.nextToken().replaceAll("\"" , ""));
			emailIds.nextToken();
		}
	}
	GeneralUserAccountVO accountVO = new GeneralUserAccountVO(createdBy,
			paymentType, emailWithExpDateList, paymentDate, numberOfUsers,
			chequeNos, totalAmountToBePaid);
	try {
		manageVO.setStatusDesc(service.renewGeneralUserPaymentManual(accountVO, validEmailList.size(),paymentType));
		manageVO.setStatusCode(StatusConstants.STATUS_SUCCESS_WITH_RESULT);
	} catch (JCTException ex) {
		LOGGER.error("UNABLE TO RENEW GENERAL USER BY CSV: "+ex.getLocalizedMessage());
	}
	return manageVO;
}

/**
 * Method adds to renew individual user. 
 * This method takes care for Manual entry.
 * @param body
 * @param request
 * @return ConfirmationMessageVO.
 */
@RequestMapping(value = ACTION_RENEW_GEN_USER_MANUAL_PAYMENT, method = RequestMethod.POST)
@ResponseBody()
public ManageUserVO renewGeneralUserPaymentManual(@RequestBody String body, HttpServletRequest request) {
	LOGGER.info(">>>>>> ManageUserController.renewGeneralUserPaymentManual");	
	ManageUserVO manageVO = new ManageUserVO();
	ArrayList<String> validEmailList = new ArrayList<String>();
	ArrayList<String> inValidEmailList = new ArrayList<String>();
	//ArrayList<String> expDatelist = new ArrayList<String>();
	
	ArrayList<String> emailWithExpDateList = new ArrayList<String>();
	
	JsonNode node = CommonUtility.getNode(body);	
	String emailIdString = node.get("emailIdString").asText();
	String paymentDate = node.get("paymentDate").asText();
	String numberOfUsers = node.get("numberOfUsers").asText();
	String chequeNos = node.get("chequeNos").asText();
	String totalAmountToBePaid = node.get("totalAmountToBePaid").asText();
	String createdBy = node.get("createdBy").asText();
	String paymentType = node.get("paymentType").asText();
	//String expDateString = node.get("expDateString").asText();
	
	StringTokenizer emailToken = new StringTokenizer(emailIdString, "~");
	while (emailToken.hasMoreTokens()) {
		String emailWithExpDuratn = (String) emailToken.nextToken().replaceAll("\"" , "");
		
		String[] initialValue = emailWithExpDuratn.split("#`#");
		emailWithExpDateList.add(emailWithExpDuratn);
		
		
		Boolean isValid = CommonUtility.isEmailValid(initialValue[0]);
		if (isValid) {
			validEmailList.add(initialValue[0]);
		} else {
			inValidEmailList.add(initialValue[0]);
		}
	}
	
	List<String> nonExistingUserList = new ArrayList<String>();
	List<String> existingUserList = null;
	List<String> inactiveUserList = null;
	// If email ids are invalid and email ids exist
	if (validEmailList.size() > 0) {
		try {
			existingUserList = service.getUserNamesOfIndUser(validEmailList);
			if(existingUserList.size() == 0) {
				nonExistingUserList = validEmailList;
			} else {
				for (int ind = 0; ind < validEmailList.size(); ind++) {
					String uName = validEmailList.get(ind).toString();
					if (!existingUserList.contains(uName)) {
						nonExistingUserList.add(uName);
					} else {
						
					}
				}
			}			
		} catch (JCTException ex) {
			LOGGER.error("UNABLE TO FETCH USER FOR RENEW BY MANUAL: "+ex.getLocalizedMessage());
		}		
		
	} else {
		existingUserList = new ArrayList<String>();
	}
	
	if (validEmailList.size() > 0) {
		try {
			inactiveUserList = service.getInactiveUserListByEmail(validEmailList);
		} catch (JCTException e) {
			LOGGER.error("UNABLE TO FETCH USER FOR RENEW BY MANUAL: "+e.getLocalizedMessage());
		}
	}
	
	if(inValidEmailList.size() > 0) {  //FOR INVALID USER
		LOGGER.info("SOME EMAILS IDS ARE INVALID");
		manageVO.setStatusCode(StatusConstants.INVALID_EMAIL_ID_FOUND_CSV_UPLD);		
		manageVO.setInvalidEmailIds(inValidEmailList);
		//return manageVO;
	} else if (nonExistingUserList.size() > 0 && inValidEmailList.size() == 0) {  // CHECK EXISTENCE OF USER
		LOGGER.info("SOME EMAILS ARE DOES NOT EXIST, HENCE RETURNING WITH ERROR MESSAGE");
		manageVO.setStatusCode(StatusConstants.DOES_NOT_EXIST);		
		manageVO.setNonExistingEmailIds(nonExistingUserList);
		//return manageVO;
	} else if (nonExistingUserList.size() == 0 && inValidEmailList.size() == 0 && inactiveUserList.size() > 0) {  // CHECK INACTIVE USER
		LOGGER.info("SOME EMAILS ARE NOT ACTIVATED, HENCE RETURNING WITH ERROR MESSAGE");
		manageVO.setStatusCode(StatusConstants.STATUS_INACTIVE);		
		manageVO.setInactiveEmailIds(inactiveUserList);
		//return manageVO;
	} else {
		GeneralUserAccountVO accountVO = new GeneralUserAccountVO(createdBy,
				paymentType, emailWithExpDateList, paymentDate, numberOfUsers,
				chequeNos, totalAmountToBePaid);
		try {
			manageVO.setStatusDesc(service.renewGeneralUserPaymentManual(accountVO, validEmailList.size(),paymentType));
			manageVO.setStatusCode(StatusConstants.STATUS_SUCCESS_WITH_RESULT);
		} catch (JCTException ex) {
			LOGGER.error("UNABLE TO SAVE GENERAL USER BY MANUAL: "+ex.getLocalizedMessage());
		}
	}
	LOGGER.info(">>>>>> ManageUserController.renewGeneralUserPaymentManual");
	return manageVO;
}

/**
 * Method is used to fetch the sample PDF
 * @param body
 * @param request
 * @return
 */
@RequestMapping(value = "/downloadIndividualRenewCSVFile/{fileName}", method = RequestMethod.GET)
public ReturnVO downloadIndividualRenewCSVFile(@PathVariable("fileName") String fileName, HttpServletRequest request, HttpServletResponse response) {
    try {
    	LOGGER.info(">>>> ManageUserController.downloadIndividualRenewCSVFile");
    	String csvPath = this.messageSource.getMessage("csv.renew.file.location",null, null);
    	final File file = new File(csvPath);
        response.setContentType("text/csv");
        response.setHeader("Content-Disposition",
                "attachment; filename=" + file.getName());
        final BufferedReader br = new BufferedReader(new FileReader(file));
        try {
            String line;
            while ((line = br.readLine()) != null) {
                response.getWriter().write(line + "\n");
            }
        } finally {
            br.close();
        }	
    
    } catch (Exception e) {
    	LOGGER.error(e.getLocalizedMessage());
	}
    LOGGER.info("<<<<<< ManageUserController.downloadIndividualRenewCSVFile");
    return null;
}

@RequestMapping(value = "/shopifyUserChecker/{username}/{userType}", method = RequestMethod.GET)
@ResponseBody()
public CrossBrowserUserQueryVO individualUserExist(@PathVariable("username") String username, 
		@PathVariable("userType") String userType, HttpServletRequest request, HttpServletResponse response) throws JCTException {
	CrossBrowserUserQueryVO vo = new CrossBrowserUserQueryVO();
	response.setHeader("Access-Control-Allow-Origin", "*");
	// Check if individual user already exists in the system.
	boolean userExists = service.shopifyUserChecker(username, userType);
	String exists = "NO";
	if (userExists == true) {
		exists = "YES";
	}
	vo.setExist(exists);
	return vo;
}
@RequestMapping(value = ACTION_VALIDATE_CSV_ENTRY, method = RequestMethod.POST)
@ResponseBody()
public ManageUserVO validateCSVEntry (@RequestBody String body, HttpServletRequest request) {
	LOGGER.info(">>>>>> ManageUserFacilitatorController.validateCSVFacilitator");
	ManageUserVO manageVO = new ManageUserVO();
	ArrayList<String> validEmailList = new ArrayList<String>();
	ArrayList<String> inValidEmailList = new ArrayList<String>();
	ArrayList<String> emailList = new ArrayList<String>();
	JsonNode node = CommonUtility.getNode(body);
	String emailIdString = node.get("emailIdString").asText();
	String selectionType = node.get("selectionType").asText();
	ArrayList<String> emailDurationList = new ArrayList<String>(); 
	StringTokenizer emailToken = new StringTokenizer(emailIdString, "~");
	while (emailToken.hasMoreTokens()) {
		String tok = (String) emailToken.nextToken().replaceAll("\"" , "");
		StringTokenizer emailIds = new StringTokenizer(tok, "#`#");
		while (emailIds.hasMoreTokens()) {
			String emailId = (String) emailIds.nextToken().replaceAll("\"" , "");
			String expDt = (String) emailIds.nextToken().replaceAll("\"" , "");
			emailList.add(emailId);
			emailDurationList.add(emailId+"`~`"+expDt);
			
		}
	}
	// Iterate over the email Ids and check for its validity
	for (int index = 0; index < emailList.size(); index++) {
		Boolean isValid = this.isEmailValid(emailList.get(index));
		if (isValid) {
			// Check if email ids exist
			try {
				boolean emailExist = service.isIndividualActiveUserExist(emailList.get(index)); // 3
				if (emailExist) {
					// Is already present in the list
					if (!validEmailList.contains(emailList.get(index))) {
						String expDt[] = emailDurationList.get(index).toString().split("`~`");
						if (Integer.parseInt(expDt[1]) > 12 || Integer.parseInt(expDt[1]) < 1) {
							inValidEmailList.add(emailList.get(index)+"#`~`#(Invalid months specified)");
						}
						
						validEmailList.add(emailList.get(index));
					} else {
						if (!inValidEmailList.contains(emailList.get(index)+"#`~`#(Listed More Than Once)")) {
							inValidEmailList.add(emailList.get(index)+"#`~`#(Listed More Than Once)");
						}
					}
				} else {
					// Check if the user
					boolean emailAgainExist = service.isIndividualInactiveUserExist(emailList.get(index)); // other than 3
					if (emailAgainExist) {
						inValidEmailList.add(emailList.get(index)+"#`~`#(Inactive User)");
					} else {
						inValidEmailList.add(emailList.get(index)+"#`~`#(Non Existing User)");
					}
				}
			} catch (JCTException e) {
				LOGGER.error("SOMETHING SERIOUSLY WRONG....");
			}
		} else {
			inValidEmailList.add(emailList.get(index)+"#`~`#(Domain Name / Email Format Incorrect)");
		}
	}
	// If all entries are valid and no duplicate or invalid email id exists
	if (inValidEmailList.size() == 0) {
		LOGGER.info("EVERYTHING IS CORRECT");
		manageVO.setStatusCode(StatusConstants.ALL_CORRECT);
		StringBuffer msg = new StringBuffer();
		if(selectionType.equalsIgnoreCase("csv")){
			msg.append(this.messageSource.getMessage("manage.user.inv.existing.csv.66",null, null));
		} else {
			msg.append(this.messageSource.getMessage("manage.user.inv.existing.manual.66",null, null));
		}
		msg.append(validEmailList.size());
		msg.append(this.messageSource.getMessage("manage.user.inv.existing.77",null, null));
		manageVO.setStatusDesc(msg.toString());
		// Append the end dates
		
		String emailString = node.get("emailIdString").asText();
		StringTokenizer token = new StringTokenizer(emailString, "~");
		List<String> finalList = new ArrayList<String>();
		while (token.hasMoreTokens()) {
			String tok = (String) token.nextToken().replaceAll("\"" , "");
			StringTokenizer emailIds = new StringTokenizer(tok, "#`#");
			while (emailIds.hasMoreTokens()) {
				String emailId = (String) emailIds.nextToken().replaceAll("\"" , "");
				String expiryMonths = (String) emailIds.nextToken().replaceAll("\"" , "");
				int expMnths = Integer.parseInt(expiryMonths);
				try {
					List<UserAccountDTO> dto = service.getUserEndDateDTO(emailId);
					UserAccountDTO userObj = dto.get(dto.size()-1);
					Date expiryDate = userObj.getJctAccountExpirationDate();
					
					Calendar calendar = Calendar.getInstance();			 
					java.util.Date now = calendar.getTime();
					java.sql.Timestamp currentTimestamp = new java.sql.Timestamp(now.getTime());
					Calendar cal = Calendar.getInstance(); 
					if (userObj.getCustomerId().substring(0,2).equals(CommonConstants.FACILITATOR_USER_CUST_CODE)) {
						if(expiryDate.getTime() < currentTimestamp.getTime()){
							expiryDate = currentTimestamp;
							cal.setTime(expiryDate);
							cal.add(Calendar.MONTH, 6);
						} 
						else{
							cal.setTime(expiryDate);
							cal.add(Calendar.MONTH, 6);
						}
					} else {
						if(expiryDate.getTime() < currentTimestamp.getTime()){
							expiryDate = currentTimestamp;
							cal.setTime(expiryDate);
							cal.add(Calendar.MONTH, expMnths);
						} 
						else{
							cal.setTime(expiryDate);
							cal.add(Calendar.MONTH, expMnths);
						}
					}
					finalList.add(emailId+"#`~`#"+cal.getTime());
					} catch (JCTException e) {
					e.printStackTrace();
				}
			}
		}
		manageVO.setValidEmailIds(finalList);
		return manageVO;
	} else {
		LOGGER.info("SOME EMAILS IDS ARE INVALID AND SOME EMAIL DOESNOT EXIST, HENCE RETURNING WITH ERROR MESSAGE");
		manageVO.setStatusCode(StatusConstants.TOTAL_CSV_UPLD_FAILURE);
		StringBuffer msg = new StringBuffer();
		if(selectionType.equalsIgnoreCase("csv")){
			msg.append(this.messageSource.getMessage("manage.user.inv.existing.csv.666",null, null));
		} else {
			msg.append(this.messageSource.getMessage("manage.user.inv.existing.manual.666",null, null));
		}	
		// Remove the duplicate entries
		for (int ind = 0; ind < inValidEmailList.size(); ind++) {
			String emailIdComp[] = inValidEmailList.get(ind).toString().split("#`~`#");
			if (validEmailList.contains(emailIdComp[0].toString().trim())) {
				validEmailList.remove(emailIdComp[0].toString().trim());
			}
		}
		msg.append(validEmailList.size());
		msg.append(this.messageSource.getMessage("manage.user.inv.existing.777",null, null));
		msg.append(inValidEmailList.size());
		if(selectionType.equalsIgnoreCase("csv")){
			msg.append(this.messageSource.getMessage("manage.user.inv.existing.csv.888",null, null));
		} else {
			msg.append(this.messageSource.getMessage("manage.user.inv.existing.manual.888",null, null));
		}		
		
		manageVO.setStatusDesc(msg.toString());
		manageVO.setValidEmailIds(validEmailList);
		manageVO.setInvalidEmailIds(inValidEmailList);
		return manageVO;
	}
}

/**
 * Method check the email id string is valid or not
 * @param emailId
 * @return true, false
 */
private Boolean isEmailValid(String emailId) {
	java.util.regex.Pattern p = java.util.regex.Pattern.compile(CommonConstants.EMAIL_PATTERN);
    java.util.regex.Matcher m = p.matcher(emailId);
    Boolean valid = false;
    if (m.matches()) {
    	String hostName = emailId.split("@")[1].toString();
    	//check if domain name is valid [number of mail servers configured]
    	try {
			int mailServerConfCount = CommonUtility.doLookup(hostName);
			if (mailServerConfCount > 0) {
				valid = true;
			}
		} catch (NamingException e) {
			
		}
    }
    return valid;
}


@RequestMapping(value = ACTION_VALIDATE_CSV_ENTRY_ADD_USER, method = RequestMethod.POST)
@ResponseBody()
public ManageUserVO validateCSVEntryOnAddUser(@RequestBody String body, HttpServletRequest request) {
	LOGGER.info(">>>>>> ManageUserFacilitatorController.validateCSVEntryOnAddUser");
	ManageUserVO manageVO = new ManageUserVO();
	ArrayList<String> validEmailList = new ArrayList<String>();
	ArrayList<String> inValidEmailList = new ArrayList<String>();
	ArrayList<String> emailList = new ArrayList<String>();
	JsonNode node = CommonUtility.getNode(body);
	String emailIdString = node.get("emailIdString").asText(); 
	String selectionType = node.get("selectionType").asText();
	String expiryDuration = node.get("expiryDuration").asText();
	StringTokenizer emailToken = new StringTokenizer(emailIdString, "~");
	while (emailToken.hasMoreTokens()) {
		

		String tok = (String) emailToken.nextToken().replaceAll("\"" , "");			
		String emailId = tok.split("#`#")[0];
		if(emailId.trim().length()>0){
			emailList.add(emailId);			
		}
		
		
		/*String emailId = (String) emailToken.nextToken().replaceAll("\"" , "");
		if(emailId.trim().length()>0){
			emailList.add(emailId);
		}*/
	}	
	
	// Iterate over the email Ids and check for its validity
	for (int index = 0; index < emailList.size(); index++) {
		Boolean isValid = this.isEmailValid(emailList.get(index));
		if (isValid) {
			// Check if email ids exist
			try {
				boolean emailExist = service.isIndividualUserExist(emailList.get(index)); // 3
				if (emailExist) {
					inValidEmailList.add(emailList.get(index)+"#`~`#(Already Existing User)");				
				} else {										
					if (!validEmailList.contains(emailList.get(index))) {
						 validEmailList.add(emailList.get(index));
					} else {
						if (!inValidEmailList.contains(emailList.get(index)+"#`~`#(Listed More Than Once)")) {
							inValidEmailList.add(emailList.get(index)+"#`~`#(Listed More Than Once)");
						}
					}				
				}
			} catch (JCTException e) {
				LOGGER.error("SOMETHING SERIOUSLY WRONG....");
			}
		} else {
			inValidEmailList.add(emailList.get(index)+"#`~`#(Domain Name / Email Format Incorrect)");
		}
	}
	// If all entries are valid and no duplicate or invalid email id exists
	if (inValidEmailList.size() == 0) {
		LOGGER.info("EVERYTHING IS CORRECT");
		manageVO.setStatusCode(StatusConstants.ALL_CORRECT);
		StringBuffer msg = new StringBuffer();
		if(selectionType.equalsIgnoreCase("csv")){
			msg.append(this.messageSource.getMessage("manage.user.inv.existing.csv.66",null, null));
		} else {
			msg.append(this.messageSource.getMessage("manage.user.inv.existing.manual.66",null, null));
		}
		msg.append(validEmailList.size());
		msg.append(this.messageSource.getMessage("manage.user.inv.existing.add.user.77",null, null));
		manageVO.setStatusDesc(msg.toString());
		// Append the end dates
				
		StringTokenizer token = new StringTokenizer(emailIdString, "~");		
		List<String> finalList = new ArrayList<String>();
		while (token.hasMoreTokens()) {						
			String mailId = (String) token.nextToken().replaceAll("\"" , "");	
			String emailId = mailId.split("#`#")[0];
			if(emailId.trim().length()>0){
				String expiryMonths = expiryDuration;
				int expMnths = Integer.parseInt(expiryMonths);								
				Calendar calendar = Calendar.getInstance();			 
				java.util.Date now = calendar.getTime();
				java.sql.Timestamp currentTimestamp = new java.sql.Timestamp(now.getTime());
				Calendar cal = Calendar.getInstance(); 				
				cal.setTime(currentTimestamp);
				cal.add(Calendar.MONTH, expMnths);					
				finalList.add(emailId+"#`~`#"+cal.getTime());				
			}
		}
		manageVO.setValidEmailIds(finalList);
		return manageVO;
	} else {
		LOGGER.info("SOME EMAILS IDS ARE INVALID AND SOME EMAIL DOESNOT EXIST, HENCE RETURNING WITH ERROR MESSAGE");
		manageVO.setStatusCode(StatusConstants.TOTAL_CSV_UPLD_FAILURE);
		StringBuffer msg = new StringBuffer();
		if(selectionType.equalsIgnoreCase("csv")){
			msg.append(this.messageSource.getMessage("manage.user.inv.existing.csv.666",null, null));
		} else {
			msg.append(this.messageSource.getMessage("manage.user.inv.existing.manual.666",null, null));
		}		
		// Remove the duplicate entries
		for (int ind = 0; ind < inValidEmailList.size(); ind++) {
			String emailIdComp[] = inValidEmailList.get(ind).toString().split("#`~`#");
			if (validEmailList.contains(emailIdComp[0].toString().trim())) {
				validEmailList.remove(emailIdComp[0].toString().trim());
			}
		}
		msg.append(validEmailList.size());
		msg.append(this.messageSource.getMessage("manage.user.inv.existing.777",null, null));
		msg.append(inValidEmailList.size());
		if(selectionType.equalsIgnoreCase("csv")){
			msg.append(this.messageSource.getMessage("manage.user.inv.existing.csv.888",null, null));
		} else {
			msg.append(this.messageSource.getMessage("manage.user.inv.manual.add.user.888",null, null));
		}				
		manageVO.setStatusDesc(msg.toString());
		manageVO.setValidEmailIds(validEmailList);
		manageVO.setInvalidEmailIds(inValidEmailList);
		return manageVO;
	}
	}

/**
 * Method fetches user's expiration dtls
 * @param body
 * @param request
 * @return return VO.
 */
@RequestMapping(value = ACTION_FETCH_INDIVIDUAL_RENEW_DETAILS, method = RequestMethod.POST)
@ResponseBody()
public ExistingUsersVO fetchIndividualDetails(@RequestBody String body, HttpServletRequest request){
	LOGGER.info(">>>>>> ManageUserController.fetchIndividualDetails");
	JsonNode node = CommonUtility.getNode(body);
	ExistingUsersVO vo = new ExistingUsersVO();
	List<ExistingUserDTO> userDTOList = new ArrayList<ExistingUserDTO>();
	int usrType = Integer.parseInt(node.get("usrType").toString().replaceAll("\"" , ""));
	String emailId = node.get("emailId").toString().replaceAll("\"" , "");
	ExistingUserDTO userDTO = new ExistingUserDTO();
	try {
		userDTO = service.getUserExpirationDtls(emailId, usrType);
		if(null != userDTO) {	//	user found
			//userDTOList.set(0, userDTO);
			userDTOList.add(userDTO);
			
			vo.setExistingUsers(userDTOList);
			vo.setStatusCode(StatusConstants.STATUS_USER_EXIST);
		} else {
			vo.setStatusCode(StatusConstants.DOES_NOT_EXIST);
		}
	} catch (JCTException e) {
		LOGGER.error(e.getLocalizedMessage());
	}
	LOGGER.info("<<<<<< ManageUserController.fetchIndividualDetails");
	return vo;
}
/**
 * Method fetches user's expiration dtls
 * @param body
 * @param request
 * @return return VO.
 */
@RequestMapping(value = ACTION_FETCH_FACILITATOR_RENEW_DETAILS, method = RequestMethod.POST)
@ResponseBody()
public ExistingFacilitatorVO fetchFacilitatorDetails(@RequestBody String body, HttpServletRequest request){
	LOGGER.info(">>>>>> ManageUserController.fetchIndividualDetails");
	JsonNode node = CommonUtility.getNode(body);
	ExistingFacilitatorVO vo = new ExistingFacilitatorVO();
	int usrType = Integer.parseInt(node.get("usrType").toString().replaceAll("\"" , ""));
	String custId = node.get("custId").toString().replaceAll("\"" , "");
	FacilitatorDetailDTO userDTO = new FacilitatorDetailDTO();
	try {
		userDTO = service.getFaciExpirationDtls(custId, usrType);
		if(null != userDTO) {	//	user found
			
			vo.setFaciUserName(userDTO.getJctFacUserName());
			vo.setJctFacSubscribeLimit(userDTO.getJctFacSubscribeLimit());
			vo.setJctFacTotalLimit(userDTO.getJctFacTotalLimit());
			vo.setStatusCode(StatusConstants.STATUS_USER_EXIST);
		} else {
			vo.setStatusCode(StatusConstants.DOES_NOT_EXIST);
		}
	} catch (JCTException e) {
		LOGGER.error(e.getLocalizedMessage());
	}
	LOGGER.info("<<<<<< ManageUserController.fetchIndividualDetails");
	return vo;
}

/**
 * Method handles reset password functionality
 * @param body
 * @param request
 * @return
 */
@RequestMapping(value = "/forgotPassword", method = RequestMethod.POST)
@ResponseBody()
public ReturnVO forgotPassword(@RequestBody String body,
		HttpServletRequest request) {
	LOGGER.info(">>>> ManageUserController.resetPassword");
	HttpSession httpSession = request.getSession();
	ReturnVO returnVO = new ReturnVO();
	JsonNode node = CommonUtility.getNode(body);
	String emailId = node.get("emailId").asText(); // This is the username	
	
	String status = null;
	try {
		status = service.forgotPassword(emailId);
		if(status.equalsIgnoreCase("success")){
			returnVO.setStatusCode(CommonConstants.SUCCESSFUL);
			returnVO.setStatusDesc(this.messageSource.getMessage("msg.forgotPass.success",null, null));
		} else {
			returnVO.setStatusCode(status);			
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
	LOGGER.info("<<<< ManageUserController.resetPassword");
	return returnVO;
}
}