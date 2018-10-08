package com.vmware.jct.controller;

import java.util.ArrayList;
import java.util.Iterator;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.Random;
import java.util.StringTokenizer;

import javax.naming.NamingException;
import javax.servlet.http.HttpServletRequest;

import org.apache.commons.lang.WordUtils;
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
import com.vmware.jct.common.utility.StatusConstants;
import com.vmware.jct.dao.dto.ExistingUserDTO;
import com.vmware.jct.dao.dto.NewUserDTO;
import com.vmware.jct.dao.dto.UserDTO;
import com.vmware.jct.dao.dto.UserGroupDTO;
import com.vmware.jct.exception.JCTException;
import com.vmware.jct.exception.MailingException;
import com.vmware.jct.service.IManageFacilitatorService;
import com.vmware.jct.service.IManageUserService;
import com.vmware.jct.service.vo.ConfirmationMessageVO;
import com.vmware.jct.service.vo.ExistingUsersVO;
import com.vmware.jct.service.vo.ManageUserVO;
import com.vmware.jct.service.vo.ReportsVO;
import com.vmware.jct.service.vo.UserGroupVO;

/**
 * 
 * <p><b>Class name:</b> ManageUserFacilitatorController.java</p>
 * <p><b>Author:</b> InterraIT</p>
 * <p><b>Purpose:</b> This class acts as a controller for all User Management on Facilitator Side..
 * ManageUserFacilitatorController extends BasicController  and has following  methods.
 * -populateSubscribedUser()
 * -populateUsersByFacilitatorId()
 * -saveManualUserFacilitator()
 * -isEmailValid()
 * -saveCSVFacilitator()
 * -fetchActiveInactiveUserListFacilitator()
 * -fetchUserListForResetPassword()
 * -changeUserStatusInBatchFacilitator()
 * -populateActiveInactiveUserList()
 * -resetPasswordInBatchFacilitator()
 * -resetMassPassword()
 * -generatePassword()
 * -encriptString()
 * -renewUserByFacilitator()
 * -populateUserDropDown()
 * -populateExistingFacilitator()
 * -updateUserToFacilitator()
 * -addNewFacilitator()
 * <p><b>Description:</b> This class acts as a controller for all User Management on Facilitator Side. </p>
 * <p><b>Copyrights:</b> All rights reserved by InterraIT and should only be used for its internal application development.</p>
 * <p><b>Created Date:</b> 14/Oct/2014</p>
 * <p><b>Revision History:</b>
 * 	<li></li>
 * </p>
 */

@Controller
@RequestMapping(value = "/manageuserFacilitator")
public class ManageUserFacilitatorController extends BasicController {

	@Autowired
	private IManageFacilitatorService service;
	
	@Autowired
	private IManageUserService userService;
	
	@Autowired  
	private MessageSource messageSource;
	
	@Autowired
	private MailNotificationHelper mailer;
	
	private static final Logger LOGGER = LoggerFactory.getLogger(ManageUserFacilitatorController.class);
	
	/**
	 * Method populates the Total Subscribed User and total Registered User
	 * @param body
	 * @param request
	 * @return ManageUserVO containing Total Subscribed User and Registered User
	 */
	@RequestMapping(value = ACTION_POPULATE_SUBSCRIBED_USER, method = RequestMethod.POST)
	@ResponseBody()
	public ManageUserVO populateSubscribedUser(@RequestBody String body, HttpServletRequest request) {
		LOGGER.info(">>>>>> ManageUserFacilitatorController.populateSubscribedUser");
		ManageUserVO manageVO = null;
		JsonNode node = CommonUtility.getNode(body);
		String emailAddress = node.get("emailId").toString().replaceAll("\"" , "");
		String type = node.get("type").toString().replaceAll("\"" , "");
		String customerId = node.get("customerId").toString().replaceAll("\"" , "");	
		try {
			manageVO = service.populateSubscribedUser(emailAddress, type, customerId);
		} catch (JCTException e) {
			manageVO = new ManageUserVO();
			manageVO.setStatusCode(StatusConstants.STATUS_FAILURE);
			manageVO.setStatusDesc(this.messageSource.getMessage("exception.dao.error",null, null));
			LOGGER.error(e.getLocalizedMessage());
		}
		LOGGER.info("<<<<<< ManageUserFacilitatorController.populateSubscribedUser");
		return manageVO;
	}
	
	/**
	 * Method populate existing users
	 * @param body
	 * @param request
	 * @return existingUserVO containing user list
	 */
	@RequestMapping(value = ACTION_POPULATE_USERS_BY_FACILITATOR, method = RequestMethod.POST)
	@ResponseBody()
	public ExistingUsersVO populateUsersByFacilitatorId(@RequestBody String body, HttpServletRequest request) throws JCTException {
		LOGGER.info(">>>>>> ManageUserController.populateExistingUsers");
		ExistingUsersVO existingUserVO = new ExistingUsersVO();
		JsonNode node = CommonUtility.getNode(body);
		//String emailAddress = node.get("emailId").toString().replaceAll("\"" , "");
		String customerId = node.get("customerId").toString().replaceAll("\"" , "");
		List<ExistingUserDTO> userList = service.getUserList(customerId);
		if ( null == userList ) {
			userList = new ArrayList<ExistingUserDTO>();
		}
		existingUserVO.setStatusCode(StatusConstants.STATUS_SUCCESS_WITH_RESULT);
		existingUserVO.setExistingUsers(userList);
		LOGGER.info("<<<<<< ManageUserController.populateExistingUsers");
		return existingUserVO;
	}
	
	
	/**
	 * Method adds new user for facilitator by manual entry
	 * @param body
	 * @param request
	 * @return ManageUserVO.
	 */
	@RequestMapping(value = ACTION_SAVE_MANUAL_USER_FACILITATOR, method = RequestMethod.POST)
	@ResponseBody()
	public ManageUserVO saveManualUserFacilitator(@RequestBody String body, HttpServletRequest request) {
		LOGGER.info(">>>>>> ManageUserFacilitatorController.saveManualUserFacilitator");
		ManageUserVO manageVO = new ManageUserVO();
		ArrayList<String> validEmailList = new ArrayList<String>();
		ArrayList<String> inValidEmailList = new ArrayList<String>();
		JsonNode node = CommonUtility.getNode(body);
		String userGroup = node.get("userGroup").toString().replaceAll("\"" , "");
		String emailIdString = node.get("emailIdString").toString().replaceAll("\"" , "");
		String createdBy = node.get("createdBy").toString().replaceAll("\"" , "");
		String expiryDate = node.get("expiryDate").toString().replaceAll("\"" , "");
		String customerId = node.get("customerId").toString().replaceAll("\"" , "");
		String newUserGroup = node.get("newUserGroup").toString().replaceAll("\"" , "");	
		String userGroupUpper = node.get("userGroup").toString().replaceAll("\"" , "").toUpperCase().trim();
		String faciName = "";
		
		StringTokenizer emailToken = new StringTokenizer(emailIdString, "~");
		while (emailToken.hasMoreTokens()) {
			
			String email = (String) emailToken.nextToken().replaceAll("\"" , "");
			validEmailList.add(email);
			//String mailId = email.split("#")[0];
			/*Boolean isValid = this.isEmailValid(email);
			if (isValid) {
				validEmailList.add(email);
			} else {
				inValidEmailList.add(email);
			}*/
		}
		try {
			
			if(newUserGroup.equals("YES")) {
				String updationSuccess = userService.validateExistenceUserGrp(userGroupUpper, 2, 3, customerId);
				if (updationSuccess.equals("success")) {
					manageVO.setStatusCode(StatusConstants.ALREADY_EXISTS);
				} else {
					manageVO.setStatusCode(StatusConstants.DOES_NOT_EXIST);
					UserGroupVO userGroupVO = new UserGroupVO();
					userGroupVO = new UserGroupVO();
					userGroupVO.setUserGroupName(userGroup);
					userGroupVO.setUserProfileName("Default Profile");
					userGroupVO.setUserProfileId(2);
					userGroupVO.setCreatedBy(createdBy);	
					userGroupVO.setCustomerId(customerId);
					userGroupVO.setRoleId(3);
					manageVO = userService.saveUserGroup(userGroupVO);
				}
			} else {
				manageVO.setStatusCode(StatusConstants.STATUS_SUCCESS_WITH_RESULT);
			}
			if(manageVO.getStatusCode() == 200) {	
				//NewUserDTO saveUserDTO = service.saveUsers(validEmailList, Integer.parseInt(grpSplit[0]), grpSplit[1], createdBy);
				NewUserDTO saveUserDTO = service.saveUsers(validEmailList, userGroup, expiryDate, createdBy, customerId, newUserGroup);
				manageVO.setStatusCode(StatusConstants.USER_ONLY_REGISTERED);
				manageVO.setStatusDesc(this.messageSource.getMessage("manage.user.registered.no.email",null, null));

				/*try{
					LOGGER.info("INTIMATING THE ADMIN...");
					mailer.intimateFacilitator(saveUserDTO, createdBy, inValidEmailList);
				} catch (MailingException e) {
					LOGGER.error("UNABLE TO SEND MAIL TO ADMIN: "+e.getLocalizedMessage());
					manageVO.setStatusCode(StatusConstants.USER_REGISTERED_MAIL_SEND_ERROR_ADMIN);
					manageVO.setStatusDesc(this.messageSource.getMessage("manage.user.registered.no.email.error",null, null));
				}*/
				if(null != saveUserDTO.getNewRegistration()){
					
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
								faciName = userService.getFaciName(createdBy);
								mailer.intimateUserByFaciliator(dto.getEmailId(), dto.getPassword(), "register", 
										WordUtils.capitalize(dto.getFirstName()), createdBy, faciName);
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
				} else {
					manageVO.setStatusCode(StatusConstants.USER_REGISTERED_MAIL_SEND_ERROR_ADMIN);
					manageVO.setStatusDesc(this.messageSource.getMessage("manage.user.registration.admin.cron.confirmation.record.invalid",null, null));
				}		
						
			} else {
				manageVO = new ManageUserVO();
				manageVO.setStatusCode(StatusConstants.ALREADY_EXISTS);
			} 
		} catch (JCTException e) {
				LOGGER.error("-----> "+e.getLocalizedMessage());
				manageVO.setStatusCode(StatusConstants.USER_NOT_REGISTERED);
				manageVO.setStatusDesc(this.messageSource.getMessage("manage.user.registration.failed",null, null));
			}
		LOGGER.info("<<<<<< ManageUserFacilitatorController.saveManualUserFacilitator");
		return manageVO;
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

	
	/**
	 * Method adds new user with CSV file
	 * @param body
	 * @param request
	 * @return ManageUserVO.
	 */
	@RequestMapping(value = ACTION_VALIDATE_CSV_FACILITATOR, method = RequestMethod.POST)
	@ResponseBody()
	public ManageUserVO validateUserDataFacilitator(@RequestBody String body, HttpServletRequest request) {
		LOGGER.info(">>>>>> ManageUserFacilitatorController.validateUserDataFacilitator");
		ManageUserVO manageVO = new ManageUserVO();
		ArrayList<String> validEmailList = new ArrayList<String>();
		ArrayList<String> inValidEmailList = new ArrayList<String>();
		ArrayList<String> emailList = new ArrayList<String>();		
		JsonNode node = CommonUtility.getNode(body);
		String emailIdString = node.get("emailIdString").toString().replaceAll("\"" , "");
		String selectionType = node.get("selectionType").asText();
		String expiryDate = node.get("expiryDate").asText();		
		StringTokenizer emailToken = new StringTokenizer(emailIdString, "~");		
		while (emailToken.hasMoreTokens()) {
			String tok = (String) emailToken.nextToken().replaceAll("\"" , "");			
			String mailId = tok.split("#`#")[0];
			emailList.add(mailId);			
		}	
		// Iterate over the email Ids and check for its validity
		for (int index = 0; index < emailList.size(); index++) {
			Boolean isValid = this.isEmailValid(emailList.get(index));
			if (isValid) {
				// Check if email ids exist
				try {
					boolean emailExist = userService.isIndividualUserExist(emailList.get(index)); // 3
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
				String emailId = (String) token.nextToken().replaceAll("\"" , "");		
				String mailId = emailId.split("#`#")[0];
				if(emailId.trim().length()>0){								
					finalList.add(mailId+"#`~`#"+expiryDate);				
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
		}
		return manageVO;
	}
	
	/**
	 * Method adds new user with CSV file
	 * @param body
	 * @param request
	 * @return ManageUserVO.
	 */
	@RequestMapping(value = ACTION_SAVE_USER_CSV_FACILITATOR, method = RequestMethod.POST)
	@ResponseBody()
	public ManageUserVO saveCSVFacilitator(@RequestBody String body, HttpServletRequest request) {
		LOGGER.info(">>>>>> ManageUserFacilitatorController.saveCSVFacilitator");
		ManageUserVO manageVO = new ManageUserVO();
		ArrayList<String> validEmailList = new ArrayList<String>();
		ArrayList<String> inValidEmailList = new ArrayList<String>();
		JsonNode node = CommonUtility.getNode(body);
		String userGroup = node.get("userGroup").toString().replaceAll("\"" , "");
		String emailIdString = node.get("emailIdString").toString().replaceAll("\"" , "");
		String createdBy = node.get("createdBy").toString().replaceAll("\"" , "");
		String expiryDate = node.get("expiryDate").toString().replaceAll("\"" , "");
		String customerId = node.get("customerId").toString().replaceAll("\"" , "");
		String newUserGroup = node.get("newUserGroup").toString().replaceAll("\"" , "");
		String userGroupUpper = node.get("userGroup").toString().replaceAll("\"" , "").toUpperCase().trim();
		String faciName = "";
		
		StringTokenizer emailToken = new StringTokenizer(emailIdString, "~");
		while (emailToken.hasMoreTokens()) {
			String email = (String) emailToken.nextToken().replaceAll("\"" , "");
			if (email.trim().length() > 5) {
				validEmailList.add(email);
			}
			
			/*String email = (String) emailToken.nextToken().replaceAll("\"" , "");
			Boolean isValid = this.isEmailValid(email);
			if (isValid) {
				validEmailList.add(email);
			} else {
				inValidEmailList.add(email);
			}*/
		}
		try {		
			if(newUserGroup.equals("YES")) {
				String updationSuccess = userService.validateExistenceUserGrp(userGroupUpper, 2, 3, customerId);
				if (updationSuccess.equals("success")) {
					manageVO.setStatusCode(StatusConstants.ALREADY_EXISTS);
				} else {
					manageVO.setStatusCode(StatusConstants.DOES_NOT_EXIST);
					UserGroupVO userGroupVO = new UserGroupVO();
					userGroupVO = new UserGroupVO();
					userGroupVO.setUserGroupName(userGroup);
					userGroupVO.setUserProfileName("Default Profile");
					userGroupVO.setUserProfileId(2);
					userGroupVO.setCreatedBy(createdBy);	
					userGroupVO.setCustomerId(customerId);
					userGroupVO.setRoleId(3);
					manageVO = userService.saveUserGroup(userGroupVO);
				}
			} else {
				manageVO.setStatusCode(StatusConstants.STATUS_SUCCESS_WITH_RESULT);
			}
			if(manageVO.getStatusCode() == 200) {								
				NewUserDTO saveUserDTO = service.saveUsers(validEmailList, userGroup, expiryDate, createdBy, customerId, newUserGroup);
				manageVO.setStatusCode(StatusConstants.USER_ONLY_REGISTERED);
				manageVO.setStatusDesc(this.messageSource.getMessage("manage.user.registered.no.email",null, null));
				
				//postpone the mail send and save it to database. Cron Job will send when triggered
				if(null != saveUserDTO.getNewRegistration()){
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
								/**
								 * Date : 22.06.2017 
								 * facilitator name added in mail content
								 */
								//mailer.intimateUser(dto.getEmailId(), dto.getPassword(), dto.getFirstName(),"register",null);
								faciName = userService.getFaciName(createdBy);
								mailer.intimateUserByFaciliator(dto.getEmailId(), dto.getPassword(), "register", 
										WordUtils.capitalize(dto.getFirstName()), createdBy, faciName);
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
				} else {
					manageVO.setStatusCode(StatusConstants.USER_REGISTERED_MAIL_SEND_ERROR_ADMIN);
					manageVO.setStatusDesc(this.messageSource.getMessage("manage.user.registration.admin.cron.confirmation.record.invalid",null, null));
				}		
			}  else {
				manageVO = new ManageUserVO();
				manageVO.setStatusCode(StatusConstants.ALREADY_EXISTS);
			}
		}catch (JCTException e) {
			LOGGER.error("-----> "+e.getLocalizedMessage());
			manageVO.setStatusCode(StatusConstants.USER_NOT_REGISTERED);
			manageVO.setStatusDesc(this.messageSource.getMessage("manage.user.registration.failed",null, null));
		}
		LOGGER.info("<<<<<< ManageUserFacilitatorController.saveCSVFacilitator");
		return manageVO;
	}
	
	/**
	 * Method to fetch the user list with active and inactive status
	 * @param body
	 * @param request
	 * @return ExistingUsersVO.
	 */
	@RequestMapping(value = ACTION_POPULATE_ACTIVE_INACTIVE_USER_FACILITATOR, method = RequestMethod.POST)
	@ResponseBody()
	public ExistingUsersVO fetchActiveInactiveUserListFacilitator(@RequestBody String body, HttpServletRequest request) throws JCTException {
		LOGGER.info(">>>>>> ManageUserFacilitatorController.specificUserList");
		JsonNode node = CommonUtility.getNode(body);
		int status = Integer.parseInt(node.get("statusToBe").toString().replaceAll("\"" , ""));
		//String emailId = node.get("emailId").asText();
		String customerId = node.get("customerId").toString().replaceAll("\"" , "");
		String userGrpVal = node.get("userGroup").toString().replaceAll("\"" , "");
		ExistingUsersVO existingUserVO = new ExistingUsersVO();
		String type = "active";
		if(status == 0) {
			type = "inactive";
		} 
		List<ExistingUserDTO> userList = service.populateActiveInactiveUserListForFacilitator(userGrpVal.split("!")[1], type, customerId);
		if ( null == userList ) {
			userList = new ArrayList<ExistingUserDTO>();
		}
		existingUserVO.setStatusCode(StatusConstants.STATUS_SUCCESS_WITH_RESULT);
		existingUserVO.setExistingUsers(userList);
		LOGGER.info("<<<<<< ManageUserFacilitatorController.specificUserList");
		return existingUserVO;
	}
	
	/**
	 * Method to fetch the user list to resets password
	 * @param body
	 * @param request
	 * @return return Map.
	 */
	@RequestMapping(value = ACTION_FETCH_USER_RESET_PASSWORD_FACILITATOR, method = RequestMethod.POST)
	@ResponseBody()
	public ExistingUsersVO fetchUserListForResetPassword(@RequestBody String body, HttpServletRequest request) 
			throws JCTException {
		LOGGER.info(">>>>>> ManageUserFacilitatorController.fetchUserListForResetPassword");
		JsonNode node = CommonUtility.getNode(body);
		ExistingUsersVO vo = new ExistingUsersVO();
		//String emailId = node.get("emailId").asText();
		String customerId = node.get("customerId").toString().replaceAll("\"" , "");
		String userGrpVal = node.get("userGroup").toString().replaceAll("\"" , "");
		List<ExistingUserDTO> userList = service.getResetPasswordUserListForFacilitator(userGrpVal.split("!")[1], customerId);
		if ( null == userList ) {
			userList = new ArrayList<ExistingUserDTO>();
		}
		vo.setStatusCode(StatusConstants.STATUS_SUCCESS_WITH_RESULT);
		vo.setExistingUsers(userList);
		LOGGER.info("<<<<<< ManageUserFacilitatorController.fetchUserListForResetPassword");
		return vo;
	}
	
	/**
	 * Method added to change the user status as activate and deactivate
	 * @param body
	 * @param request
	 * @return return Map.
	 */
	@RequestMapping(value = ACTION_CHANGE_BATCH_USER_STATUS_FACILITATOR, method = RequestMethod.POST)
	@ResponseBody()
	public ExistingUsersVO changeUserStatusInBatchFacilitator(@RequestBody String body, HttpServletRequest request) 
			throws JCTException {
		LOGGER.info(">>>>>> ManageUserFacilitatorController.changeUserStatusInBatchFacilitator");
		ExistingUsersVO existingUserVO = new ExistingUsersVO();
		JsonNode node = CommonUtility.getNode(body);
		String emailIdString = node.get("emailIdString").toString().replaceAll("\"" , "");
		int status = Integer.parseInt(node.get("statusToBe").toString().replaceAll("\"" , ""));
		//String facilitatorId = node.get("facilitatorId").asText();
		String customerId = node.get("customerId").toString().replaceAll("\"" , "");
		String userGrpVal = node.get("userGroup").toString().replaceAll("\"" , "");
		String updationSuccess = service.updateActivationStatusInBatchFacilitator(emailIdString , status, customerId);
		if (updationSuccess.equals("success")) {
			existingUserVO.setStatusCode(StatusConstants.USER_UPDATION_SUCCESS);
		} else {
			existingUserVO.setStatusCode(StatusConstants.USER_UPDATION_FAILURE);
		}
		String type = "inactive";
		if(status == 0) {
			type = "active";
		} 
		existingUserVO = populateActiveInactiveUserList(userGrpVal.split("!")[1], type, customerId);
		LOGGER.info("<<<<<< ManageUserFacilitatorController.changeUserStatusInBatchFacilitator");
		return existingUserVO;
	}
	/**
	 * Method to populate user list with active and inactive status
	 * @param body
	 * @param request
	 * @return ExistingUsersVO.
	 */
	private ExistingUsersVO populateActiveInactiveUserList(String userGroup, String type, String customerId) throws JCTException {
		LOGGER.info(">>>>>> ManageUserFacilitatorController.populateActiveInactiveUserList");
		ExistingUsersVO existingUserVO = new ExistingUsersVO();
		List<ExistingUserDTO> userList = service.populateActiveInactiveUserListForFacilitator(userGroup, type, customerId);
		if ( null == userList ) {
			userList = new ArrayList<ExistingUserDTO>();
		}
		existingUserVO.setStatusCode(StatusConstants.STATUS_SUCCESS_WITH_RESULT);
		existingUserVO.setExistingUsers(userList);
		LOGGER.info("<<<<<< ManageUserFacilitatorController.populateActiveInactiveUserList");
		return existingUserVO;
	}
	
	/**
	 * Method added to reset the password for selected user
	 * @param body
	 * @param request
	 * @return ManageUserVO.
	 */
	@RequestMapping(value = ACTION_RESET_PASSWORD_IN_BATCH_FACILITATOR, method = RequestMethod.POST)
	@ResponseBody()
	public ManageUserVO resetPasswordInBatchFacilitator(@RequestBody String body, HttpServletRequest request) {
		LOGGER.info(">>>>>> ManageUserFacilitatorController.resetPasswordInBatchFacilitator");
		ManageUserVO manageVO = new ManageUserVO();
		JsonNode node = CommonUtility.getNode(body);
		String emailIdString = node.get("emailIdString").toString().replaceAll("\"" , "");
		String updatedBy = node.get("updatedBy").toString().replaceAll("\"" , "");
		//String customerId = node.get("customerId").toString().replaceAll("\"" , "");
		StringTokenizer token = new StringTokenizer(emailIdString, "~");
		int totalEmailCount = token.countTokens();
		LOGGER.info("TOTAL NUMBER OF EMAILS: "+totalEmailCount);
		while (token.hasMoreTokens()) {
			try {
				manageVO = resetMassPassword((String) token.nextToken() , totalEmailCount, manageVO, updatedBy);
			} catch (JCTException e) {
				LOGGER.error("PASSWORD RESET FAILURE: "+e.getLocalizedMessage());
				manageVO.setStatusCode(StatusConstants.USER_PASSWORD_RESET_FAILURE);
			}
		}
		LOGGER.info("<<<<<< ManageUserFacilitatorController.resetPasswordInBatchFacilitator");
		return manageVO;
	}
	
	/**
	 * Method added to reset the password and send mail to selected user
	 * @param emailId
	 * @param totalEmailCount
	 * @param manageVO
	 * @param updatedBy
	 * @return manageVO
	 */
	private ManageUserVO resetMassPassword(String emailId, int totalEmailCount, ManageUserVO manageVO, String updatedBy) throws JCTException {
		LOGGER.info(">>>>>> ManageUserFacilitatorController.resetMassPassword");
		String generatedPassword = this.generatePassword();
		String encryptedPassword = this.encriptString(generatedPassword);
		String passwordResetStatus = service.resetPasswordFacilitator(emailId, encryptedPassword, updatedBy);
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
					fName = userService.getFirstName(emailId, 1);
					mailer.intimateUser(emailId, generatedPassword, fName, "reset",null);
				} else {
					//Store in the db
					String status = service.updateCronUsersForFacilitator(emailId, generatedPassword, updatedBy);
					LOGGER.info("====== STATUS OF UPDATE: "+status+" =======");
				}
			} catch (MailingException e) {
				manageVO.setStatusCode(StatusConstants.USER_REGISTERED_MAIL_SEND_ERROR_USERS);
				LOGGER.error(e.getLocalizedMessage());
			}
		} else if (manageVO.getStatusCode() != StatusConstants.USER_PASSWORD_RESET_CANNOT_PROCEED) {			
			manageVO.setStatusCode(StatusConstants.USER_PASSWORD_RESET_FAILURE);
		}
		LOGGER.info("<<<<<< ManageUserFacilitatorController.resetMassPassword");
		return manageVO;
	}
	
	/**
	 * Method generates random password
	 * @return generated password
	 */
	private String generatePassword() {
		LOGGER.info(">>>>>> ManageUserFacilitatorController.generatePassword");
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
		LOGGER.info("<<<<<< ManageUserFacilitatorController.generatePassword");
		return sb.toString();
	}
	
	/**
	 * Method encripts the string [Custom encryption]
	 * @param string
	 * @return
	 */
	private String encriptString(String string){
		LOGGER.info(">>>>>> ManageUserFacilitatorController.encriptString");
		//String ref = "1&(9%2^*37@86^54&&^%&^0CKj=mo(%np(^qs$YTuQ%^rxyd=ahg)&ei=lBfWzC@*KJM*&%ioON=PQajpSTUVR~XYDA*&^*HGE=ILBFWZ";
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
			LOGGER.info("<<<<<< ManageUserController.encriptString");
		return result.toString();
	}
	
	
	/**
	 * Method to renew user account for next 6 months
	 * @param body
	 * @param request
	 * @return ManageUserVO
	 */
	@RequestMapping(value = ACTION_RENEW_USER_FACILITATOR, method = RequestMethod.POST)
	@ResponseBody()
	public ManageUserVO renewUserByFacilitator(@RequestBody String body, HttpServletRequest request) 
			throws JCTException {
		LOGGER.info(">>>>>> ManageUserFacilitatorController.renewUserByFacilitator");
		ManageUserVO manageVO = new ManageUserVO();
		JsonNode node = CommonUtility.getNode(body);
		int userId = Integer.parseInt(node.get("userId").toString().replaceAll("\"" , ""));		
		String userEmail = node.get("userEmail").toString().replaceAll("\"" , "");
		String facilitatorEmail = node.get("facilitatorEmail").toString().replaceAll("\"" , "");
		int usersToRenew = Integer.parseInt(node.get("usersToRenew").toString().replaceAll("\"" , ""));		
		String customerId = node.get("customerId").toString().replaceAll("\"" , "");
			try {
				String updationSuccess = service.renewUserByFacilitator(userId, userEmail, facilitatorEmail, usersToRenew, customerId);
				if (updationSuccess.equals("success")) {
					manageVO.setStatusCode(StatusConstants.STATUS_SUCCESS_WITH_RESULT);
					manageVO.setStatusDesc(this.messageSource.getMessage("manage.facilitator.user.renew.success",null, null));
				} else {
					manageVO.setStatusCode(StatusConstants.STATUS_SUCCESS_WITH_NO_RESULT);
					manageVO.setStatusDesc(this.messageSource.getMessage("manage.facilitator.user.renew.failed",null, null));
				}
			} catch (JCTException e) {
				LOGGER.error("PASSWORD RESET FAILURE: "+e.getLocalizedMessage());
				manageVO.setStatusCode(StatusConstants.USER_PASSWORD_RESET_FAILURE);
			}

		LOGGER.info("<<<<<< ManageUserFacilitatorController.renewUserByFacilitator");
		return manageVO;
	}
	
	/**
	 * Method populate existing user to populate drop down list
	 * @param body
	 * @param request
	 * @return return Map.
	 */
	@RequestMapping(value = ACTION_POPULATE_USER_DROP_DOWN, method = RequestMethod.POST)
	@ResponseBody()
	public ReportsVO populateUserDropDown(@RequestBody String body, HttpServletRequest request) throws JCTException {
		LOGGER.info(">>>>>> ManageUserFacilitatorController.populateUserDropDown");
		ReportsVO reportVO = new ReportsVO();
		JsonNode node = CommonUtility.getNode(body);			
		String facilitatorEmail = node.get("facilitatorEmail").toString().replaceAll("\"" , "");
		String customerId = node.get("customerId").toString().replaceAll("\"" , "");
		List<UserDTO> userList = null;
		Map<Integer, String> userMap = new LinkedHashMap<Integer, String>();
		userList = service.getUserDropDownList(customerId, facilitatorEmail);
		int j = 0;
		for (int i=1; i<=userList.size(); i++) {
			userMap.put(userList.get(j).getUserId(), userList.get(j).getUserEmail());
			j++;
		}
		reportVO.setStatusCode(StatusConstants.STATUS_SUCCESS_WITH_RESULT);
		reportVO.setUserNameMap(userMap);
		LOGGER.info("<<<<<< ManageUserFacilitatorController.populateUserDropDown");
		return reportVO;
	}
	
	/**
	 * Method populate existing facilitator list
	 * @param body
	 * @param request
	 * @return existingUserVO containig facilitator list
	 */
	@RequestMapping(value = ACTION_POPULATE_EXISTING_FACILITATOR, method = RequestMethod.POST)
	@ResponseBody()
	public ExistingUsersVO populateExistingFacilitator(@RequestBody String body, HttpServletRequest request) throws JCTException {
		LOGGER.info(">>>>>> ManageUserFacilitatorController.populateExistingFacilitator");
		ExistingUsersVO existingUserVO = new ExistingUsersVO();
		JsonNode node = CommonUtility.getNode(body);
		String customerId = node.get("customerId").toString().replaceAll("\"" , "");
		List<ExistingUserDTO> facilitatorList = service.getFacilitatorList(customerId);
		if ( null == facilitatorList ) {
			facilitatorList = new ArrayList<ExistingUserDTO>();
		}
		existingUserVO.setStatusCode(StatusConstants.STATUS_SUCCESS_WITH_RESULT);
		existingUserVO.setExistingFacilitators(facilitatorList);
		LOGGER.info("<<<<<< ManageUserFacilitatorController.populateExistingFacilitator");
		return existingUserVO;
	}
	
	/**
	 * Method to update any existing user to facilitator
	 * @param body
	 * @param request
	 * @return ManageUserVO
	 */
	@RequestMapping(value = ACTION_UPDATE_USER_TO_FACILITATOR, method = RequestMethod.POST)
	@ResponseBody()
	public ManageUserVO updateUserToFacilitator(@RequestBody String body, HttpServletRequest request) {
		LOGGER.info(">>>>>> ManageUserFacilitatorController.updateUserToFacilitator");
		ManageUserVO manageVO = new ManageUserVO();
		JsonNode node = CommonUtility.getNode(body);
		String userVal = node.get("userVal").toString().replaceAll("\"" , "");
		String facilitatorEmail = node.get("facilitatorEmail").toString().replaceAll("\"" , "");	
		String customerId = node.get("customerId").toString().replaceAll("\"" , "");	
		String[] userSplit = userVal.split("!"); //0 is the id and 1 is the value	
		int status = 0;
		try {
			String updationSuccess = service.validateUserExpiryDate(Integer.parseInt(userSplit[0]), userSplit[1]);
			if (updationSuccess.equals("success")) {
				status = service.changeUserRoleForFacilitator(Integer.parseInt(userSplit[0]), userSplit[1], facilitatorEmail, customerId);				
				if (status == StatusConstants.STATUS_SUCCESS) {
					manageVO.setStatusCode(StatusConstants.STATUS_SUCCESS_WITH_RESULT);
					manageVO.setStatusDesc(this.messageSource.getMessage("user.updated.to.facilitator.msg",null, null));
				} else {
					manageVO.setStatusCode(StatusConstants.USER_REGISTERED_MAIL_SEND_ERROR_ADMIN);
					manageVO.setStatusDesc(this.messageSource.getMessage("manage.user.registered.no.email.error",null, null));
				}
			} else {
				manageVO.setStatusCode(999);
				manageVO.setStatusDesc(this.messageSource.getMessage("warning.account.expire.user.msg",null, null));
			}
		} catch (JCTException e) {
				manageVO = new ManageUserVO();
				manageVO.setStatusCode(StatusConstants.STATUS_FAILURE);
				manageVO.setStatusDesc(this.messageSource.getMessage("exception.dao.error",null, null));
				LOGGER.error(e.getLocalizedMessage());
			}
		LOGGER.info("<<<<<< ManageUserFacilitatorController.updateUserToFacilitator");
		return manageVO;
	}
	
	/**
	 * Method to add new facilitator
	 * @param body
	 * @param request
	 * @return ManageUserVO
	 */
	@RequestMapping(value = ACTION_ADD_FACILITATOR_TO_CHANGE_ROLE, method = RequestMethod.POST)
	@ResponseBody()
	public ConfirmationMessageVO addNewFacilitator(@RequestBody String body, HttpServletRequest request) {
		LOGGER.info(">>>>>> ManageUserFacilitatorController.addNewFacilitator");
		//ManageUserVO manageVO = new ManageUserVO();
		JsonNode node = CommonUtility.getNode(body);
		String newFacilitatorVal = node.get("newFacilitatorVal").toString().replaceAll("\"" , "");
		String facilitatorEmail = node.get("facilitatorEmail").toString().replaceAll("\"" , "");		
		String customerId = node.get("customerId").toString().replaceAll("\"" , "");	
		Boolean isValid = this.isEmailValid(newFacilitatorVal);
		int status = 0;
		String result = "";
		ConfirmationMessageVO vo = new ConfirmationMessageVO();
		try {
			if (isValid) {
				result = service.searchByEmail(newFacilitatorVal);
				if(result.equalsIgnoreCase("notexist")){
					status = service.addNewFacilitator( newFacilitatorVal, facilitatorEmail, customerId);
					if (status == StatusConstants.STATUS_SUCCESS) {
						StringBuilder msgBuilder = new StringBuilder(this.messageSource.getMessage("msg.registered1",null, null));
						msgBuilder.append(newFacilitatorVal);
						msgBuilder.append(this.messageSource.getMessage("msg.registered2",null, null));
						msgBuilder.append(this.messageSource.getMessage("msg.registered3",null, null));
						vo.setMessage(msgBuilder.toString());
					} else if (status == StatusConstants.USER_REGISTERED_MAIL_SEND_ERROR) {
						StringBuilder msgBuilder = new StringBuilder(this.messageSource.getMessage("msg.registered1",null, null));
						msgBuilder.append(newFacilitatorVal);
						msgBuilder.append(this.messageSource.getMessage("msg.registered2",null, null));
						msgBuilder.append(this.messageSource.getMessage("msg.registered4",null, null));
						vo.setMessage(msgBuilder.toString());
					} else if (status == StatusConstants.TABLE_OR_DATA_RELATED_ERROR) {
						StringBuilder msgBuilder = new StringBuilder(this.messageSource.getMessage("msg.registered5",null, null));
						vo.setMessage(msgBuilder.toString());
					} else { //FAILED
						StringBuilder msgBuilder = new StringBuilder(this.messageSource.getMessage("msg.registered5",null, null));
						vo.setMessage(msgBuilder.toString());
					}
				} else {
					StringBuilder msgBuilder = new StringBuilder(this.messageSource.getMessage("msg.already.register.user",null, null));
					vo.setMessage(msgBuilder.toString());
					vo.setStatusCode(999);
				}
			} else {
				StringBuilder msgBuilder = new StringBuilder(this.messageSource.getMessage("msg.registered6",null, null));
				vo.setMessage(msgBuilder.toString());
				vo.setStatusCode(888);
			}
		} catch (JCTException e) {
			LOGGER.error(e.getLocalizedMessage());
		}
		LOGGER.info("<<<<<< ManageUserFacilitatorController.addNewFacilitator");
		return vo;
	}
	
	/**
	 * Method Default Profile Id
	 * @param body
	 * @param request
	 * @return ManageUserVO containing default user profile id.
	 */
	@RequestMapping(value = ACTION_POPULATE_DEFAULT_PROFILE_ID, method = RequestMethod.POST)
	@ResponseBody()
	public ManageUserVO fetchDefaultProfileId(@RequestBody String body, HttpServletRequest request) {
		LOGGER.info(">>>>>> ManageUserFacilitatorController.fetchDefaultProfileId");
		ManageUserVO manageVO = new ManageUserVO();
		try {
			manageVO.setDeafultProfileId(service.fetchDefaultProfileId());
			manageVO.setStatusCode(StatusConstants.STATUS_SUCCESS_WITH_RESULT);
		} catch (JCTException e) {
			manageVO = new ManageUserVO();
			manageVO.setStatusCode(StatusConstants.STATUS_FAILURE);
			manageVO.setStatusDesc(this.messageSource.getMessage("exception.dao.error",null, null));
			LOGGER.error(e.getLocalizedMessage());
		}
		LOGGER.info("<<<<<< ManageUserFacilitatorController.fetchDefaultProfileId");
		return manageVO;
	}
	
	/**
	 * Method populate existing user group
	 * @param body
	 * @param request
	 * @return return Map.
	 */
	@RequestMapping(value = ACTION_POPULATE_USER_GROUP_FACILITATOR, method = RequestMethod.POST)
	@ResponseBody()
	public ReportsVO populateUserGroupDroDwnFacilitator(@RequestBody String body, HttpServletRequest request) throws JCTException {
		LOGGER.info(">>>>>> ManageUserFacilitatorController.populateUserGroupDroDwnFacilitator");
		ReportsVO reportVO = new ReportsVO();
		JsonNode node = CommonUtility.getNode(body);
		String customerId = node.get("customerId").toString().replaceAll("\"" , "");
		List<UserGroupDTO> userGroupList = null;
		Map<Integer, String> userGroupMap = new LinkedHashMap<Integer, String>();
		userGroupList = service.getUserGroupDropDown(customerId);
			int j = 0;
			for (int i=1; i<=userGroupList.size(); i++) {
				userGroupMap.put(userGroupList.get(j).getJctUserGroupId(), userGroupList.get(j).getUserGroupName());
				j++;
			}
			reportVO.setStatusCode(StatusConstants.STATUS_SUCCESS_WITH_RESULT);
			reportVO.setUserGroupMap(userGroupMap);
			LOGGER.info("<<<<<< ManageUserFacilitatorController.populateUserGroupDroDwnFacilitator");
		return reportVO;
	}
	
	/**
	 * Method to renew user account for next 6 months
	 * @param body
	 * @param request
	 * @return ManageUserVO
	 */
	@RequestMapping(value = ACTION_VALIDATE_BULK_RENEW_FACILITATOR, method = RequestMethod.POST)
	@ResponseBody()
	public ManageUserVO validateBulkRenewCSV(@RequestBody String body, HttpServletRequest request) 
			throws JCTException {
		LOGGER.info(">>>>>> ManageUserFacilitatorController.validateBulkRenewCSV");
			ManageUserVO manageVO = new ManageUserVO();
			ArrayList<String> validEmailList = new ArrayList<String>();
			ArrayList<String> inValidEmailList = new ArrayList<String>();
			JsonNode node = CommonUtility.getNode(body);
			String emailIdString = node.get("emailIdString").toString().replaceAll("\"" , "");	
			String customerId = node.get("customerId").toString().replaceAll("\"" , "");
			StringTokenizer emailToken = new StringTokenizer(emailIdString, "~");
			while (emailToken.hasMoreTokens()) {
				String email = (String) emailToken.nextToken().replaceAll("\"" , "");
				Boolean isValid = this.isEmailValid(email);
				if (isValid) {
					validEmailList.add(email);
				} else {
					inValidEmailList.add(email);
				}
			}
			try {
				List<String> nonExistingUserList = new ArrayList<String>();
				List<String> existingUserList = null;
				// If email ids are invalid and email ids exist
				if (validEmailList.size() > 0) {
					//existingUsernameList = userService.getUserNamesOfIndUser(validEmailList);
					existingUserList = userService.getUserNamesForFacilitator(validEmailList, customerId);
					/*if(existingUserList.size() > 0) {
						for(int i = 0; i < validEmailList.size(); i++) {
							
						}
					}*/
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
					
				} else {
					existingUserList = new ArrayList<String>();
				}
				if (inValidEmailList.size() > 0 && nonExistingUserList.size() > 0) {
					LOGGER.info("SOME EMAILS IDS ARE INVALID AND SOME EMAIL IDS DOES NOT EXIST, HENCE RETURNING WITH ERROR MESSAGE");
					manageVO.setStatusCode(StatusConstants.TOTAL_CSV_UPLD_FAILURE);
					StringBuffer msg = new StringBuffer();				
					msg.append(this.messageSource.getMessage("manage.user.renew.by.facilitator.1",null, null));
					msg.append(nonExistingUserList.size());
					msg.append(this.messageSource.getMessage("manage.user.renew.by.facilitator.2",null, null));
					msg.append(inValidEmailList.size());
					msg.append(this.messageSource.getMessage("manage.user.renew.by.facilitator.3",null, null));
					manageVO.setStatusDesc(msg.toString());
					manageVO.setInvalidEmailIds(inValidEmailList);
					manageVO.setNonExistingEmailIds(nonExistingUserList);
					return manageVO;
				} else if (inValidEmailList.size() == 0 && nonExistingUserList.size() > 0) { // If no invalid email ids but username does not exists
					LOGGER.info("SOME EMAILS ARE DOES NOT EXIST, HENCE RETURNING WITH ERROR MESSAGE");
					manageVO.setStatusCode(StatusConstants.ALREADY_REG_EMAIL_CSV_UPLD);
					StringBuffer msg = new StringBuffer();
					msg.append(this.messageSource.getMessage("manage.user.renew.by.facilitator.1",null, null));
					msg.append(nonExistingUserList.size());
					msg.append(this.messageSource.getMessage("manage.user.renew.by.facilitator.4",null, null));
					manageVO.setStatusDesc(msg.toString());
					manageVO.setNonExistingEmailIds(nonExistingUserList);
					return manageVO;
				}  else if (inValidEmailList.size() > 0 && nonExistingUserList.size() == 0) { // If invalid email ids but none user exists exists
					LOGGER.info("SOME EMAILS ARE FOUND INVALID, HENCE RETURNING WITH ERROR MESSAGE");
					manageVO.setStatusCode(StatusConstants.INVALID_EMAIL_ID_FOUND_CSV_UPLD);
					StringBuffer msg = new StringBuffer();
					msg.append(this.messageSource.getMessage("manage.user.renew.by.facilitator.1",null, null));
					msg.append(inValidEmailList.size());
					msg.append(this.messageSource.getMessage("manage.user.renew.by.facilitator.5",null, null));
					manageVO.setStatusDesc(msg.toString());
					//manageVO.setExistingEmailIds(inValidEmailList);
					manageVO.setInvalidEmailIds(inValidEmailList);
					return manageVO;
				} else if (validEmailList.size() > 0){ // Everything is correct
					LOGGER.info("EVERYTHING IS CORRECT");
					manageVO.setStatusCode(StatusConstants.ALL_CORRECT);
					StringBuffer msg = new StringBuffer();
					msg.append(this.messageSource.getMessage("manage.user.renew.by.facilitator.6",null, null));
					msg.append(validEmailList.size());
					msg.append(this.messageSource.getMessage("manage.user.renew.by.facilitator.7",null, null));
					manageVO.setStatusDesc(msg.toString());
					return manageVO;
				}						
			}   catch (JCTException e) {
				LOGGER.error("-----> "+e.getLocalizedMessage());
				manageVO.setStatusCode(StatusConstants.USER_NOT_REGISTERED);
				manageVO.setStatusDesc(this.messageSource.getMessage("manage.user.registration.failed",null, null));
			}
			LOGGER.info("<<<<<< ManageUserFacilitatorController.validateBulkRenewCSV");
			return manageVO;
	}
	
	/**
	 * Method to renew user account for next 6 months
	 * @param body
	 * @param request
	 * @return ManageUserVO
	 */
	@RequestMapping(value = ACTION_BULK_RENEW_BY_FACILITATOR, method = RequestMethod.POST)
	@ResponseBody()
	public ManageUserVO renewBulkUserByFacilitator(@RequestBody String body, HttpServletRequest request) 
			throws JCTException {
		LOGGER.info(">>>>>> ManageUserFacilitatorController.renewBulkUserByFacilitator");
		ManageUserVO manageVO = new ManageUserVO();
		ArrayList<String> validEmailList = new ArrayList<String>();
		JsonNode node = CommonUtility.getNode(body);		
		String emailIdString = node.get("emailIdString").toString().replaceAll("\"" , "");	
		int totalUsers = Integer.parseInt(node.get("totalUsers").toString().replaceAll("\"" , ""));	
		String facilitatorEmail = node.get("createdBy").toString().replaceAll("\"" , "");	
		String customerId = node.get("customerId").toString().replaceAll("\"" , "");
		StringTokenizer emailToken = new StringTokenizer(emailIdString, "~");
			while (emailToken.hasMoreTokens()) {
				String email = (String) emailToken.nextToken().replaceAll("\"" , "");
					validEmailList.add(email);		
			}	
			try {
				String updationSuccess = service.renewBulkUserByFacilitator(validEmailList, facilitatorEmail, customerId,totalUsers);
				
				if (updationSuccess.equals("success")) {
					manageVO.setStatusCode(StatusConstants.STATUS_SUCCESS_WITH_RESULT);
					manageVO.setStatusDesc(this.messageSource.getMessage("manage.facilitator.user.renew.success",null, null));
				} else {
					manageVO.setStatusCode(StatusConstants.STATUS_SUCCESS_WITH_NO_RESULT);
					manageVO.setStatusDesc(this.messageSource.getMessage("manage.facilitator.user.renew.failed",null, null));
				}
			} catch (JCTException e) {
				LOGGER.error("PASSWORD RESET FAILURE: "+e.getLocalizedMessage());
				manageVO.setStatusCode(StatusConstants.USER_PASSWORD_RESET_FAILURE);
			}

		LOGGER.info("<<<<<< ManageUserFacilitatorController.renewBulkUserByFacilitator");
		return manageVO;
	}
	
	/**
	 * Method populate existing users
	 * @param body
	 * @param request
	 * @return existingUserVO containing user list
	 */
	@RequestMapping(value = ACTION_POPULATE_RENEWED_USERS_BY_FACILITATOR, method = RequestMethod.POST)
	@ResponseBody()
	public ExistingUsersVO populateRenewedUsersByFacilitator(@RequestBody String body, HttpServletRequest request) throws JCTException {
		LOGGER.info(">>>>>> ManageUserController.populateExistingUsers");
		ExistingUsersVO existingUserVO = new ExistingUsersVO();
		ArrayList<String> validEmailList = new ArrayList<String>();
		JsonNode node = CommonUtility.getNode(body);
		//String emailAddress = node.get("emailId").toString().replaceAll("\"" , "");
		String customerId = node.get("customerId").toString().replaceAll("\"" , "");
		String emailIdString = node.get("emailIdString").asText();	
		
		StringTokenizer emailToken = new StringTokenizer(emailIdString, "~");
		while (emailToken.hasMoreTokens()) {
			String tok = (String) emailToken.nextToken().replaceAll("\"" , "");
			StringTokenizer emailIds = new StringTokenizer(tok, "#`#");
			while (emailIds.hasMoreTokens()) {
				/*String email = (String) emailToken.nextToken().replaceAll("\"" , "");
				validEmailList.add(email);	*/
				validEmailList.add((String) emailIds.nextToken().replaceAll("\"" , ""));
				emailIds.nextToken();
			}
		}
		List<ExistingUserDTO> userList = service.getRenewedUsersList(customerId, validEmailList);
		
		if ( null == userList ) {
			userList = new ArrayList<ExistingUserDTO>();
		}
		existingUserVO.setStatusCode(StatusConstants.STATUS_SUCCESS_WITH_RESULT);
		existingUserVO.setExistingUsers(userList);
		LOGGER.info("<<<<<< ManageUserController.populateExistingUsers");
		return existingUserVO;
	}
	

	/**
	 * Method to renew user account for next 6 months
	 * @param body
	 * @param request
	 * @return ManageUserVO
	 */
	@RequestMapping(value = ACTION_BULK_ASSIGN_GROUP_BY_FACILITATOR, method = RequestMethod.POST)
	@ResponseBody()
	public ManageUserVO assignGroupInBatchFacilitator(@RequestBody String body, HttpServletRequest request) 
			throws JCTException {
		LOGGER.info(">>>>>> ManageUserFacilitatorController.renewBulkUserByFacilitator");
		ManageUserVO manageVO = new ManageUserVO();
		ArrayList<String> validEmailList = new ArrayList<String>();
		JsonNode node = CommonUtility.getNode(body);		
		String emailIdString = node.get("emailIdString").toString().replaceAll("\"" , "");	
		int totalUsers = Integer.parseInt(node.get("totalUsers").toString().replaceAll("\"" , ""));	
		String facilitatorEmail = node.get("createdBy").toString().replaceAll("\"" , "");	
		String customerId = node.get("customerId").toString().replaceAll("\"" , "");
		String userGroup = node.get("userGroup").toString().replaceAll("\"" , "");		
		String newUserGroup = node.get("newUserGroup").toString().replaceAll("\"" , "");	
		String userGroupUpper = node.get("userGroup").toString().replaceAll("\"" , "").toUpperCase().trim();
		
		StringTokenizer emailToken = new StringTokenizer(emailIdString, "~");
			while (emailToken.hasMoreTokens()) {
				String email = (String) emailToken.nextToken().replaceAll("\"" , "");
					validEmailList.add(email);		
			}	
			try {				
				if(newUserGroup.equals("YES")) {
					String updationSuccess = userService.validateExistenceUserGrp(userGroupUpper, 2, 3, customerId);
					if (updationSuccess.equals("success")) {
						manageVO.setStatusCode(StatusConstants.ALREADY_EXISTS);
					} else {
						manageVO.setStatusCode(StatusConstants.DOES_NOT_EXIST);
						UserGroupVO userGroupVO = new UserGroupVO();
						userGroupVO = new UserGroupVO();
						userGroupVO.setUserGroupName(userGroup);
						userGroupVO.setUserProfileName("Default Profile");
						userGroupVO.setUserProfileId(2);
						userGroupVO.setCreatedBy(facilitatorEmail);	
						userGroupVO.setCustomerId(customerId);
						userGroupVO.setRoleId(3);
						manageVO = userService.saveUserGroup(userGroupVO);
					}
				} else {
					manageVO.setStatusCode(StatusConstants.STATUS_SUCCESS_WITH_RESULT);
				}		
				if(manageVO.getStatusCode() == 200) {	
					
					String updationStatus = service.assignNewGroupInBatch(validEmailList, facilitatorEmail, customerId, totalUsers, userGroup, newUserGroup);
					
					if (updationStatus.equals("success")) {
						manageVO.setStatusCode(StatusConstants.STATUS_SUCCESS_WITH_RESULT);
						manageVO.setStatusDesc(this.messageSource.getMessage("manage.facilitator.assign.group.success",null, null));
					} else {
						manageVO.setStatusCode(StatusConstants.STATUS_SUCCESS_WITH_NO_RESULT);
						manageVO.setStatusDesc(this.messageSource.getMessage("manage.facilitator.assign.group.failed",null, null));
					}					
				} else {
					manageVO = new ManageUserVO();
					manageVO.setStatusCode(StatusConstants.ALREADY_EXISTS);
				} 
			} catch (JCTException e) {
				LOGGER.error("PASSWORD RESET FAILURE: "+e.getLocalizedMessage());
				manageVO.setStatusCode(StatusConstants.USER_PASSWORD_RESET_FAILURE);
			}

		LOGGER.info("<<<<<< ManageUserFacilitatorController.renewBulkUserByFacilitator");
		return manageVO;
	}
}
