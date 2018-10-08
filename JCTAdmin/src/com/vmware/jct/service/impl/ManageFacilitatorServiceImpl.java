package com.vmware.jct.service.impl;

import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Date;
import java.util.Iterator;
import java.util.List;
import java.util.Random;
import java.util.StringTokenizer;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.context.MessageSource;
import org.springframework.stereotype.Service;
import org.springframework.transaction.annotation.Propagation;
import org.springframework.transaction.annotation.Transactional;

import com.vmware.jct.common.utility.CommonConstants;
import com.vmware.jct.common.utility.CommonUtility;
import com.vmware.jct.common.utility.MailNotificationHelper;
import com.vmware.jct.common.utility.StatusConstants;
import com.vmware.jct.dao.IAuthenticatorDAO;
import com.vmware.jct.dao.ICommonDao;
import com.vmware.jct.dao.IManageServiceFacilitatorDAO;
import com.vmware.jct.dao.dto.ExistingUserDTO;
import com.vmware.jct.dao.dto.FacilitatorDetailDTO;
import com.vmware.jct.dao.dto.NewUserDTO;
import com.vmware.jct.dao.dto.UserDTO;
import com.vmware.jct.dao.dto.UserGroupDTO;
import com.vmware.jct.exception.DAOException;
import com.vmware.jct.exception.JCTException;
import com.vmware.jct.exception.MailingException;
import com.vmware.jct.model.JctEmailDetails;
import com.vmware.jct.model.JctUser;
import com.vmware.jct.model.JctUserDetails;
import com.vmware.jct.model.JctUserGroup;
import com.vmware.jct.model.JctUserProfile;
import com.vmware.jct.model.JctUserRole;
import com.vmware.jct.service.IManageFacilitatorService;
import com.vmware.jct.service.vo.ManageUserVO;
import com.vmware.jct.service.vo.UserVO;
/**
 * 
 * <p><b>Class name:</b> ManageFacilitatorServiceImpl.java</p>
 * <p><b>Author:</b> InterraIT</p>
 * <p><b>Purpose:</b> This is a ManageFacilitatorServiceImpl class. This artifact is Business layer artifact.
 * ManageFacilitatorServiceImpl implement IManageFacilitatorService interface and override the following  methods.
 * -populateSubscribedUser()
 * -alreadyRegUserCount()
 * -saveUsers()
 * -saveCronUsers()
 * -generatePassword()
 * -encriptString()
 * -prepareUserObject()
 * -getUserList()
 * -populateActiveInactiveUserListForFacilitator()
 * -getResetPasswordUserListForFacilitator()
 * -updateActivationStatusInBatchFacilitator()
 * -getEmailList()
 * -resetPasswordFacilitator()
 * -updateCronUsersForFacilitator()
 * -renewUserByFacilitator()
 * -getUserDropDownList()
 * -getFacilitatorList()
 * -validateUserExpiryDate()
 * -changeUserRoleForFacilitator()
 * -addNewFacilitator()
 * -searchByEmail()
 * -decriptString()
 * </p>
 * <p><b>Description:</b> This class used to perform insert, modify ,fetch data from db and delete operation. Save, update, delete operations are performed in DataAccessObject class </p>
 * <p><b>Revision History:</b>
 * 	<li></li>
 * </p>
 */
@Service
public class ManageFacilitatorServiceImpl implements IManageFacilitatorService{

	@Autowired
	private IManageServiceFacilitatorDAO serviceDAO;
	
	@Autowired  
	private MessageSource messageSource;
	
	@Autowired
	private IAuthenticatorDAO authenticatorDAO;
	
	@Autowired
	private ICommonDao commonDaoImpl;
	
	@Autowired
	private MailNotificationHelper mailer;

	private static final Logger LOGGER = LoggerFactory.getLogger(ManageFacilitatorServiceImpl.class);
	
	/**
	 * Method populates the Total Subscribed User and total Registered User
	 * @throws JCTException
	 * @return ManageUserVO Total Subscribed User and Registered User
	 */
	@Transactional(propagation = Propagation.REQUIRED)
	public ManageUserVO populateSubscribedUser(String emailAddress, String type, String customerId) throws JCTException {
		LOGGER.info(">>>>>> ManageFacilitatorServiceImpl.populateSubscribedUser");
		StringBuilder builder = new StringBuilder("");
		ManageUserVO vo = new ManageUserVO();
		Calendar cal = Calendar.getInstance(); 
		cal.add(Calendar.MONTH, 6);	
		DateFormat dateFormat = new SimpleDateFormat("MM/dd/yyyy ");
		try {
			List<FacilitatorDetailDTO> dtoList = serviceDAO.populateSubscribedUser(emailAddress, type, customerId);
			if(null == dtoList) {
				if( type == "AD"){
					vo.setStatusCode(StatusConstants.NO_FACILITATOR_FOR_ADD);
					vo.setStatusDesc(this.messageSource.getMessage("warning.no.data.facilitator.add",null, null));
				} else {
					vo.setStatusCode(StatusConstants.NO_FACILITATOR_FOR_RENEW);
					vo.setStatusDesc(this.messageSource.getMessage("warning.no.data.facilitator.renew",null, null));
				}					
			} else if (dtoList.size() > 0) {
				Iterator<FacilitatorDetailDTO> itr = dtoList.iterator();
				while (itr.hasNext()) {
					FacilitatorDetailDTO dto = (FacilitatorDetailDTO) itr.next();
					builder.append(dto.getJctFacTotalLimit());
					builder.append("~");
					builder.append(dto.getJctFacSubscribeLimit());
					builder.append("~");
					builder.append(dto.getJctFacTotalLimit() - dto.getJctFacSubscribeLimit());
					builder.append("~");
					builder.append(this.messageSource.getMessage("label.default.user.group", null , null));
					builder.append("~");
					builder.append(dateFormat.format(cal.getTime()));
					builder.append("~");
					
				}
				vo.setSubscribedUsersList(builder.toString());
				vo.setStatusCode(StatusConstants.STATUS_SUCCESS_WITH_RESULT);
			} else {
				LOGGER.info("---- NO RECORDS FOUND -----");
				vo.setStatusCode(StatusConstants.STATUS_SUCCESS_WITH_NO_RESULT);
				vo.setStatusDesc(this.messageSource.getMessage("waring.no.subscribed.user",null, null));
				
			}
		} catch (DAOException e) {
			LOGGER.error("----"+e.getLocalizedMessage()+" ----");
			throw new JCTException(e.getLocalizedMessage());
		}
		LOGGER.info("<<<<<< ManageFacilitatorServiceImpl.populateSubscribedUser");
		return vo;
	}

	/**
	 * Method the number of already registered users users
	 * @param emailIdList
	 * @return count
	 */
	@Transactional(propagation=Propagation.REQUIRED)
	public int alreadyRegUserCount(List<String> emailIdList) throws JCTException {
		LOGGER.info(">>>>>> ManageFacilitatorServiceImpl.alreadyRegUserCount");
		int regUserCount = 0;
		try {
			regUserCount = authenticatorDAO.checkUserList(emailIdList).size();
		} catch (DAOException e) {
			LOGGER.error("UNABLE TO FIND COUNT OF USERS: "+e.getLocalizedMessage());
		}
		LOGGER.info("<<<<<< ManageFacilitatorServiceImpl.alreadyRegUserCount");
		return regUserCount;
	}
	
	
	/**
	 * Method saves new user
	 * @param validEmailIds
	 * @param userGrpName
	 * @param expiryDate
	 * @param createdBy
	 * @param customerId
	 * @return NewUserDTO
	 */
	@Transactional(propagation=Propagation.REQUIRED)
	public NewUserDTO saveUsers(List<String> validEmailIds, String userGrpName, String expiryDate, String createdBy, String customerId, String newUserGroup) throws JCTException {
		LOGGER.info(">>>>>> ManageFacilitatorServiceImpl.saveUsers");
		List<NewUserDTO> newRegistration = new ArrayList<NewUserDTO>();
		List<NewUserDTO> alreadyRegisteredUsers = new ArrayList<NewUserDTO>();
		JctUser user = null;
		NewUserDTO returnDTO = new NewUserDTO();
		Iterator<String> emailItr = validEmailIds.iterator();
		while (emailItr.hasNext()) {
			try {
				String fullString = (String) emailItr.next();//checkUser
				String emailId = fullString.split("#`#")[0];
				String fName = fullString.split("#`#")[1];
				String lName = fullString.split("#`#")[2];
				UserVO userVO = new UserVO();
				userVO.setEmail(emailId);
				//List<JctUser> userLst = authenticatorDAO.checkUser(userVO);
				List<JctUser> userLst = authenticatorDAO.checkUserByEmailAndRole(userVO, 1);
				if (userLst.size() == 0) {
					String generatedPassword = this.generatePassword();
					String encryptedPassword = this.encriptString(generatedPassword);
					user = prepareUserObject(emailId, encryptedPassword, createdBy, userGrpName, expiryDate, customerId);
					
					//Fetch the user group
					JctUserGroup userGrp = null;
					if(newUserGroup.equals("NO")) {
						if(Integer.parseInt(userGrpName.split("!")[0]) == 2) {
							userGrp = serviceDAO.fetchUserGroupId(Integer.parseInt(userGrpName.split("!")[0]), "ALL", userGrpName);	
						} else {
							userGrp = serviceDAO.fetchUserGroupId(Integer.parseInt(userGrpName.split("!")[0]), customerId, userGrpName);	
						}
					} else {
						userGrp = serviceDAO.fetchUserGroupId(0, customerId, userGrpName);	
					}
					
					JctUserProfile userProfile = userGrp.getJctUserProfile();
					
					//Create the user details object
					JctUserDetails userDetails = new JctUserDetails();
					userDetails.setJctUserDetailsProfileId(userProfile.getJctUserProfile());
					userDetails.setJctUserDetailsProfileName(userProfile.getJctUserProfileDesc());
					//set the user group
					userDetails.setJctUserDetailsGroupId(userGrp.getJctUserGroup());
					userDetails.setJctUserDetailsGroupName(userGrp.getJctUserGroupDesc());
					
					userDetails.setJctUserDetailsAdminId(0);		
					userDetails.setJctUserDetailsFacilitatorId(serviceDAO.fetchFacilitatorId(createdBy));

					userDetails.setJctUserDetailsFirstName(fName);
					userDetails.setJctUserDetailsLastName(lName);
					
					userDetails.setJctUser(user);
					user.setJctUserDetails(userDetails);
					
					String statusMsg = authenticatorDAO.createNewUser(user);					
					if (statusMsg.equals("success")) {
						NewUserDTO dto = new NewUserDTO();
						dto.setEmailId(emailId);
						dto.setFirstName(fName);
						dto.setPassword(generatedPassword);
						dto.setGroupName(userGrpName);
						newRegistration.add(dto);
					}
				} else {
					NewUserDTO dto = new NewUserDTO();
					dto.setEmailId(emailId);
					dto.setFirstName(fName);
					alreadyRegisteredUsers.add(dto);
				}
				returnDTO.setNewRegistration(newRegistration);
				returnDTO.setAlreadyRegistered(alreadyRegisteredUsers);
			} catch (DAOException ex) {
				LOGGER.error("ERROR: UNABLE TO CREATE NEW USER. REASON: "+ex.getLocalizedMessage());
				throw new JCTException("ERROR: UNABLE TO CREATE NEW USER. REASON: "+ex.getLocalizedMessage());
			}
		}
		try {
			//this method to update the number of registered user based on number of user added
			if(null != returnDTO.getNewRegistration()) {
				serviceDAO.updateFacilitatorDetails(returnDTO.getNewRegistration().size(), customerId, "AD");
			}	
		} catch (DAOException e) {
			LOGGER.error("ERROR: UNABLE TO UPDATE FACILITATOR DETAIL TABLE. REASON: "+e.getLocalizedMessage());
		}
		LOGGER.info("<<<<<< ManageFacilitatorServiceImpl.saveUsers");
		return returnDTO;
	}
	
	/**
	 * Method saves list of email ids of users so that the cron job 
	 * can pick it up and can send the registration mail to the user later. 
	 * @param registeredEmailIds
	 * @param createdBy
	 * @return status - String
	 */
	@Transactional(propagation=Propagation.REQUIRED)
	public String saveCronUsers(List<NewUserDTO> registeredEmailIds,
			String createdBy) throws JCTException {
		String status = "failure";
		JctEmailDetails emailDtlsObj = null;
		Iterator<NewUserDTO> itr = registeredEmailIds.iterator();
		while (itr.hasNext()) {
			NewUserDTO dto = (NewUserDTO) itr.next();
			emailDtlsObj = new JctEmailDetails();
			try {
				emailDtlsObj.setJctRegistrationId(commonDaoImpl.generateKey("jct_email_dtls_id"));
				emailDtlsObj.setJctAccountCreatedBy(createdBy);
				emailDtlsObj.setJctAccountCreatedDate(new Date());
				emailDtlsObj.setJctMailSentDate(new Date()); // This is for hibernate fix
				emailDtlsObj.setJctMailDispatched(StatusConstants.USER_EMAIL_NOT_SENT);
				emailDtlsObj.setJctRegisteredMail(dto.getEmailId());
				emailDtlsObj.setJctRegisteredPassword(dto.getPassword());
				emailDtlsObj.setJctUserGroupDesc(dto.getGroupName());
				status = serviceDAO.cronPersister(emailDtlsObj);
			} catch (DAOException e) {
				status = "failure";
				throw new JCTException(e.getLocalizedMessage());
			}
		}
		return status;
	}
	
	/**
	 * Method generates random password
	 * @return generated password
	 */
	private String generatePassword() {
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
		
		return sb.toString();
	}
	
	/**
	 * Method encripts the string [Custom encryption]
	 * @param string
	 * @return
	 */
	private String encriptString(String string){
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
		return result.toString();
	}
	
	/**
	 * Method to set the data to user object
	 * @param emailId
	 * @param encryptedPassword
	 * @param createdBy
	 * @param userGrpName
	 * @param expiryDate
	 * @param customerId
	 * @return user entity
	 */
	private JctUser prepareUserObject(String emailId, String encryptedPassword, String createdBy, String userGrpName, String expiryDate, String customerId) throws DAOException {
		JctUser user = new JctUser();
		user.setJctUserId(commonDaoImpl.generateKey("jct_user_id"));
		user.setJctUserEmail(emailId);
		user.setJctUserName(emailId);
		user.setJctPassword(encryptedPassword);
		user.setJctCreatedBy(createdBy);
		user.setJctLastmodifiedBy(createdBy);
		user.setJctCreatedTs(new Date());
		user.setLastmodifiedTs(new Date());
		user.setJctActiveYn(CommonConstants.CREATED);
		JctUserRole role = new JctUserRole();
		role.setJctRoleId(new Integer(1));
		user.setJctUserRole(role);
		//set the expiry date		
		Date expryDate = CommonUtility.convertToDateObj(expiryDate);
		user.setJctAccountExpirationDate(expryDate);
		user.setJctUserCustomerId(customerId);	
		
		//Fetch the user profile
		//JctUserGroup userGrp = serviceDAO.fetchUserGroup(userGrpId);
		//JctUserProfile userProfile = userGrp.getJctUserProfile();
		//user.setJctProfileId(userProfile.getJctUserProfile());
		//user.setJctProfileName(userProfile.getJctUserProfileDesc());
		
		//set the user group
		//user.setJctUserGroupId(userGrp.getJctUserGroup());
		//user.setJctUserGroupName(userGrp.getJctUserGroupDesc());
		return user;
	}


	/**
	 * Method returns existing active and deactivate user list 
	 * @param customerId
	 * @return userList
	 */
	@Transactional(propagation = Propagation.REQUIRED)
	public List<ExistingUserDTO> getUserList(String customerId)
			throws JCTException {
		LOGGER.info(">>>>>> ManageFacilitatorServiceImpl.getUserList");
		List<ExistingUserDTO> userList = null;
		try {
			userList= serviceDAO.populateUserList(customerId);
		} catch(DAOException daoException ) {
			LOGGER.error(daoException.getLocalizedMessage());
			throw new JCTException(daoException.getMessage());
		}
		LOGGER.info("<<<<<< ManageFacilitatorServiceImpl.getUserList");
		return userList;
	}

	/**
	 * Method to fetch the user list with active and inactive status
	 * @param activeInactive
	 * @param customerId
	 * @return ExistingUsersVO.
	 */
	@Transactional(propagation = Propagation.REQUIRED)
	public List<ExistingUserDTO> populateActiveInactiveUserListForFacilitator (
			String userGroup, String activeInactive, String customerId) throws JCTException {
		LOGGER.info(">>>>>> ManageFacilitatorServiceImpl.populateActiveInactiveUserListForFacilitator");
		List<ExistingUserDTO> userList = null;
		try {
			userList= serviceDAO.populateActiveInactiveUserListForFacilitator(userGroup, activeInactive, customerId);
		} catch(DAOException daoException ) {
			LOGGER.error(daoException.getLocalizedMessage());
			throw new JCTException(daoException.getMessage());
		}
		LOGGER.info("<<<<<< ManageFacilitatorServiceImpl.populateActiveInactiveUserListForFacilitator");
		return userList;
	}

	/**
	 * Method to fetch the user list to resets password
	 * @param customerId
	 * @return userList
	 */
	@Transactional(propagation = Propagation.REQUIRED)
	public List<ExistingUserDTO> getResetPasswordUserListForFacilitator (
			String userGroup, String customerId) throws JCTException {
		LOGGER.info(">>>>>> ManageFacilitatorServiceImpl.getResetPasswordUserListForFacilitator");
		List<ExistingUserDTO> userList = null;
		try {
			userList= serviceDAO.populateUserListResetPasswordForFacilitator(userGroup, customerId);
		} catch(DAOException daoException ) {
			LOGGER.error(daoException.getLocalizedMessage());
			throw new JCTException(daoException.getMessage());
		}
		LOGGER.info("<<<<<< ManageFacilitatorServiceImpl.getResetPasswordUserListForFacilitator");
		return userList;
	}

	/**
	 * Method updates existing user's activation status in batch
	 * @param customerId
	 * @param softDelete
	 * @param emailIdString
	 * @return status
	 */
	@Transactional(propagation = Propagation.REQUIRED)
	public String updateActivationStatusInBatchFacilitator (
			String emailIdString, int softDelete, String customerId) throws JCTException {
		String statusMessage = "failure";
		try {
			List<String> emailList = this.getEmailList(emailIdString);
			statusMessage = serviceDAO.updateUserStatusInBatchFacilitator(emailList, softDelete, customerId);
		} catch (DAOException e) {
			LOGGER.error("UNABLE TO FIND / CHANGE STAUS USER: "+e.getLocalizedMessage());
			statusMessage = "failure";
		}
		return statusMessage;
	}
	
	/**
	 * USER WILL NOT BE DELETED. ONLY ACTIVATED OR INACTIVATED
	 */
	private List<String> getEmailList(String batchEmail) {
		List<String> emailList = new ArrayList<String>();
		StringTokenizer token = new StringTokenizer(batchEmail,"~");
		while(token.hasMoreTokens()) {
			emailList.add((String) token.nextToken());
		}
		return emailList;
	}

	/**
	 * Method updates the mail sent status of the existing email id
	 * @param registeredEmailId
	 * @param generatedPassword
	 * @param updatedBy
	 * @return status - String
	 */
	@Transactional(propagation = Propagation.REQUIRED)
	public String resetPasswordFacilitator(String emailId, String encryptedPassword, String updatedBy)
			throws JCTException {
		LOGGER.info(">>>>>> ManageFacilitatorServiceImpl.resetPasswordFacilitator");
		String status = "failure";
		UserVO userVO = new UserVO();
		userVO.setEmail(emailId);
		try {
			List<JctUser> userList = serviceDAO.checkUserToResetPassword(userVO, updatedBy);
			JctUser userObj = (JctUser) userList.get(0);			
			if (userObj.getJctActiveYn() == CommonConstants.CREATED) {
				userObj.setJctPassword(encryptedPassword);
				status = serviceDAO.updateUserToResetPasswordFacilitator(userObj);
			} else {
				status = "changed";
			}
		} catch (DAOException e) {
			LOGGER.error("UNABLE TO STORE USER PASSWORD: "+e.getLocalizedMessage());
			status = "failure";
		}
		LOGGER.info("<<<<<< ManageFacilitatorServiceImpl.resetPasswordFacilitator");
		return status;
	}

	/**
	 * Method updates the mail sent status of the existing email id
	 * @param registeredEmailId
	 * @param generatedPassword
	 * @param updatedBy
	 * @return status - String
	 */
	@Transactional(propagation = Propagation.REQUIRED)
	public String updateCronUsersForFacilitator(String registeredEmailId,
			String generatedPassword, String updatedBy) throws JCTException {
		LOGGER.info(">>>>>> ManageFacilitatorServiceImpl.updateCronUsersForFacilitator");
		String status = "failure";
		try {
			JctEmailDetails object = null;
			//if existing in JctEmailDetails then update
			//else insert
			List<JctEmailDetails> list = serviceDAO.getEmailObjectFacilitator(registeredEmailId, updatedBy);
			if (list.size() > 0 ) {
				object = (JctEmailDetails) list.get(0);
				object.setJctRegisteredPassword(generatedPassword);
				object.setJctAccountCreatedBy(updatedBy);
				object.setJctAccountCreatedDate(new Date());
				object.setJctMailDispatched(StatusConstants.USER_EMAIL_NOT_SENT);
				//status = serviceDAO.cronUpdater(object);
			} else {
				//status = "failure";
				object = new JctEmailDetails();
				object.setJctRegistrationId(commonDaoImpl.generateKey("jct_email_dtls_id"));
				object.setJctAccountCreatedBy(updatedBy);
				object.setJctAccountCreatedDate(new Date());
				object.setJctMailDispatched(StatusConstants.USER_EMAIL_NOT_SENT);
				object.setJctRegisteredMail(registeredEmailId);
				object.setJctRegisteredPassword(generatedPassword);
			}
			status = serviceDAO.cronUpdaterFacilitator(object);
		} catch (DAOException ex) {
			status = "failure";
			LOGGER.error("UNABLE TO FETCH OR MODIFY THE EMAIL OBJECT. REASON: "+ex.getLocalizedMessage());
			throw new JCTException(ex.getLocalizedMessage());
		}
		LOGGER.info("<<<<<< ManageFacilitatorServiceImpl.updateCronUsersForFacilitator");
		return status;
	}

	/**
	 * Method to renew particular user account for next 6 months
	 * @param userId
	 * @param userEmail
	 * @param facilitatorEmail
	 * @param usersToRenew
	 * @param customerId
	 * @return status in string
	 */
	@Transactional(propagation = Propagation.REQUIRED)
	public String renewUserByFacilitator(int userId, String userEmail, String facilitatorEmail, int usersToRenew, String customerId)
			throws JCTException {
		LOGGER.info(">>>>>> ManageFacilitatorServiceImpl.renewUserByFacilitator");
		String statusMessage = "failure";
		int status = StatusConstants.STATUS_FAILED;
		Date expiryDate;
		try {
			expiryDate = serviceDAO.fetchUserExpiryDate(userId);	
			Calendar calendar = Calendar.getInstance();			 
			java.util.Date now = calendar.getTime();
			java.sql.Timestamp currentTimestamp = new java.sql.Timestamp(now.getTime());
			Calendar cal = Calendar.getInstance(); 
			if(expiryDate.getTime() < currentTimestamp.getTime()){
				expiryDate = currentTimestamp;
				cal.setTime(expiryDate);
				cal.add(Calendar.MONTH, 6);
			} 
			else{
				cal.setTime(expiryDate);
				cal.add(Calendar.MONTH, 6);
			}
			statusMessage = serviceDAO.renewUserByFacilitator(userId, userEmail, facilitatorEmail, cal.getTime());
			String result = serviceDAO.updateFacilitatorDetails(1, customerId, "RW");
			
			if (statusMessage.equalsIgnoreCase("success") && result.equalsIgnoreCase("Success")) {
				try {
					mailer.sendMailToRenewedUser(userEmail, cal.getTime(), null);						
				} catch (MailingException e) {
					LOGGER.error("FACILITATOR REGISTERED....COULD NOT SEND EMAIL...  REASON: "+e.getLocalizedMessage());
					status = StatusConstants.USER_REGISTERED_MAIL_SEND_ERROR;
				}
			} else {
				LOGGER.error("UNABLE TO RENEW USER:");
				status = StatusConstants.USER_REGISTERED_MAIL_SEND_ERROR;
			}
			
		} catch (DAOException e) {
			LOGGER.error("UNABLE TO RENEW USER: "+e.getLocalizedMessage());
			statusMessage = "failure";
		}
		LOGGER.info("<<<<<< ManageFacilitatorServiceImpl.renewUserByFacilitator");
		return statusMessage;
	}

	/**
	 * Method returns list of user 
	 * @param customerId
	 * @return list of UserDTO
	 */
	@Transactional(propagation = Propagation.REQUIRED)
	public List<UserDTO> getUserDropDownList(String customerId, String facilitatorEmail) throws JCTException {
		LOGGER.info(">>>>>> ManageFacilitatorServiceImpl.getUserDropDownList");
		List<UserDTO> userList=null;
		try{
			userList= serviceDAO.populateUser(customerId, facilitatorEmail);
		}catch(DAOException daoException ){
			LOGGER.error(daoException.getLocalizedMessage());
			throw new JCTException(daoException.getMessage());
		}
		LOGGER.info("<<<<<< ManageFacilitatorServiceImpl.getUserDropDownList");
		return userList;
	}

	/**
	 * Method populates existing facilitator list
	 * @throws JCTException
	 * @return ExistingUserDTO
	 */
	@Transactional(propagation = Propagation.REQUIRED)
	public List<ExistingUserDTO> getFacilitatorList(String customerId) throws JCTException {
		LOGGER.info(">>>>>> ManageFacilitatorServiceImpl.getFacilitatorList");
		List<ExistingUserDTO> facilitatorList = null;
		try {
			facilitatorList= serviceDAO.populateFacilitatorList(customerId);
		} catch(DAOException daoException ) {
			LOGGER.error(daoException.getLocalizedMessage());
			throw new JCTException(daoException.getMessage());
		}
		LOGGER.info("<<<<<< ManageFacilitatorServiceImpl.getFacilitatorList");
		return facilitatorList;
	}

	/**
	 * Method to validate account expire date for user
	 * @param userId
	 * @param userName
	 * @return status in String 
	 * @throws JCTException
	 */
	@Transactional(propagation = Propagation.REQUIRED)
	public String validateUserExpiryDate(int userId, String userName)
			throws JCTException {
		LOGGER.info(">>>>>> ManageFacilitatorServiceImpl.validateUserExpiryDate");
		String statusMessage = "";
		Date expiryDate;
		try {
			expiryDate = serviceDAO.fetchUserExpiryDate(userId);	
			Calendar calendar = Calendar.getInstance();			 
			java.util.Date now = calendar.getTime();
			java.sql.Timestamp currentTimestamp = new java.sql.Timestamp(now.getTime());

			if(expiryDate.getTime() < currentTimestamp.getTime()){
				expiryDate = currentTimestamp;
				statusMessage = "failure";
			} else {
				statusMessage = "success";
			}							
		} catch (DAOException e) {
			LOGGER.error("UNABLE TO RENEW USER: "+e.getLocalizedMessage());
			statusMessage = "failure";
		}
		LOGGER.info("<<<<<< ManageFacilitatorServiceImpl.validateUserExpiryDate");
		return statusMessage;
	}


	/**
	 * Method to update the existing user to facilitator
	 * @param userId
	 * @param userEmail
	 * @param customerId
	 * @param facilitatorEmail
	 */
	@Transactional(propagation = Propagation.REQUIRED)
	public int changeUserRoleForFacilitator(int userId,
			String userEmail, String facilitatorEmail , String customerId) throws JCTException {
		LOGGER.info(">>>>>> ManageFacilitatorServiceImpl.changeUserRoleForFacilitator");		
		int status = StatusConstants.STATUS_FAILED;
		try {
			//fetch the user object
			JctUser user = serviceDAO.fetchUserData(userId);
			//fetch the facilitator object
			JctUser facilitator = serviceDAO.fetchFacilitatorData(facilitatorEmail, customerId);
			facilitator.setJctUserSoftDelete(1);
			facilitator.getJctUserDetails().setJctUserDetailsSoftDelete(1);
			facilitator.setJctActiveYn(CommonConstants.CREATED);
			try {
				// Create User Object
				JctUser userNew = new JctUser();
				userNew.setJctUserId(commonDaoImpl.generateKey("jct_user_id"));
				userNew.setJctUserName(user.getJctUserEmail());
				userNew.setJctActiveYn(CommonConstants.CREATED);
				userNew.setJctCreatedBy(user.getJctCreatedBy());
				userNew.setJctCreatedTs(user.getJctCreatedTs());
				userNew.setJctLastmodifiedBy(user.getJctLastmodifiedBy());
				userNew.setJctPassword(user.getJctPassword());
				userNew.setJctUserEmail(user.getJctUserEmail());
				
				JctUserRole role = new JctUserRole();
				role.setJctRoleId(CommonConstants.FACILITATOR_USER);
				userNew.setJctUserRole(role);
				
				//userNew.setJctUserRole(CommonConstants.FACILITATOR_USER);
				userNew.setJctVersion(user.getJctVersion());
				userNew.setLastmodifiedTs(new Date());
				userNew.setJctUserCustomerId(customerId);
				userNew.setJctAccountExpirationDate(new Date());
				
				// Create User Details Object
				JctUserDetails userDetails = new JctUserDetails();
				userDetails.setJctUserDetailsFirstName(user.getJctUserDetails().getJctUserDetailsFirstName());
				userDetails.setJctUserDetailsLastName(user.getJctUserDetails().getJctUserDetailsLastName());
				userDetails.setJctUserDetailsGender(user.getJctUserDetails().getJctUserDetailsGender());
				userDetails.setJctUserDetailsRegion(user.getJctUserDetails().getJctUserDetailsRegion());
				userDetails.setJctUserDetailsFunctionGroup(user.getJctUserDetails().getJctUserDetailsFunctionGroup());
				userDetails.setJctUserDetailsLevels(user.getJctUserDetails().getJctUserDetailsLevels());
				userDetails.setJctUserDetailsSupervisePeople(user.getJctUserDetails().getJctUserDetailsSupervisePeople());
				userDetails.setJctUserDetailsTenure(user.getJctUserDetails().getJctUserDetailsTenure());
				userDetails.setJctUserDetailsGroupId(user.getJctUserDetails().getJctUserDetailsGroupId());
				userDetails.setJctUserDetailsGroupName(user.getJctUserDetails().getJctUserDetailsGroupName());
				userDetails.setJctUserDetailsProfileId(user.getJctUserDetails().getJctUserDetailsProfileId());
				userDetails.setJctUserDetailsProfileName(user.getJctUserDetails().getJctUserDetailsProfileName());
				userDetails.setJctUserDetailsAdminId(user.getJctUserDetails().getJctUserDetailsAdminId());
				userDetails.setJctUserDetailsFacilitatorId(user.getJctUserDetails().getJctUserDetailsFacilitatorId());
				userDetails.setJctUser(userNew);
				userNew.setJctUserDetails(userDetails);
				String save = serviceDAO.saveFacilitator(userNew);
				
				String result = serviceDAO.updateUser(facilitator);
								
				if (save.equals("success") && result.equals("success")) {
					try {						
						mailer.intimateUserToFacilitator(userNew.getJctUserEmail(), facilitatorEmail);
					} catch (MailingException e) {
						LOGGER.error("FACILITATOR REGISTERED....COULD NOT SEND EMAIL...  REASON: "+e.getLocalizedMessage());
						status = StatusConstants.USER_REGISTERED_MAIL_SEND_ERROR;
					}
				} else {
					LOGGER.error("SAVING TO PAYMENT HEADER AND FACILITATOR DETAILS: "+save+". SAVE TO PAYMENT DETAILS: "+result);
					status = StatusConstants.TABLE_OR_DATA_RELATED_ERROR;
				}
				status = StatusConstants.STATUS_SUCCESS;
				
			} catch (Exception e) {
				LOGGER.error("----"+e.getLocalizedMessage()+" ----");
			}
			
		} catch (DAOException e) {
			LOGGER.error("----"+e.getLocalizedMessage()+" ----");
			throw new JCTException(e.getLocalizedMessage());
		}
		LOGGER.info("<<<<<< ManageFacilitatorServiceImpl.changeUserRoleForFacilitator");
		return status;
	}

	/**
	 * Method to add new facilitator to change role
	 * @param newFacilitatorVal
	 * @param facilitatorEmail
	 * @param customerId
	 * @return ManageUserVO
	 * @throws JCTException
	 */
	@Transactional(propagation = Propagation.REQUIRED)
	public int addNewFacilitator(String newFacilitatorVal,
			String facilitatorEmail, String customerId) throws JCTException {
		LOGGER.info(">>>>>> ManageFacilitatorServiceImpl.addNewFacilitator");
		int status = StatusConstants.STATUS_FAILED;
		String fName = "";
		try {
			
			JctUser facilitator = serviceDAO.fetchFacilitatorData(facilitatorEmail, customerId);
			facilitator.setJctUserSoftDelete(1);
			facilitator.getJctUserDetails().setJctUserDetailsSoftDelete(1);
			facilitator.setJctActiveYn(CommonConstants.CREATED);
			
			try {
				// Create User Object
				JctUser userNew = new JctUser();
				userNew.setJctUserId(commonDaoImpl.generateKey("jct_user_id"));
				userNew.setJctUserName(newFacilitatorVal);
				userNew.setJctActiveYn(CommonConstants.CREATED);
				userNew.setJctCreatedBy(facilitator.getJctCreatedBy());
				userNew.setJctCreatedTs(facilitator.getJctCreatedTs());
				userNew.setJctLastmodifiedBy(facilitator.getJctLastmodifiedBy());
				userNew.setJctPassword(facilitator.getJctPassword());
				userNew.setJctUserEmail(newFacilitatorVal);
				
				JctUserRole role = new JctUserRole();
				role.setJctRoleId(CommonConstants.FACILITATOR_USER);
				userNew.setJctUserRole(role);
				
				//userNew.setJctUserRole(CommonConstants.FACILITATOR_USER);
				userNew.setJctVersion(facilitator.getJctVersion());
				userNew.setLastmodifiedTs(new Date());
				userNew.setJctUserCustomerId(customerId);
				userNew.setJctAccountExpirationDate(new Date());
				
				// Create User Details Object
				JctUserDetails userDetails = new JctUserDetails();
				userDetails.setJctUserDetailsFirstName(facilitator.getJctUserDetails().getJctUserDetailsFirstName());
				userDetails.setJctUserDetailsLastName(facilitator.getJctUserDetails().getJctUserDetailsLastName());
				//userDetails.setJctUserDetailsGender(facilitator.getJctUserDetails().getJctUserDetailsGender());
				userDetails.setJctUserDetailsRegion(facilitator.getJctUserDetails().getJctUserDetailsRegion());
				userDetails.setJctUserDetailsFunctionGroup(facilitator.getJctUserDetails().getJctUserDetailsFunctionGroup());
				userDetails.setJctUserDetailsLevels(facilitator.getJctUserDetails().getJctUserDetailsLevels());
				userDetails.setJctUserDetailsSupervisePeople(facilitator.getJctUserDetails().getJctUserDetailsSupervisePeople());
				userDetails.setJctUserDetailsTenure(facilitator.getJctUserDetails().getJctUserDetailsTenure());
				userDetails.setJctUserDetailsGroupId(facilitator.getJctUserDetails().getJctUserDetailsGroupId());
				userDetails.setJctUserDetailsGroupName(facilitator.getJctUserDetails().getJctUserDetailsGroupName());
				userDetails.setJctUserDetailsProfileId(facilitator.getJctUserDetails().getJctUserDetailsProfileId());
				userDetails.setJctUserDetailsProfileName(facilitator.getJctUserDetails().getJctUserDetailsProfileName());
				userDetails.setJctUserDetailsAdminId(facilitator.getJctUserDetails().getJctUserDetailsAdminId());
				userDetails.setJctUserDetailsFacilitatorId(facilitator.getJctUserDetails().getJctUserDetailsFacilitatorId());
				userDetails.setJctUser(userNew);
				userNew.setJctUserDetails(userDetails);
				String save = serviceDAO.saveFacilitator(userNew);
				
				String result = serviceDAO.updateUser(facilitator);
								
				if (save.equals("success") && result.equals("success")) {
					try {
						mailer.intimateNewFacilitator(newFacilitatorVal, 
								decriptString(facilitator.getJctPassword()), userDetails.getJctUserDetailsFirstName(), 0,
								"N",facilitator.getJctUserCustomerId(),null);						
					} catch (MailingException e) {
						LOGGER.error("FACILITATOR REGISTERED....COULD NOT SEND EMAIL...  REASON: "+e.getLocalizedMessage());
						status = StatusConstants.USER_REGISTERED_MAIL_SEND_ERROR;
					}
				} else {
					LOGGER.error("SAVING TO PAYMENT HEADER AND FACILITATOR DETAILS: "+save+". SAVE TO PAYMENT DETAILS: "+result);
					status = StatusConstants.TABLE_OR_DATA_RELATED_ERROR;
				}
				status = StatusConstants.STATUS_SUCCESS;
				
			} catch (Exception e) {
				LOGGER.error("----"+e.getLocalizedMessage()+" ----");
			}
		} catch (Exception ex) {
			LOGGER.error(ex.getLocalizedMessage());
			throw new JCTException();
		}
		LOGGER.info("<<<<<< ManageFacilitatorServiceImpl.addNewFacilitator");
		return status;
	}


	/**
	 * Method to search the email id is exist in database
	 * @param newFacilitatorVal
	 * @return String
	 * @throws JCTException
	 */
	@Transactional(propagation = Propagation.REQUIRED)
	public String searchByEmail(String newFacilitatorVal) throws JCTException {
		LOGGER.info(">>>>>> ManageFacilitatorServiceImpl.searchByEmail");
		String status = "";
		try {
			status = serviceDAO.searchByEmail(newFacilitatorVal);
		} catch (DAOException e) {
			LOGGER.error("UNABLE TO FIND COUNT OF USERS: "+e.getLocalizedMessage());
		}	
		LOGGER.info("<<<<<< ManageFacilitatorServiceImpl.searchByEmail");
		return status;
	}

	/**
	 * Method decript the password string 
	 * @param string
	 * @return string
	 */
	private String decriptString(String string){
       // String ref = "1&(9%2^*37@86^54&&^%&^0CKj=mo(%np(^qs$YTuQ%^rxyd=ahg)&ei=lBfWzC@*KJM*&%ioON=PQajpSTUVR~XYDA*&^*HGE=ILBFWZ";
		//String ref = "P!QR#S$T%U&V'W(X)Y*Z+[,-].^/_0`1a2b3c4d5e6f7g8h9i:j;k<l=m>n?o@pAqBrCsDtEuFvGwHxIyJzK{L|M}N~O";
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

	@Transactional(propagation = Propagation.REQUIRED)
	public int fetchDefaultProfileId() throws JCTException {
		LOGGER.info(">>>>>> ManageFacilitatorServiceImpl.fetchDefaultProfileId");
		int userProfileId = 0;
		try {
			userProfileId = serviceDAO.fetchDefaultProfileId();
		} catch (DAOException e) {
			LOGGER.error("UNABLE TO FIND COUNT OF USERS: "+e.getLocalizedMessage());
		}
		LOGGER.info("<<<<<< ManageFacilitatorServiceImpl.fetchDefaultProfileId");
		return userProfileId;
	}

	@Transactional(propagation = Propagation.REQUIRED)
	public List<UserGroupDTO> getUserGroupDropDown(String customerId)
			throws JCTException {
		LOGGER.info(">>>>>> ManageFacilitatorServiceImpl.getUserGroupDropDown");
		List<UserGroupDTO> userGroupList=null;
		try{
			userGroupList= serviceDAO.getUserGroupDropDown(customerId);
		}catch(DAOException daoException ){
			LOGGER.error(daoException.getLocalizedMessage());
			throw new JCTException(daoException.getMessage());
		}
		LOGGER.info("<<<<<< ManageFacilitatorServiceImpl.getUserGroupDropDown");
		return userGroupList;
	}

	@Transactional(propagation = Propagation.REQUIRED)
	public String renewBulkUserByFacilitator(List<String> validEmailIds,
			String facilitatorEmail, String customerId, int totalUsers) throws JCTException {
		LOGGER.info(">>>>>> ManageFacilitatorServiceImpl.renewBulkUserByFacilitator");
		String statusMessage = "failure";
		int status = StatusConstants.STATUS_FAILED;
		Date expiryDate;		
		Iterator<String> emailItr = validEmailIds.iterator();
		while (emailItr.hasNext()) {
			String emailId = (String) emailItr.next();//checkUser
		//	List<JctUser> userLst = authenticatorDAO.checkUser(userVO);
			List<JctUser> userLst;
			try {
				userLst = serviceDAO.fetchUserObj(emailId, customerId);
				if (userLst.size() > 0) {
					expiryDate = userLst.get(0).getJctAccountExpirationDate();	
					Calendar calendar = Calendar.getInstance();			 
					java.util.Date now = calendar.getTime();
					java.sql.Timestamp currentTimestamp = new java.sql.Timestamp(now.getTime());
					Calendar cal = Calendar.getInstance(); 
					if(expiryDate.getTime() < currentTimestamp.getTime()){
						expiryDate = currentTimestamp;
						cal.setTime(expiryDate);
						cal.add(Calendar.MONTH, 6);
					} 
					else{
						cal.setTime(expiryDate);
						cal.add(Calendar.MONTH, 6);
					}
					statusMessage = serviceDAO.renewUserByFacilitator(userLst.get(0).getJctUserId(), userLst.get(0).getJctUserEmail(), facilitatorEmail, cal.getTime());
					String result = serviceDAO.updateFacilitatorDetails(1, customerId, "RW");
					
					if (statusMessage.equalsIgnoreCase("success") && result.equalsIgnoreCase("Success")) {
						try {
							mailer.sendMailToRenewedUser(userLst.get(0).getJctUserEmail(), cal.getTime(), null);						
						} catch (MailingException e) {
							LOGGER.error("FACILITATOR REGISTERED....COULD NOT SEND EMAIL...  REASON: "+e.getLocalizedMessage());
							status = StatusConstants.USER_REGISTERED_MAIL_SEND_ERROR;
						}
					} else {
						LOGGER.error("UNABLE TO RENEW USER:");
						status = StatusConstants.USER_REGISTERED_MAIL_SEND_ERROR;
					}																							
				}
			} catch (DAOException e) {
				LOGGER.error("UNABLE TO RENEW USER: "+e.getLocalizedMessage());
				statusMessage = "failure";
			}
		}
				
		LOGGER.info("<<<<<< ManageFacilitatorServiceImpl.renewBulkUserByFacilitator");
		return statusMessage;
	}

	@Transactional(propagation = Propagation.REQUIRED)
	public List<ExistingUserDTO> getRenewedUsersList(String customerId,
			List<String> validEmailIds) throws JCTException {
		LOGGER.info(">>>>>> ManageFacilitatorServiceImpl.getRenewedUsersList");
		List<ExistingUserDTO> userList = null;
		try {
			userList= serviceDAO.populateRenewedUserList(customerId, validEmailIds);
		} catch(DAOException daoException ) {
			LOGGER.error(daoException.getLocalizedMessage());
			throw new JCTException(daoException.getMessage());
		}
		LOGGER.info("<<<<<< ManageFacilitatorServiceImpl.getRenewedUsersList");
		return userList;
	}

	@Transactional(propagation = Propagation.REQUIRED)
	public String assignNewGroupInBatch(ArrayList<String> validEmailIds,
			String facilitatorEmail, String customerId, int totalUsers,
			String userGroup, String newUserGroup) {
		LOGGER.info(">>>>>> ManageFacilitatorServiceImpl.assignNewGroupInBatch");
		String statusMessage = "failure";
		Iterator<String> emailItr = validEmailIds.iterator();
		while (emailItr.hasNext()) {
			String emailId = (String) emailItr.next();//checkUser
		//	List<JctUser> userLst = authenticatorDAO.checkUser(userVO);
			List<JctUser> userLst;
			try {
				userLst = serviceDAO.fetchUserObj(emailId, customerId);
				if (userLst.size() > 0) {
					
					//Fetch the user group
					JctUserGroup userGrp = null;
					if(newUserGroup.equals("NO")) {
						//userGrp = serviceDAO.fetchUserGroupId(Integer.parseInt(userGroup.split("!")[0]), customerId, userGroup);	
						if(Integer.parseInt(userGroup.split("!")[0]) == 2) {
							userGrp = serviceDAO.fetchUserGroupId(Integer.parseInt(userGroup.split("!")[0]), "ALL", userGroup);	
						} else {
							userGrp = serviceDAO.fetchUserGroupId(Integer.parseInt(userGroup.split("!")[0]), customerId, userGroup);	
						}
					} else {
						userGrp = serviceDAO.fetchUserGroupId(0, customerId, userGroup);	
					}
					
					statusMessage = serviceDAO.assignNewUserGroupInBatch(userLst.get(0).getJctUserId(), userGrp.getJctUserGroup(), userGrp.getJctUserGroupDesc(), facilitatorEmail, customerId);
					//String result = serviceDAO.updateFacilitatorDetails(1, customerId, "RW");
					
					/*
					if (statusMessage.equalsIgnoreCase("success")) {
						try {
							mailer.sendMailToRenewedUser(userLst.get(0).getJctUserEmail(), cal.getTime(), null);						
						} catch (MailingException e) {
							LOGGER.error("FACILITATOR REGISTERED....COULD NOT SEND EMAIL...  REASON: "+e.getLocalizedMessage());
							status = StatusConstants.USER_REGISTERED_MAIL_SEND_ERROR;
						}
					} else {
						LOGGER.error("UNABLE TO RENEW USER:");
						status = StatusConstants.USER_REGISTERED_MAIL_SEND_ERROR;
					}	*/																						
				}
			} catch (DAOException e) {
				LOGGER.error("UNABLE TO RENEW USER: "+e.getLocalizedMessage());
				statusMessage = "failure";
			}
		}
		LOGGER.info("<<<<<< ManageFacilitatorServiceImpl.assignNewGroupInBatch");
		return statusMessage;
	}
}