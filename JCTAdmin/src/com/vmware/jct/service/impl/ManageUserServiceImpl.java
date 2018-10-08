package com.vmware.jct.service.impl;

import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Date;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.Random;
import java.util.StringTokenizer;
import java.util.TreeMap;

import org.apache.commons.lang.WordUtils;
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
import com.vmware.jct.dao.IFacilitatorAccountServiceDAO;
import com.vmware.jct.dao.IManageServiceDAO;
import com.vmware.jct.dao.IManageServiceFacilitatorDAO;
import com.vmware.jct.dao.IPaymentDAO;
import com.vmware.jct.dao.dto.ExistingUserDTO;
import com.vmware.jct.dao.dto.FacilitatorDetailDTO;
import com.vmware.jct.dao.dto.NewUserDTO;
import com.vmware.jct.dao.dto.PaymentDetailsDTO;
import com.vmware.jct.dao.dto.UserAccountDTO;
import com.vmware.jct.dao.dto.UserGroupDTO;
import com.vmware.jct.dao.dto.UserProfileDTO;
import com.vmware.jct.exception.DAOException;
import com.vmware.jct.exception.JCTException;
import com.vmware.jct.exception.MailingException;
import com.vmware.jct.model.JctCheckPaymentUserDetails;
import com.vmware.jct.model.JctEmailDetails;
import com.vmware.jct.model.JctFacilitatorDetails;
import com.vmware.jct.model.JctPaymentDetails;
import com.vmware.jct.model.JctPaymentHeader;
import com.vmware.jct.model.JctUser;
import com.vmware.jct.model.JctUserDetails;
import com.vmware.jct.model.JctUserGroup;
import com.vmware.jct.model.JctUserProfile;
import com.vmware.jct.model.JctUserRole;
import com.vmware.jct.service.IManageUserService;
import com.vmware.jct.service.vo.FacilitatorAccountVO;
import com.vmware.jct.service.vo.GeneralUserAccountVO;
import com.vmware.jct.service.vo.ManageUserVO;
import com.vmware.jct.service.vo.UserGroupVO;
import com.vmware.jct.service.vo.UserVO;
/**
 * 
 * <p><b>Class name:</b> ManageUserServiceImpl.java</p>
 * <p><b>Author:</b> InterraIT</p>
 * <p><b>Purpose:</b> This is a ManageUserServiceImpl class. This artifact is Business layer artifact.
 * ManageUserServiceImpl implement IManageUserService interface and override the following  methods.
 * -populateExistingUserProfile()
 * -saveUserProfile(String userProfile, String createdBy)
 * -populateUserProfile()
 * -saveUserGroup(UserGroupVO userGroupVO)
 * -updateUserGroup(UserGroupVO userGroupVO, String dist)
 * -populateExistingUserGroup()
 * </p>
 * <p><b>Description:</b> This class used to perform insert, modify ,fetch data from db and delete operation. Save, update, delete operations are performed in DataAccessObject class </p>
 * <p><b>Revision History:</b>
 * 	<li></li>
 * </p>
 */
@Service
public class ManageUserServiceImpl implements IManageUserService {
	
	@Autowired
	private IManageServiceDAO serviceDAO;
	
	@Autowired  
	private MessageSource messageSource;
	
	@Autowired
	private IAuthenticatorDAO authenticatorDAO;
	
	@Autowired
	private ICommonDao commonDaoImpl;
	
	@Autowired
	private IPaymentDAO paymentDAO;
	
	@Autowired
	private IFacilitatorAccountServiceDAO facilitatorAccountServiceDAO;
	
	@Autowired
	private IManageServiceFacilitatorDAO facilitatorServiceDAO;
	
	@Autowired
	private MailNotificationHelper mailer;
	
	/*@Autowired
	private CommonUtility commonUtility;*/

	private static final Logger LOGGER = LoggerFactory.getLogger(ManageUserServiceImpl.class);

	private static final String String = null;
	
	/**
	 * Method populates existing user profile.
	 * @throws JCTException
	 * @return ManageUserVO containing the String list
	 */
	@Transactional(propagation = Propagation.REQUIRED)
	public ManageUserVO populateExistingUserProfile() throws JCTException {
		LOGGER.info(">>>>>> ManageUserServiceImpl.populateExistingUserProfile");
		StringBuilder builder = new StringBuilder("");
		ManageUserVO vo = new ManageUserVO();
		try {
			List<UserProfileDTO> dtoList = serviceDAO.populateExistingUserProfile();
			if (dtoList.size() > 0) {
				Iterator<UserProfileDTO> itr = dtoList.iterator();
				while (itr.hasNext()) {
					UserProfileDTO dto = (UserProfileDTO) itr.next();
					builder.append(dto.getUserProfileDesc());
					builder.append("~");
				}
				vo.setExistingUserProfileList(builder.toString());
				vo.setStatusCode(StatusConstants.STATUS_SUCCESS_WITH_RESULT);
			} else {
				LOGGER.info("---- NO RECORDS FOUND -----");
				vo.setStatusCode(StatusConstants.STATUS_SUCCESS_WITH_NO_RESULT);
				vo.setStatusDesc(this.messageSource.getMessage("waring.no.user.profile",null, null));
				
			}
		} catch (DAOException e) {
			LOGGER.error("----"+e.getLocalizedMessage()+" ----");
			throw new JCTException(e.getLocalizedMessage());
		}
		LOGGER.info("<<<<<< ManageUserServiceImpl.populateExistingUserProfile");
		return vo;
	}
	/**
	 * Method saves user profile
	 * @param userProfile
	 * @param created By
	 * @return ManageUserVO
	 * @throws JCTException
	 */
	@Transactional(propagation = Propagation.REQUIRED)
	public ManageUserVO saveUserProfile(String userProfile, String createdBy) throws JCTException {
		LOGGER.info(">>>>>> ManageUserServiceImpl.saveUserProfile");
		JctUserProfile profile = new JctUserProfile();
		ManageUserVO manageVO = new ManageUserVO();
		try {
			profile.setJctUserProfile(commonDaoImpl.generateKey("jct_user_profile"));
			profile.setJctUserProfileDesc(userProfile);
			profile.setJctCreatedBy(createdBy);
			profile.setJctCreatedTs(new Date());
			profile.setJctLastmodifiedBy(createdBy);
			profile.setJctLastmodifiedTs(new Date());
			profile.setJctSoftDelete(CommonConstants.ENTRY_NOT_DELETED);
			profile.setVersion(1);
			String retString = serviceDAO.saveUserProfile(profile);
			if (!retString.equals("success")) {
				manageVO.setStatusCode(StatusConstants.STATUS_FAILURE);
				manageVO.setStatusDesc(this.messageSource.getMessage("exception.dao.user.profile.save",null, null));
			} else {
				manageVO.setStatusCode(StatusConstants.STATUS_SUCCESS_WITH_RESULT);
				manageVO.setStatusDesc(retString);
			}
		} catch (DAOException e) {
			LOGGER.error("----"+e.getLocalizedMessage()+" ----");
			manageVO.setStatusCode(StatusConstants.STATUS_FAILURE);
			throw new JCTException(e.getLocalizedMessage());
		}
		LOGGER.info("<<<<<< ManageUserServiceImpl.saveUserProfile");
		return manageVO;
	}
	/**
	 * Method populates user profile
	 * @return Map
	 * @throws JCTException
	 */
	@Transactional(propagation = Propagation.REQUIRED)
	public Map<Integer, String> populateUserProfile() throws JCTException {
		LOGGER.info(">>>> ManageUserServiceImpl.populateUserProfile");
		Map<Integer, String> userProfileMap = new TreeMap<Integer, String>();
			List<UserProfileDTO> dtoList;
			try {
				dtoList = serviceDAO.populateExistingUserProfileWithId();
				for (int index = 0; index < dtoList.size(); index++) {
					UserProfileDTO dto = (UserProfileDTO) dtoList.get(index);
					userProfileMap.put(dto.getUserProfileId(), dto.getUserProfileDesc());
				}
			} catch (DAOException e) {
				LOGGER.error("----"+e.getLocalizedMessage()+" ----");
			}
			LOGGER.info("<<<<<< ManageUserServiceImpl.populateUserProfile");	
		return userProfileMap;
	}
	/**
	 * Method saves user group
	 * @param UserGroupVO
	 * @return ManageUserVO
	 * @throws JCTException
	 */
	@Transactional(propagation = Propagation.REQUIRED)
	public ManageUserVO saveUserGroup(UserGroupVO userGroupVO)
			throws JCTException {
		LOGGER.info(">>>>>> ManageUserServiceImpl.saveUserGroup");
		JctUserGroup userGroup = new JctUserGroup();
		ManageUserVO manageVO = new ManageUserVO();
		try {
			userGroup.setJctUserGroup(commonDaoImpl.generateKey("jct_user_group"));
			userGroup.setJctCreatedBy(userGroupVO.getCreatedBy());
			userGroup.setJctCreatedTs(new Date());
			userGroup.setJctLastmodifiedBy(userGroupVO.getCreatedBy());
			userGroup.setJctLastmodifiedTs(new Date());
			userGroup.setJctSoftDelete(CommonConstants.ENTRY_NOT_DELETED);
			userGroup.setJctUserGroupDesc(userGroupVO.getUserGroupName());
			userGroup.setJctUserProfileDesc(userGroupVO.getUserProfileName());
			JctUserProfile profile = new JctUserProfile();
			profile.setJctUserProfile(userGroupVO.getUserProfileId());
			userGroup.setJctUserProfile(profile);
			userGroup.setVersion(1);
			userGroup.setJctActiveStatus(StatusConstants.STATUS_ACTIVE);
			userGroup.setJctUserCustomerId(userGroupVO.getCustomerId());
			userGroup.setJctUserRoleId(userGroupVO.getRoleId());
			String retString = serviceDAO.saveUserGroup(userGroup);
			if (!retString.equals("success")) {
				manageVO.setStatusCode(StatusConstants.STATUS_FAILURE);
				manageVO.setStatusDesc(this.messageSource.getMessage("exception.dao.user.group.save",null, null));
			} else {
				manageVO.setStatusCode(StatusConstants.STATUS_SUCCESS_WITH_RESULT);
				manageVO.setStatusDesc(retString);
			}
		} catch (DAOException e) {
			LOGGER.error("----"+e.getLocalizedMessage()+" ----");
			manageVO.setStatusCode(StatusConstants.STATUS_FAILURE);
			throw new JCTException(e.getLocalizedMessage());
		}
		LOGGER.info("<<<<<< ManageUserServiceImpl.saveUserGroup");
		return manageVO;
	}
	/**
	 * Method updates user group
	 * @param userGroup
	 * @param dist
	 * @return ManageUserVO
	 * @throws JCTException
	 */
	@Transactional(propagation = Propagation.REQUIRED)
	public ManageUserVO updateUserGroup(UserGroupVO userGroupVO, String dist)
			throws JCTException {
		LOGGER.info(">>>>>> ManageUserServiceImpl.updateUserGroup");
		ManageUserVO manageVO = new ManageUserVO();
		try {
			JctUserGroup userGroup = serviceDAO.fetchUserGroup(userGroupVO.getPrimaryKeyVal());
			userGroup.setJctLastmodifiedTs(new Date());
			if (null == dist) {
				userGroup.setJctUserGroupDesc(userGroupVO.getUserGroupName());
				userGroup.setJctUserProfileDesc(userGroupVO.getUserProfileName());
				JctUserProfile profile = new JctUserProfile();
				profile.setJctUserProfile(userGroupVO.getUserProfileId());
				userGroup.setJctUserProfile(profile);
			}
			if (null != dist) {
				userGroup.setJctActiveStatus(userGroupVO.getActiveStatus());
			}
			String retString = serviceDAO.updateUserGroup(userGroup);
			if (!retString.equals("success")) {
				manageVO.setStatusCode(StatusConstants.STATUS_FAILURE);
				manageVO.setStatusDesc(this.messageSource.getMessage("exception.dao.user.group.update",null, null));
			} else {
				//If success then modify the corresponding users for that group...
				List<JctUser> userList = authenticatorDAO.getAllUsersOfUserGroup(userGroupVO.getUserGroupName());
				if (userList.size() > 0) {
					Iterator<JctUser> userItr = userList.iterator();
					while (userItr.hasNext()) {
						JctUser user = (JctUser) userItr.next();
						if (user.getJctUserSoftDelete() == 0) {
							user.setJctUserSoftDelete(1);
						} else {
							user.setJctUserSoftDelete(0);
						}
						authenticatorDAO.updateUserStatus(user);
					}
				}
				manageVO.setStatusCode(StatusConstants.STATUS_SUCCESS_WITH_RESULT);
				manageVO.setStatusDesc(retString);
			}
		} catch (DAOException e) {
			LOGGER.error("----"+e.getLocalizedMessage()+" ----");
			manageVO.setStatusCode(StatusConstants.STATUS_FAILURE);
			throw new JCTException(e.getLocalizedMessage());
		}
		LOGGER.info("<<<<<< ManageUserServiceImpl.updateUserGroup");
		return manageVO;
	}
	/**
	 * Method populates existing user Group
	 * @param userProfile Id
	 * @return ManageUserVO
	 * @throws JCTException
	 */
	@Transactional(propagation = Propagation.REQUIRED)
	public ManageUserVO populateExistingUserGroup(Integer profileId, String customerId, Integer roleId) throws JCTException {
		LOGGER.info(">>>>>> ManageUserServiceImpl.populateExistingUserGroup");
		StringBuilder builder = new StringBuilder("");
		ManageUserVO vo = new ManageUserVO();
		try {
			List<UserGroupDTO> dtoList = serviceDAO.populateExistingUserGroup(profileId, customerId, roleId);
			LOGGER.info("FOUND "+dtoList.size()+" USER GROUP ITEMS...");
			if (dtoList.size() > 0) {
				Iterator<UserGroupDTO> itr = dtoList.iterator();
				while(itr.hasNext()) {
					UserGroupDTO dto = (UserGroupDTO) itr.next();
					builder.append(dto.getUserGroupName());
					builder.append("~");
					builder.append(dto.getUserProfileDesc());
					builder.append("~");
					builder.append(dto.getActiveStatus());
					builder.append("~");
					builder.append(dto.getJctUserGroupId());
					builder.append("~");
					builder.append(dto.getUserProfileId());
					builder.append("$$$");
				}
				vo.setExistingUserGroupList(builder.toString());
				vo.setStatusCode(StatusConstants.STATUS_SUCCESS_WITH_RESULT);
			} else {
				LOGGER.info("---- NO RECORDS FOUND -----");
				vo.setStatusCode(StatusConstants.STATUS_SUCCESS_WITH_NO_RESULT);
				vo.setStatusDesc(this.messageSource.getMessage("waring.no.user.profile",null, null));
			}
		} catch (DAOException e) {
			LOGGER.error("----"+e.getLocalizedMessage()+" ----");
			throw new JCTException(e.getLocalizedMessage());
		}
		LOGGER.info("<<<<<< ManageUserServiceImpl.populateExistingUserGroup");
		return vo;
	}
	/**
	 * Method fetches list user group
	 * @return List
	 * @throws JCTException
	 */
	@Transactional(propagation=Propagation.REQUIRED)
	public List<UserGroupDTO> getUserGroupList() throws JCTException {
		LOGGER.info(">>>>>> ManageUserServiceImpl.getUserGroupList");
		List<UserGroupDTO> userGroupList=null;
		try{
			userGroupList= serviceDAO.populateDistinctUserGroup();
		}catch(DAOException daoException ){
			LOGGER.error(daoException.getLocalizedMessage());
			throw new JCTException(daoException.getMessage());
		}
		LOGGER.info("<<<<<< ManageUserServiceImpl.getUserGroupList");
		return userGroupList;
	}
	
	@Transactional(propagation=Propagation.REQUIRED)
	public List<ExistingUserDTO> getUserList(String userGroup) throws JCTException {
		LOGGER.info(">>>>>> ManageUserServiceImpl.getUserList");
		List<ExistingUserDTO> userList = null;
		try {
			userList= serviceDAO.populateUserList(userGroup);
		} catch(DAOException daoException ) {
			LOGGER.error(daoException.getLocalizedMessage());
			throw new JCTException(daoException.getMessage());
		}
		LOGGER.info("<<<<<< ManageUserServiceImpl.getUserList");
		return userList;
	}
	
	@Transactional(propagation=Propagation.REQUIRED)
	public List<ExistingUserDTO> getUserListByUserTypeAndGroup(String userGroup, String userType) throws JCTException {
		LOGGER.info(">>>>>> ManageUserServiceImpl.getUserListByUserTypeAndGroup");
		List<ExistingUserDTO> userList = null;
		try {
			userList= serviceDAO.populateUserListByUserTypeAndGroup(userGroup, userType);
		} catch(DAOException daoException ) {
			LOGGER.error(daoException.getLocalizedMessage());
			throw new JCTException(daoException.getMessage());
		}
		LOGGER.info("<<<<<< ManageUserServiceImpl.getUserListByUserTypeAndGroup");
		return userList;
	}
	
	@Transactional(propagation=Propagation.REQUIRED)
	public List<ExistingUserDTO> getUserListForAdmin(String userGroup, String emailId, int userType) throws JCTException {
		LOGGER.info(">>>>>> ManageUserServiceImpl.getUserListForAdmin");
		List<ExistingUserDTO> userList = null;
		try {
			userList= serviceDAO.populateUserListForAdmin(userGroup, emailId, userType);
		} catch(DAOException daoException ) {
			LOGGER.error(daoException.getLocalizedMessage());
			throw new JCTException(daoException.getMessage());
		}
		LOGGER.info("<<<<<< ManageUserServiceImpl.getUserListForAdmin");
		return userList;
	}

	@Transactional(propagation=Propagation.REQUIRED)
	public NewUserDTO saveUsers(List<String> validEmailIds, int userGrpId,
			String userGrpName, String createdBy, String paymentType) throws JCTException {
		LOGGER.info(">>>>>> ManageUserServiceImpl.saveUsers");
		List<NewUserDTO> newRegistration = new ArrayList<NewUserDTO>();
		List<NewUserDTO> alreadyRegisteredUsers = new ArrayList<NewUserDTO>();
		JctUser user = null;
		NewUserDTO returnDTO = new NewUserDTO();
		Iterator<String> emailItr = validEmailIds.iterator();
		while (emailItr.hasNext()) {
			try {
				String emailId = (String) emailItr.next();//checkUser
				UserVO userVO = new UserVO();
				userVO.setEmail(emailId);
				List<JctUser> userLst = authenticatorDAO.checkUser(userVO);
				if (userLst.size() == 0) {
					String generatedPassword = this.generatePassword();
					String encryptedPassword = this.encriptString(generatedPassword);
					user = prepareUserObject(null, emailId, encryptedPassword, createdBy, userGrpId, 1, 6);
					String statusMsg = authenticatorDAO.createNewUser(user);
					if (statusMsg.equals("success")) {
						NewUserDTO dto = new NewUserDTO();
						dto.setEmailId(emailId);
						dto.setPassword(generatedPassword);
						dto.setGroupName(userGrpName);
						newRegistration.add(dto);
					}
				} else {
					NewUserDTO dto = new NewUserDTO();
					dto.setEmailId(emailId);
					alreadyRegisteredUsers.add(dto);
				}
				returnDTO.setNewRegistration(newRegistration);
				returnDTO.setAlreadyRegistered(alreadyRegisteredUsers);
			} catch (DAOException ex) {
				LOGGER.error("ERROR: UNABLE TO CREATE NEW USER. REASON: "+ex.getLocalizedMessage());
				throw new JCTException("ERROR: UNABLE TO CREATE NEW USER. REASON: "+ex.getLocalizedMessage());
			}
		}
		LOGGER.info("<<<<<< ManageUserServiceImpl.saveUsers");
		return returnDTO;
	}
	
	private JctUser prepareUserObject(String paymentType,String emailId, String encryptedPassword, String createdBy, int userGrpId, int roleId, int toExpMonth) throws DAOException{
		JctUser user = new JctUser();
		user.setJctUserId(commonDaoImpl.generateKey("jct_user_id"));
		user.setJctUserEmail(emailId);
		user.setJctUserName(emailId);
		user.setJctPassword(encryptedPassword);
		user.setJctCreatedBy(createdBy);
		user.setJctLastmodifiedBy(createdBy);
		user.setJctCreatedTs(new Date());
		user.setLastmodifiedTs(new Date());
		//user.setJctActiveYn(CommonConstants.CREATED);
		JctUserRole role = new JctUserRole();
		role.setJctRoleId(roleId);
		user.setJctUserRole(role);
		
		// Check the user type
		 if (roleId == 1) {
			 if (null != paymentType) {
				 if(paymentType.equals("CHECK")){
						user.setJctActiveYn(CommonConstants.CHECK_PAYMENT_NOT_ACTIVATED);
					} else {
						user.setJctActiveYn(CommonConstants.CREATED);
					}
			 } else {
					user.setJctActiveYn(CommonConstants.CREATED);
				}
			 
			 SimpleDateFormat sdf = new SimpleDateFormat("MM/dd/yyyy");	
				Calendar calendar = Calendar.getInstance();
				//add number of months as selected by Admin
				calendar.add(Calendar.MONTH, toExpMonth);
				user.setJctAccountExpirationDate(CommonUtility.convertToDateObj(sdf.format(calendar.getTime())));
		 } else { // Facilitator
			 user.setJctAccountExpirationDate(new Date());
			 if (null != paymentType) {
				 if(paymentType.equals("CHECK")){
						user.setJctActiveYn(CommonConstants.CHECK_PAYMENT_NOT_ACTIVATED);
					} else {
						user.setJctActiveYn(CommonConstants.FACI_ACCOUNT_JUST_CREATED);
					}
			 } else {
					user.setJctActiveYn(CommonConstants.FACI_ACCOUNT_JUST_CREATED);
				}
			 
		 }
			
		return user;
	}
	
	/**
	 * Method encripts the string [Custom encryption]
	 * @param string
	 * @return
	 */
	private String encriptString(String string){
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
	@Transactional(propagation=Propagation.REQUIRED)
	public String deleteUser(String emailId, int userType) throws JCTException {
		String statusMessage = "failure";
		UserVO userVO = new UserVO();
		userVO.setEmail(emailId);
		try {
			List<JctUser> userList = authenticatorDAO.checkUserForDeletion(userVO, userType);
			JctUser user = (JctUser) userList.get(0);
			int userId = user.getJctUserId();
			//List<Integer> userIdList = new ArrayList<Integer>();
			if (user.getJctUserRole().getJctRoleId() == 3) { // Facilitator
				// get all the users which he created
				List<JctUserDetails> list = authenticatorDAO.checkUsersCreatedByFacilitator(userId);
				for (int index = 0; index < list.size(); index++) {
					JctUserDetails obj = (JctUserDetails) list.get(index);
					int id = obj.getJctUser().getJctUserId();
					String flag = authenticatorDAO.deleteUserByUserId(id);
					LOGGER.info("DELETION "+flag+" FOR USER: "+id);
				}
				String flag = authenticatorDAO.deleteUserByUserId(userId);
				LOGGER.info("DELETION "+flag+" FOR USER: "+userId);
			} else {
				String flag = authenticatorDAO.deleteUserByUserId(userId);
				LOGGER.info("DELETION "+flag+" FOR USER: "+userId);
			}
			statusMessage = "success";
		} catch (DAOException e) {
			LOGGER.error("UNABLE TO FIND / DELETE USER: "+emailId+". REASON: "+e.getLocalizedMessage());
			statusMessage = "failure";
		}
		return statusMessage;
	}

	@Transactional(propagation=Propagation.REQUIRED)
	public String updateUserActivationStatus(String emailId, int softDelete)
			throws JCTException {
		String statusMessage = "failure";
		try {
			statusMessage = authenticatorDAO.updateUserSoftDeleteStatus(emailId, softDelete);
		} catch (DAOException e) {
			LOGGER.error("UNABLE TO FIND / DELETE USER: "+emailId+". REASON: "+e.getLocalizedMessage());
			statusMessage = "failure";
		}
		return statusMessage;
	}

	@Transactional(propagation=Propagation.REQUIRED)
	public String resetPassword(String emailId, String encryptedPassword, int roleId)
			throws JCTException {
		String status = "failure";
		UserVO userVO = new UserVO();
		userVO.setEmail(emailId);
		try {
			List<JctUser> userList = authenticatorDAO.checkUserByEmailAndRole(userVO, roleId);
			JctUser userObj = (JctUser) userList.get(0);
			//If profile is 3 - Facilitator: directly change
			if (roleId == 3) {
				userObj.setJctPassword(encryptedPassword);
				status = serviceDAO.updateUser(userObj);
			} else {
				userObj.setJctPassword(encryptedPassword);
				if (userObj.getJctActiveYn() == CommonConstants.CREATED) {
					status = serviceDAO.updateUser(userObj);
				} else {
					status = "changed";
				}
			}
		} catch (DAOException ex) {
			LOGGER.error("UNABLE TO STORE USER PASSWORD: "+ex.getLocalizedMessage());
			status = "failure";
		}
		return status;
	}

	/**
	 * USER WILL NOT BE DELETED. ONLY ACTIVATED OR INACTIVATED
	 */
	/*@Transactional(propagation=Propagation.REQUIRED)
	public String deleteUserInBatch(String emailIdString) throws JCTException {
		List<String> emailList = this.getEmailList(emailIdString);
		String status = "";
		try {
			status = serviceDAO.deleteBatchUser(emailList);
		} catch (DAOException e) {
			LOGGER.error("UNABLE TO STORE USER PASSWORD: "+e.getLocalizedMessage());
			status = "failure";
		}
		return status;
	}*/
	
	private List<String> getEmailList(String batchEmail) {
		List<String> emailList = new ArrayList<String>();
		StringTokenizer token = new StringTokenizer(batchEmail,"~");
		while(token.hasMoreTokens()) {
			emailList.add((String) token.nextToken().split("`!!`")[0]);
		}
		return emailList;
	}

	@Transactional(propagation=Propagation.REQUIRED)
	public String updateUserActivationStatusInBatch(String emailIdString, int softDelete, int userType)
			throws JCTException {
		String statusMessage = "failure";
		try {
			List<String> emailList = this.getEmailList(emailIdString);
			//statusMessage = authenticatorDAO.updateUserSoftDeleteStatusInBatch(emailList, softDelete);
			statusMessage = authenticatorDAO.updateUserSoftDeleteStatusInBatch(emailList, softDelete, userType);
		} catch (DAOException e) {
			LOGGER.error("UNABLE TO FIND / CHANGE STAUS USER: "+e.getLocalizedMessage());
			statusMessage = "failure";
		}
		return statusMessage;
	}

	@Transactional(propagation=Propagation.REQUIRED)
	public List<ExistingUserDTO> populateActiveInactiveUserList(
			String userGroup, String activeInactive, int userType) throws JCTException {
		LOGGER.info(">>>>>> ManageUserServiceImpl.populateActiveInactiveUserList");
		List<ExistingUserDTO> userList = null;
		try {
			//userList= serviceDAO.populateSelectedUserList(userGroup, activeInactive);
			userList= serviceDAO.populateSelectedUserListAll(userGroup, activeInactive, userType);
		} catch(DAOException daoException ) {
			LOGGER.error(daoException.getLocalizedMessage());
			throw new JCTException(daoException.getMessage());
		}
		LOGGER.info("<<<<<< ManageUserServiceImpl.populateActiveInactiveUserList");
		return userList;
	}
	
	@Transactional(propagation=Propagation.REQUIRED)
	public List<ExistingUserDTO> populateActiveInactiveUserListForAdmin(
			String userGroup, String activeInactive, String emailId, int roleId) throws JCTException {
		LOGGER.info(">>>>>> ManageUserServiceImpl.populateActiveInactiveUserListForAdmin");
		List<ExistingUserDTO> userList = null;
		try {
			/** THIS IS DONE FOR THE PUBLIC VERSION **/
			//userList= serviceDAO.populateSelectedUserList(userGroup, activeInactive);
			userList= serviceDAO.populateSelectedUserListForAdmin(userGroup, activeInactive, emailId, roleId);
			/*****************************************/
		} catch(DAOException daoException ) {
			LOGGER.error(daoException.getLocalizedMessage());
			throw new JCTException(daoException.getMessage());
		}
		LOGGER.info("<<<<<< ManageUserServiceImpl.populateActiveInactiveUserListForAdmin");
		return userList;
	}

	@Transactional(propagation=Propagation.REQUIRED)
	public String updateCronUsers(String registeredEmailId, String generatedPassword, String updatedBy)
			throws JCTException {
		LOGGER.info(">>>>>> ManageUserServiceImpl.updateCronUsers");
		String status = "failure";
		try {
			JctEmailDetails object = null;
			//if existing in JctEmailDetails then update
			//else insert
			List<JctEmailDetails> list = serviceDAO.getEmailObject(registeredEmailId);
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
			status = serviceDAO.cronUpdater(object);
		} catch (DAOException ex) {
			status = "failure";
			LOGGER.error("UNABLE TO FETCH OR MODIFY THE EMAIL OBJECT. REASON: "+ex.getLocalizedMessage());
			throw new JCTException(ex.getLocalizedMessage());
		}
		LOGGER.info("<<<<<< ManageUserServiceImpl.updateCronUsers");
		return status;
	}
	
	@Transactional(propagation=Propagation.REQUIRED)
	public String validateExistence(String userProfile)
			throws JCTException {
		String statusMessage = "failure";
		try {
			statusMessage = serviceDAO.validateUserProfile(userProfile);
		} catch (DAOException e) {
			LOGGER.error("UNABLE TO FIND / DELETE USER: "+userProfile+". REASON: "+e.getLocalizedMessage());
			statusMessage = "failure";
		}
		return statusMessage;
	}
	
	
	@Transactional(propagation=Propagation.REQUIRED)
	public String validateExistenceUserGrp(String userGroup, int userProfile, int roleId, String customerId)
			throws JCTException {
		String statusMessage = "failure";
		try {
			statusMessage = serviceDAO.validateUserGroup(userGroup, userProfile, roleId, customerId);
		} catch (DAOException e) {
			LOGGER.error("UNABLE TO FIND / DELETE USER: "+userGroup+". REASON: "+e.getLocalizedMessage());
			statusMessage = "failure";
		}
		return statusMessage;
	}

	@Transactional(propagation=Propagation.REQUIRED)
	public int alreadyRegUserCount(List<String> emailIdList) throws JCTException {
		int regUserCount = 0;
		try {
			regUserCount = authenticatorDAO.checkUserList(emailIdList).size();
		} catch (DAOException e) {
			LOGGER.error("UNABLE TO FIND COUNT OF USERS: "+e.getLocalizedMessage());
		}
		return regUserCount;
	}
	
	@Transactional(propagation = Propagation.REQUIRED)
	public List<JctEmailDetails> getEmailDetails (int maxResultsToFetch) throws JCTException {
		List<JctEmailDetails> list = null;
		try {
			list = serviceDAO.getEmailDetails(maxResultsToFetch);
		} catch (DAOException ex) {
			LOGGER.error("----"+ex.getLocalizedMessage()+" ----");
		}
		return list;
	}
	
	@Transactional(propagation = Propagation.REQUIRED)
	public String updateEmailDetails(JctEmailDetails details) throws JCTException {
		String status = "failure";
		try {
			status = serviceDAO.updateEmailDetails(details);
		} catch (DAOException ex) {
			LOGGER.error("----"+ex.getLocalizedMessage()+" ----");
		}
		return status;
	}

	@Transactional(propagation = Propagation.REQUIRED)
	public int saveFacilitatorViaChequePayment(
			FacilitatorAccountVO facilitatorChequeVO, String paymentType) throws JCTException {
		LOGGER.info(">>>>>>>> ManageUserServiceImpl.saveFacilitatorViaChequePayment");
		int status = StatusConstants.STATUS_FAILED;
		String genUserSave = "";
		String transId = "";
		String userSave = "failure";
		JctUser facilitatorUser = null;
		try {
			// Check if the facilitator exists
			List<JctUser> faciList = authenticatorDAO.getExistingUserByEmailAndRole(facilitatorChequeVO.getFacilitatorEmail(), 3);
			if (faciList.size() == 0) {
				//Create user first
				String pwd1 = this.generatePassword();
				String encryptedPassword = this.encriptString(pwd1);
				String pwd2 = this.generatePassword();
				String encryptedPasswordUser = this.encriptString(pwd2);
				String customerId = commonDaoImpl.generateUniqueCustomerId(3);
				
				//insert user group for facilitator							
				/*JctUserGroup userGroup = new JctUserGroup();
				userGroup.setJctUserGroup(commonDaoImpl.generateKey("jct_user_group"));
				userGroup.setJctCreatedBy(facilitatorChequeVO.getFacilitatorEmail());
				userGroup.setJctCreatedTs(new Date());
				userGroup.setJctLastmodifiedBy(facilitatorChequeVO.getFacilitatorEmail());
				userGroup.setJctLastmodifiedTs(new Date());
				userGroup.setJctSoftDelete(CommonConstants.ENTRY_NOT_DELETED);
				userGroup.setJctUserGroupDesc("Default User Group");
				
				UserProfileDTO usrProfileDTO =  serviceDAO.getDefaultProfileObj("Default Profile");
				
				userGroup.setJctUserProfileDesc(usrProfileDTO.getUserProfileDesc());
				JctUserProfile profile = new JctUserProfile();
						
				profile.setJctUserProfile(usrProfileDTO.getUserProfileId());
				userGroup.setJctUserProfile(profile);
				
				userGroup.setVersion(1);
				userGroup.setJctActiveStatus(StatusConstants.STATUS_ACTIVE);
				userGroup.setJctUserCustomerId(customerId);
				userGroup.setJctUserRoleId(3);
				String retString = serviceDAO.saveUserGroup(userGroup);*/
				String retString = "success";
				
				if (retString.equalsIgnoreCase("success")) {		//	admin creation
					 facilitatorUser = prepareUserObject(paymentType,facilitatorChequeVO.getFacilitatorEmail(), 
							encryptedPassword, 
							facilitatorChequeVO.getCreatedBy(), 
							Integer.parseInt(facilitatorChequeVO.getUserGroup().split("!")[0]), 
							3, 0);
					
					facilitatorUser.setJctUserCustomerId(customerId);
					//Fetch the user profile
					JctUserGroup userGrp = serviceDAO.fetchUserGroup(Integer.parseInt(facilitatorChequeVO.getUserGroup().split("!")[0]));
					JctUserProfile userProfile = userGrp.getJctUserProfile();
					
					//Create the user details object
					JctUserDetails userDetails = new JctUserDetails();
					userDetails.setJctUserDetailsProfileId(userProfile.getJctUserProfile());
					userDetails.setJctUserDetailsProfileName(userProfile.getJctUserProfileDesc());
					//set the user group
					userDetails.setJctUserDetailsGroupId(userGrp.getJctUserGroup());
					userDetails.setJctUserDetailsGroupName(userGrp.getJctUserGroupDesc());
					if (!paymentType.equals("ECOMMERCE")) {
						userDetails.setJctUserDetailsAdminId(authenticatorDAO.getUserId(facilitatorChequeVO.getCreatedBy(), 2));
					} else {
						userDetails.setJctUserDetailsAdminId(2);
					}
					userDetails.setJctUserDetailsFacilitatorId(0);
					
					userDetails.setJctUserDetailsFirstName(facilitatorChequeVO.getFacilitatorFirstName());
					userDetails.setJctUserDetailsLastName(facilitatorChequeVO.getFacilitatorLastName());
					//String userSave = authenticatorDAO.createNewUser(facilitatorUser);
					//Save the child 1 - 1
					userDetails.setJctUser(facilitatorUser);
					facilitatorUser.setJctUserDetails(userDetails);
					userSave = authenticatorDAO.createNewUser(facilitatorUser);
					
					
					//If the facilitator will have his own account in the tool then create another for the same
					if (facilitatorChequeVO.getFacToHaveAccount().equals("Y")) {
						
						// Check if the user exists
						List<JctUser> userList = authenticatorDAO.getExistingUserByEmailAndRole(facilitatorChequeVO.getFacilitatorEmail(), 1);
						if (userList.size() == 0) {		//	 user creation
							//  as the user is under the facilitator,thus created by = facilitator 
							JctUser user = prepareUserObject(paymentType,facilitatorChequeVO.getFacilitatorEmail(), 
									encryptedPasswordUser, 
									facilitatorChequeVO.getFacilitatorEmail(), 
									Integer.parseInt(facilitatorChequeVO.getUserGroup().split("!")[0]), 
									1, 6);
							user.setJctUserCustomerId(customerId);
							
							//Create the user details object
							JctUserDetails userDetails2 = new JctUserDetails();
							userDetails2.setJctUserDetailsProfileId(userProfile.getJctUserProfile());
							userDetails2.setJctUserDetailsProfileName(userProfile.getJctUserProfileDesc());
							
							//fetch user default user group for this facilitator
							UserGroupDTO userGroupDTO =  serviceDAO.getDefaultGroupObj("Default User Group", customerId, 3);														
							//set the user group
							userDetails2.setJctUserDetailsGroupId(userGroupDTO.getJctUserGroupId());
							userDetails2.setJctUserDetailsGroupName(userGroupDTO.getUserGroupName());
							
							if (!paymentType.equals("ECOMMERCE")) {
								userDetails2.setJctUserDetailsAdminId(authenticatorDAO.getUserId(facilitatorChequeVO.getCreatedBy(), 2));
							} else {
								userDetails2.setJctUserDetailsAdminId(2);
							}
							
							userDetails2.setJctUserDetailsFacilitatorId(facilitatorUser.getJctUserId());
							
							//Save the child 1 - 1
							userDetails2.setJctUser(user);
							user.setJctUserDetails(userDetails2);
							authenticatorDAO.createNewUser(user);
						} else {
							genUserSave = "gen-user-exists";
						}
					}		
					
				}				
				if (userSave.equals("success")) {
					transId = new CommonUtility().generateTransactionCode(commonDaoImpl);
										
					JctPaymentHeader paymentHeader = new JctPaymentHeader();
					paymentHeader.setJctPmtHdrCreatedBy(facilitatorChequeVO.getCreatedBy());
					paymentHeader.setJctPmtHdrCreatedTs(new Date());
					paymentHeader.setJctPmtHdrEmailId(facilitatorChequeVO.getCreatedBy());
					paymentHeader.setJctPmtHdrModifiedBy(facilitatorChequeVO.getCreatedBy());
					paymentHeader.setJctPmtHdrModifiedTs(new Date());
					paymentHeader.setJctPmtHdrUserId(authenticatorDAO.getUserId(facilitatorChequeVO.getFacilitatorEmail(), 3));
					paymentHeader.setJctPmtHdrTotalAmt(Double.parseDouble(facilitatorChequeVO.getTotalAmountToBePaid()));
					paymentHeader.setJctUserCustomerId(customerId);
					
					JctFacilitatorDetails details = new JctFacilitatorDetails();
					details.setJctFacCreatedBy(facilitatorChequeVO.getCreatedBy());
					details.setJctFacCreatedTs(new Date());
					//If the facilitator will have his own account in the tool
					if (facilitatorChequeVO.getFacToHaveAccount().equals("Y")) {
						details.setJctFacSubscribeLimit(1);
					} else {
						details.setJctFacSubscribeLimit(0);
					}					
					//details.setJctFacSubscribeLimit(0);
					details.setJctFacSubscriptionDt(new Date());
					details.setJctFacTotalLimit(Integer.parseInt(facilitatorChequeVO.getNumberOfUsers()));
					details.setJctFacType(CommonConstants.NEW_FACILITATOR_CREATION);
					details.setJctFacUserName(facilitatorChequeVO.getFacilitatorEmail());
					details.setJctFacUserId(authenticatorDAO.getUserId(facilitatorChequeVO.getFacilitatorEmail(), 3));
					details.setJctPaymentHeader(paymentHeader);
					paymentHeader.setJctFacilitatorDetails(details);
					details.setJctUserCustomerId(customerId);					
					String save = facilitatorAccountServiceDAO.savePaymentHeader(paymentHeader);
					
					//Populate payment details table
					JctPaymentDetails paymentDetails = new JctPaymentDetails();
				//	paymentDetails.setJctPmtDtlsBankName(facilitatorChequeVO.getBankDetails());
					paymentDetails.setJctPmtDtlsChequeNos(facilitatorChequeVO.getChequeNos());
					paymentDetails.setJctPmtDtlsCreatedBy(facilitatorChequeVO.getCreatedBy());
					paymentDetails.setJctPmtDtlsCreatedTs(new Date());
					paymentDetails.setJctPmtDtlsDate(CommonUtility.convertToDateObj(facilitatorChequeVO.getPaymentDate()));
					paymentDetails.setJctPmtDtlsModifiedBy(facilitatorChequeVO.getCreatedBy());
					paymentDetails.setJctPmtDtlsModifiedTs(new Date());
					paymentDetails.setJctPmtDtlsPmtTypDesc(facilitatorChequeVO.getPaymentType());
					paymentDetails.setJctPmtDtlsPmtTransNos(transId);
					paymentDetails.setJctPmtHdrId(paymentHeader);
					String result = paymentDAO.savePaymentDetails(paymentDetails);
					
					if(paymentType.equalsIgnoreCase("CHECK")) {
						// populate jct_check_payment_user_details table
						JctCheckPaymentUserDetails checkDetails = new JctCheckPaymentUserDetails();
						checkDetails.setJctCheckPaymentCheckDate(CommonUtility.convertToDateObj(facilitatorChequeVO.getPaymentDate()));
						checkDetails.setJctCheckPaymentCheckNo(facilitatorChequeVO.getChequeNos());
						checkDetails.setJctCheckPaymentCreatedBy(facilitatorChequeVO.getCreatedBy());
						checkDetails.setJctCheckPaymentCreatedTs(new Date());
						checkDetails.setJctCheckPaymentCustomerId(customerId);
						checkDetails.setJctCheckPaymentIsHonored(0);
						checkDetails.setJctCheckPaymentmModifiedBy(facilitatorChequeVO.getCreatedBy());
						checkDetails.setJctCheckPaymentModifiedTs(new Date());
						checkDetails.setJctPmtdtlsPmtTransNos(transId);
						checkDetails.setJctUserId(facilitatorUser.getJctUserId());
						if (facilitatorChequeVO.getFacToHaveAccount().equals("Y")) {
							checkDetails.setJctCheckPaymentRoleId(5);
							checkDetails.setJctCheckPaymentExpiryDuration(6);
						} else {
							checkDetails.setJctCheckPaymentRoleId(3);
							checkDetails.setJctCheckPaymentExpiryDuration(0);
						}					
						checkDetails.setJctCheckPaymentUserName(facilitatorChequeVO.getFacilitatorEmail());
						checkDetails.setJctCheckPaymentHeaderId(paymentHeader);
						checkDetails.setJctCheckPaymentDetailsId(paymentDetails);
						checkDetails.setJctCheckPaymentUserType("NEW_FACILITATOR");				
						facilitatorAccountServiceDAO.saveCheckPaymentUserDetails(checkDetails);
					}					
					
					if (save.equals("success") && result.equals("success") && genUserSave.equals("")) {
						String fName = WordUtils.capitalize(facilitatorChequeVO.getFacilitatorFirstName().toLowerCase());
						try {
							
								mailer.intimateNewFacilitator(facilitatorChequeVO.getFacilitatorEmail(),pwd1, 
										fName, Integer.parseInt(facilitatorChequeVO.getNumberOfUsers()),
										facilitatorChequeVO.getFacToHaveAccount(),
										facilitatorUser.getJctUserCustomerId() ,paymentType);								
														
							/*if (facilitatorChequeVO.getFacToHaveAccount().equals("Y")) {
								mailer.intimateUser(facilitatorChequeVO.getFacilitatorEmail(), fName, pwd2, "register",paymentType);
							}*/
							
						} catch (MailingException e) {
							LOGGER.error("FACILITATOR REGISTERED....COULD NOT SEND EMAIL...  REASON: "+e.getLocalizedMessage());
							status = StatusConstants.USER_REGISTERED_MAIL_SEND_ERROR;
						}
					} else if (save.equals("success") && result.equals("success") && genUserSave.equals("gen-user-exists")) { 
						try {
							//facilitatorChequeVO.getFacilitatorID()
							mailer.intimateNewFacilitator(facilitatorChequeVO.getFacilitatorEmail(),pwd1, facilitatorChequeVO.getFacilitatorFirstName(),
									Integer.parseInt(facilitatorChequeVO.getNumberOfUsers()),
									"USER-EXIST",facilitatorUser.getJctUserCustomerId() ,paymentType);							
							status = StatusConstants.STATUS_GEN_USER_EXIST;
						} catch (MailingException e) {
							LOGGER.error("FACILITATOR REGISTERED....COULD NOT SEND EMAIL...  REASON: "+e.getLocalizedMessage());
							status = StatusConstants.USER_REGISTERED_MAIL_SEND_ERROR;
						}
					} else {
						LOGGER.error("SAVING TO PAYMENT HEADER AND FACILITATOR DETAILS: "+save+". SAVE TO PAYMENT DETAILS: "+result);
						status = StatusConstants.TABLE_OR_DATA_RELATED_ERROR;
					}
				}
				status = StatusConstants.STATUS_SUCCESS;
			} else {
				status = StatusConstants.STATUS_USER_EXIST;
			}
			
			
		} catch (DAOException ex) {
			LOGGER.error(ex.getLocalizedMessage());
			throw new JCTException();
		}
		LOGGER.info("<<<<<<<< ManageUserServiceImpl.saveFacilitatorViaChequePayment");
		if (genUserSave.equals("gen-user-exists")) { 
			status = StatusConstants.STATUS_GEN_USER_EXIST;
		}
		return status;
	}

	@Transactional(propagation = Propagation.REQUIRED)
	public String saveGeneralUserViaChequePayment(GeneralUserAccountVO accountVO, int validEmailListSize,String paymentType)
			throws JCTException {
		LOGGER.info(">>>>>>>> ManageUserServiceImpl.saveGeneralUserViaChequePayment");
		List<NewUserDTO> newRegistration = new ArrayList<NewUserDTO>();
		List<NewUserDTO> alreadyRegisteredUsers = new ArrayList<NewUserDTO>();
		JctUser user = null;
		StringBuffer sb = new StringBuffer("");
		Iterator<String> emailItr = accountVO.getValidEmailList().iterator();
		int mailsToUserFailedCount = 0;
		int alreadyRegisteredUsersCount = 0;
		StringBuffer registeredPaymentHdrId = new StringBuffer("");
		String transId = new CommonUtility().generateTransactionCode(commonDaoImpl);
		while (emailItr.hasNext()) {
			try {				
				String fullString = (String) emailItr.next();//checkUser
				String emailId = fullString.split("#`#")[0];
				String fName = fullString.split("#`#")[1];
				String lName = fullString.split("#`#")[2];
				
				UserVO userVO = new UserVO();
				userVO.setEmail(emailId);
				List<JctUser> userLst = authenticatorDAO.checkUserByEmailAndRole(userVO, 1);
				if (userLst.size() == 0) { // User doesn't exist.
					String generatedPassword = this.generatePassword();
					String encryptedPassword = this.encriptString(generatedPassword);
					user = prepareUserObject(paymentType, emailId, 
							encryptedPassword, 
							accountVO.getCreatedBy(), 
							Integer.parseInt(accountVO.getUserGroup().split("!")[0]), 
							1, 
							Integer.parseInt(accountVO.getExpiryDuration()));
					String customerId = commonDaoImpl.generateUniqueCustomerId(1);
					user.setJctUserCustomerId(customerId);
					
					//Fetch the user profile
					JctUserGroup userGrp = serviceDAO.fetchUserGroup(Integer.parseInt(accountVO.getUserGroup().split("!")[0]));
					JctUserProfile userProfile = userGrp.getJctUserProfile();
					
					//Create the user details object
					JctUserDetails userDetails = new JctUserDetails();
					userDetails.setJctUserDetailsProfileId(userProfile.getJctUserProfile());
					userDetails.setJctUserDetailsProfileName(userProfile.getJctUserProfileDesc());
					//set the user group
					userDetails.setJctUserDetailsGroupId(userGrp.getJctUserGroup());
					userDetails.setJctUserDetailsGroupName(userGrp.getJctUserGroupDesc());
					//userDetails.setJctUserDetailsAdminId(authenticatorDAO.getUserId(accountVO.getCreatedBy(), 2));
					if (!paymentType.equals("ECOMMERCE")) {
						userDetails.setJctUserDetailsAdminId(authenticatorDAO.getUserId(accountVO.getCreatedBy(), 2));
					} else {
						userDetails.setJctUserDetailsAdminId(2);
					}
					
					userDetails.setJctUserDetailsFacilitatorId(0);
					
					userDetails.setJctUserDetailsFirstName(fName);
					userDetails.setJctUserDetailsLastName(lName);
					//String userSave = authenticatorDAO.createNewUser(facilitatorUser);
					//Save the child 1 - 1
					userDetails.setJctUser(user);
					user.setJctUserDetails(userDetails);
					String statusMsg = authenticatorDAO.createNewUser(user);
					if (statusMsg.equals("success")) {						
						JctPaymentHeader paymentHeader = new JctPaymentHeader();
						paymentHeader.setJctPmtHdrCreatedBy(accountVO.getCreatedBy());
						paymentHeader.setJctPmtHdrCreatedTs(new Date());
						paymentHeader.setJctPmtHdrEmailId(accountVO.getCreatedBy());
						paymentHeader.setJctPmtHdrModifiedBy(accountVO.getCreatedBy());
						paymentHeader.setJctPmtHdrModifiedTs(new Date());
						paymentHeader.setJctPmtHdrTotalAmt(Double.parseDouble(accountVO.getTotalAmountToBePaid()));
						paymentHeader.setJctPmtHdrUserId(authenticatorDAO.getUserId(emailId, 1));
						paymentHeader.setJctUserCustomerId(customerId);
						
						facilitatorAccountServiceDAO.savePaymentHeader(paymentHeader);
						
						registeredPaymentHdrId.append(paymentHeader.getJctPmtHdrId()+"#");
						//Populate payment details table
						JctPaymentDetails paymentDetails = new JctPaymentDetails();
						//paymentDetails.setJctPmtDtlsBankName(accountVO.getBankDetails());
						paymentDetails.setJctPmtDtlsChequeNos(accountVO.getChequeNos());
						paymentDetails.setJctPmtDtlsCreatedBy(accountVO.getCreatedBy());
						paymentDetails.setJctPmtDtlsCreatedTs(new Date());
						paymentDetails.setJctPmtDtlsDate(CommonUtility.convertToDateObj(accountVO.getPaymentDate()));
						paymentDetails.setJctPmtDtlsModifiedBy(accountVO.getCreatedBy());
						paymentDetails.setJctPmtDtlsModifiedTs(new Date());
						paymentDetails.setJctPmtDtlsPmtTypDesc(accountVO.getPaymentType());
						paymentDetails.setJctPmtHdrId(paymentHeader);
						paymentDetails.setJctPmtDtlsPmtTransNos(transId);
						paymentDAO.savePaymentDetails(paymentDetails);
						
						if(paymentType.equalsIgnoreCase("CHECK")) {
							// populate jct_check_payment_user_details table
							JctCheckPaymentUserDetails checkDetails = new JctCheckPaymentUserDetails();
							checkDetails.setJctCheckPaymentCheckDate(CommonUtility.convertToDateObj(accountVO.getPaymentDate()));
							checkDetails.setJctCheckPaymentCheckNo(accountVO.getChequeNos());
							checkDetails.setJctCheckPaymentCreatedBy(accountVO.getCreatedBy());
							checkDetails.setJctCheckPaymentCreatedTs(new Date());
							checkDetails.setJctCheckPaymentCustomerId(customerId);
							checkDetails.setJctCheckPaymentIsHonored(0);
							checkDetails.setJctCheckPaymentmModifiedBy(accountVO.getCreatedBy());
							checkDetails.setJctCheckPaymentModifiedTs(new Date());
							checkDetails.setJctCheckPaymentRoleId(1);
							checkDetails.setJctCheckPaymentUserName(emailId);
							checkDetails.setJctCheckPaymentHeaderId(paymentHeader);
							checkDetails.setJctCheckPaymentDetailsId(paymentDetails);
							checkDetails.setJctCheckPaymentUserType("NEW_USER");
							checkDetails.setJctPmtdtlsPmtTransNos(transId);
							checkDetails.setJctUserId(user.getJctUserId());
							checkDetails.setJctCheckPaymentExpiryDuration(Integer.parseInt(accountVO.getExpiryDuration()));
							facilitatorAccountServiceDAO.saveCheckPaymentUserDetails(checkDetails);
						}
						
						
						NewUserDTO dto = new NewUserDTO();
						dto.setEmailId(emailId);
						dto.setPassword(generatedPassword);
						dto.setGroupName(accountVO.getUserGroup().split("!")[1]);
						newRegistration.add(dto);
						
						if (validEmailListSize <= 10) { 
							int status = sendEmailToGeneralUser(emailId, generatedPassword, WordUtils.capitalize(fName),								
									accountVO.getUserGroup().split("!")[1], accountVO.getCreatedBy(),paymentType);
							if (status == 1) {
								mailsToUserFailedCount = mailsToUserFailedCount + 1;
							}
						}
						
						//	update payment dtls for payment amount 
						try {
							double costPerUser = Double.parseDouble(accountVO.getTotalAmountToBePaid()) / Double.parseDouble(accountVO.getNumberOfUsers());
							//int updatedRows = serviceDAO.updateAmountInPaymentHdr(registeredPaymentHdrId.toString(), costPerUser);
							serviceDAO.updateAmountInPaymentHdr(registeredPaymentHdrId.toString(), costPerUser);
						} catch (DAOException ex) {
							LOGGER.error("ERROR: UNABLE TO UPDATE USER PAYMENT AMOUNT: "+ex.getLocalizedMessage());			
						}
					}
				} else { // User exists.
					NewUserDTO dto = new NewUserDTO();
					dto.setEmailId(emailId);
					alreadyRegisteredUsers.add(dto);
					alreadyRegisteredUsersCount = alreadyRegisteredUsersCount + 1;
				}
			} catch (DAOException ex) {
				LOGGER.error("ERROR: UNABLE TO CREATE NEW USER. REASON: "+ex.getLocalizedMessage());
				throw new JCTException("ERROR: UNABLE TO CREATE NEW USER. REASON: "+ex.getLocalizedMessage());
			}
		}	//	WHILE
		sendEmailToAdminForGeneralUserRegistration(newRegistration, accountVO.getCreatedBy(), accountVO.getInValidEmailList());
		if (validEmailListSize > 10) {
			LOGGER.info("AS NUMBER OF VALID REGISTERED USER IS MORE THAN 10, CRON WILL BE TRIGGERED TO SEND MAIL TO REGISTERED USER");
			saveCronUsers(newRegistration, accountVO.getCreatedBy());
			
			// Prepare the message
			if ((accountVO.getInValidEmailList().size() == 0) && (accountVO.getValidEmailList().size() == newRegistration.size())) {
				sb.append(this.messageSource.getMessage("manage.user.registration.admin.cron.confirmation.record.matched",null, null));
			}  else {
				int userExistsCount = accountVO.getValidEmailList().size() - newRegistration.size();
				sb.append(newRegistration.size()+" "+
						this.messageSource.getMessage("manage.user.registration.admin.cron.confirmation.record.mismatched",null, null) + 
						"\n"+accountVO.getInValidEmailList()+" "+this.messageSource.getMessage("manage.user.registration.admin.cron.confirmation.record.invalid",null, null) + 
						"\n"+userExistsCount+" "+this.messageSource.getMessage("manage.user.registration.admin.cron.confirmation.record.exists",null, null) +
						"\n\n"+this.messageSource.getMessage("manage.user.registration.admin.cron.confirmation.footer",null, null));
			}
		} else {
			if ((accountVO.getInValidEmailList().size() == 0) && (accountVO.getValidEmailList().size() == newRegistration.size())) {
				sb.append(this.messageSource.getMessage("manage.user.registration.admin.manual.success",null, null));
			} else {
				int userExistsCount = accountVO.getValidEmailList().size() - newRegistration.size();				
				sb.append(newRegistration.size()+" "+
						this.messageSource.getMessage("manage.user.registration.admin.cron.confirmation.record.mismatched",null, null) + 
						"\n"+accountVO.getInValidEmailList().size()+" "+this.messageSource.getMessage("manage.user.registration.admin.cron.confirmation.record.invalid",null, null) + 
						"\n"+userExistsCount+" "+this.messageSource.getMessage("manage.user.registration.admin.cron.confirmation.record.exists",null, null) +
						"\n"+mailsToUserFailedCount+" "+this.messageSource.getMessage("manage.user.registration.admin.cron.confirmation.email.failed",null, null) +
						"\n\n"+this.messageSource.getMessage("manage.user.registration.admin.cron.confirmation.footer",null, null));
			}
		}
		LOGGER.info("<<<<<<<< ManageUserServiceImpl.saveGeneralUserViaChequePayment");		
		
		return sb.toString();
	}	

	private String sendEmailToAdminForGeneralUserRegistration (List<NewUserDTO> newUserDTOList, String createdBy, List<String> inValidEmailList) {
		String status = "failure";
		try{
			LOGGER.info("INTIMATING THE ADMIN...");
			mailer.intimateAdminGeneralUserRegistration(newUserDTOList, createdBy, (ArrayList<String>)inValidEmailList);
			status = "success";
		} catch (MailingException e) {
			LOGGER.error("UNABLE TO SEND MAIL TO ADMIN: "+e.getLocalizedMessage());
		}
		return status;
	}
	
	private int sendEmailToGeneralUser(String emailId, String generatedPassword, String fName,
			String groupName, String createdBy, String paymentType) {
		int failed = 0;
		try{
			LOGGER.info("SENDING MAIL TO: "+emailId);
			mailer.intimateUser(emailId, generatedPassword, fName, "register",paymentType);
		} catch (MailingException e) {
			LOGGER.error("UNABLE TO SEND MAIL TO USER: "+e.getLocalizedMessage());
			LOGGER.error("INSERTING IN CRON TABLE: "+e.getLocalizedMessage());
			failed = 1;
			//Insert in the cron
			try {
				NewUserDTO dto = new NewUserDTO();
				List<NewUserDTO> registeredEmailIds = new ArrayList<NewUserDTO>();
				dto.setEmailId(emailId);
				dto.setPassword(generatedPassword);
				dto.setGroupName(groupName);
				saveCronUsers(registeredEmailIds, createdBy);
			} catch (JCTException e1) {
				LOGGER.error("UNABLE TO INSERT IN CRON TABLE: "+e1.getLocalizedMessage());
			}
		}
		return failed;
	}
	
	@Transactional(propagation = Propagation.REQUIRED)
	public int renewSubscribeViaChequePayment(
			FacilitatorAccountVO facilitatorChequeVO, String inputRenewSub) throws JCTException {
		
		LOGGER.info(">>>>>>>> ManageUserServiceImpl.renewSubscribeViaChequePayment");
		int status = StatusConstants.STATUS_FAILED;
		String genUserSave = "";
		
		try {
			// Check if the facilitator exists
			List<Object[]> existingFaci = authenticatorDAO.getFacilitatorByMailId(facilitatorChequeVO.getFacilitatorID());
			
						
			if (existingFaci.size() != 0) {		//	faci exists	
				String transId = new CommonUtility().generateTransactionCode(commonDaoImpl);
				
					JctPaymentHeader paymentHeader = new JctPaymentHeader();					
					paymentHeader.setJctPmtHdrEmailId(facilitatorChequeVO.getCreatedBy());
					paymentHeader.setJctPmtHdrCreatedTs(new Date());
					paymentHeader.setJctPmtHdrCreatedBy(facilitatorChequeVO.getCreatedBy());
					paymentHeader.setJctPmtHdrModifiedBy(facilitatorChequeVO.getCreatedBy());
					paymentHeader.setJctPmtHdrModifiedTs(new Date());
					paymentHeader.setJctPmtHdrTotalAmt(Double.parseDouble(facilitatorChequeVO.getTotalAmountToBePaid()));
					paymentHeader.setJctPmtHdrUserId(Integer.parseInt((existingFaci.get(0)[0].toString())));	
					paymentHeader.setJctUserCustomerId(facilitatorChequeVO.getFacilitatorID());
					String save = facilitatorAccountServiceDAO.savePaymentHeader(paymentHeader);
					
					JctFacilitatorDetails details = new JctFacilitatorDetails();
					details.setJctFacCreatedBy(facilitatorChequeVO.getCreatedBy());
					details.setJctFacCreatedTs(new Date());
					details.setJctFacSubscribeLimit(0);
					details.setJctFacSubscriptionDt(new Date());
					details.setJctFacTotalLimit(Integer.parseInt(facilitatorChequeVO.getNumberOfUsers()));										
					details.setJctFacUserName(facilitatorChequeVO.getFacilitatorEmail());					
					details.setJctFacUserId(authenticatorDAO.getUserId( facilitatorChequeVO.getFacilitatorEmail(), 3 ));					
					details.setJctPaymentHeader(paymentHeader);
					if( facilitatorChequeVO.getPaymentType().equals("CHECK")){
						details.setJctFacType(inputRenewSub+"_CHECK");	//	check is not cleared						
					} else {
						details.setJctFacType(inputRenewSub);						
					}					
					paymentHeader.setJctFacilitatorDetails(details);
					details.setJctUserCustomerId((existingFaci.get(0)[2].toString()));
										
					//Populate payment details table
					JctPaymentDetails paymentDetails = new JctPaymentDetails();
					paymentDetails.setJctPmtDtlsChequeNos(facilitatorChequeVO.getChequeNos());
					paymentDetails.setJctPmtDtlsCreatedBy(facilitatorChequeVO.getCreatedBy());
					paymentDetails.setJctPmtDtlsCreatedTs(new Date());
					paymentDetails.setJctPmtDtlsPmtTypDesc(facilitatorChequeVO.getPaymentType());
					paymentDetails.setJctPmtDtlsDate(CommonUtility.convertToDateObj(facilitatorChequeVO.getPaymentDate()));
					paymentDetails.setJctPmtDtlsModifiedBy(facilitatorChequeVO.getCreatedBy());
					paymentDetails.setJctPmtDtlsModifiedTs(new Date());
					paymentDetails.setJctPmtHdrId(paymentHeader);
					paymentDetails.setJctPmtDtlsPmtTransNos(transId);
					String result = paymentDAO.savePaymentDetails(paymentDetails);
					
					if( facilitatorChequeVO.getPaymentType().equals("CHECK")){
						// populate jct_check_payment_user_details table
						JctCheckPaymentUserDetails checkDetails = new JctCheckPaymentUserDetails();
						checkDetails.setJctCheckPaymentCheckDate(CommonUtility.convertToDateObj(facilitatorChequeVO.getPaymentDate()));
						checkDetails.setJctCheckPaymentCheckNo(facilitatorChequeVO.getChequeNos());
						checkDetails.setJctCheckPaymentCreatedBy(facilitatorChequeVO.getCreatedBy());
						checkDetails.setJctCheckPaymentCreatedTs(new Date());
						checkDetails.setJctCheckPaymentCustomerId(facilitatorChequeVO.getFacilitatorID());
						checkDetails.setJctCheckPaymentIsHonored(0);
						checkDetails.setJctCheckPaymentmModifiedBy(facilitatorChequeVO.getCreatedBy());
						checkDetails.setJctCheckPaymentModifiedTs(new Date());
						checkDetails.setJctCheckPaymentRoleId(3);
						checkDetails.setJctCheckPaymentUserName(facilitatorChequeVO.getFacilitatorEmail());
						checkDetails.setJctCheckPaymentHeaderId(paymentHeader);
						checkDetails.setJctCheckPaymentDetailsId(paymentDetails);
						checkDetails.setJctCheckPaymentUserType(inputRenewSub);
						checkDetails.setJctPmtdtlsPmtTransNos(transId);
						checkDetails.setJctUserId(Integer.parseInt((existingFaci.get(0)[0].toString())));
						
						facilitatorAccountServiceDAO.saveCheckPaymentUserDetails(checkDetails);			
						}
										
					if (save.equals("success") && result.equals("success") && genUserSave.equals("")) {
						try {							
							mailer.renewSubscribeFacilitator(facilitatorChequeVO.getFacilitatorEmail(), 
								facilitatorChequeVO.getFacilitatorID(),
								Integer.parseInt(facilitatorChequeVO.getNumberOfUsers()),
								inputRenewSub, facilitatorChequeVO.getPaymentType());												
								
						} catch (MailingException e) {
								LOGGER.error("FACILITATOR RENEWED / SUBSCRIBED....COULD NOT SEND EMAIL...  REASON: "+e.getLocalizedMessage());
								status = StatusConstants.USER_REGISTERED_MAIL_SEND_ERROR;
						}
					} else {
						LOGGER.error("SAVING TO PAYMENT HEADER AND FACILITATOR DETAILS: "+save+". SAVE TO PAYMENT DETAILS: "+result);
						status = StatusConstants.TABLE_OR_DATA_RELATED_ERROR;
					}
				status = StatusConstants.STATUS_SUCCESS;
			} else {
				status = StatusConstants.USER_NOT_REGISTERED;
			}
		} catch (DAOException ex) {
			LOGGER.error(ex.getLocalizedMessage());
			throw new JCTException();
		}
		LOGGER.info("<<<<<<<< ManageUserServiceImpl.renewSubscribeViaChequePayment");
		return status;
	}
	@Transactional(propagation = Propagation.REQUIRED)
	public String getFacilitatorEmailByID(String facilitatorID) throws JCTException {
	
		StringBuilder facilitatorEmail = new StringBuilder("");
		try{
			List<Object[]> faciExist = authenticatorDAO.getFacilitatorEmailByID(facilitatorID);
			 if( null != faciExist && faciExist.size() >0) {
				 facilitatorEmail.append(faciExist.get(0)[1].toString());				 
			 } else {
				 return null;
			 }		
			
		} catch(DAOException ex){
			LOGGER.error(ex.getLocalizedMessage());			
		} 
		return facilitatorEmail.toString();
	}
	
	@Transactional(propagation = Propagation.REQUIRED)
	public int renewSubscribeViaCMPPayment(
			FacilitatorAccountVO facilitatorChequeVO, String inputRenewSub) throws JCTException {
		
		LOGGER.info(">>>>>>>> ManageUserServiceImpl.renewSubscribeViaCMPPayment");
		int status = StatusConstants.STATUS_FAILED;
		String genUserSave = "";
		try {
			// Check if the facilitator exists
			List<Object[]> existingFaci = authenticatorDAO.getFacilitatorEmailByID(facilitatorChequeVO.getFacilitatorID());
			
			if (existingFaci.size() != 0) {			//	facilitator exists		
				String transId = new CommonUtility().generateTransactionCode(commonDaoImpl);
					JctPaymentHeader paymentHeader = new JctPaymentHeader();					
					paymentHeader.setJctPmtHdrEmailId(facilitatorChequeVO.getCreatedBy());
					paymentHeader.setJctPmtHdrCreatedTs(new Date());
					paymentHeader.setJctPmtHdrCreatedBy(facilitatorChequeVO.getCreatedBy());
					paymentHeader.setJctPmtHdrModifiedBy(facilitatorChequeVO.getCreatedBy());
					paymentHeader.setJctPmtHdrModifiedTs(new Date());
					paymentHeader.setJctPmtHdrTotalAmt(Double.parseDouble(facilitatorChequeVO.getTotalAmountToBePaid()));
					paymentHeader.setJctPmtHdrUserId(1);	
					paymentHeader.setJctUserCustomerId(facilitatorChequeVO.getFacilitatorID());
					
					JctFacilitatorDetails details = new JctFacilitatorDetails();
					details.setJctFacCreatedBy(facilitatorChequeVO.getCreatedBy());
					details.setJctFacCreatedTs(new Date());
					details.setJctFacSubscribeLimit(0);
					details.setJctFacSubscriptionDt(new Date());
					details.setJctFacTotalLimit(Integer.parseInt(facilitatorChequeVO.getNumberOfUsers()));										
					details.setJctFacType(CommonConstants.RENEW_SUBSCRIBE_FACILITATOR);										
					details.setJctFacUserName(facilitatorChequeVO.getFacilitatorEmail());					
					details.setJctFacUserId(authenticatorDAO.getUserId( facilitatorChequeVO.getFacilitatorEmail(), 3 ));					
					details.setJctPaymentHeader(paymentHeader);
					details.setJctFacType(inputRenewSub);
					paymentHeader.setJctFacilitatorDetails(details);
					details.setJctUserCustomerId(facilitatorChequeVO.getFacilitatorID());
					paymentHeader.setJctPmtHdrUserId(Integer.parseInt(existingFaci.get(0)[0].toString()));				
					String save = facilitatorAccountServiceDAO.savePaymentHeader(paymentHeader);
					
					//Populate payment details table
					JctPaymentDetails paymentDetails = new JctPaymentDetails();
				//	paymentDetails.setJctPmtDtlsBankName(facilitatorChequeVO.getBankDetails());
					paymentDetails.setJctPmtDtlsChequeNos(facilitatorChequeVO.getChequeNos());
					paymentDetails.setJctPmtDtlsCreatedBy(facilitatorChequeVO.getCreatedBy());
					paymentDetails.setJctPmtDtlsCreatedTs(new Date());
					paymentDetails.setJctPmtDtlsDate(CommonUtility.convertToDateObj(facilitatorChequeVO.getPaymentDate()));
					paymentDetails.setJctPmtDtlsModifiedBy(facilitatorChequeVO.getCreatedBy());
					paymentDetails.setJctPmtDtlsModifiedTs(new Date());
					paymentDetails.setJctPmtDtlsPmtTypDesc(facilitatorChequeVO.getPaymentType());
					paymentDetails.setJctPmtHdrId(paymentHeader);
					paymentDetails.setJctPmtDtlsPmtTransNos(transId);
					String result = paymentDAO.savePaymentDetails(paymentDetails);
					
					if (save.equals("success") && result.equals("success") && genUserSave.equals("")) {
						try {							
							mailer.renewSubscribeFacilitator(facilitatorChequeVO.getFacilitatorEmail(), 
								facilitatorChequeVO.getFacilitatorID(),
								Integer.parseInt(facilitatorChequeVO.getNumberOfUsers()), 
								inputRenewSub, facilitatorChequeVO.getPaymentType());												
							
						} catch (MailingException e) {
							LOGGER.error("FACILITATOR RENEWED / SUBSCRIBED....COULD NOT SEND EMAIL...  REASON: "+e.getLocalizedMessage());
							status = StatusConstants.USER_REGISTERED_MAIL_SEND_ERROR;
						}
					} else {
						LOGGER.error("SAVING TO PAYMENT HEADER AND FACILITATOR DETAILS: "+save+". SAVE TO PAYMENT DETAILS: "+result);
						status = StatusConstants.TABLE_OR_DATA_RELATED_ERROR;
					}
				status = StatusConstants.STATUS_SUCCESS;
			} else {
				status = StatusConstants.USER_NOT_REGISTERED;
			}
		} catch (DAOException ex) {
			LOGGER.error(ex.getLocalizedMessage());
			throw new JCTException();
		}
		LOGGER.info("<<<<<<<< ManageUserServiceImpl.renewSubscribeViaCMPPayment");
		return status;
	}	
/***
 * Method to fetch / search the existing users having check payment 
 */
	@Transactional(propagation=Propagation.REQUIRED)
	public String fetchExistingChequeUsers(String fetchType, String chequeNum, String userName) throws JCTException {
		LOGGER.info(">>>> ManageUserServiceImpl.fetchChequeDetails");
		List<PaymentDetailsDTO> userUniqueFieldsList = null;
		List<PaymentDetailsDTO> userRemainingFieldsList = null;
		StringBuilder chkDetails = new StringBuilder("");
		try {
			if(fetchType.equals("BY_USERNAME")){
				String chkNo = paymentDAO.getCHkNoByUserName(userName);	//	need to get check no using user name to fetch payment details
				userUniqueFieldsList= paymentDAO.fetchExistingChequeUsersUniqueFields(fetchType, chkNo, userName);
			} else if (fetchType.equals("BY_CHECK")){
				//chkNo = chequeNum;
				userUniqueFieldsList= paymentDAO.fetchExistingChequeUsersUniqueFields(fetchType, chequeNum, userName);
			} else if(fetchType.equals("BY_ALL")) {
				userUniqueFieldsList= paymentDAO.fetchExistingChequeUsersUniqueFields(fetchType, chequeNum, userName);
			} else {	// on page load fetchType = ""
				userUniqueFieldsList= paymentDAO.fetchExistingChequeUsersUniqueFields(fetchType, chequeNum, userName);				
			}
			//chkNo = fetchType.equals("BY_USERNAME") ? paymentDAO.getCHkNoByUserName(userName) : chequeNum;
			
			if(userUniqueFieldsList.size() != 0){
				SimpleDateFormat dateFormat = new SimpleDateFormat("dd MMMM, yyyy");
				for (int index = 0; index < userUniqueFieldsList.size(); index++) {
					PaymentDetailsDTO innerObj = (PaymentDetailsDTO)userUniqueFieldsList.get(index);
					chkDetails.append(innerObj.getChequeNos()+"~");					
					chkDetails.append(dateFormat.format(innerObj.getDate())+"~");
					chkDetails.append(innerObj.getJctPmtdtlsPmtTransNos()+"$$");
					try {
						userRemainingFieldsList = paymentDAO.fetchuserRemainingFields(innerObj.getJctPmtdtlsPmtTransNos());	//	fetch remaining details		
						for(int i = 0; i < userRemainingFieldsList.size(); i++){
							PaymentDetailsDTO objAll = (PaymentDetailsDTO)userRemainingFieldsList.get(i);
							chkDetails.append(objAll.getUserName()+"~");
							chkDetails.append("!");
						}						
					chkDetails.append("#");				
					} catch(DAOException daoEx){
						LOGGER.error(daoEx.getLocalizedMessage());
						throw new JCTException(daoEx.getMessage());
					}				
				}
			}
			
		} catch(DAOException daoException ) {
			LOGGER.error(daoException.getLocalizedMessage());
			throw new JCTException(daoException.getMessage());
		}
		LOGGER.info("<<<< ManageUserServiceImpl.fetchChequeDetails");
		return chkDetails.toString();
	}

	@Transactional(propagation=Propagation.REQUIRED)
	public int honoredCheck(String tranId,String createdBy) throws JCTException {
		LOGGER.info(">>>> ManageUserServiceImpl.honoredCheck");
		int status = 0;	//	failure		
		String payDtlStatus = "failure";
		try {
			List<JctCheckPaymentUserDetails> paymntDtlList = paymentDAO.searchByTranId(tranId);
				//Subscribe/Renew
				for (int index = 0; index < paymntDtlList.size(); index++) {
					JctCheckPaymentUserDetails innerObj = (JctCheckPaymentUserDetails)paymntDtlList.get(index);
					
					if(innerObj.getJctCheckPaymentUserType().equalsIgnoreCase("NEW_FACILITATOR") 
							|| innerObj.getJctCheckPaymentUserType().equalsIgnoreCase("NEW_USER")){		//		NEW REGISTRATION												
						//Update user's account expiration date
						int expDuration = innerObj.getJctCheckPaymentExpiryDuration();											
						Calendar calendar = Calendar.getInstance();			 
						java.util.Date now = calendar.getTime();
						java.sql.Timestamp currentTimestamp = new java.sql.Timestamp(now.getTime());
						Calendar cal = Calendar.getInstance(); 
						cal.setTime(currentTimestamp);
						cal.add(Calendar.MONTH, expDuration);
																
						// activate the user in user table
						status = serviceDAO.updateUserToHonor(innerObj.getJctUserId(),innerObj.getJctCheckPaymentRoleId(), cal.getTime());
						//update jct_payment_details				
						payDtlStatus = serviceDAO.updatePaymentDtlToHonor(innerObj.getJctCheckPaymentDetailsId().getJctPmtDtlsId(), createdBy);					
						if(status == 1 && payDtlStatus.equalsIgnoreCase("success") ) {
							//update jct_check_payment_user_details	
							status = serviceDAO.updateCheckPaymentDetails(innerObj.getJctUserId());	
						}
					} else {	//	SUBSCRIBE / RENEW MORE USER
						//update facilitator details - set-> AD / RW
						if(innerObj.getJctCheckPaymentRoleId() == 3) { //FACILITATOR
							status = serviceDAO.updateFacilitatorDtlToHonor(innerObj.getJctCheckPaymentHeaderId().getJctPmtHdrId(),
									innerObj.getJctCheckPaymentCustomerId(), innerObj.getJctCheckPaymentCreatedBy());		
						} else { //INDIVIDUAL USER RENEW
							int expDuration = innerObj.getJctCheckPaymentExpiryDuration();
							JctUser userObj = facilitatorServiceDAO.fetchUserData(innerObj.getJctUserId());
							String customerId = userObj.getJctUserCustomerId();
							
							Date expiryDate = userObj.getJctAccountExpirationDate();	
							Calendar calendar = Calendar.getInstance();			 
							java.util.Date now = calendar.getTime();
							java.sql.Timestamp currentTimestamp = new java.sql.Timestamp(now.getTime());
							Calendar cal = Calendar.getInstance(); 
							if (customerId.substring(0,2).equals(CommonConstants.FACILITATOR_USER_CUST_CODE)) {
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
									cal.add(Calendar.MONTH, expDuration);
								} 
								else{
									cal.setTime(expiryDate);
									cal.add(Calendar.MONTH, expDuration);
								}
							}							
							String statusMsg = facilitatorServiceDAO.renewUserByFacilitator(userObj.getJctUserId(), userObj.getJctUserEmail(), createdBy, cal.getTime());
							try {
								mailer.sendMailToRenewedUser(userObj.getJctUserEmail(), cal.getTime(), null);						
							} catch (MailingException e) {
								LOGGER.error("FACILITATOR REGISTERED....COULD NOT SEND EMAIL...  REASON: "+e.getLocalizedMessage());								
							};
							if(statusMsg.equalsIgnoreCase("success")) {
								status = 1;
							} else {
								status = 0;
							}						
						}
										
						//update payment details
						payDtlStatus = serviceDAO.updatePaymentDtlToHonor(innerObj.getJctCheckPaymentDetailsId().getJctPmtDtlsId(), createdBy);
						
						if((status == 1) && payDtlStatus.equalsIgnoreCase("success")) {
							//update jct_check_payment_user_details
							status = serviceDAO.updateCheckPaymentDetails(innerObj.getJctUserId());
						}
					}
					
				}
			
		} catch (DAOException e) {
			LOGGER.error(e.getLocalizedMessage());
		}					
		LOGGER.info("<<<< ManageUserServiceImpl.honoredCheck");
		return status;
	}
	
/**
 * method to dishonored payment check
 *  @param tranId
 *  @return status 
 *  **/	
@Transactional(propagation=Propagation.REQUIRED)
public String dishonoredCheck(String tranId, String createdBy)
		throws JCTException {
	LOGGER.info(">>>> ManageUserServiceImpl.dishonoredCheck");
	String status = "failure";		
	String payDtlStatus = "failure";
	String payHdrStatus = "failure";
	StringBuffer userDtlsStr = new StringBuffer("");
	String fName = "";	
	try {
		List<JctCheckPaymentUserDetails> paymntDtlList = paymentDAO.searchByTranId(tranId);	//	search user from check table
			if(paymntDtlList.size() != 0){
				//Subscribe/Renew
				for (int index = 0; index < paymntDtlList.size(); index++) {
					JctCheckPaymentUserDetails innerObj = (JctCheckPaymentUserDetails)paymntDtlList.get(index);
					fName = getFirstName(innerObj.getJctCheckPaymentUserName(), innerObj.getJctCheckPaymentRoleId());
					if(innerObj.getJctCheckPaymentUserType().equalsIgnoreCase("NEW_FACILITATOR") 
							|| innerObj.getJctCheckPaymentUserType().equalsIgnoreCase("NEW_USER")){		//		NEW REGISTRATION
						
						status = serviceDAO.deleteNewUser(innerObj.getJctUserId(),innerObj.getJctCheckPaymentRoleId());	//	hard delete user from user table				
					}
					//	soft delete -> payment details & header, for both cases new registration and sub-renew 			
					payDtlStatus = serviceDAO.updatePaymentDtlToDishonor(innerObj.getJctCheckPaymentDetailsId().getJctPmtDtlsId(),createdBy);					
					
					payHdrStatus = serviceDAO.updatePaymentHdrToDishonor(innerObj.getJctCheckPaymentHeaderId().getJctPmtHdrId(), createdBy);
					
					SimpleDateFormat dateFormat = new SimpleDateFormat("MMMM d, yyyy");					
					if(innerObj.getJctCheckPaymentUserType().equalsIgnoreCase("NEW_FACILITATOR") 
							|| innerObj.getJctCheckPaymentUserType().equalsIgnoreCase("NEW_USER")) {	//		NEW REGISTRATION
						
						if (payDtlStatus.equalsIgnoreCase("success") && payHdrStatus.equalsIgnoreCase("success") && status.equalsIgnoreCase("success")){
							status = serviceDAO.updateCheckPaymentDetailsToDishonor(innerObj.getJctCheckPaymentUserId());	
							userDtlsStr.append(innerObj.getJctCheckPaymentUserName()+"~"+innerObj.getJctCheckPaymentCheckNo()
									+"~"+innerObj.getJctCheckPaymentRoleId()+"~"+innerObj.getJctCheckPaymentUserType()
									+"~"+dateFormat.format(innerObj.getJctCheckPaymentCheckDate())
									+"~"+ fName+"#");
						}						
					} else {
						if (payDtlStatus.equalsIgnoreCase("success") && payHdrStatus.equalsIgnoreCase("success")){
							status = serviceDAO.updateCheckPaymentDetailsToDishonor(innerObj.getJctCheckPaymentUserId());
							userDtlsStr.append(innerObj.getJctCheckPaymentUserName()+"~"+innerObj.getJctCheckPaymentCheckNo()
									+"~"+innerObj.getJctCheckPaymentRoleId()+"~"+innerObj.getJctCheckPaymentUserType()
									+"~"+dateFormat.format(innerObj.getJctCheckPaymentCheckDate())+"~"+ fName+"#");
						}						
					}
				}	//	for
			}		
	} catch (DAOException e) {
		LOGGER.error(e.getLocalizedMessage());
	}
	LOGGER.info("<<<< ManageUserServiceImpl.dishonoredCheck");
	return userDtlsStr.toString();
}

	@Transactional(propagation=Propagation.REQUIRED)
	public List<PaymentDetailsDTO> getExistingUserByTranId(String tranId) throws JCTException {
		LOGGER.info(">>>> ManageUserServiceImpl.getExistingUserByEmailAndRole");
		List<PaymentDetailsDTO> userList = null;
		try {
			 userList = authenticatorDAO.getExistingUserByTranId(tranId);
		} catch (DAOException e) {
			LOGGER.error(e.getLocalizedMessage());
		}
		LOGGER.info("<<<<< ManageUserServiceImpl.getExistingUserByEmailAndRole");
		return userList;
	}	
	
	@Transactional(propagation = Propagation.REQUIRED)
	public String getFacilitatorActiveYNByID(String facilitatorID) throws JCTException {
	
		StringBuilder facilitatorEmail = new StringBuilder("");
		try{
			List<Object[]> faciExist = authenticatorDAO.getFacilitatorEmailByID(facilitatorID);
			 if( null != faciExist) {
				 facilitatorEmail.append(faciExist.get(0).toString());
			 } else {
				 return null;
			 }		
			
		} catch(DAOException ex){
			LOGGER.error(ex.getLocalizedMessage());			
		} 
		return facilitatorEmail.toString();
	}

	@Transactional(propagation = Propagation.REQUIRED)
	public List<java.lang.String> getUserNamesOfIndUser(
			List<java.lang.String> emailIdList) throws JCTException {
		List<String> usernameList = new ArrayList<String>();
		try{
			usernameList = authenticatorDAO.getUserNamesOfIndUser(emailIdList); 	
		} catch(DAOException ex){
			LOGGER.error(ex.getLocalizedMessage());			
		} 
		return usernameList;
	}
	
	@Transactional(propagation = Propagation.REQUIRED)
	public List<String> getUserNamesForFacilitator (
			List<String> emailIdList, String customerId) throws JCTException {
		List<String> usernameList = new ArrayList<String>();
		try{
			usernameList = authenticatorDAO.getUserNamesForFacilitator(emailIdList, customerId); 	
		} catch(DAOException ex){
			LOGGER.error(ex.getLocalizedMessage());			
		} 
		return usernameList;
	}

	@Transactional(propagation = Propagation.REQUIRED)
	public String renewGeneralUserPaymentManual(
			GeneralUserAccountVO accountVO, int validEmailListSize, String paymentType) throws JCTException {
		LOGGER.info(">>>>>>>> ManageUserServiceImpl.renewGeneralUserPaymentManual");	
		String statusMsg = "success";
		StringBuffer sb = new StringBuffer("");		
		StringBuffer registeredPaymentHdrId = new StringBuffer("");
		String transId = new CommonUtility().generateTransactionCode(commonDaoImpl);
		Iterator<String> emailItr = accountVO.getEmailWithExpDateList().iterator();
		while (emailItr.hasNext()) {
			try {
				String emailWithExpDuratn = (String) emailItr.next();
				String[] initialValue = emailWithExpDuratn.split("#`#");
				String emailId = initialValue[0];
				int expDuration = Integer.parseInt(initialValue[1]);
				
				UserVO userVO = new UserVO();
				userVO.setEmail(emailId);
				
				JctUser userObj = serviceDAO.checkIndividualForRenew(userVO, 1);			
				String customerId = userObj.getJctUserCustomerId();
				Date expiryDate = userObj.getJctAccountExpirationDate();	
				Calendar calendar = Calendar.getInstance();			 
				java.util.Date now = calendar.getTime();
				java.sql.Timestamp currentTimestamp = new java.sql.Timestamp(now.getTime());
				Calendar cal = Calendar.getInstance(); 
				/**
				 * Expiration date will not be modified if payment mode is check
				 */
					if(!paymentType.equalsIgnoreCase("CHECK")) {
						if (customerId.substring(0,2).equals(CommonConstants.FACILITATOR_USER_CUST_CODE)) {
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
								cal.add(Calendar.MONTH, expDuration);
							} 
							else{
								cal.setTime(expiryDate);
								cal.add(Calendar.MONTH, expDuration);
							}
						}
						statusMsg = facilitatorServiceDAO.renewUserByFacilitator(userObj.getJctUserId(), userObj.getJctUserEmail(), accountVO.getCreatedBy(), cal.getTime());
					}
					
					
										
					if (statusMsg.equals("success")) {						
						JctPaymentHeader paymentHeader = new JctPaymentHeader();
						paymentHeader.setJctPmtHdrCreatedBy(accountVO.getCreatedBy());
						paymentHeader.setJctPmtHdrCreatedTs(new Date());
						paymentHeader.setJctPmtHdrEmailId(accountVO.getCreatedBy());
						paymentHeader.setJctPmtHdrModifiedBy(accountVO.getCreatedBy());
						paymentHeader.setJctPmtHdrModifiedTs(new Date());
						paymentHeader.setJctPmtHdrTotalAmt(Double.parseDouble(accountVO.getTotalAmountToBePaid()));
						paymentHeader.setJctPmtHdrUserId(userObj.getJctUserId());
						paymentHeader.setJctUserCustomerId(userObj.getJctUserCustomerId());
						
						facilitatorAccountServiceDAO.savePaymentHeader(paymentHeader);
						
						registeredPaymentHdrId.append(paymentHeader.getJctPmtHdrId()+"#");
						//Populate payment details table
						JctPaymentDetails paymentDetails = new JctPaymentDetails();
						//paymentDetails.setJctPmtDtlsBankName(accountVO.getBankDetails());
						paymentDetails.setJctPmtDtlsChequeNos(accountVO.getChequeNos());
						paymentDetails.setJctPmtDtlsCreatedBy(accountVO.getCreatedBy());
						paymentDetails.setJctPmtDtlsCreatedTs(new Date());
						paymentDetails.setJctPmtDtlsDate(CommonUtility.convertToDateObj(accountVO.getPaymentDate()));
						paymentDetails.setJctPmtDtlsModifiedBy(accountVO.getCreatedBy());
						paymentDetails.setJctPmtDtlsModifiedTs(new Date());
						paymentDetails.setJctPmtDtlsPmtTypDesc(accountVO.getPaymentType());
						paymentDetails.setJctPmtHdrId(paymentHeader);
						paymentDetails.setJctPmtDtlsPmtTransNos(transId);
						paymentDAO.savePaymentDetails(paymentDetails);
						
						if(paymentType.equalsIgnoreCase("CHECK")) {
							// populate jct_check_payment_user_details table
							JctCheckPaymentUserDetails checkDetails = new JctCheckPaymentUserDetails();
							checkDetails.setJctCheckPaymentCheckDate(CommonUtility.convertToDateObj(accountVO.getPaymentDate()));
							checkDetails.setJctCheckPaymentCheckNo(accountVO.getChequeNos());
							checkDetails.setJctCheckPaymentCreatedBy(accountVO.getCreatedBy());
							checkDetails.setJctCheckPaymentCreatedTs(new Date());
							checkDetails.setJctCheckPaymentCustomerId(userObj.getJctUserCustomerId());
							checkDetails.setJctCheckPaymentIsHonored(0);
							checkDetails.setJctCheckPaymentmModifiedBy(accountVO.getCreatedBy());
							checkDetails.setJctCheckPaymentModifiedTs(new Date());
							checkDetails.setJctCheckPaymentRoleId(1);
							checkDetails.setJctCheckPaymentUserName(emailId);
							checkDetails.setJctCheckPaymentHeaderId(paymentHeader);
							checkDetails.setJctCheckPaymentDetailsId(paymentDetails);
							checkDetails.setJctCheckPaymentUserType("RENEW");
							checkDetails.setJctPmtdtlsPmtTransNos(transId);
							checkDetails.setJctUserId(userObj.getJctUserId());
							checkDetails.setJctCheckPaymentExpiryDuration(expDuration);
							facilitatorAccountServiceDAO.saveCheckPaymentUserDetails(checkDetails);
						} 

						try {
							mailer.sendMailToRenewedUser(userObj.getJctUserEmail(), cal.getTime(), paymentType);						
						} catch (MailingException e) {
							LOGGER.error("FACILITATOR REGISTERED....COULD NOT SEND EMAIL...  REASON: "+e.getLocalizedMessage());								
						}
						//	update payment dtls for payment amount 
						try {
							double costPerUser = Double.parseDouble(accountVO.getTotalAmountToBePaid()) / Double.parseDouble(accountVO.getNumberOfUsers());
							serviceDAO.updateAmountInPaymentHdr(registeredPaymentHdrId.toString(), costPerUser);
						} catch (DAOException ex) {
							LOGGER.error("ERROR: UNABLE TO UPDATE USER PAYMENT AMOUNT: "+ex.getLocalizedMessage());			
						}
					}			
			} catch (DAOException ex) {
				LOGGER.error("ERROR: UNABLE TO CREATE NEW USER. REASON: "+ex.getLocalizedMessage());
				throw new JCTException("ERROR: UNABLE TO CREATE NEW USER. REASON: "+ex.getLocalizedMessage());
			}
		}	
		LOGGER.info("<<<<<<<< ManageUserServiceImpl.saveGeneralUserViaChequePayment");		
		
		return sb.toString();
		}
	
	@Transactional(propagation = Propagation.REQUIRED)
	public List<String> getInactiveUserListByEmail(
			List<String> emailIdList) throws JCTException {
		LOGGER.info(">>>>>>>> ManageUserServiceImpl.getInactiveUserListByEmail");
		List<String> usernameList = new ArrayList<String>();
		try{
			usernameList = serviceDAO.getInactiveUserListByEmail(emailIdList); 	
		} catch(DAOException ex){
			LOGGER.error(ex.getLocalizedMessage());			
		} 
		LOGGER.info("<<<<<<<< ManageUserServiceImpl.getInactiveUserListByEmail");		
		return usernameList;	
	}
	
	@Transactional(propagation = Propagation.REQUIRED)
	public boolean isIndividualUserExist(java.lang.String emailId)
			throws JCTException {
		boolean userExists = false;
		try {
			List<String> userList = authenticatorDAO.isIndividualUserExist(emailId);
			if (userList.size() > 0) {
				userExists = true;
			}
		} catch (DAOException dao) {
			LOGGER.error(dao.getLocalizedMessage());
		}
		return userExists;
	}
	
	@Transactional(propagation = Propagation.REQUIRED)
	public boolean shopifyUserChecker (String emailId, String userType) 
			throws JCTException {
		boolean userExists = false;
		try {
			List<String> userList = authenticatorDAO.shopifyUserChecker(emailId, Integer.parseInt(userType));
			if (userList.size() > 0) {
				userExists = true;
			}
		} catch (DAOException dao) {
			LOGGER.error(dao.getLocalizedMessage());
		}
		return userExists;
	}
	
	@Transactional(propagation = Propagation.REQUIRED)
	public boolean isIndividualActiveUserExist(java.lang.String emailId)
			throws JCTException {
		boolean userExists = false;
		try {
			List<String> userList = authenticatorDAO.isIndividualActiveUserExist(emailId);
			if (userList.size() > 0) {
				userExists = true;
			}
		} catch (DAOException dao) {
			LOGGER.error(dao.getLocalizedMessage());
		}
		return userExists;
	}
	
	@Transactional(propagation = Propagation.REQUIRED)
	public boolean isIndividualInactiveUserExist(java.lang.String emailId)
			throws JCTException {
		boolean userExists = false;
		try {
			List<String> userList = authenticatorDAO.isIndividualInactiveUserExist(emailId);
			if (userList.size() > 0) {
				userExists = true;
			}
		} catch (DAOException dao) {
			LOGGER.error(dao.getLocalizedMessage());
		}
		return userExists;
	}

	@Transactional(propagation = Propagation.REQUIRED)
	public List<UserAccountDTO> getUserEndDateDTO(java.lang.String emailId) throws JCTException {
		List<UserAccountDTO> dto = null;
		try {
			dto = authenticatorDAO.getUserEndDateDTO(emailId);
		} catch (DAOException dao) {
			LOGGER.error(dao.getLocalizedMessage());
		}
		return dto;
	}

	@Transactional(propagation = Propagation.REQUIRED)
	public JctUser getPassword(String emailId, int role) throws JCTException {
		LOGGER.info(">>>>>>>> ManageUserServiceImpl.getInactiveUserListByEmail");
		JctUser userObj = new JctUser();
		try{
			userObj = serviceDAO.getPasswordObject(emailId, role); 	
		} catch(DAOException ex){
			LOGGER.error(ex.getLocalizedMessage());			
		} 
		LOGGER.info("<<<<<<<< ManageUserServiceImpl.getInactiveUserListByEmail");		
		return userObj;	
	}
	@Transactional(propagation = Propagation.REQUIRED)
	public String getFirstName(String emailId, int role){
		LOGGER.info(">>>>>>>> ManageUserServiceImpl.getFirstName");
		String fName = "";
		try {
			fName = serviceDAO.getFirstName(emailId, role).toLowerCase();
		} catch (JCTException e) {
			LOGGER.error(e.getLocalizedMessage());	
		}
		LOGGER.info("<<<<<<<< ManageUserServiceImpl.getFirstName");		
		return WordUtils.capitalize(fName);	
	}
	@Transactional(propagation = Propagation.REQUIRED)
	public ExistingUserDTO getUserExpirationDtls(String emailId, int role) {
		LOGGER.info(">>>>>>> ManageUserServiceImpl.getUserExpirationDtls");
		ExistingUserDTO userObj = null;		
		try {
			userObj = serviceDAO.getUserExpirationDtls(emailId, role);			
		} catch (JCTException ex) {
			LOGGER.error(ex.getLocalizedMessage());	
		} 
		LOGGER.info("<<<<<<<< ManageUserServiceImpl.getUserExpirationDtls");		
		return userObj;
	}
	@Transactional(propagation = Propagation.REQUIRED)
	public FacilitatorDetailDTO getFaciExpirationDtls(String custId, int role) throws JCTException {
		LOGGER.info(">>>>>>> ManageUserServiceImpl.getFaciExpirationDtls");
		FacilitatorDetailDTO userObj = null;		
		try {
			userObj = serviceDAO.getFaciExpirationDtls(custId);
		} catch (JCTException ex) {
			LOGGER.error(ex.getLocalizedMessage());	
		} 
		LOGGER.info("<<<<<<<< ManageUserServiceImpl.getFaciExpirationDtls");		
		return userObj;
	}
	@Transactional(propagation = Propagation.REQUIRED)
	public String getFaciName(String userName)
			throws JCTException {
		LOGGER.info(">>>>>>>> ManageUserServiceImpl.getFaciExpirationDtls");
		JctUser userObj = new JctUser();
		StringBuffer fullName = new StringBuffer();
		try {
			userObj = serviceDAO.getFullName(userName);
			fullName.append(WordUtils.capitalize(userObj.getJctUserDetails().getJctUserDetailsFirstName().toLowerCase()));
			fullName.append(" ");
			fullName.append(WordUtils.capitalize(userObj.getJctUserDetails().getJctUserDetailsLastName().toLowerCase()));			
		} catch (JCTException ex) {
			LOGGER.error(ex.getLocalizedMessage());	

		}		
		LOGGER.info("<<<<<<<< ManageUserServiceImpl.getFaciExpirationDtls");
		return fullName.toString();
	}

	public String forgotPassword(String emailId) throws JCTException {
		LOGGER.info(">>>> ManageUserServiceImpl.forgotPassword");
		String rtn = "failure";
		JctUser userObj = new JctUser();
		String generatedPassword = this.generatePassword();
		String encryptedPassword = this.encriptString(generatedPassword);
		
		try {
			userObj = (JctUser) serviceDAO.getFullName(emailId);
			if( null != userObj ){
				userObj.setJctPassword(encryptedPassword);	
				userObj.setJctActiveYn(CommonConstants.FACI_PASSWORD_RESET);
				rtn = serviceDAO.updateUser(userObj);
				if(rtn.equalsIgnoreCase("success")){
					try {
						mailer.intimateFaciPassword(emailId,generatedPassword,
								WordUtils.capitalize(userObj.getJctUserDetails().getJctUserDetailsFirstName()));
					} catch (MailingException e) {
						LOGGER.error("FACILITATOR REGISTERED....COULD NOT SEND EMAIL...  REASON: "+e.getLocalizedMessage());
						rtn = StatusConstants.USER_REGISTERED_MAIL_SEND_ERROR.toString();
					}
					
				}
				
			} else {
				rtn = CommonConstants.USER_NOT_FOUND;
			}
		} catch (DAOException e) {
			LOGGER.error(e.getLocalizedMessage());
		}		
		LOGGER.info("<<<< ManageUserServiceImpl.forgotPassword");
        return rtn;
	} 
}
