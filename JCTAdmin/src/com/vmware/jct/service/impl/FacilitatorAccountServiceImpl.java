package com.vmware.jct.service.impl;

import java.util.Date;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.context.MessageSource;
import org.springframework.stereotype.Service;
import org.springframework.transaction.annotation.Propagation;
import org.springframework.transaction.annotation.Transactional;

import com.vmware.jct.common.utility.StatusConstants;
import com.vmware.jct.dao.ICommonDao;
import com.vmware.jct.dao.IFacilitatorAccountServiceDAO;
import com.vmware.jct.exception.DAOException;
import com.vmware.jct.exception.JCTException;
import com.vmware.jct.model.JctUser;
import com.vmware.jct.model.JctUserDetails;
import com.vmware.jct.service.IFacilitatorAccountService;
import com.vmware.jct.service.vo.MyAccountVO;
/**
 * 
 * <p><b>Class name:</b> FacilitatorAccountServiceImpl.java</p>
 * <p><b>Author:</b> InterraIT</p>
 * <p><b>Purpose:</b> This is a FacilitatorAccountServiceImpl class. This artifact is Business layer artifact.
 * FacilitatorAccountServiceImpl implement IFacilitatorAccountService interface and override the following  methods.
 * -populateExistingFacilitatorData()
 * -updateUserNameFacilitator()
 * -updateEmailIdFacilitator()
 * -updatePasswordFacilitator()
 * -prepareUserObject()
 * </p>
 * <p><b>Description:</b> This class used to perform insert, modify ,fetch data from db and delete operation. Save, update, delete operations are performed in DataAccessObject class </p>
 * <p><b>Revision History:</b>
 * 	<li>InterraIT - 01/Oct/2014 - Implement Exception </li>
 * <li>InterraIT - 31/Oct/2014 - Implement comments, introduce Named query, Transactional attribute </li>
 * </p>
 */
@Service
public class FacilitatorAccountServiceImpl implements IFacilitatorAccountService {
	
	@Autowired
	private IFacilitatorAccountServiceDAO serviceDAO;
	
	@Autowired  
	private MessageSource messageSource;
	
	@Autowired
	private ICommonDao commonDaoImpl;

	private static final Logger LOGGER = LoggerFactory.getLogger(FacilitatorAccountServiceImpl.class);
	
	/**
	 * Method populates facilitator data based on email
	 * @param facilitatorEmail
	 * @return MyAccountVO
	 */
	@Transactional(propagation = Propagation.REQUIRED)
	public MyAccountVO populateExistingFacilitatorData(String facilitatorEmail) throws JCTException {
		LOGGER.info(">>>>>> FacilitatorAccountServiceImpl.populateExistingFacilitatorData");
		StringBuilder builder = new StringBuilder("");
		MyAccountVO vo = new MyAccountVO();
		try {
			JctUser user = serviceDAO.populateExistingFacilitatorData(facilitatorEmail);
			if (null != user) {						
				builder.append(user.getJctUserName());
				builder.append("~");
				builder.append(user.getJctPassword());
				builder.append("~");
				builder.append(user.getJctUserEmail());
				builder.append("~");
				builder.append(user.getJctUserId());
				builder.append("~");
				builder.append(user.getJctUserCustomerId());
				builder.append("~");
				vo.setExistingUserDataList(builder.toString());
				vo.setStatusCode(StatusConstants.STATUS_SUCCESS_WITH_RESULT);
			} else {
				LOGGER.info("---- NO RECORDS FOUND -----");
				vo.setStatusCode(StatusConstants.STATUS_SUCCESS_WITH_NO_RESULT);
				vo.setStatusDesc(this.messageSource.getMessage("waring.no.user.profile", null, null));
			}
		} catch (DAOException e) {
			LOGGER.error("----"+e.getLocalizedMessage()+" ----");
			throw new JCTException(e.getLocalizedMessage());
		}
		LOGGER.info("<<<<<< FacilitatorAccountServiceImpl.populateExistingFacilitatorData");
		return vo;
	}
	
	/**
	 * Method updates the user name of the facilitator
	 * @param valueDesc
	 * @param tablePkId
	 * @return MyAccountVO
	 */
	@Transactional(propagation = Propagation.REQUIRED)
	public MyAccountVO updateUserNameFacilitator(String valueDesc, Integer primaryKey, String facilitatorEmail) throws JCTException {
		LOGGER.info(">>>>>> FacilitatorAccountServiceImpl.updateUserNameFacilitator");
		MyAccountVO myAccountVO = new MyAccountVO();
		try {			
			JctUser user = serviceDAO.fetchUserData(primaryKey);					
			user.setJctUserSoftDelete(1);
			user.setJctUserDispIdentifier(1);
			user.getJctUserDetails().setJctUserDetailsSoftDelete(1);
			try {
				JctUser userNew = new JctUser();
				String retString  = prepareUserObject(user, userNew, valueDesc, "USERNAME" ,facilitatorEmail);

				serviceDAO.updateUserNameFacilitator(user);
				myAccountVO.setUserName(userNew.getJctUserName());
				myAccountVO.setPrimaryKey(userNew.getJctUserId());
				if (!retString.equals("success")) {					
					myAccountVO.setStatusCode(StatusConstants.STATUS_FAILURE);
					myAccountVO.setStatusDesc(this.messageSource.
							getMessage("exception.dao.user.name.update", null, null));
				} else {
					myAccountVO.setStatusCode(StatusConstants.STATUS_SUCCESS_WITH_RESULT);
					myAccountVO.setStatusDesc(this.messageSource.
							getMessage("success.msg.user.name.update", null, null));
				}
			} catch (Exception e){
				LOGGER.error("----"+e.getLocalizedMessage()+" ----");
			}
		} catch (DAOException e) {
			LOGGER.error("----"+e.getLocalizedMessage()+" ----");
			myAccountVO.setStatusCode(StatusConstants.STATUS_FAILURE);
			throw new JCTException(e.getLocalizedMessage());
		}
		LOGGER.info("<<<<<< FacilitatorAccountServiceImpl.updateUserNameFacilitator");
		return myAccountVO;
	}
	
	/**
	 * Method updates the email address of the facilitator
	 * @param valueDesc
	 * @param primaryKey
	 * @return MyAccountVO
	 */
	@Transactional(propagation = Propagation.REQUIRED)
	public MyAccountVO updateEmailIdFacilitator(String valueDesc, Integer primaryKey, String facilitatorEmail) throws JCTException {
		LOGGER.info(">>>>>> FacilitatorAccountServiceImpl.updateEmailIdFacilitator");
		MyAccountVO myAccountVO = new MyAccountVO();
		try {			
			JctUser user = serviceDAO.fetchUserData(primaryKey);					
			user.setJctUserSoftDelete(1);
			user.setJctUserDispIdentifier(1);
			user.getJctUserDetails().setJctUserDetailsSoftDelete(1);
			try {
				// Create User Object
				JctUser userNew = new JctUser();				
				String retString  = prepareUserObject(user, userNew, valueDesc, "EMAIL", facilitatorEmail);

				serviceDAO.updateEmailIdFacilitator(user);
				myAccountVO.setUserEmailId(userNew.getJctUserEmail());
				myAccountVO.setPrimaryKey(userNew.getJctUserId());
				if (!retString.equals("success")) {					
					myAccountVO.setStatusCode(StatusConstants.STATUS_FAILURE);
					myAccountVO.setStatusDesc(this.messageSource.
							getMessage("exception.dao.user.group.save", null, null));
				} else {
					myAccountVO.setStatusCode(StatusConstants.STATUS_SUCCESS_WITH_RESULT);
					myAccountVO.setStatusDesc(retString);
				}
			} catch (Exception e){
				LOGGER.error("----"+e.getLocalizedMessage()+" ----");
			}
		} catch (DAOException e) {
			LOGGER.error("----"+e.getLocalizedMessage()+" ----");
			myAccountVO.setStatusCode(StatusConstants.STATUS_FAILURE);
			throw new JCTException(e.getLocalizedMessage());
		}
		LOGGER.info("<<<<<< FacilitatorAccountServiceImpl.updateEmailIdFacilitator");
		return myAccountVO;
	}
	
	/**
	 * Method updates the password of the facilitator
	 * @param valueDesc
	 * @param tablePkId
	 * @return MyAccountVO
	 */
	@Transactional(propagation = Propagation.REQUIRED)
	public MyAccountVO updatePasswordFacilitator(String valueDesc, Integer primaryKey, String facilitatorEmail) throws JCTException {
		LOGGER.info(">>>>>> FacilitatorAccountServiceImpl.updatePasswordFacilitator");
		MyAccountVO myAccountVO = new MyAccountVO();
		try {			
			JctUser user = serviceDAO.fetchUserData(primaryKey);					
			user.setJctUserSoftDelete(1);
			user.setJctUserDispIdentifier(1);
			user.getJctUserDetails().setJctUserDetailsSoftDelete(1);
			try {
				// Create User Object
				JctUser userNew = new JctUser();
				String retString  = prepareUserObject(user, userNew, valueDesc, "PASSWORD", facilitatorEmail );
				
				serviceDAO.updatePasswordFacilitator(user);
				myAccountVO.setUserPassword(userNew.getJctPassword());
				myAccountVO.setPrimaryKey(userNew.getJctUserId());
				if (!retString.equals("success")) {					
					myAccountVO.setStatusCode(StatusConstants.STATUS_FAILURE);
					myAccountVO.setStatusDesc(this.messageSource.
							getMessage("exception.dao.password.update", null, null));
				} else {
					myAccountVO.setStatusCode(StatusConstants.STATUS_SUCCESS_WITH_RESULT);
					myAccountVO.setStatusDesc(this.messageSource.
							getMessage("success.msg.password.update", null, null));
				}
			} catch (Exception e){
				LOGGER.error("----"+e.getLocalizedMessage()+" ----");
			}
		} catch (DAOException e) {
			LOGGER.error("----"+e.getLocalizedMessage()+" ----");
			myAccountVO.setStatusCode(StatusConstants.STATUS_FAILURE);
			throw new JCTException(e.getLocalizedMessage());
		}
		LOGGER.info("<<<<<< FacilitatorAccountServiceImpl.updatePasswordFacilitator");
		return myAccountVO;
	}
	
	/**
	 * Method to prepare user object
	 * @param user
	 * @param userNew
	 * @param valueDesc
	 * @param action
	 * @param facilitatorEmail
	 * @return status
	 * @throws DAOException
	 */
	private String prepareUserObject(JctUser user, JctUser userNew, String valueDesc,
			String action, String facilitatorEmail) throws DAOException {
		LOGGER.info(">>>>>> FacilitatorAccountServiceImpl.prepareUserObject");
		String retString = "";
		userNew.setJctUserId(commonDaoImpl.generateKey("jct_user_id"));
		if(action.equals("USERNAME")){
			userNew.setJctUserName(valueDesc);
			//userNew.setJctUserEmail(user.getJctUserEmail());
			userNew.setJctUserEmail(valueDesc);
			userNew.setJctPassword(user.getJctPassword());
		} else if (action.equals("EMAIL")) {
			userNew.setJctUserEmail(valueDesc);
			userNew.setJctUserName(user.getJctUserName());	
			userNew.setJctPassword(user.getJctPassword());
		} else {
			userNew.setJctPassword(valueDesc);
			userNew.setJctUserEmail(user.getJctUserEmail());
			userNew.setJctUserName(user.getJctUserName());
		}
		
		userNew.setJctActiveYn(user.getJctActiveYn());
		userNew.setJctCreatedBy(user.getJctCreatedBy());
		userNew.setJctCreatedTs(user.getJctCreatedTs());
		userNew.setJctLastmodifiedBy(facilitatorEmail);				
		userNew.setJctUserRole(user.getJctUserRole());
		userNew.setJctVersion(user.getJctVersion());
		userNew.setLastmodifiedTs(new Date());
		userNew.setJctUserCustomerId(user.getJctUserCustomerId());
		userNew.setJctAccountExpirationDate(user.getJctAccountExpirationDate());
		
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
		retString = serviceDAO.saveFacilitator(userNew);		
		LOGGER.info("<<<<<<< FacilitatorAccountServiceImpl.prepareUserObject");
		return retString;
	}

	/**
	 * Method search the user id is exist in database or not
	 * @param facilitatorEmail
	 * @return status
	 */
	@Transactional(propagation = Propagation.REQUIRED)
	public String searchUserIdInDB(String facilitatorEmail) throws JCTException {
		LOGGER.info(">>>>>> FacilitatorAccountServiceImpl.searchUserIdInDB");
		String statusMessage = "failure";
		try {
			statusMessage = serviceDAO.searchUserIdInDB(facilitatorEmail);
		} catch (DAOException e) {
			LOGGER.error("UNABLE TO FIND USER : "+facilitatorEmail);
			statusMessage = "failure";
		}
		LOGGER.info("<<<<<< FacilitatorAccountServiceImpl.searchUserIdInDB");
		return statusMessage;
	}

}