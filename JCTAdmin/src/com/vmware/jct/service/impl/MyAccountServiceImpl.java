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
import com.vmware.jct.dao.IMyAccountServiceDAO;
import com.vmware.jct.exception.DAOException;
import com.vmware.jct.exception.JCTException;
import com.vmware.jct.model.JctUser;
import com.vmware.jct.model.JctUserDetails;
import com.vmware.jct.service.IMyAccountService;
import com.vmware.jct.service.vo.MyAccountVO;
/**
 * 
 * <p><b>Class name:</b> MyAccountServiceImpl.java</p>
 * <p><b>Author:</b> InterraIT</p>
 * <p><b>Purpose:</b> This is a MyAccountServiceImpl class. This artifact is Business layer artifact.
 * MyAccountServiceImpl implement IMyAccountService interface and override the following  methods.
 * -populateExistingUserData()
 * -updateFirstName()
 * -updateLastName()
 * -updateEmailId()
 * -updatePassword()
 * </p>
 * <p><b>Description:</b> This class used to perform insert, modify ,fetch data from db and delete operation. Save, update, delete operations are performed in DataAccessObject class </p>
 * <p><b>Revision History:</b>
 * 	<li>InterraIT - 19/Jan/2014 - Implement Exception </li>
 * <li>InterraIT - 31/Jan/2014 - Implement comments, introduce Named query, Transactional attribute </li>
 * </p>
 */
@Service
public class MyAccountServiceImpl implements IMyAccountService{
	
	@Autowired
	private IMyAccountServiceDAO serviceDAO;
	
	@Autowired  
	private MessageSource messageSource;
	
	@Autowired
	private ICommonDao commonDaoImpl;

	private static final Logger LOGGER = LoggerFactory.getLogger(MyAccountServiceImpl.class);
	
	@Transactional(propagation = Propagation.REQUIRED)
	public MyAccountVO populateExistingUserData() throws JCTException {
		LOGGER.info(">>>>>> MyAccountServiceImpl.populateExistingUserData");
		StringBuilder builder = new StringBuilder("");
		MyAccountVO vo = new MyAccountVO();
		try {
			JctUser user = serviceDAO.populateExistingUserData();
			if (null != user) {		
				if(null == user.getJctUserDetails().getJctUserDetailsFirstName() || user.getJctUserDetails().getJctUserDetailsFirstName() == ""){
					builder.append("NA");
				} else {
					builder.append(user.getJctUserDetails().getJctUserDetailsFirstName());
				}				
				builder.append("~");
				if(null == user.getJctUserDetails().getJctUserDetailsLastName() || user.getJctUserDetails().getJctUserDetailsLastName() == ""){
					builder.append("NA");
				} else {
					builder.append(user.getJctUserDetails().getJctUserDetailsLastName());
				}	
				builder.append("~");
				builder.append(user.getJctPassword());
				builder.append("~");
				builder.append(user.getJctUserEmail());
				builder.append("~");
				builder.append(user.getJctUserId());
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
		LOGGER.info("<<<<<< MyAccountServiceImpl.populateExistingUserData");
		return vo;
	}
	
	@Transactional(propagation = Propagation.REQUIRED)
	public MyAccountVO updateFirstName(String valueDesc, Integer primaryKey) throws JCTException {
		LOGGER.info(">>>>>> MyAccountServiceImpl.updateFirstName");
		MyAccountVO myAccountVO = new MyAccountVO();
		try {			
			JctUser user = serviceDAO.fetchUserData(primaryKey);					
			user.setJctUserSoftDelete(1);
			user.getJctUserDetails().setJctUserDetailsSoftDelete(1);
			try {
				// Create User Object
				JctUser userNew = new JctUser();
				userNew.setJctUserId(commonDaoImpl.generateKey("jct_user_id"));
				userNew.setJctActiveYn(user.getJctActiveYn());
				userNew.setJctCreatedBy(user.getJctCreatedBy());
				userNew.setJctCreatedTs(user.getJctCreatedTs());
				userNew.setJctLastmodifiedBy(user.getJctLastmodifiedBy());
				userNew.setJctPassword(user.getJctPassword());
				userNew.setJctUserEmail(user.getJctUserEmail());
				userNew.setJctUserRole(user.getJctUserRole());
				userNew.setJctVersion(user.getJctVersion());
				userNew.setLastmodifiedTs(new Date());
				userNew.setJctUserCustomerId(user.getJctUserCustomerId());
				userNew.setJctAccountExpirationDate(user.getJctAccountExpirationDate());
				userNew.setJctUserName(user.getJctUserName());
				
				// Create User Details Object
				JctUserDetails userDetails = new JctUserDetails();
				userDetails.setJctUserDetailsFirstName(valueDesc);
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
				String retString = serviceDAO.saveUser(userNew);
				
				serviceDAO.updateFirstName(user);
				myAccountVO.setUserFirstName(userNew.getJctUserDetails().getJctUserDetailsFirstName());
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
		LOGGER.info("<<<<<< MyAccountServiceImpl.updateFirstName");
		return myAccountVO;
	}
	
	@Transactional(propagation = Propagation.REQUIRED)
	public MyAccountVO updateLastName(String valueDesc, Integer primaryKey) throws JCTException {
		LOGGER.info(">>>>>> MyAccountServiceImpl.updateLastName");
		MyAccountVO myAccountVO = new MyAccountVO();
		try {			
			JctUser user = serviceDAO.fetchUserData(primaryKey);					
			user.setJctUserSoftDelete(1);
			user.getJctUserDetails().setJctUserDetailsSoftDelete(1);
			try {
				// Create User Object
				JctUser userNew = new JctUser();
				userNew.setJctUserId(commonDaoImpl.generateKey("jct_user_id"));
				userNew.setJctActiveYn(user.getJctActiveYn());
				userNew.setJctCreatedBy(user.getJctCreatedBy());
				userNew.setJctCreatedTs(user.getJctCreatedTs());
				userNew.setJctLastmodifiedBy(user.getJctLastmodifiedBy());
				userNew.setJctPassword(user.getJctPassword());
				userNew.setJctUserEmail(user.getJctUserEmail());
				userNew.setJctUserRole(user.getJctUserRole());
				userNew.setJctVersion(user.getJctVersion());
				userNew.setLastmodifiedTs(new Date());
				userNew.setJctUserCustomerId(user.getJctUserCustomerId());
				userNew.setJctAccountExpirationDate(user.getJctAccountExpirationDate());
				userNew.setJctUserName(user.getJctUserName());
				
				// Create User Details Object
				JctUserDetails userDetails = new JctUserDetails();
				userDetails.setJctUserDetailsFirstName(user.getJctUserDetails().getJctUserDetailsFirstName());
				userDetails.setJctUserDetailsLastName(valueDesc);
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
				String retString = serviceDAO.saveUser(userNew);
				
				serviceDAO.updateLastName(user);
				myAccountVO.setUserLastName(user.getJctUserDetails().getJctUserDetailsLastName());
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
		LOGGER.info("<<<<<< MyAccountServiceImpl.updateLastName");
		return myAccountVO;
	}
	
	@Transactional(propagation = Propagation.REQUIRED)
	public MyAccountVO updateEmailId(String valueDesc, Integer primaryKey) throws JCTException {
		LOGGER.info(">>>>>> MyAccountServiceImpl.updateEmailId");
		MyAccountVO myAccountVO = new MyAccountVO();
		try {			
			JctUser user = serviceDAO.fetchUserData(primaryKey);					
			user.setJctUserSoftDelete(1);
			user.getJctUserDetails().setJctUserDetailsSoftDelete(1);
			try {
				// Create User Object
				JctUser userNew = new JctUser();
				userNew.setJctUserId(commonDaoImpl.generateKey("jct_user_id"));
				userNew.setJctActiveYn(user.getJctActiveYn());
				userNew.setJctCreatedBy(user.getJctCreatedBy());
				userNew.setJctCreatedTs(user.getJctCreatedTs());
				userNew.setJctLastmodifiedBy(user.getJctLastmodifiedBy());
				userNew.setJctPassword(user.getJctPassword());
				userNew.setJctUserEmail(valueDesc);
				userNew.setJctUserRole(user.getJctUserRole());
				userNew.setJctVersion(user.getJctVersion());
				userNew.setLastmodifiedTs(new Date());
				userNew.setJctUserCustomerId(user.getJctUserCustomerId());
				userNew.setJctAccountExpirationDate(user.getJctAccountExpirationDate());
				userNew.setJctUserName(user.getJctUserName());
				
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
				String retString = serviceDAO.saveUser(userNew);
				serviceDAO.updateEmailId(user);
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
		LOGGER.info("<<<<<< MyAccountServiceImpl.updateEmailId");
		return myAccountVO;
	}
	
	@Transactional(propagation = Propagation.REQUIRED)
	public MyAccountVO updatePassword(String valueDesc, Integer primaryKey) throws JCTException {
		LOGGER.info(">>>>>> MyAccountServiceImpl.updatePassword");
		MyAccountVO myAccountVO = new MyAccountVO();
		try {			
			JctUser user = serviceDAO.fetchUserData(primaryKey);					
			user.setJctUserSoftDelete(1);
			user.getJctUserDetails().setJctUserDetailsSoftDelete(1);
			try {
				// Create User Object
				JctUser userNew = new JctUser();
				userNew.setJctUserId(commonDaoImpl.generateKey("jct_user_id"));
				userNew.setJctActiveYn(user.getJctActiveYn());
				userNew.setJctCreatedBy(user.getJctCreatedBy());
				userNew.setJctCreatedTs(user.getJctCreatedTs());
				userNew.setJctLastmodifiedBy(user.getJctLastmodifiedBy());
				userNew.setJctPassword(valueDesc);
				userNew.setJctUserEmail(user.getJctUserEmail());
				userNew.setJctUserRole(user.getJctUserRole());
				userNew.setJctVersion(user.getJctVersion());
				userNew.setLastmodifiedTs(new Date());
				userNew.setJctUserCustomerId(user.getJctUserCustomerId());
				userNew.setJctAccountExpirationDate(user.getJctAccountExpirationDate());
				userNew.setJctUserName(user.getJctUserName());
				
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
				String retString = serviceDAO.saveUser(userNew);
				
				serviceDAO.updatePassword(user);
				myAccountVO.setUserPassword(userNew.getJctPassword());
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
		LOGGER.info("<<<<<< MyAccountServiceImpl.updatePassword");
		return myAccountVO;
	}
}