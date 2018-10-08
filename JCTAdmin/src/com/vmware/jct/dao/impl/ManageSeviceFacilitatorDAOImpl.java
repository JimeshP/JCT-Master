package com.vmware.jct.dao.impl;

import java.util.Date;
import java.util.List;

import org.hibernate.SessionFactory;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Repository;
import org.springframework.transaction.annotation.Propagation;
import org.springframework.transaction.annotation.Transactional;

import com.vmware.jct.dao.DataAccessObject;
import com.vmware.jct.dao.IManageServiceFacilitatorDAO;
import com.vmware.jct.dao.dto.ExistingUserDTO;
import com.vmware.jct.dao.dto.FacilitatorDetailDTO;
import com.vmware.jct.dao.dto.UserDTO;
import com.vmware.jct.dao.dto.UserGroupDTO;
import com.vmware.jct.exception.DAOException;
import com.vmware.jct.model.JctEmailDetails;
import com.vmware.jct.model.JctFacilitatorDetails;
import com.vmware.jct.model.JctUser;
import com.vmware.jct.model.JctUserGroup;
import com.vmware.jct.service.vo.UserVO;
/**
 * 
 * <p><b>Class name:</b> ManageSeviceFacilitatorDAOImpl.java</p>
 * <p><b>Author:</b> InterraIT</p>
 * <p><b>Purpose:</b> This is a ManageSeviceFacilitatorDAOImpl class. This artifact is Persistence Manager layer artifact.
 * ManageSeviceFacilitatorDAOImpl implement IManageServiceFacilitatorDAO interface and override the following  methods.
 * -populateSubscribedUser()
 * -cronPersister()
 * -fetchUserGroupId()
 * -fetchFacilitatorId()
 * -updateFacilitatorDetails()
 * -populateUserList()
 * -populateActiveInactiveUserListForFacilitator()
 * -populateUserListResetPasswordForFacilitator()
 * -updateUserStatusInBatchFacilitator()
 * -checkUserToResetPassword()
 * -updateUserToResetPasswordFacilitator()
 * -getEmailObjectFacilitator()
 * -cronUpdaterFacilitator()
 * -renewUserByFacilitator()
 * -fetchUserExpiryDate()
 * -populateUser()
 * -populateFacilitatorList()
 * -fetchUserData()
 * -saveFacilitator()
 * -updateUser()
 * -fetchFacilitatorData()
 * -searchByEmail()
 * </p>
 * <p><b>Description:</b> This class used to perform insert, modify ,fetch data from db and delete operation. Save, update, delete operations are performed in DataAccessObject class </p>
 * <p><b>Revision History:</b>
 * 	<li>InterraIT - 01/Oct/2014 - Implement Exception </li>
 * <li>InterraIT -  18/Oct/2014 - Implement comments, introduce Named query, Transactional attribute </li>
 * <li>InterraIT - 31/Oct/2014 - Updated comments </li>
 * </p>
 */
@Repository
public class ManageSeviceFacilitatorDAOImpl extends DataAccessObject implements IManageServiceFacilitatorDAO {

	@Autowired
	private SessionFactory sessionFactory;
	
	private static final Logger LOGGER = LoggerFactory.getLogger(ManageSeviceFacilitatorDAOImpl.class);
	
	/**
	 * Method populates the Total Subscribed User and Registered User
	 * @param emailAddress
	 * @param type
	 * @param customerId
	 */
	@Transactional(propagation = Propagation.REQUIRED)
	public List<FacilitatorDetailDTO> populateSubscribedUser(String emailAddress, String type, String customerId)
			throws DAOException {
		LOGGER.info(">>>>>> ManageSeviceFacilitatorDAOImpl.populateSubscribedUser");
		List<FacilitatorDetailDTO> list = null;
		try {
			list = sessionFactory.getCurrentSession().getNamedQuery("fetchSubscribedUserADD")
					.setParameter("customerId", customerId).list();	
		} catch (Exception e) {
			list = null;
		}
		LOGGER.info("<<<<<< ManageSeviceFacilitatorDAOImpl.populateSubscribedUser");
		return list;
	}
	
	/**
	 * Method persists JctEmailDetails by the cron service
	 * @param dtls
	 * @return status
	 */
	@Transactional(propagation=Propagation.REQUIRED)
	public String cronPersister(JctEmailDetails dtls) throws DAOException {
		return save(dtls);
	}

	/**
	 * Method fetches user group Id based on user group name
	 * @param userGrpName
	 * @return JctUserGroup entity
	 */
	@Transactional(propagation = Propagation.REQUIRED)
	public JctUserGroup fetchUserGroupId(int userGrpId, String customerId, String userGrpName) throws DAOException {
		LOGGER.info(">>>>>> ManageSeviceFacilitatorDAOImpl.fetchUserGroupId");
		JctUserGroup userGrpObj = null;
		if(userGrpId != 0) {
			userGrpObj = (JctUserGroup) sessionFactory.getCurrentSession().getNamedQuery("fetchUserGroupId")
					.setParameter("userGrpId", userGrpId)
					.setParameter("customerId", customerId).list().get(0);
		} else {
			userGrpObj = (JctUserGroup) sessionFactory.getCurrentSession().getNamedQuery("fetchUserGroupByGrpName")
					.setParameter("userGrpDesc", userGrpName)
					.setParameter("customerId", customerId).list().get(0);
		}
		return userGrpObj;
	}

	/**
	 * Method fetches facilitator ID based on facilitator email address
	 * @param emailAddress
	 * @return facilitatorId
	 */
	@Transactional(propagation = Propagation.REQUIRED)
	public Integer fetchFacilitatorId(String emailAddress) throws DAOException {
		LOGGER.info(">>>>>> ManageSeviceFacilitatorDAOImpl.fetchFacilitatorId");
		return (Integer) sessionFactory.getCurrentSession().getNamedQuery("fetchFacilitatorID")
				.setParameter("role", 3).setParameter("emailAddress", emailAddress).uniqueResult();
	}

	/**
	 * Method to update the number of registered user based on number of user added
	 * @param userCount
	 * @param customerId
	 * @param type
	 * @return status
	 */
	@Transactional(propagation = Propagation.REQUIRED)
	public String updateFacilitatorDetails(int userCount, String customerId, String type) throws DAOException {
		LOGGER.info(">>>>>> ManageSeviceFacilitatorDAOImpl.updateFacilitatorDetails");
		String statusDesc = "";
		String newType = "";
		//JctFacilitatorDetails facilitatorDetailObj = null;
		List<JctFacilitatorDetails> facilitatorDetailObj = null;
		if(type.equalsIgnoreCase("AD")) {
			newType = "RW";
		} else {
			newType = "AD";
		}
		for(int i = 0; i < userCount; i++) {
			try{
				 facilitatorDetailObj = sessionFactory.getCurrentSession().getNamedQuery("fetchFacilitatorObj")
						.setParameter("customerId", customerId)
						.setParameter("type", type).list();

				 if(facilitatorDetailObj.size() == 0) {
					 facilitatorDetailObj = sessionFactory.getCurrentSession().getNamedQuery("fetchFacilitatorObj")
								.setParameter("customerId", customerId)
								.setParameter("type", newType).list();
				 }
				 
				/*if(facilitatorDetailObj.getJctFacTotalLimit() == facilitatorDetailObj.getJctFacSubscribeLimit()) {
				  facilitatorDetailObj = (JctFacilitatorDetails) sessionFactory.getCurrentSession().getNamedQuery("fetchFacilitatorObj")
							.setParameter("customerId", customerId)
							.setParameter("type", newType).list().get(0);
				}*/
				facilitatorDetailObj.get(0).setJctFacSubscribeLimit(facilitatorDetailObj.get(0).getJctFacSubscribeLimit() + 1);			
				update(facilitatorDetailObj.get(0));
				statusDesc = "Success";
			} catch (Exception e) {
				statusDesc = "failure";
				LOGGER.error("UNABLE TO UPDATE FACILITATOR DETAILS TABLE");
				LOGGER.error(e.getLocalizedMessage());
			}		
		}	
		LOGGER.info("<<<<<< ManageSeviceFacilitatorDAOImpl.updateFacilitatorDetails");
		return statusDesc;
	}

	/**
	 * Method returns existing active and deactivate user list 
	 * @param customerId
	 * @return list
	 */
	@Transactional(propagation = Propagation.REQUIRED)
	public List<ExistingUserDTO> populateUserList(String customerId)
			throws DAOException {
		LOGGER.info(">>>>>> ManageSeviceFacilitatorDAOImpl.populateUserList");
		List<ExistingUserDTO> list = null;
		list = sessionFactory.getCurrentSession().getNamedQuery("fetchUserByFacilitatorId").setParameter("customerId", customerId).list();
		LOGGER.info("<<<<<< ManageSeviceFacilitatorDAOImpl.populateUserList");
		return list;
	}

	/**
	 * Method to fetch the user list with active and inactive status
	 * @param type
	 * @param customerId
	 */
	@Transactional(propagation = Propagation.REQUIRED)
	public List<ExistingUserDTO> populateActiveInactiveUserListForFacilitator (
			String userGroup, String type, String customerId) throws DAOException {
		LOGGER.info(">>>>>> ManageSeviceFacilitatorDAOImpl.populateActiveInactiveUserListForFacilitator");
		List<ExistingUserDTO> list = null;		
		int softDelete = 0;
		if (type.equals("active")) {
			softDelete = 1;
		} 
		list = sessionFactory.getCurrentSession().getNamedQuery("fetchActivateInactivateUserForFacilitator")
				.setParameter("softDelete", softDelete)
				.setParameter("customerId", customerId)
				.setParameter("userGroup", userGroup).list();
		LOGGER.info("<<<<<< ManageSeviceFacilitatorDAOImpl.populateActiveInactiveUserListForFacilitator");
		return list;
	}

	/**
	 * Method to fetch the user list to resets password
	 * @param customerId
	 * @return ExistingUserDTO
	 */
	@Transactional(propagation = Propagation.REQUIRED)
	public List<ExistingUserDTO> populateUserListResetPasswordForFacilitator (
			String userGroup, String customerId) throws DAOException {
		LOGGER.info(">>>>>> ManageSeviceFacilitatorDAOImpl.populateUserListResetPasswordForFacilitator");
		List<ExistingUserDTO> list = null;
		list = sessionFactory.getCurrentSession()
				.getNamedQuery("fetchUserToResetPasswordFacilitator")
				.setParameter("customerId", customerId)
				.setParameter("userGroup", userGroup).list();
		LOGGER.info("<<<<<< ManageSeviceFacilitatorDAOImpl.populateUserListResetPasswordForFacilitator");
		return list;
	}

	/**
	 * Method updates existing user's activation status in batch
	 * @param customerId
	 * @param softDelete
	 * @param emailIdString
	 * @return status
	 */
	@Transactional(propagation = Propagation.REQUIRED)
	public String updateUserStatusInBatchFacilitator(List<String> emailIdList,
			int softDelete, String customerId) throws DAOException {
		LOGGER.info(">>>> ManageSeviceFacilitatorDAOImpl.updateUserStatusInBatchFacilitator");
		String status = "failure";			
		try{
			sessionFactory.getCurrentSession().getNamedQuery("updateSoftDeleteStatusInBatchFacilitator")
				.setParameter("softDelete", softDelete)
				.setParameter("customerId", customerId)
				.setParameterList("emailIdList",emailIdList).executeUpdate();
			status = "success";
		}catch(Exception ez){
			status = "failure";
			LOGGER.error(ez.getLocalizedMessage());
		}
		LOGGER.info("<<<< ManageSeviceFacilitatorDAOImpl.updateUserStatusInBatchFacilitator");
		return status;
	}
	
	/**
	 * This will fetch the user object for which the password will be reset
	 * @param It will take User VO object as a input parameter
	 * @param updatedBy
	 * @return It returns list of all user data from database
	 */
	@Transactional(propagation = Propagation.REQUIRED)
	public List<JctUser> checkUserToResetPassword(UserVO userVO,
			String updatedBy) throws DAOException {
		LOGGER.info(">>>>>> ManageSeviceFacilitatorDAOImpl.checkUserToResetPassword");
		int facilitatorId = 0;
		facilitatorId = (Integer) sessionFactory.getCurrentSession().getNamedQuery("fetchFacilitatorID").
				setParameter("role", 3).
				setParameter("emailAddress", updatedBy).uniqueResult();			
		LOGGER.info("<<<<<< ManageSeviceFacilitatorDAOImpl.checkUserToResetPassword");
		return sessionFactory.getCurrentSession().getNamedQuery("userToResetPasswordFacilitator")
				.setParameter("facilitatorId", facilitatorId)
				.setParameter("emailId", userVO.getEmail()).list();
	}

	/**
	 * Method updates existing JctUser
	 * @param user
	 * @return status
	 */
	@Transactional(propagation = Propagation.REQUIRED)
	public String updateUserToResetPasswordFacilitator(JctUser user)
			throws DAOException {
		LOGGER.info(">>>>>> ManageSeviceFacilitatorDAOImpl.updateUserToResetPasswordFacilitator");
		return update(user);
	}

	/**
	 * Method returns list of JctEmailDetails
	 * @param emailId
	 * @return List of JctEmailDetails
	 */
	@Transactional(propagation = Propagation.REQUIRED)
	public List<JctEmailDetails> getEmailObjectFacilitator(String emailId, String updatedBy)
			throws DAOException {
		LOGGER.info(">>>>>> ManageSeviceFacilitatorDAOImpl.getEmailObject");
		return sessionFactory.getCurrentSession().getNamedQuery("jctEmailObjByEmail").setParameter("emailId", emailId).list();
	}

	@Transactional(propagation = Propagation.REQUIRED)
	public String cronUpdaterFacilitator(JctEmailDetails dtls) throws DAOException {
		LOGGER.info(">>>>>> ManageSeviceFacilitatorDAOImpl.updateUserToResetPasswordFacilitator");
		return update(dtls);
	}

	/**
	 * Method to update the account expire date for a particular user
	 * @param userId
	 * @param userEmails
	 * @param facilitatorEmail
	 * @param expiryDate
	 * @return status
	 */
	@Transactional(propagation = Propagation.REQUIRED)
	public String renewUserByFacilitator(int userId, String userEmails, String facilitatorEmail, Date expiryDate)
			throws DAOException {
		LOGGER.info(">>>>>> ManageSeviceFacilitatorDAOImpl.renewUserByFacilitator");
		String status = "failure";
		try{
			sessionFactory.getCurrentSession().getNamedQuery("renewUserByFacilitator")
				.setParameter("userId", userId)
				.setParameter("modifiedBy", facilitatorEmail)
				.setParameter("modifiedTs", new Date())
				.setParameter("expiryDate", expiryDate).executeUpdate();
			status = "success";
		}catch(Exception ez){
			status = "failure";
			LOGGER.error(ez.getLocalizedMessage());
		}
		LOGGER.info("<<<<<< ManageSeviceFacilitatorDAOImpl.renewUserByFacilitator");
		return status;
	}

	/**
	 * Method to fetch account expire date for user based on user ID
	 * @param userId
	 * @return Date
	 */
	@Transactional(propagation = Propagation.REQUIRED)
	public Date fetchUserExpiryDate(int userId) throws DAOException {
		LOGGER.info(">>>>>> ManageSeviceFacilitatorDAOImpl.fetchUserExpiryDate");
		return (Date) sessionFactory.getCurrentSession().getNamedQuery("fetchExpiryDateToRenew")
				.setParameter("role", 1)
				.setParameter("userId", userId).uniqueResult();
	}

	/**
	 * Method populates distinct user 
	 * @param customerId
	 * @return List of UserDTO
	 */
	@Transactional(propagation = Propagation.REQUIRED)
	public List<UserDTO> populateUser(String customerId, String facilitatorEmail) throws DAOException {
		LOGGER.info(">>>>>> ManageSeviceFacilitatorDAOImpl.populateUser");
		List<UserDTO> list = sessionFactory.getCurrentSession().getNamedQuery("fetchAllDistinctUser")
				.setParameter("customerId", customerId).setParameter("facilitatorEmail", facilitatorEmail).list();
		LOGGER.info("<<<<<< ManageSeviceFacilitatorDAOImpl.populateUser");
		return list;
	}

	/**
	 * Method populates existing facilitator list
	 * @return ExistingUserDTO
	 */
	@Transactional(propagation = Propagation.REQUIRED)
	public List<ExistingUserDTO> populateFacilitatorList(String customerId) throws DAOException {
		LOGGER.info(">>>>>> ManageSeviceDAOImpl.populateDistinctUserGroup");
		List<ExistingUserDTO> list = sessionFactory.getCurrentSession()
				.getNamedQuery("fetchAllFacilitatorList")
				.setParameter("customerId", customerId)
				.list();		
		LOGGER.info("<<<<<< ManageSeviceFacilitatorDAOImpl.populateDistinctUserGroup");
		return list;	
	}

	/**
	 * Method populates existing user data based on primary key
	 * @param primaryKey
	 * @return JctUser entity
	 */
	@Transactional(propagation = Propagation.REQUIRED)
	public JctUser fetchUserData(Integer primaryKey) throws DAOException {
		LOGGER.info(">>>>>> ManageSeviceFacilitatorDAOImpl.fetchUserData");
		return (JctUser) sessionFactory.getCurrentSession().
				getNamedQuery("fetchUserByPk")
				.setParameter("pk", primaryKey).list().get(0);	
	}

	/**
	 * Method save the user data
	 * @param user
	 * @return status
	 * @throws DAOException
	 */
	@Transactional(propagation = Propagation.REQUIRED)
	public String saveFacilitator(JctUser user) throws DAOException {
		LOGGER.info(">>>>>> ManageSeviceFacilitatorDAOImpl.saveFacilitator");
		return save(user);
	}

	/**
	 * Method updates user object set soft delete 1
	 * @param user
	 * @return status
	 * @throws DAOException
	 */
	@Transactional(propagation = Propagation.REQUIRED)
	public String updateUser(JctUser user) throws DAOException {
		LOGGER.info(">>>>>> ManageSeviceFacilitatorDAOImpl.updateUser");
		return update(user);
	}

	/**
	 * Method populates existing facilitator data based on emailAddress
	 * @param primaryKey
	 * @return JctUser entity
	 */
	@Transactional(propagation = Propagation.REQUIRED)
	public JctUser fetchFacilitatorData(String emailId, String customerId) throws DAOException {	
		LOGGER.info(">>>>>> ManageSeviceFacilitatorDAOImpl.fetchFacilitatorData");
		return (JctUser) sessionFactory.getCurrentSession().
				getNamedQuery("fetchFacilitatorByEmailId")
				.setParameter("customerId", customerId)
				.setParameter("emailId", emailId).list().get(0);
	}

	/**
	 * Method to search the email id is exist in database
	 * @param emailId
	 * @return status
	 */
	@Transactional(propagation = Propagation.REQUIRED)
	public String searchByEmail(String emailId) throws DAOException {
		LOGGER.info(">>>>>> ManageSeviceFacilitatorDAOImpl.searchByEmail");
		String status = "failure";
		List list = null;
		try{
			list = sessionFactory.getCurrentSession().getNamedQuery("fetchJctUserByEmailId")
					.setParameter("emailId", emailId).list();
			if(list.size() == 0){
				status = "notexist";
			} else {
				status = "exist";
			}			
		}catch(Exception ez){
			status = "notexist";
			LOGGER.error(ez.getLocalizedMessage());
		}
		LOGGER.info("<<<<<< ManageSeviceFacilitatorDAOImpl.searchByEmail");
		return status;
	}

	@Transactional(propagation = Propagation.REQUIRED)
	public int fetchDefaultProfileId() throws DAOException {
		LOGGER.info(">>>>>> ManageSeviceFacilitatorDAOImpl.fetchDefaultProfileId");
		return (Integer) sessionFactory.getCurrentSession().getNamedQuery("fetchDefaultProfileId")
				.setParameter("profileName", "Default Profile").uniqueResult();
	}

	@Transactional(propagation = Propagation.REQUIRED)
	public List<UserGroupDTO> getUserGroupDropDown(String customerId)
			throws DAOException {
		LOGGER.info(">>>>>> ManageSeviceFacilitatorDAOImpl.getUserGroupDropDown");
		List<UserGroupDTO> list = sessionFactory.getCurrentSession()
				.getNamedQuery("fetchUserGroupByCustId")
				.setParameter("customerId", customerId).list();
		LOGGER.info("<<<<<< ManageSeviceFacilitatorDAOImpl.getUserGroupDropDown");
		return list;
	}

	@Transactional(propagation = Propagation.REQUIRED)
	public List<JctUser> fetchUserObj(String email, String customerId)
			throws DAOException {
		LOGGER.info(">>>>>> ManageSeviceFacilitatorDAOImpl.fetchUserObj");
		List<JctUser> list =sessionFactory.getCurrentSession().getNamedQuery("getUserIdByUsernameCustIdRoleForSingleUser")
				.setParameter("userName", email)
				.setParameter("customerId", customerId)
				.setParameter("role", 1).list();
		LOGGER.info("<<<<<< ManageSeviceFacilitatorDAOImpl.fetchUserObj");
		return list;
	}

	@Transactional(propagation = Propagation.REQUIRED)
	public List<ExistingUserDTO> populateRenewedUserList(String customerId,
			List<String> validEmailIds) throws DAOException {
		LOGGER.info(">>>>>> ManageSeviceFacilitatorDAOImpl.populateRenewedUserList");
		List<ExistingUserDTO> list = null;
		list = sessionFactory.getCurrentSession().getNamedQuery("fetchRenewedUserByFacilitatorId")
				.setParameterList("validEmailIds", validEmailIds).list();
		LOGGER.info("<<<<<< ManageSeviceFacilitatorDAOImpl.populateUserList");
		return list;
	}

	@Transactional(propagation = Propagation.REQUIRED)
	public String assignNewUserGroupInBatch(int jctUserId,
			int jctUserGroup, String jctUserGroupDesc, String facilitatorEmail,
			String customerId) {
		LOGGER.info(">>>>>> ManageSeviceFacilitatorDAOImpl.assignNewUserGroupInBatch");
		String status = "failure";
		try{
			
			JctUser obj = (JctUser) sessionFactory.getCurrentSession().getNamedQuery("fetchUserByPk")
					.setParameter("pk", jctUserId).uniqueResult();
			obj.getJctUserDetails().setJctUserDetailsGroupId(jctUserGroup);
			obj.getJctUserDetails().setJctUserDetailsGroupName(jctUserGroupDesc);
			obj.getJctUserDetails().setJctUserDetailsLastmodifiedTs(new Date());
			update(obj);
			status = "success";
		}catch(Exception ez){
			status = "failure";
			LOGGER.error(ez.getLocalizedMessage());
		}
		LOGGER.info("<<<<<< ManageSeviceFacilitatorDAOImpl.assignNewUserGroupInBatch");
		return status;
	}

}