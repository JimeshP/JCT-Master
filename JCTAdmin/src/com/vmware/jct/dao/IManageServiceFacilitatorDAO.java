package com.vmware.jct.dao;

import java.util.Date;
import java.util.List;

import com.vmware.jct.dao.dto.ExistingUserDTO;
import com.vmware.jct.dao.dto.FacilitatorDetailDTO;
import com.vmware.jct.dao.dto.UserDTO;
import com.vmware.jct.dao.dto.UserGroupDTO;
import com.vmware.jct.exception.DAOException;
import com.vmware.jct.model.JctEmailDetails;
import com.vmware.jct.model.JctUser;
import com.vmware.jct.model.JctUserGroup;
import com.vmware.jct.service.vo.UserVO;

/**
 * 
 * <p><b>Class name:</b>IManageServiceFacilitatorDAO.java</p>
 * <p><b>@author:</b> InterraIT </p>
 * <p><b>Purpose:</b> 		This DAO interface represents a IAuthenticatorDAO interface
 * <p><b>Description:</b> 	This will have all methods mapped to the database and fetching data from database to be used for Manage user service. </p>
 * <p><b>Note:</b> 			This methods implemented by DAOImpl classes and execute the mapped queries and get the records.</p>  
 * <p><b>Copyrights:</b> 	All rights reserved by Interra IT and should only be used for its internal application development.</p>
 *  <pre>
 * <b>Revision History:</b>
 * <li> </li>
 * </pre>
 */
public interface IManageServiceFacilitatorDAO {

	/**
	 * Method populates the Total Subscribed User and Registered User
	 * @param emailAddress
	 * @param type
	 * @param customerId
	 * @return List of FacilitatorDetailDTO
	 * @throws DAOException
	 */
	public List<FacilitatorDetailDTO> populateSubscribedUser(String emailAddress, String type, String customerId) throws DAOException;
	
	/**
	 * Method returns existing active and deactivate user list 
	 * @param customerId
	 * @return List of ExistingUserDTO
	 * @throws DAOException
	 */
	public List<ExistingUserDTO> populateUserList(String customerId) throws DAOException;
	
	/**
	 * Method persists JctEmailDetails by the cron service
	 * @param dtls
	 * @return status
	 * @throws DAOException
	 */
	public String cronPersister(JctEmailDetails dtls) throws DAOException;
	/**
	 * Method fetches user group Id based on user group name
	 * @param userGrpName
	 * @return JctUserGroup entity
	 * @throws DAOException
	 */
	public JctUserGroup fetchUserGroupId(int userGrpId, String customerId, String userGrpName) throws DAOException;	
	/**
	 * Method fetches facilitator ID based on facilitator email address
	 * @param emailId
	 * @return facilitatorId
	 * @throws DAOException
	 */
	public Integer fetchFacilitatorId(String emailId) throws DAOException;
	/**
	 * Method to update the number of registered user based on number of user added
	 * @param userCount
	 * @param customerId
	 * @param type
	 * @return status
	 * @throws DAOException
	 */
	public String updateFacilitatorDetails(int userCount, String customerId, String type) throws DAOException;
	
	/**
	 * Method to fetch the user list with active and inactive status
	 * @param type
	 * @param emailId
	 * @return ExistingUserDTO
	 * @throws DAOException
	 */
	public List<ExistingUserDTO> populateActiveInactiveUserListForFacilitator(String userGroup, String type, String customerId) throws DAOException;
	
	/**
	 * Method to fetch the user list to resets password
	 * @param customerId
	 * @return ExistingUserDTO
	 * @throws DAOException
	 */
	public List<ExistingUserDTO> populateUserListResetPasswordForFacilitator(String userGroup, String customerId) throws DAOException;
	
	/**
	 * This method will update the soft delete status.
	 * 0 to 1 and vice versa based on email id
	 * @param customerId
	 * @param softDelete
	 * @return status
	 * @throws DAOException
	 */
	public String updateUserStatusInBatchFacilitator(List<String> emailIdList, int softDelete, String customerId) throws DAOException;
	/**
	 * @param It will take User VO object as a input parameter
	 * @return It returns list of all user data from database
	 * @throws DAOException -it throws DAOException
	 */
	public List<JctUser> checkUserToResetPassword(UserVO userVO, String updatedBy) throws DAOException;
	
	/**
	 * Method updates existing JctUser
	 * @param user
	 * @return status
	 * @throws DAOException
	 */
	public String updateUserToResetPasswordFacilitator(JctUser user) throws DAOException;
	
	/**
	 * Method returns list of JctEmailDetails
	 * @param emailId
	 * @return List of JctEmailDetails
	 * @throws DAOException
	 */
	public List<JctEmailDetails> getEmailObjectFacilitator(String emailId, String updatedBy) throws DAOException;
	/**
	 * Method updates existing JctEmailDetails by cron service
	 * @param dtls
	 * @return status
	 * @throws DAOException
	 */
	public String cronUpdaterFacilitator(JctEmailDetails dtls) throws DAOException;
	/**
	 * Method to update the account expire date for a particular user
	 * @param userId
	 * @param userEmails
	 * @param facilitatorEmail
	 * @param expiryDate
	 * @return status
	 * @throws DAOException
	 */
	public String renewUserByFacilitator(int userId, String userEmails, String facilitatorEmail, Date expiryDate) throws DAOException;	
	/**
	 * Method to fetch account expire date for user based on user ID
	 * @param userId
	 * @return Date
	 * @throws DAOException
	 */
	public Date fetchUserExpiryDate(int userId) throws DAOException;
	/**
	 * Method populates distinct user 
	 * @param customerId
	 * @return List of UserDTO
	 * @throws DAOException
	 */
	public List<UserDTO> populateUser(String customerId, String facilitatorEmail) throws DAOException;
	/**
	 * Method populates existing facilitator
	 * @param null
	 * @return List of ExistingUserDTO
	 * @throws DAOException
	 */
	public List<ExistingUserDTO> populateFacilitatorList(String customerId) throws DAOException;
	/**
	 * Method populates existing user data based on primary key
	 * @param primaryKey
	 * @return JctUser entity
	 * @throws DAOException
	 */
	public JctUser fetchUserData(Integer primaryKey) throws DAOException;
	/**
	 * Method save the user data
	 * @param user
	 * @return status
	 * @throws DAOException
	 */
	public String saveFacilitator(JctUser user) throws DAOException;
	/**
	 * Method updates user object set soft delete 1
	 * @param user
	 * @return status
	 * @throws DAOException
	 */
	public String updateUser(JctUser user) throws DAOException;
	/**
	 * Method populates existing facilitator data based emailAddress and customerId
	 * @param emailId
	 * @param customerId
	 * @return JctUser entity
	 * @throws DAOException
	 */
	public JctUser fetchFacilitatorData(String emailId, String customerId) throws DAOException;
	/**
	 * Method to search the email id is exist in db
	 * @param newFacilitatorVal
	 * @return ManageUserVO
	 * @throws DAOException
	 */
	public String searchByEmail(String emailId) throws DAOException;
	/**
	 * Method to fetch Default Profile Id	
	 * @return Default Profile Id	
	 * @throws DAOException
	 */
	public int fetchDefaultProfileId() throws DAOException;
	/**
	 * Method populates distinct user groups
	 * @return List of UserGroupDTO
	 * @throws DAOException
	 */
	public List<UserGroupDTO> getUserGroupDropDown(String customerId) throws DAOException;
	/**
	 * 
	 * @param email
	 * @param customerId
	 * @return
	 * @throws DAOException
	 */
	public List<JctUser> fetchUserObj(String email, String customerId) throws DAOException;
	/**
	 * Method returns existing active and deactivate user list 
	 * @param customerId
	 * @return List of ExistingUserDTO
	 * @throws DAOException
	 */
	public List<ExistingUserDTO> populateRenewedUserList(String customerId, List<String> validEmailIds) throws DAOException;
	/**
	 * 
	 * @param jctUserId
	 * @param jctUserGroup
	 * @param jctUserGroupDesc
	 * @param facilitatorEmail
	 * @param customerId
	 * @return
	 */
	public String assignNewUserGroupInBatch(int jctUserId, int jctUserGroup, String jctUserGroupDesc, String facilitatorEmail, String customerId);
}
