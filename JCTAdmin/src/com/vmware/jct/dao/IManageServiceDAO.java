package com.vmware.jct.dao;

import java.util.Date;
import java.util.List;

import com.vmware.jct.dao.dto.ExistingUserDTO;
import com.vmware.jct.dao.dto.FacilitatorDetailDTO;
import com.vmware.jct.dao.dto.UserGroupDTO;
import com.vmware.jct.dao.dto.UserProfileDTO;
import com.vmware.jct.exception.DAOException;
import com.vmware.jct.exception.JCTException;
import com.vmware.jct.model.JctEmailDetails;
import com.vmware.jct.model.JctUser;
import com.vmware.jct.model.JctUserGroup;
import com.vmware.jct.model.JctUserProfile;
import com.vmware.jct.service.vo.ExistingFacilitatorVO;
import com.vmware.jct.service.vo.UserVO;

/**
 * 
 * <p><b>Class name:</b>IManageServiceDAO.java</p>
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
public interface IManageServiceDAO {
	/**
	 * Method populates Existing user profiles
	 * @return List of UserProfileDTO
	 * @throws DAOException
	 */
	public List<UserProfileDTO> populateExistingUserProfile() throws DAOException;
	/**
	 * Method saves new user profile
	 * @param userProfile
	 * @return statau
	 * @throws DAOException
	 */
	public String saveUserProfile(JctUserProfile userProfile) throws DAOException;
	/**
	 * Method populates existing user profiles with id
	 * @return List of UserProfileDTO
	 * @throws DAOException
	 */
	public List<UserProfileDTO> populateExistingUserProfileWithId() throws DAOException;
	/**
	 * Method saves new user group
	 * @param userGroup
	 * @return status
	 * @throws DAOException
	 */
	public String saveUserGroup(JctUserGroup userGroup) throws DAOException;
	/**
	 * Method updates existing user group
	 * @param userGroup
	 * @return status
	 * @throws DAOException
	 */
	public String updateUserGroup(JctUserGroup userGroup) throws DAOException;
	/**
	 * Method returns existing user group
	 * @return List of UserGroupDTO
	 * @throws DAOException
	 */
	public List<UserGroupDTO> populateExistingUserGroup(Integer profileId, String customerId, Integer roleId) throws DAOException;
	/**
	 * Method fetches user group based on primary key
	 * @param pkId
	 * @return JctUserGroup entity
	 * @throws DAOException
	 */
	public JctUserGroup fetchUserGroup(int pkId) throws DAOException;
	/**
	 * Method populates distinct user groups
	 * @return List of UserGroupDTO
	 * @throws DAOException
	 */
	public List<UserGroupDTO> populateDistinctUserGroup() throws DAOException;
	/**
	 * Method populates existing users based on user group
	 * @param userGroup
	 * @return List of ExistingUserDTO
	 * @throws DAOException
	 */
	public List<ExistingUserDTO> populateUserList(String userGroup) throws DAOException;
	/**
	 * Method populates existing users based on user group and user type
	 * @param userGroup
	 * @param userType
	 * @return
	 * @throws DAOException
	 */
	public List<ExistingUserDTO> populateUserListByUserTypeAndGroup(String userGroup, String userType) throws DAOException;
	/**
	 * Method populates existing users based on user group
	 * @param userGroup
	 * @param created by
	 * @param userType
	 * @return List of ExistingUserDTO
	 * @throws DAOException
	 */
	public List<ExistingUserDTO> populateUserListForAdmin(String userGroup, String emailId, int userType) throws DAOException;
	/**
	 * Method persists JctEmailDetails by the cron service
	 * @param dtls
	 * @return status
	 * @throws DAOException
	 */
	public String cronPersister(JctEmailDetails dtls) throws DAOException;
	/**
	 * Method updates existing JctEmailDetails by cron service
	 * @param dtls
	 * @return status
	 * @throws DAOException
	 */
	public String cronUpdater(JctEmailDetails dtls) throws DAOException;
	/**
	 * Method updates existing JctUser
	 * @param user
	 * @return status
	 * @throws DAOException
	 */
	public String updateUser(JctUser user) throws DAOException;
	/**
	 * Method populates user list based on user group and type
	 * @param userGroup
	 * @param type
	 * @return List of ExistingUserDTO
	 * @throws DAOException
	 */
	public List<ExistingUserDTO> populateSelectedUserList(String userGroup, String type) throws DAOException;
	/**
	 * Method populates user list based on user group and type
	 * @param userGroup
	 * @param type
	 * @return List of ExistingUserDTO
	 * @throws DAOException
	 */
	public List<ExistingUserDTO> populateSelectedUserListAll(String userGroup, String type, int userType) throws DAOException;
	/**
	 * Method populates user list based on user group and roleId 
	 * @param userGroup
	 * @param type
	 * @param emailId
	 * @param roleId
	 * @return
	 * @throws DAOException
	 */
	public List<ExistingUserDTO> populateSelectedUserListForAdmin(String userGroup, String type, String emailId, int roleId) throws DAOException;
	/**
	 * Method returns list of JctEmailDetails
	 * @param emailId
	 * @return List of JctEmailDetails
	 * @throws DAOException
	 */
	public List<JctEmailDetails> getEmailObject(String emailId) throws DAOException;
	/**
	 * Method validates user profile
	 * @param userProfile
	 * @return status
	 * @throws DAOException
	 */
	public String validateUserProfile(String userProfile) throws DAOException;
	/**
	 * Method validates user group
	 * @param userGroup
	 * @param userProfile
	 * @return status
	 * @throws DAOException
	 */
	public String validateUserGroup(String userGroup, int userProfile, int roleId, String customerId) throws DAOException;
	/**
	 * Method fetches the email details object
	 * @return list of email details
	 * @throws DAOException
	 */
	public List<JctEmailDetails> getEmailDetails (int maxResultsToFetch) throws DAOException;
	/**
	 * Method merges the latest obj
	 * @param details
	 * @return
	 * @throws DAOException
	 */
	public String updateEmailDetails (JctEmailDetails details) throws DAOException;
	/**
	 * Method to honor the check for new user
	 * @param jctUserid
	 * @param role
	 * @param expiryDate
	 * @throws DAOException
	 */
	public int updateUserToHonor (int jctUserid, int role, Date expiryDate) throws DAOException;
	/**
	 * Method to hard delete new user after check dishonor
	 * @param jctUserId
	 * @param role
	 * @throws DAOException
	 */
	public String deleteNewUser(int jctUserId, int role) throws DAOException;
	/**
	 * Method to update facilitator details table to after check realization
	 * @param userName
	 * @param customerId
	 * @throws DAOException
	 */
	public int updateFacilitatorDtlToHonor(int headerId, String custId ,String createdBy) throws DAOException;	

	/**
	 * Method to update payment details table after check realization
	 * @param detailId
	 * @param createdBy
	 * @throws DAOException
	 */
	public String updatePaymentDtlToHonor(int detailId, String createdBy) throws DAOException;
	

	/**
	 * Method to update check more user subscription renew table to after check realization
	 * @param jctUserId	
	 * @throws DAOException
	 */
	public int updateCheckPaymentDetails(int jctUserId) throws DAOException;
	
	/**
	 * Method to update check payment details after check dishonored
	 * @param  jctCheckPaymentUserId
	 * @throws DAOException 
	 * **/
	public String updateCheckPaymentDetailsToDishonor(int jctCheckPaymentUserId) throws DAOException ;
	
	/**
	 * Method to update payment details table to after check dishonored
	 * @param detailId
	 * @param createdBy
	 * @throws DAOException
	 */
	public String updatePaymentDtlToDishonor(int detailId, String createdBy) throws DAOException;
	/**
	 * Method to update payment header table to after check dishonored
	 * @param headerId
	 * @param createdBy
	 * @throws DAOException
	 */	
	public String updatePaymentHdrToDishonor(int headerId, String createdBy) throws DAOException;
	
	/**
	 * Method to update payment header table by the amount paid for single user registration
	 * @param registeredPaymentHdrId
	 * @param costPerUser
	 * @throws DAOException
	 */	
	public int updateAmountInPaymentHdr(String registeredPaymentHdrId, double costPerUser) throws DAOException;
	/**
	 * Method fetch the user profile object against Default Profile
	 * @param userProfileName
	 * @return
	 * @throws DAOException
	 */
	public UserProfileDTO getDefaultProfileObj(String userProfileName) throws DAOException;
	/**
	 * Method fetch the user group object against Default Profile
	 * @param userGroupName
	 * @param customerId
	 * @param role
	 * @return
	 * @throws DAOException
	 */
	public UserGroupDTO getDefaultGroupObj(String userGroupName, String customerId, int role) throws DAOException;	
	/**
	 * Method fetch user object to renew individual user
	 * @param userVO
	 * @param roleId
	 * @return
	 * @throws DAOException
	 */
	public JctUser checkIndividualForRenew(UserVO userVO, int roleId) throws DAOException;
	/**
	 * 
	 * @param emailIdList
	 * @return
	 * @throws JCTException
	 */
	public List<String> getInactiveUserListByEmail(List<String> emailIdList) throws DAOException;
	/**
	 * 
	 * @param emailId
	 * @return
	 * @throws JCTException
	 */
	public JctUser getPasswordObject(String emailId, int role) throws DAOException;
	/**
	 * Method to fetch user's expiration dtls using email & role
	 * @param emailId
	 * @param role
	 * @return
	 * @throws JCTException
	 */
	public ExistingUserDTO getUserExpirationDtls(String userName, int role) throws JCTException;
	/**
	 * Method to fetch user's expiration dtls using email & role
	 * @param custId
	 * @return FacilitatorDetailDTO
	 * @throws JCTException
	 */
	public FacilitatorDetailDTO getFaciExpirationDtls(String custId) throws JCTException;
	/**
	 * Method to fetch user's first name
	 * @param emailId
	 * @param role
	 * @return String
	 * @throws JCTException
	 */
	public String getFirstName(String emailId, int role) throws JCTException;
	/**
	 * Method to fetch facilitator's full name
	 * @param emailId
	 * @return JctUser
	 * @throws JCTException
	 */
	public JctUser getFullName(String emailId) throws JCTException;
	
}
