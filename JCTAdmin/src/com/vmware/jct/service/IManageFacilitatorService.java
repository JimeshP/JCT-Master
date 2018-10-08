package com.vmware.jct.service;

import java.util.ArrayList;
import java.util.List;

import com.vmware.jct.dao.dto.ExistingUserDTO;
import com.vmware.jct.dao.dto.NewUserDTO;
import com.vmware.jct.dao.dto.UserDTO;
import com.vmware.jct.dao.dto.UserGroupDTO;
import com.vmware.jct.exception.JCTException;
import com.vmware.jct.service.vo.ManageUserVO;

/**
 * 
 * <p><b>Class name:</b>IManageFacilitatorService.java</p>
 * <p><b>@author:</b> InterraIT </p>
 * <p><b>Purpose:</b> 		This Service interface represents a IManageFacilitatorService interface
 * <p><b>Description:</b> 	This will have all methods where we implement all business logic and all service layer </p>
 * <p><b>Copyrights:</b> 	All rights reserved by InterraIT and should only be used for its internal application development.</p>
 *  <pre>
 * <b>Revision History:</b>
 * InterraIT - 10/Oct/2014 - Implement exception through out the application
 * </pre>
 */

public interface IManageFacilitatorService {
	/**
	 * Method populates total subscribed user and total registered user
	 * @param emailAddress
	 * @param type
	 * @param customerId
	 * @return ManageUserVO
	 * @throws JCTException
	 */
	public ManageUserVO populateSubscribedUser(String emailAddress, String type, String customerId) throws JCTException;
	/**
	 * Method returns existing active and deactivate user list 
	 * @param customerId
	 * @return ExistingUserDTO
	 * @throws JCTException
	 */
	public List<ExistingUserDTO> getUserList(String customerId) throws JCTException;	
	/**
	 * Method returns existing active and deactivate user list 
	 * @param customerId
	 * @return ExistingUserDTO
	 * @throws JCTException
	 */
	public List<ExistingUserDTO> getRenewedUsersList(String customerId,List<String> validEmailIds) throws JCTException;	
	/**
	 * Method the number of already registered users users
	 * @param emailIdList
	 * @return count
	 * @throws JCTException
	 */
	public int alreadyRegUserCount(List<String> emailIdList) throws JCTException;	
	/**
	 * Method saves new user
	 * @param validEmailIds
	 * @param userGrpName
	 * @param expiryDate
	 * @param createdBy
	 * @param customerId
	 * @return NewUserDTO
	 * @throws JCTException
	 */
	public NewUserDTO saveUsers(List<String> validEmailIds, String userGrpName, String expiryDate, String createdBy, String customerId, String newUserGroup) throws JCTException;
	/**
	 * Method saves list of email ids of users so that the cron job 
	 * can pick it up and can send the registration mail to the user later. 
	 * @param registeredEmailIds
	 * @param createdBy
	 * @return status - String
	 * @throws JCTException
	 */
	public String saveCronUsers(List<NewUserDTO> registeredEmailIds, String createdBy) throws JCTException;
	/**
	 * Method to populate Active and Inactive User List
	 * @param activeInactive
	 * @param customerId
	 * @return ExistingUserDTO
	 * @throws JCTException
	 */
	public List<ExistingUserDTO> populateActiveInactiveUserListForFacilitator(String userGroup, String activeInactive, String customerId) throws JCTException;
	/**
	 * Method to fetch the user list to resets password
	 * @param customerId
	 * @return ExistingUserDTO
	 * @throws JCTException
	 */
	public List<ExistingUserDTO> getResetPasswordUserListForFacilitator(String userGroup, String customerId) throws JCTException;
	/**
	 * Method updates existing user's activation status in batch
	 * @param customerId
	 * @param softDelete
	 * @param emailIdString
	 * @return status
	 * @throws JCTException
	 */
	public String updateActivationStatusInBatchFacilitator(String emailIdString, int softDelete, String customerId) throws JCTException;
	/**
	 * Method resets the password only if the activation status is 1
	 * @param emailId
	 * @param encryptedPassword
	 * @return status
	 * @throws JCTException
	 */
	public String resetPasswordFacilitator(String emailId, String encryptedPassword, String updatedBy) throws JCTException;
	/**
	 * Method updates the mail sent status of the existing email id
	 * @param registeredEmailId
	 * @param generatedPassword
	 * @param updatedBy
	 * @return status - String
	 * @throws JCTException
	 */
	public String updateCronUsersForFacilitator(String registeredEmailId, String generatedPassword, String updatedBy) throws JCTException;	
	/**
	 * Method to renew the account for particular user
	 * @param userId
	 * @param userEmail
	 * @param facilitatorEmail
	 * @param usersToRenew
	 * @param customerId
	 * @return status in string
	 * @throws JCTException
	 */
	public String renewUserByFacilitator(int userId, String userEmail, String facilitatorEmail, int usersToRenew, String customerId) throws JCTException;
	
	/**
	 * Method returns list of user 
	 * @param customerId
	 * @return list of UserDTO
	 * @throws JCTException
	 */
	public List<UserDTO> getUserDropDownList(String customerId, String facilitatorEmail) throws JCTException;	
	/**
	 * Method returns existing facilitator list
	 * @param null
	 * @return ExistingUserDTO
	 * @throws JCTException
	 */
	public List<ExistingUserDTO> getFacilitatorList(String customerId) throws JCTException;
	/**
	 * Method to validate account expire date for user
	 * @param userId
	 * @param userName
	 * @return status in String 
	 * @throws JCTException
	 */
	public String validateUserExpiryDate(int userId, String userName) throws JCTException;	
	/**
	 * Method to update the existing user to facilitator
	 * @param userId
	 * @param userName
	 * @param facilitatorEmail
	 * @param customerId
	 * @return ManageUserVO
	 * @throws JCTException
	 */
	public int changeUserRoleForFacilitator(int userId, String userName, String facilitatorEmail, String customerId) throws JCTException;
	
	/**
	 * Method to add new facilitator to change role
	 * @param newFacilitatorVal
	 * @param facilitatorEmail
	 * @param customerId
	 * @return ManageUserVO
	 * @throws JCTException
	 */
	public int addNewFacilitator(String newFacilitatorVal, String facilitatorEmail, String customerId) throws JCTException;
	
	/**
	 * Method to search the email id is exist in db
	 * @param newFacilitatorVal
	 * @return ManageUserVO
	 * @throws JCTException
	 */
	public String searchByEmail(String newFacilitatorVal) throws JCTException;
	
	/**
	 * Method to fetch Default Profile Id
	 * @return Default Profile Id
	 * @throws JCTException
	 */
	public int fetchDefaultProfileId() throws JCTException;
	/**
	 * Method returns list of user group
	 * @return list of UserGroupDTO
	 * @throws JCTException
	 */
	public List<UserGroupDTO> getUserGroupDropDown(String customerId) throws JCTException;
	/**
	 *  Method to renew bulk users
	 * @param emailIdString
	 * @param facilitatorEmail
	 * @param customerId
	 * @param totalUsers
	 * @return
	 * @throws JCTException
	 */
	public String renewBulkUserByFacilitator(List<String> validEmailIds, String facilitatorEmail, String customerId, int totalUsers) throws JCTException;
	/**
	 * 
	 * @param validEmailList
	 * @param facilitatorEmail
	 * @param customerId
	 * @param totalUsers
	 * @param userGroup
	 * @param newUserGroup
	 * @return
	 */
	public String assignNewGroupInBatch(ArrayList<String> validEmailList, String facilitatorEmail, String customerId, int totalUsers, String userGroup, String newUserGroup);
}
