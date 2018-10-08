package com.vmware.jct.service;

import java.util.List;
import java.util.Map;

import com.vmware.jct.dao.dto.ExistingUserDTO;
import com.vmware.jct.dao.dto.FacilitatorDetailDTO;
import com.vmware.jct.dao.dto.NewUserDTO;
import com.vmware.jct.dao.dto.PaymentDetailsDTO;
import com.vmware.jct.dao.dto.UserAccountDTO;
import com.vmware.jct.dao.dto.UserGroupDTO;
import com.vmware.jct.exception.JCTException;
import com.vmware.jct.model.JctEmailDetails;
import com.vmware.jct.model.JctUser;
import com.vmware.jct.service.vo.FacilitatorAccountVO;
import com.vmware.jct.service.vo.GeneralUserAccountVO;
import com.vmware.jct.service.vo.ManageUserVO;
import com.vmware.jct.service.vo.UserGroupVO;
/**
 * 
 * <p><b>Class name:</b>IManageUserService.java</p>
 * <p><b>@author:</b> InterraIT </p>
 * <p><b>Purpose:</b> 		This Service interface represents a IManageUserService interface
 * <p><b>Description:</b> 	This will have all methods where we implement all business logic and all service layer </p>
 * <p><b>Copyrights:</b> 	All rights reserved by InterraIT and should only be used for its internal application development.</p>
 *  <pre>
 * <b>Revision History:</b>
 * InterraIT - 31/Jan/2014 - Implement exception through out the application
 * </pre>
 */
public interface IManageUserService {
	/**
	 * Method populates existing user profile
	 * @return ManageUserVO
	 * @throws JCTException
	 */
	public ManageUserVO populateExistingUserProfile() throws JCTException;
	/**
	 * Method saves new user profile
	 * @param userProfile
	 * @param createdBy
	 * @return ManageUserVO
	 * @throws JCTException
	 */
	public ManageUserVO saveUserProfile(String userProfile, String createdBy) throws JCTException;
	/**
	 * Method populates existing user profile in key value fashion
	 * @return Map
	 * @throws JCTException
	 */
	public Map<Integer, String> populateUserProfile() throws JCTException;
	/**
	 * Method saves new user group
	 * @param userGroupVO
	 * @return ManageUserVO
	 * @throws JCTException
	 */
	public ManageUserVO saveUserGroup(UserGroupVO userGroupVO) throws JCTException;
	/**
	 * Method updates existing user group
	 * @param userGroupVO
	 * @param dist
	 * @return ManageUserVO
	 * @throws JCTException
	 */
	public ManageUserVO updateUserGroup(UserGroupVO userGroupVO, String dist) throws JCTException;
	/**
	 * Method populates existing user group
	 * @param profileId
	 * @param customerId
	 * @param roleId
	 * @return ManageUserVO
	 * @throws JCTException
	 */
	public ManageUserVO populateExistingUserGroup(Integer profileId, String customerId, Integer roleId) throws JCTException;
	/**
	 * Method returns list of user group
	 * @return list of UserGroupDTO
	 * @throws JCTException
	 */
	public List<UserGroupDTO> getUserGroupList() throws JCTException;
	/**
	 * Method saves new user
	 * @param validEmailIds
	 * @param userGrpId
	 * @param userGrpName
	 * @param createdBy
	 * @return NewUserDTO
	 * @throws JCTException
	 */
	public NewUserDTO saveUsers(List<String> validEmailIds, int userGrpId, String userGrpName, String createdBy, String paymentType) throws JCTException;
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
	 * Method updates the mail sent status of the existing email idd
	 * @param registeredEmailId
	 * @param generatedPassword
	 * @param updatedBy
	 * @return status - String
	 * @throws JCTException
	 */
	public String updateCronUsers(String registeredEmailId, String generatedPassword, String updatedBy) throws JCTException;
	/**
	 * Method returns existing active user list based on user group
	 * @param userGroup
	 * @return ExistingUserDTO
	 * @throws JCTException
	 */
	public List<ExistingUserDTO> getUserList(String userGroup) throws JCTException;
	/**
	 * Method returns user list by user type and group
	 * @param userGroup
	 * @param userType
	 * @return
	 * @throws JCTException
	 */
	public List<ExistingUserDTO> getUserListByUserTypeAndGroup(String userGroup, String userType) throws JCTException;
	/**
	 * Method returns existing active user list based on user group
	 * @param userGroup
	 * @return ExistingUserDTO
	 * @throws JCTException
	 */
	public List<ExistingUserDTO> getUserListForAdmin(String userGroup, String emailId, int userType) throws JCTException;
	/**
	 * Method returns existing all user list based on user group
	 * @param userGroup
	 * @return ExistingUserDTO
	 * @throws JCTException
	 */
	public List<ExistingUserDTO> populateActiveInactiveUserList(String userGroup, String activeInactive, int userType) throws JCTException;
	/**
	 * 
	 * @param userGroup
	 * @param activeInactive
	 * @param emailId
	 * @return
	 * @throws JCTException
	 */
	public List<ExistingUserDTO> populateActiveInactiveUserListForAdmin(String userGroup, String activeInactive, String emailId, int roleId) throws JCTException;
	/**
	 * Method deletes existing user based on email id
	 * @param emailId
	 * @return status
	 * @throws JCTException
	 */
	public String deleteUser(String emailId, int userType) throws JCTException;
	/**
	 * Method updates existing user's activation status
	 * @param emailId
	 * @param softDelete
	 * @return status
	 * @throws JCTException
	 */
	public String updateUserActivationStatus(String emailId, int softDelete) throws JCTException;
	/**
	 * Method updates existing user's activation status in batch
	 * @param emailId
	 * @param softDelete
	 * @return status
	 * @throws JCTException
	 */
	public String updateUserActivationStatusInBatch(String emailIdString, int softDelete, int userType) throws JCTException;
	/**
	 * Method resets the password only if the activation status is 1
	 * @param emailId
	 * @param encryptedPassword
	 * @param profileId
	 * @return status
	 * @throws JCTException
	 */
	public String resetPassword(String emailId, String encryptedPassword, int profileId) throws JCTException;
	/**
	 * Method checks for the existence of provided user profile
	 * @param userProfile
	 * @return status
	 * @throws JCTException
	 */
	public String validateExistence(String userProfile) throws JCTException;
	/**
	 * Method checks for the existence of provided user group and user profile id
	 * @param userProfile
	 * @return status
	 * @throws JCTException
	 */
	public String validateExistenceUserGrp(String userGroup,int userProfile,int roleId, String customerId) throws JCTException;
	/**
	 * Method the number of already registered users users
	 * @param emailIdList
	 * @return count
	 * @throws JCTException
	 */
	public int alreadyRegUserCount(List<String> emailIdList) throws JCTException;
	/**
	 * Method fetches the email details of unsent emails
	 * @return
	 * @throws JCTException
	 */
	public List<JctEmailDetails> getEmailDetails (int maxResultsToFetch) throws JCTException;
	/**
	 * Method updates Email Details
	 * @param details
	 * @return
	 * @throws JCTException
	 */
	public String updateEmailDetails(JctEmailDetails details) throws JCTException;
	/**
	 * Method stores Facilitator who have been registered through <b>CHECK PAYMENT AND COMPLEMENTARY ONLY</b>.
	 * @param facilitatorChequeVO
	 * @return ConfirmationMessageVO
	 * @throws JCTException
	 */
	public int saveFacilitatorViaChequePayment (FacilitatorAccountVO facilitatorChequeVO, String paymentType) throws JCTException;
	/**
	 * Method stores general user who have been registered through <b>CHECK PAYMENT AND COMPLEMENTARY ONLY</b>.
	 * @param GeneralUserAccountVO
	 * @param validEmailListSize
	 * @return message
	 * @throws JCTException
	 */
	public String saveGeneralUserViaChequePayment (GeneralUserAccountVO accountVO, int validEmailListSize,String paymentType) throws JCTException;
	/**
	 * Method stores Facilitator who have been renewed through <b>CHECK PAYMENT AND COMPLEMENTARY ONLY</b>.
	 * @param facilitatorChequeVO
	 * @return ConfirmationMessageVO
	 * @throws JCTException
	 */
	public int renewSubscribeViaChequePayment (FacilitatorAccountVO facilitatorChequeVO,String inputRenewSub) throws JCTException;
	/**
	 * Method to search existing Facilitaro Email by Facilitator ID
	 * @throws JCTException
	 */
	public String getFacilitatorEmailByID (String facilitatorID) throws JCTException;
	/**
	 * Method fetches all users having cheque payment
	 * @param fetchType
	 * @param chequeNum
	 * @param userName
	 * @return
	 * @throws JCTException
	 */
	public String fetchExistingChequeUsers(String fetchType, String chequeNum, String userName) throws JCTException;
	
	/**
	 * Method to honor the check
	 * @param tranId
	 * @param createdBy
	 * @throws JCTException
	 */
	public int honoredCheck(String tranId,String createdBy) throws JCTException;
	/**
	 * Method to dishonor the check
	 * @param tranId
	 * @param createdBy
	 * @throws JCTException
	 */
	public String dishonoredCheck(String tranId, String createdBy) throws JCTException;
	/**
	 * Method fetches user ID based on facilitator email address
	 * @param tranId
	 * @throws JCTException
	 */
	public List<PaymentDetailsDTO> getExistingUserByTranId(String tranIdm) throws JCTException;
	/**
	 * 
	 * @param emailIdList
	 * @return
	 * @throws JCTException
	 */
	public List<String> getUserNamesOfIndUser(List<String> emailIdList) throws JCTException;
	/**
	 * 
	 * @param emailIdList
	 * @return
	 * @throws JCTException
	 */
	public List<String> getUserNamesForFacilitator(List<String> emailIdList, String customerId) throws JCTException;
	/**
	 * Method renew individual users
	 * @param GeneralUserAccountVO
	 * @param validEmailListSize
	 * @return message
	 * @throws JCTException
	 */
	public String renewGeneralUserPaymentManual (GeneralUserAccountVO accountVO, int validEmailListSize,String paymentType) throws JCTException;
	/**
	 * 
	 * @param emailIdList
	 * @return
	 * @throws JCTException
	 */
	public List<String> getInactiveUserListByEmail(List<String> emailIdList) throws JCTException;
	/**
	 * @param emailId
	 * @return
	 * @throws JCTException
	 */
	public boolean isIndividualUserExist (String emailId) throws JCTException;
	/**
	 * 
	 * @param emailId
	 * @param userType
	 * @return
	 * @throws JCTException
	 */
	public boolean shopifyUserChecker (String emailId, String userType) throws JCTException;
	/**
	 * @param emailId
	 * @return
	 * @throws JCTException
	 */
	public boolean isIndividualActiveUserExist (String emailId) throws JCTException;
	/**
	 * @param emailId
	 * @return
	 * @throws JCTException
	 */
	public boolean isIndividualInactiveUserExist (String emailId) throws JCTException;
	/**
	 * @param emailId
	 * @return
	 * @throws JCTException
	 */
	public List<UserAccountDTO> getUserEndDateDTO (String emailId) throws JCTException;
	
	/**
	 * @param emailId
	 * @return
	 * @throws JCTException
	 */
	public JctUser getPassword(String userName, int role) throws JCTException;
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
	 * @param role
	 * @return
	 * @throws JCTException
	 */
	public FacilitatorDetailDTO getFaciExpirationDtls(String custId, int role) throws JCTException;
	/**
	 * Method to fetch user's First name from user dtls
	 * @param userName 
	 * @param role
	 * @return String
	 * @throws JCTException
	 */
	public String getFirstName(String userName, int role) throws JCTException;
	/**
	 * Method to fetch faci's full name from user dtls
	 * @param userName 
	 * @param role
	 * @return String
	 * @throws JCTException
	 */
	public String getFaciName(String userName) throws JCTException;/**
	 * Method to reset facilitator password on demand
	 * @param userName
	 * @return String
	 * @throws JCTException
	 */
	public String forgotPassword(String userName) throws JCTException;
}
