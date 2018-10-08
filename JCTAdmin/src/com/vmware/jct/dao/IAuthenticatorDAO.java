package com.vmware.jct.dao;

import java.util.List;

import com.vmware.jct.dao.dto.FunctionDTO;
import com.vmware.jct.dao.dto.LevelDTO;
import com.vmware.jct.dao.dto.OccupationListDTO;
import com.vmware.jct.dao.dto.PaymentDetailsDTO;
import com.vmware.jct.dao.dto.UserAccountDTO;
import com.vmware.jct.dao.dto.UserProfileDTO;
import com.vmware.jct.exception.DAOException;
import com.vmware.jct.exception.JCTException;
import com.vmware.jct.model.JctUser;
import com.vmware.jct.model.JctUserDetails;
import com.vmware.jct.model.JctUserProfile;
import com.vmware.jct.service.vo.UserVO;

/**
 * 
 * <p><b>Class name:</b>IAuthenticatorDAO.java</p>
 * <p><b>@author:</b> InterraIT </p>
 * <p><b>Purpose:</b> 		This DAO interface represents a IAuthenticatorDAO interface
 * <p><b>Description:</b> 	This will have all methods mapped to the database and fetching data from database to be used across the application </p>
 * <p><b>Note:</b> 			This methods implemented by DAOImpl classes and execute the mapped queries and get the records.</p>  
 * <p><b>Copyrights:</b> 	All rights reserved by Interra IT and should only be used for its internal application development.</p>
 *  <pre>
 * <b>Revision History:</b>
 * InterraIT - 19/Jan/2014 - Implement exception through out the appliaction
 * </pre>
 */

public interface IAuthenticatorDAO {
	/**
	 * Method authenticates user
	 * @param It will take User VO object as a input parameter
	 * @return It returns list of all user data from database
	 * @throws DAOException -it throws DAOException
	 */
	public List<JctUser> authenticateUser(UserVO userVO) throws DAOException;
	/**
	 * Mthod checks user
	 * @param It will take User VO object as a input parameter
	 * @return It returns list of all user data from database
	 * @throws DAOException -it throws DAOException
	 */
	public List<JctUser> checkUser(UserVO userVO) throws DAOException;
	/**
	 * Method check user by email id and role
	 * @param userVO
	 * @param roleId
	 * @return
	 * @throws DAOException
	 */
	public List<JctUser> checkUserByEmailAndRole(UserVO userVO, int roleId) throws DAOException;
	/**
	 * method checks user list
	 * @param It will take user list object as a input parameter
	 * @return It returns list of all user data from database
	 * @throws DAOException -it throws DAOException
	 */
	public List<JctUser> checkUserList(List<String> userList) throws DAOException;
	/**
	 * method updates login info
	 * @param It will take Job reference number and row id as a input parameters
	 * @return It will returns string object
	 * @throws DAOException -it throws DAOException
	 */
	public String updateLoginInfo(String jobRefNo, int rowId, String lastActivePage) throws DAOException;
	/**
	 * This will return list of function DTOs
	 * @return list of function dto
	 * @throws DAOException
	 */
	public List<FunctionDTO> getJctFunctionList() throws DAOException;
	/**
	 * This will return list of Level DTOs
	 * @return list of level dto
	 * @throws DAOException
	 */
	public List<LevelDTO> getJctLevelList() throws DAOException;
	/**
	 * This method will create new user
	 * @param userObj
	 * @return status
	 * @throws DAOException
	 */
	public String createNewUser(JctUser userObj) throws DAOException;
	/**
	 * This method will delete existing user
	 * @param userObj
	 * @return status
	 * @throws DAOException
	 */
	public String deleteUser(JctUser userObj) throws DAOException;
	/**
	 * This method will update the soft delete status.
	 * 0 to 1 and vice versa based on email id
	 * @param emailId
	 * @param softDelete
	 * @return status
	 * @throws DAOException
	 */
	public String updateUserSoftDeleteStatus(String emailId, int softDelete) throws DAOException;
	/**
	 * This method will update the soft delete status in batch.
	 * 0 to 1 and vice versa based on list of email id
	 * @param emailId
	 * @param softDelete
	 * @return status
	 * @throws DAOException
	 */
	public String updateUserSoftDeleteStatusInBatch(List<String> emailIdList, int softDelete, int userType) throws DAOException;
	/**
	 * Method returns list of users based on group name
	 * @param groupName
	 * @return list of users
	 * @throws DAOException
	 */
	public List<JctUser> getAllUsersOfUserGroup(String groupName) throws DAOException;
	/**
	 * This will update user status (inactive or active)
	 * @param user
	 * @return status
	 * @throws DAOException
	 */
	public String updateUserStatus (JctUser user) throws DAOException;
	/**
	 * This will return list of user profile DTOs
	 * @return list of user profile dto
	 * @throws DAOException
	 */
	public List<UserProfileDTO> getJctUserProfileList() throws DAOException;
	/**
	 * This will return user id based on email id and role
	 * @return user Id
	 * @throws DAOException
	 */
	public int getUserId(String emailId, int roleId) throws DAOException;
	/**
	 * Method returns existing users based on email id and role
	 * @param tranId
	 * @return
	 * @throws DAOException
	 */
	public List<PaymentDetailsDTO> getExistingUserByTranId(String tranId) throws DAOException;	
	/**
	 * Method gets existing facilitators by id
	 * @param facilitatorEmail
	 * @param role
	 * @return
	 * @throws DAOException
	 */
	public List<JctUser> getExistingFacilitatorByID (String facilitatorID) throws DAOException;
	/**
	 * Method gets facilitator mail by id
	 * @param facilitatorID
	 * @return List<Object[]>
	 * @throws DAOException
	 */
	public List<Object[]> getFacilitatorEmailByID (String facilitatorID) throws DAOException;
	
	public List<Object[]> getFacilitatorByMailId (String facilitatorID) throws DAOException;
	/**
	 * Method fetches survey questions
	 * @return
	 * @throws DAOException
	 */
	public List<Object[]> getTextSurveyQuestion() throws DAOException;
	/**
	 * Method fetches survey main questions
	 * @return
	 * @throws DAOException
	 */
	public List<Object[]> getSurveyMainQuestion(int ansType) throws DAOException;
	/**
	 * Method fetches survey sub questions
	 * @param ansType
	 * @param mainQtn
	 * @return
	 * @throws DAOException
	 */
	public List<String> getAllSubQtns(int ansType, String mainQtn) throws DAOException;	
	/**
	 * @param emailIdList
	 * @return
	 */
	public List<String> getUserNamesOfIndUser(List<java.lang.String> emailIdList) throws DAOException;

	
	public List<JctUser> getExistingUserByEmailAndRole(String facilitatorEmail, int i) throws DAOException;
	/**
	 * 
	 * @param emailIdList
	 * @return
	 */
	public List<String> getUserNamesForFacilitator(List<String> emailIdList, String customerId) throws DAOException;
	/**
	 * 
	 * @param emailId
	 * @return
	 * @throws JCTException
	 */
	public List<String> isIndividualUserExist(java.lang.String emailId) throws DAOException;
	/**
	 * 
	 * @param emailId
	 * @param userType
	 * @return
	 * @throws DAOException
	 */
	public List<String> shopifyUserChecker(java.lang.String emailId, int userType) throws DAOException;
	/**
	 * 
	 * @param emailId
	 * @return
	 * @throws JCTException
	 */
	public List<String> isIndividualActiveUserExist(java.lang.String emailId) throws DAOException;
	/**
	 * 
	 * @param emailId
	 * @return
	 * @throws JCTException
	 */
	public List<String> isIndividualInactiveUserExist(java.lang.String emailId) throws DAOException;
	
	public List<UserAccountDTO> getUserEndDateDTO(java.lang.String emailId) throws DAOException;
	
	public List<OccupationListDTO> searchOccupationList (String searchString) throws DAOException;
	
	public String updateUser(String password, int activeFlag, String email) throws DAOException;
	
	public List<JctUser> getUser(String email) throws DAOException;
	
	public List<JctUser> checkUserForDeletion(UserVO userVO, int userType) throws DAOException;
	
	public List<JctUserDetails> checkUsersCreatedByFacilitator(int userId) throws DAOException;
	
	public String deleteUserByUserId(int userId) throws DAOException;
	
	/**
	 * Method to fetch user profile using profile id
	 * @param userProfileId 
	 * @throws DAOException 
	 * **/
	public JctUserProfile getUserProfileByPk(int userProfileId) throws DAOException;
	/**
	 * Method to fetch terms & conditions for facilitator by user profile
	 * @param userProfile
	 * @param userType
	 * @return String 
	 * @throws DAOException 
	 * **/	
	public String getTermsAndConditions(JctUserProfile userProfile, int userType) throws DAOException;
	
	public List<Object[]> getAllSurveyQuestion() throws DAOException;
}
