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

import com.vmware.jct.common.utility.CommonConstants;
import com.vmware.jct.dao.DataAccessObject;
import com.vmware.jct.dao.IAuthenticatorDAO;
import com.vmware.jct.dao.dto.FunctionDTO;
import com.vmware.jct.dao.dto.LevelDTO;
import com.vmware.jct.dao.dto.OccupationListDTO;
import com.vmware.jct.dao.dto.PaymentDetailsDTO;
import com.vmware.jct.dao.dto.UserAccountDTO;
import com.vmware.jct.dao.dto.UserProfileDTO;
import com.vmware.jct.exception.DAOException;
import com.vmware.jct.model.JctUser;
import com.vmware.jct.model.JctUserDetails;
import com.vmware.jct.model.JctUserProfile;
import com.vmware.jct.service.vo.UserVO;

/**
 * 
 * <p><b>Class name:</b> AuthenticatorDAOImpl.java</p>
 * <p><b>Author:</b> InterraIT</p>
 * <p><b>Purpose:</b> This is a AuthenticatorDAOImpl class. This artifact is Persistence Manager layer artifact.
 * AuthenticatorDAOImpl implement IAuthenticatorDAO interface and override the following  methods.
 * -authenticateUser(UserVO userVO)
 * -checkUser(UserVO userVO)
 * -checkUsers(UserVO userVO)
 * -register(JctUser user)
 * -resetPassword(JctUser user)
 * -merge(JctUser user)</p>
 * <p><b>Description:</b> This class used to perform insert, modify ,fetch data from db and delete operation. Save, update, delete operations are performed in DataAccessObject class </p>
 * <p><b>Revision History:</b>
 * 	<li>InterraIT - 19/Jan/2014 - Implement Exception </li>
 * <li>InterraIT - 31/Jan/2014 - Implement comments, introduce Named query, Transactional attribute </li>
 * </p>
 */

@Repository
public class AuthenticatorDAOImpl extends DataAccessObject implements IAuthenticatorDAO {

	@Autowired
	private SessionFactory sessionFactory;
	
	private static final Logger LOGGER = LoggerFactory.getLogger(AuthenticatorDAOImpl.class);
	
	/**
	 *<p><b>Description:</b>  This method verify user id, password, active status and return user object</p>
	 */
	@Transactional(propagation=Propagation.REQUIRED)
	public List<JctUser> authenticateUser(UserVO userVO) throws DAOException  {
		return sessionFactory.getCurrentSession().getNamedQuery("fetchJctUser").setString("email", userVO.getEmail())
				.setString("password", userVO.getPassword())
				.setInteger("active", CommonConstants.ACTIVATED)
				.setInteger("inactive", CommonConstants.CREATED)
				.setInteger("justCreated", CommonConstants.FACI_ACCOUNT_JUST_CREATED)
				.setInteger("chequeProb", CommonConstants.CHECK_PAYMENT_NOT_ACTIVATED)
				.setInteger("resetPassword", CommonConstants.FACI_PASSWORD_RESET)
				.list();
	}
	
	/**
	 *<p><b>Description:</b>  Bases on user id this method return user object</p>
	 */
	@Transactional(propagation=Propagation.REQUIRED,readOnly = true)
	public List<JctUser> checkUser(UserVO userVO) throws DAOException {
		return sessionFactory.getCurrentSession().getNamedQuery("fetchJctUserByEmailId").setString("emailId", userVO.getEmail()).list();
	}
	
	/**
	 *<p><b>Description:</b>  Bases on user id this method return user object</p>
	 */
	@Transactional(propagation=Propagation.REQUIRED,readOnly = true)
	public List<JctUser> checkUserForDeletion(UserVO userVO, int userType) throws DAOException {
		return sessionFactory.getCurrentSession().getNamedQuery("fetchJctUserByEmailIdAndRoleId")
				.setString("emailId", userVO.getEmail())
				.setParameter("roleId", userType).list();
	}
	/**
	 *<p><b>Description:</b>  Bases on user id this method return user object</p>
	 */
	@Transactional(propagation=Propagation.REQUIRED,readOnly = true)
	public List<JctUser> checkUserByEmailAndRole(UserVO userVO, int roleId) throws DAOException {
		return sessionFactory.getCurrentSession().getNamedQuery("fetchJctUserByEmailIdAndRoleId")
				.setString("emailId", userVO.getEmail())
				.setParameter("roleId", roleId).list();
	}
	
	/**
	 *<p><b>Description:</b>  Bases on user id and password this method return user object</p>
	 */
	@Transactional(propagation=Propagation.REQUIRED,readOnly = true)
	public List<JctUser> checkUsers(UserVO userVO) throws DAOException {
		return sessionFactory.getCurrentSession().getNamedQuery("fetchJctUserByEmailIdPassword")
				.setString("email", userVO.getEmail())
				.setString("password", userVO.getInitialPassword()).list();
	}


	/**
	 *<p><b>Description:</b>  This method is used for update user info data</p>
	 */
	@Transactional(propagation=Propagation.REQUIRED)
	public String updateLoginInfo(String jobRefNo, int rowId, String lastActivePage) throws DAOException {
		LOGGER.info(">>>> AuthenticatorDAOImpl.updateLoginInfo");
		String status = "failure";
		try{
			sessionFactory.getCurrentSession().getNamedQuery("updateUserInfo")
				.setParameter("endTime", new Date())
				.setParameter("lastActivePage",lastActivePage)
				.setParameter("job", jobRefNo)
				.setParameter("rowId", rowId).executeUpdate();
			status = "success";
			LOGGER.info("<<<< AuthenticatorDAOImpl.updateLoginInfo");
		}catch(Exception e){
			LOGGER.error(e.getLocalizedMessage());
		}
		return status;
	}

	/**
	 *<p><b>Description:</b>  This method return list of function </p>
	 */
	@Transactional(propagation=Propagation.REQUIRED)
	public List<FunctionDTO> getJctFunctionList() throws DAOException{
		return sessionFactory.getCurrentSession().getNamedQuery("fetchFunction").list();
	}
	
	/**
	 *<p><b>Description:</b>  This method return list of level </p>
	 */
	@Transactional(propagation=Propagation.REQUIRED)
	public List<LevelDTO> getJctLevelList() throws DAOException{
		return sessionFactory.getCurrentSession().getNamedQuery("fetchJobLevel").list();
	}

	/**
	 *<p><b>Description:</b>  This method creates new user </p>
	 */
	@Transactional(propagation=Propagation.REQUIRED)
	public String createNewUser(JctUser userObj) throws DAOException {
		return save(userObj);
	}	
	
	/**
	 *<p><b>Description:</b>  This method deletes existing users </p>
	 */
	@Transactional(propagation=Propagation.REQUIRED)
	public String deleteUser(JctUser userObj) throws DAOException {
		LOGGER.info(">>>> AuthenticatorDAOImpl.deleteUser");
		String result = "success";
		try {
			Delete(userObj);
		} catch (DAOException ex) {
			result = "success";
			LOGGER.error(ex.getLocalizedMessage());
		}
		LOGGER.info("<<<< AuthenticatorDAOImpl.deleteUser");
		return result;
	}
	
	/**
	 *<p><b>Description:</b>  This method return list of users </p>
	 */
	@Transactional(propagation=Propagation.REQUIRED)
	public List<JctUser> getAllUsersOfUserGroup(String groupName) throws DAOException {
		return (List<JctUser>)sessionFactory.getCurrentSession().getNamedQuery("getUsrByUsrGrp").setParameter("userGroupName", groupName).list();
	}
	
	/**
	 *<p><b>Description:</b>  This method is used for update user info data</p>
	 */
	@Transactional(propagation=Propagation.REQUIRED)
	public String updateUserSoftDeleteStatus(String emailId, int softDelete) throws DAOException {
		String status = "failure";
		try{
			sessionFactory.getCurrentSession().getNamedQuery("updateSoftDeleteStatus")
				.setParameter("softDelete", softDelete)
				.setParameter("emailId",emailId).executeUpdate();
			status = "success";
		}catch(Exception ez){
			status = "failure";
		}
		return status;
	}

	/**
	 *<p><b>Description:</b>  This method updates the soft delete status </p>
	 */
	@Transactional(propagation=Propagation.REQUIRED)
	public String updateUserSoftDeleteStatusInBatch(List<String> emailIdList,
			int softDelete, int userType) throws DAOException {
		LOGGER.info(">>>> AuthenticatorDAOImpl.updateUserSoftDeleteStatusInBatch");
		String status = "failure";
		try{
			sessionFactory.getCurrentSession().getNamedQuery("updateSoftDeleteStatusInBatch")
				.setParameter("softDelete", softDelete)
				.setParameterList("emailIdList",emailIdList)
				.setParameter("roleId", userType).executeUpdate();
			status = "success";
		}catch(Exception ez){
			status = "failure";
			LOGGER.error(ez.getLocalizedMessage());
		}
		LOGGER.info("<<<< AuthenticatorDAOImpl.updateUserSoftDeleteStatusInBatch");
		return status;
	}
	
	/**
	 *<p><b>Description:</b>  This method updates user status </p>
	 */
	@Transactional(propagation=Propagation.REQUIRED)
	public String updateUserStatus (JctUser user) throws DAOException {
		return update(user);
	}

	/**
	 *<p><b>Description:</b>  This method return list of user profile </p>
	 */
	@Transactional(propagation=Propagation.REQUIRED)
	public List<UserProfileDTO> getJctUserProfileList() throws DAOException {
		return sessionFactory.getCurrentSession().getNamedQuery("fetchUserProfile").list();
	}

	@Transactional(propagation=Propagation.REQUIRED)
	public List<JctUser> checkUserList(List<String> userList)
			throws DAOException {
		return sessionFactory.getCurrentSession().getNamedQuery("fetchJctUserByEmailIdList").setParameterList("emailIdList", userList).list();
	}
	
	@Transactional(propagation=Propagation.REQUIRED)
	public int getUserId(String emailId, int roleId)
			throws DAOException {
		return (Integer) sessionFactory.getCurrentSession().getNamedQuery("fetchFacilitatorID")
				.setParameter("emailAddress", emailId)
				.setParameter("role", roleId).uniqueResult();
	}	
	

	@Transactional(propagation=Propagation.REQUIRED)
	public List<PaymentDetailsDTO> getExistingUserByTranId(String tranId) throws DAOException {
		LOGGER.info(">>>>>> AuthenticatorDAOImpl.getExistingUserByTranId");		
		return sessionFactory.getCurrentSession().getNamedQuery("getExistingUserByTranIdFromJctChk")
				.setParameter("tranId", tranId).list();
	}
	
	@Transactional(propagation=Propagation.REQUIRED)
	public List<JctUser> getExistingFacilitatorByID(String facilitatorID) throws DAOException {
		return sessionFactory.getCurrentSession().getNamedQuery("getExistingFacilitatorByID")
				.setParameter("facilitatorID", facilitatorID).list();
	}
	
	/**
	 *<p><b>Description:</b>  This method return Facilitator Email using customerID </p>
	 */
	
	@Transactional(propagation=Propagation.REQUIRED)
	public List<Object[]> getFacilitatorEmailByID(String facilitatorID) throws DAOException {
		return sessionFactory.getCurrentSession().getNamedQuery("getFacilitatorEmailByID")	
				.setParameter("facilitatorID", facilitatorID).list();
	}
	
	/**
	 *<p><b>Description:</b>  This method return Facilitator Email using customerID </p>
	 */
	
	@Transactional(propagation=Propagation.REQUIRED)
	public List<Object[]> getFacilitatorByMailId(String facilitatorID) throws DAOException {
		return sessionFactory.getCurrentSession().getNamedQuery("getFacilitatorByMailId")	
				.setParameter("facilitatorID", facilitatorID).list();
	}

	@Transactional(propagation=Propagation.REQUIRED)
	public List<Object[]> getTextSurveyQuestion() throws DAOException {
		return sessionFactory.getCurrentSession().getNamedQuery("getTextSurveyQuestion").list();
	}

	@Transactional(propagation=Propagation.REQUIRED)
	public List<Object[]> getSurveyMainQuestion(int ansType)
			throws DAOException {
		return sessionFactory.getCurrentSession().getNamedQuery("getMainQuestionForMultiple")
				.setParameter("ansType", ansType).list();
	}

	@Transactional(propagation=Propagation.REQUIRED)
	public List<String> getAllSubQtns(int ansType, String mainQtn)
			throws DAOException {
		return sessionFactory.getCurrentSession().getNamedQuery("getSubQuestions")
				.setParameter("ansType", ansType)
				.setParameter("mainQtn", mainQtn).list();
	}

	@Transactional(propagation=Propagation.REQUIRED)
    public List<JctUser> getExistingUserByEmailAndRole(String facilitatorEmail, int role) 
    		throws DAOException {
           if( role == 5) {     //     when facilitator also have tool access
                  return sessionFactory.getCurrentSession().getNamedQuery("getExistingUserFacilitatorByEmail")
                               .setParameter("email", facilitatorEmail).list();
                               //.setParameter("roleId", role).list();                
           } else {
                  return sessionFactory.getCurrentSession().getNamedQuery("getExistingUserByEmailAndRoleId")
                               .setParameter("email", facilitatorEmail)
                               .setParameter("roleId", role).list();                  
           }    
    }
	@Transactional(propagation=Propagation.REQUIRED)
	public List<String> getUserNamesOfIndUser(List<String> emailIdList) {
		return sessionFactory.getCurrentSession().getNamedQuery("fetchIndiUsersByEmailList")
				.setParameterList("emailIdList", emailIdList).list();
	}

	@Transactional(propagation=Propagation.REQUIRED)
	public List<String> getUserNamesForFacilitator(List<String> emailIdList,
			String customerId) throws DAOException {
		return sessionFactory.getCurrentSession().getNamedQuery("fetchUsersForFacilitatorEmailList")
				.setParameterList("emailIdList", emailIdList)
				.setParameter("customerId", customerId).list();
	}

	@Transactional(propagation=Propagation.REQUIRED)
	public List<String> isIndividualUserExist(String emailId)
			throws DAOException {
		return sessionFactory.getCurrentSession().getNamedQuery("isIndividualUserExist")
				.setParameter("emailId", emailId).list();
	}
	
	@Transactional(propagation=Propagation.REQUIRED)
	public List<String> shopifyUserChecker(java.lang.String emailId, int userType)
			throws DAOException {
		return sessionFactory.getCurrentSession().getNamedQuery("shopifyUserCheckerQry")
				.setParameter("emailId", emailId)
				.setParameter("userType", userType).list();
	}

	@Transactional(propagation=Propagation.REQUIRED)
	public List<String> isIndividualActiveUserExist(String emailId)
			throws DAOException {
		return sessionFactory.getCurrentSession().getNamedQuery("isIndividualActiveUserExist")
				.setParameter("emailId", emailId).list();
	}
	
	@Transactional(propagation=Propagation.REQUIRED)
	public List<String> isIndividualInactiveUserExist(String emailId)
			throws DAOException {
		return sessionFactory.getCurrentSession().getNamedQuery("isIndividualInactiveUserExist")
				.setParameter("emailId", emailId).list();
	}
	
	@Transactional(propagation=Propagation.REQUIRED)
	public List<UserAccountDTO> getUserEndDateDTO(String emailId) throws DAOException {
		return sessionFactory.getCurrentSession().getNamedQuery("getAccExpDateByEmailId")
				.setParameter("emailId", emailId).list();
	}
	
	@Transactional(propagation = Propagation.REQUIRED)
	public List<OccupationListDTO> searchOccupationList(String searchString)
			throws DAOException {
		return sessionFactory.getCurrentSession().getNamedQuery("getAllActiveOccupationList")
				.setParameter("searchString1", searchString)
				.setParameter("searchString2", "%"+searchString)
				.setParameter("searchString3", "%"+searchString+"%")
				.setParameter("searchString4", searchString+"%").list();
	}

	@Transactional(propagation = Propagation.REQUIRED)
	public String updateUser (String password, int activeFlag, String email) throws DAOException {
		String status = "failure";
		try{
			sessionFactory.getCurrentSession().getNamedQuery("activateFaciAndResetPwd")
				.setParameter("password", password)
				.setParameter("activeFlag",activeFlag)
				.setParameter("role",3)
				.setParameter("emailId", email).executeUpdate();
			status = "success";
			LOGGER.info("<<<< AuthenticatorDAOImpl.updateLoginInfo");
		} catch(Exception e){
			LOGGER.error(e.getLocalizedMessage());
		}
		return status;
	}

	@Transactional(propagation=Propagation.REQUIRED)
	public List<JctUser> getUser(String email) throws DAOException {
		return sessionFactory.getCurrentSession().getNamedQuery("fetchFacilitatorByEmailForActivation")
				.setParameter("email", email).list();
	}

	@Transactional(propagation=Propagation.REQUIRED)
	public List<JctUserDetails> checkUsersCreatedByFacilitator(int userId)
			throws DAOException {
		return sessionFactory.getCurrentSession().getNamedQuery("checkUsersCreatedByFacilitator")
				.setParameter("userId", userId).list();
	}

	@Transactional(propagation=Propagation.REQUIRED)
	public String deleteUserByUserId(int userId) throws DAOException {
		String result = "success";
		Integer successFailureFlag = 0; // 0: success
		try {
			successFailureFlag = (Integer) sessionFactory.getCurrentSession()
					.createSQLQuery("SELECT jct_delete_users(" + userId + ")")
					.uniqueResult();
			if (successFailureFlag == 1) {
				result = "failure";
			}
		} catch (Exception ex) {
			result = "failure";
		}
		return result;
	}
	
	@Transactional(propagation = Propagation.REQUIRED)
	public JctUserProfile getUserProfileByPk(int userProfileId) throws DAOException {		
		LOGGER.info(">>>> AuthenticatorDAOImpl.getUserProfileByPk");
		return (JctUserProfile) sessionFactory.getCurrentSession().getNamedQuery("fetchProfileIdByPk")
				.setParameter("userProfileId", userProfileId).uniqueResult();	
	}	
	
	@Transactional(propagation = Propagation.REQUIRED)
	public String getTermsAndConditions(JctUserProfile userProfile, int userType) 
			throws DAOException{
		LOGGER.info(">>>> AuthenticatorDAOImpl.getTermsAndConditions");
		return (String) sessionFactory.getCurrentSession().getNamedQuery("fetchTnCByProfileId")
				.setParameter("userProfile", userProfile.getJctUserProfile())
				.setParameter("userType", userType)
				.uniqueResult();
	}
	
	@Transactional(propagation = Propagation.REQUIRED)
	public List<Object[]> getAllSurveyQuestion() throws DAOException {
		return sessionFactory.getCurrentSession().getNamedQuery("getAllSurveyFaciQuestions").list();
	}
}