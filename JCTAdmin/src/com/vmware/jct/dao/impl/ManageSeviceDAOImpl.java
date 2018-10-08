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
import com.vmware.jct.dao.IManageServiceDAO;
import com.vmware.jct.dao.dto.ExistingUserDTO;
import com.vmware.jct.dao.dto.FacilitatorDetailDTO;
import com.vmware.jct.dao.dto.UserGroupDTO;
import com.vmware.jct.dao.dto.UserProfileDTO;
import com.vmware.jct.exception.DAOException;
import com.vmware.jct.exception.JCTException;
import com.vmware.jct.model.JctEmailDetails;
import com.vmware.jct.model.JctPaymentHeader;
import com.vmware.jct.model.JctUser;
import com.vmware.jct.model.JctUserGroup;
import com.vmware.jct.model.JctUserProfile;
import com.vmware.jct.service.vo.UserVO;
/**
 * 
 * <p><b>Class name:</b> ManageSeviceDAOImpl.java</p>
 * <p><b>Author:</b> InterraIT</p>
 * <p><b>Purpose:</b> This is a ContentConfigDAOImpl class. This artifact is Persistence Manager layer artifact.
 * ContentConfigDAOImpl implement IManageServiceDAO interface and override the following  methods.
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
 * -updateUserToHonor()
 * -updateUser()
 * -deleteNewUser()
 * 
 * </p>
 * <p><b>Description:</b> This class used to perform insert, modify ,fetch data from db and delete operation. Save, update, delete operations are performed in DataAccessObject class </p>
 * <p><b>Revision History:</b>
 * 	<li>InterraIT - 01/Oct/2014 - Implement Exception </li>
 * <li>InterraIT -  10/Oct/2014 - Implement comments, introduce Named query, Transactional attribute </li>
 * <li>InterraIT - 31/Oct/2014 - Updated comments </li>
 * </p>
 */
@Repository
public class ManageSeviceDAOImpl extends DataAccessObject implements IManageServiceDAO {

	@Autowired
	private SessionFactory sessionFactory;
	
	private static final Logger LOGGER = LoggerFactory.getLogger(ManageSeviceDAOImpl.class);
	
	@Transactional(propagation = Propagation.REQUIRED)
	public List<UserProfileDTO> populateExistingUserProfile()
			throws DAOException {
		LOGGER.info(">>>>>> ManageSeviceDAOImpl.populateExistingUserProfile");
		List<UserProfileDTO> list = sessionFactory.getCurrentSession().getNamedQuery("fetchAllUserProfile").list();
		LOGGER.info("<<<<<< ManageSeviceDAOImpl.populateExistingUserProfile");
		return list;
	}
	
	@Transactional(propagation = Propagation.REQUIRED)
	public List<UserProfileDTO> populateExistingUserProfileWithId() throws DAOException {
		LOGGER.info(">>>>>> ManageSeviceDAOImpl.populateExistingUserProfileWithId");
		List<UserProfileDTO> list = sessionFactory.getCurrentSession().getNamedQuery("fetchAllUserProfileWithId").list();
		LOGGER.info("<<<<<< ManageSeviceDAOImpl.populateExistingUserProfileWithId");
		return list;
	}

	@Transactional(propagation = Propagation.REQUIRED)
	public String saveUserProfile(JctUserProfile userProfile) throws DAOException {
		LOGGER.info(">>>>>> ManageSeviceDAOImpl.saveUserProfile");
		return save(userProfile);
	}
	
	@Transactional(propagation = Propagation.REQUIRED)
	public String saveUserGroup(JctUserGroup userGroup) throws DAOException {
		LOGGER.info(">>>>>> ManageSeviceDAOImpl.saveUserGroup");
		return save(userGroup);
	}
	
	@Transactional(propagation = Propagation.REQUIRED)
	public JctUserGroup fetchUserGroup(int pkId) throws DAOException {
		LOGGER.info(">>>>>> ManageSeviceDAOImpl.fetchUserGroup");
		return (JctUserGroup) sessionFactory.getCurrentSession().getNamedQuery("fetchUserGroup")
				.setParameter("pkId", pkId).list().get(0);
	}

	@Transactional(propagation = Propagation.REQUIRED)
	public List<UserGroupDTO> populateExistingUserGroup(Integer profileId, String customerId, Integer roleId) throws DAOException {
		LOGGER.info(">>>>>> ManageSeviceDAOImpl.populateExistingUserGroup");
		List<UserGroupDTO> list = null;
		if(0 == profileId) {
			list = sessionFactory.getCurrentSession().getNamedQuery("fetchAllUserGroup").list();
		} else {
			if(roleId == 2) {
				list = sessionFactory.getCurrentSession().getNamedQuery("fetchUserGroupByProfileId")
						.setParameter("profileId", profileId).list();
			} else {
				list = sessionFactory.getCurrentSession().getNamedQuery("fetchUserGroupForFacilitator")
						.setParameter("profileId", profileId)
						.setParameter("customerId", customerId).list();
			}
			
		}		
		//List<UserGroupDTO> list = sessionFactory.getCurrentSession().getNamedQuery("fetchAllUserGroup").list();
		LOGGER.info("<<<<<< ManageSeviceDAOImpl.populateExistingUserGroup");
		return list;
	}
	
	@Transactional(propagation = Propagation.REQUIRED)
	public List<UserGroupDTO> populateDistinctUserGroup() throws DAOException {
		LOGGER.info(">>>>>> ManageSeviceDAOImpl.populateDistinctUserGroup");
		List<UserGroupDTO> list = sessionFactory.getCurrentSession().getNamedQuery("fetchAllDistinctUserGroup").list();
		LOGGER.info("<<<<<< ManageSeviceDAOImpl.populateDistinctUserGroup");
		return list;
	}
	
	@Transactional(propagation = Propagation.REQUIRED)
	public String updateUserGroup(JctUserGroup userGroup) throws DAOException {
		LOGGER.info(">>>>>> ManageSeviceDAOImpl.updateUserGroup");
		return update(userGroup);
	}
	
	@Transactional(propagation=Propagation.REQUIRED)
	public String cronPersister(JctEmailDetails dtls) throws DAOException {
		return save(dtls);
	}
	
	@Transactional(propagation=Propagation.REQUIRED)
	public List<JctEmailDetails> getEmailObject(String emailId) throws DAOException {
		return sessionFactory.getCurrentSession().getNamedQuery("jctEmailObjByEmail").setParameter("emailId", emailId).list();
	}
	
	@Transactional(propagation=Propagation.REQUIRED)
	public String cronUpdater(JctEmailDetails dtls) throws DAOException {
		return update(dtls);
	}

	@Transactional(propagation=Propagation.REQUIRED)
	public List<ExistingUserDTO> populateUserList(String userGroup) throws DAOException {
		LOGGER.info(">>>>>> ManageSeviceDAOImpl.populateDistinctUserGroup");
		List<ExistingUserDTO> list = null;
		if (null == userGroup) {
			list = sessionFactory.getCurrentSession().getNamedQuery("fetchAllUserList").list();
		} else {
			list = sessionFactory.getCurrentSession().getNamedQuery("fetchAllUserListByUserGroup").setParameter("grpName", userGroup).list();
		}
		LOGGER.info("<<<<<< ManageSeviceDAOImpl.populateDistinctUserGroup");
		return list;
	}
	
	@Transactional(propagation=Propagation.REQUIRED)
	public List<ExistingUserDTO> populateUserListByUserTypeAndGroup(String userGroup, String userType) throws DAOException {
		LOGGER.info(">>>>>> ManageSeviceDAOImpl.populateUserListByUserTypeAndGroup");
		List<ExistingUserDTO> list = null;
		if (userGroup.trim().equals("UGNS") && userType.trim().equals("UTNS")) {		  // None selected
			list = sessionFactory.getCurrentSession().getNamedQuery("fetchAllUserList").list();
		} else if (userGroup.trim().equals("UGNS") && !userType.trim().equals("UTNS")) {  // Only User Type Selected
			int type = Integer.parseInt(userType);
			if (type==2) {
				type = 3;
			}
			list = sessionFactory.getCurrentSession().getNamedQuery("fetchAllUserListOnlyType")
					.setParameter("uType", type).list();
		}  else if (!userGroup.trim().equals("UGNS") && userType.trim().equals("UTNS")) { // Only User Group Selected
			list = sessionFactory.getCurrentSession().getNamedQuery("fetchAllUserListOnlyGroup")
					.setParameter("uGrpId", Integer.parseInt(userGroup.split("!")[0])).list();
		} else { // Both selected
			int type = Integer.parseInt(userType);
			if (type==2) {
				type = 3;
			}
			list = sessionFactory.getCurrentSession().getNamedQuery("fetchAllUserListGroupTypeSelc")
					.setParameter("uGrpId", Integer.parseInt(userGroup.split("!")[0]))
					.setParameter("uType", type).list();
		}
		LOGGER.info("<<<<<< ManageSeviceDAOImpl.populateUserListByUserTypeAndGroup");
		return list;
	}
	
	@Transactional(propagation=Propagation.REQUIRED)
	public List<ExistingUserDTO> populateUserListForAdmin(String userGroup, String emailId, int userType) throws DAOException {
		LOGGER.info(">>>>>> ManageSeviceDAOImpl.populateUserListForAdmin");
		List<ExistingUserDTO> list = null;
		list = sessionFactory.getCurrentSession().getNamedQuery("fetchAllUserListByUserGroupForAdmin")
				.setParameter("grpName", userGroup)
				.setParameter("roleId", userType).list();
		LOGGER.info("<<<<<< ManageSeviceDAOImpl.populateUserListForAdmin");
		return list;
	}
	
	@Transactional(propagation = Propagation.REQUIRED)
	public String updateUser(JctUser user) throws DAOException {
		LOGGER.info(">>>>>> ManageSeviceDAOImpl.updateUser");
		return update(user);
	}

	/**
	 * USER WILL NOT BE DELETED. ONLY ACTIVATED OR INACTIVATED
	 */
	@Transactional(propagation = Propagation.REQUIRED)
	public List<ExistingUserDTO> populateSelectedUserList(String userGroup, String type) throws DAOException {
		LOGGER.info(">>>>>> ManageSeviceDAOImpl.populateSelectedUserList");
		List<ExistingUserDTO> list = null;
		int softDelete = 0;
		if (type.equals("active")) {
			softDelete = 1;
		} 
		list = sessionFactory.getCurrentSession().getNamedQuery("fetchAllUserListByUserGroupAndSoftDelete")
				.setParameter("grpName", userGroup).setParameter("softDelete", softDelete).list();
		LOGGER.info("<<<<<< ManageSeviceDAOImpl.populateSelectedUserList");
		return list;
	}
	
	@Transactional(propagation = Propagation.REQUIRED)
	public List<ExistingUserDTO> populateSelectedUserListAll(String userGroup, String type, int userType) throws DAOException {
		LOGGER.info(">>>>>> ManageSeviceDAOImpl.populateSelectedUserListAll");
		List<ExistingUserDTO> list = null;
		int softDelete = 0;
		if (type.equals("active")) {
			softDelete = 1;
		} 
		list = sessionFactory.getCurrentSession().getNamedQuery("fetchAllUserListByUserGroupAndSoftDeleteAll")
				.setParameter("grpName", userGroup)
				.setParameter("softDelete", softDelete)
				.setParameter("userRoleId", userType).list();
		LOGGER.info("<<<<<< ManageSeviceDAOImpl.populateSelectedUserListAll");
		return list;
	}
	
	@Transactional(propagation = Propagation.REQUIRED)
	public List<ExistingUserDTO> populateSelectedUserListForAdmin(String userGroup, String type, String emailId, int roleId) throws DAOException {
		LOGGER.info(">>>>>> ManageSeviceDAOImpl.populateSelectedUserListForAdmin");
		List<ExistingUserDTO> list = null;
		int softDelete = 0;
		if (type.equals("active")) {
			softDelete = 1;
		} 
		list = sessionFactory.getCurrentSession().getNamedQuery("fetchAllUserListByUserGroupAndSoftDeleteForAdmin")
				.setParameter("grpName", userGroup)
				.setParameter("softDelete", softDelete)
				.setParameter("createdBy", emailId)
				.setParameter("roleId", roleId).list();
		LOGGER.info("<<<<<< ManageSeviceDAOImpl.populateSelectedUserListForAdmin");
		return list;
	}
	
	@Transactional(propagation=Propagation.REQUIRED)
	public String validateUserProfile(String userProfile) throws DAOException {
		LOGGER.info(">>>>>> ManageSeviceDAOImpl.validateUserProfile");
		List<JctUserProfile> list = null;
		String status = "failure";
		try{
			list = sessionFactory.getCurrentSession().getNamedQuery("fetchUserProfileByDesc")
					.setParameter("userProfile", userProfile).list();
			if(list.size() == 0) {
				status = "failure";
			} else {
				status = "success";
			}			
		}catch(Exception ez){
			status = "failure";
			LOGGER.error(ez.getLocalizedMessage());
		}
		LOGGER.info("<<<<<< ManageSeviceDAOImpl.validateUserProfile");
		return status;
	}
	
	@Transactional(propagation=Propagation.REQUIRED)
	public String validateUserGroup(String userGroup, int userProfile, int roleId, String customerId) throws DAOException {
		LOGGER.info(">>>>>> ManageSeviceDAOImpl.validateUserGroup");
		List<JctUserProfile> list = null;
		String status = "failure";
		try{
			if(roleId == 2) {
				list = sessionFactory.getCurrentSession().getNamedQuery("fetchUserGroupByDesc")
						.setParameter("userGroup", userGroup)
						.setParameter("userProfile", userProfile).list();
			} else {
				list = sessionFactory.getCurrentSession().getNamedQuery("fetchUserGroupByDescForFacilitator")
						.setParameter("userGroup", userGroup)
						.setParameter("userProfile", userProfile)
						.setParameter("customerId", customerId).list();
			}		
			if(list.size() == 0) {
				status = "failure";
			} else {
				status = "success";
			}			
		}catch(Exception ez){
			status = "failure";
			LOGGER.error(ez.getLocalizedMessage());
		}
		LOGGER.info("<<<<<< ManageSeviceDAOImpl.validateUserGroup");
		return status;
	}
	
	@Transactional(propagation=Propagation.REQUIRED)
	public List<JctEmailDetails> getEmailDetails (int maxResultsToFetch) throws DAOException {
		LOGGER.info(">>>>>> ManageSeviceDAOImpl.getEmailDetails");
		return sessionFactory.getCurrentSession().getNamedQuery("jctEmailsByEmail").setMaxResults(maxResultsToFetch).list();
	}
	
	@Transactional(propagation=Propagation.REQUIRED)
	public String updateEmailDetails(JctEmailDetails details) throws DAOException {
		LOGGER.info(">>>>>> ManageSeviceDAOImpl.updateEmailDetails");
		return update(details);
	}

	@Transactional(propagation=Propagation.REQUIRED)
	public int updateUserToHonor(int jctUserId, int role, Date expiryDate)
			throws DAOException {
		LOGGER.info(">>>>>> ManageSeviceDAOImpl.updateUserToHonor");
		int status = 0; // failure
		try {
			if(role == 5){	//	when faci also have tool access
				sessionFactory.getCurrentSession().getNamedQuery("updateNewUserToHonorForFaciUser")
				.setParameter("jctUserId", jctUserId)
				.setParameter("expiryDate", expiryDate)
				.executeUpdate();
			}
			if(role == 3) {
				sessionFactory.getCurrentSession().getNamedQuery("updateNewFacilitatorToHonorForSingleUser")
				.setParameter("jctUserId", jctUserId)
				.setParameter("expiryDate", expiryDate)
				.executeUpdate();
				status = 1;
			} else {
				sessionFactory.getCurrentSession().getNamedQuery("updateNewUserToHonorForSingleUser")
				.setParameter("jctUserId", jctUserId)
				.setParameter("expiryDate", expiryDate)
				.executeUpdate();
				status = 1;
			}				
		} catch(Exception ez) {
			status = 0;
		}
		LOGGER.info("<<<<<< ManageSeviceDAOImpl.updateUserToHonor");
		return status;
	}
	/**
	 * delete disHonored user
	 * @param jctUserId
	 * @param role
	 * @return status
	 * **/
	@Transactional(propagation=Propagation.REQUIRED)
	public String deleteNewUser(int jctUserId, int role)
			throws DAOException {
		LOGGER.info(">>>>>> ManageSeviceDAOImpl.updateUserToDishonor");		
		String status = "failure";
		try{	// hard delete
			if(role == 5){	//	when faci also have tool access
				Delete(sessionFactory.getCurrentSession().getNamedQuery("getFaciUserObjForDishonor")
						.setParameter("jctUserId", jctUserId).uniqueResult());
			}
			Delete(sessionFactory.getCurrentSession().getNamedQuery("fetchUserByPk")
					.setParameter("pk", jctUserId).uniqueResult());
			status = "success";				
		} catch(Exception ez) {
			status = "failure";
			throw new DAOException(ez.getMessage());
		}
		LOGGER.info("<<<<<< ManageSeviceDAOImpl.updateUserToDishonor");
		return status;
	}
	/**
	 * update <b>Honored</b> user
	 * @param headerId
	 * @param custId
	 * @param createdBy
	 * @return status
	 * **/
	@Transactional(propagation=Propagation.REQUIRED)
	public int updateFacilitatorDtlToHonor(int headerId, String custId, String createdBy) throws DAOException {
		LOGGER.info(">>>>>> ManageSeviceDAOImpl.updateFacilitatorDtlToHonor");
		int status = 0; //"failure";
		try{
			int successFailureFlag = (Integer) sessionFactory.getCurrentSession()
					.createSQLQuery("SELECT jct_update_facilitator_type_after_chk_realization("+headerId+", '"+custId+"', '"+createdBy+"')").uniqueResult();
			if (successFailureFlag == 0)
				status = 1;	// "success";
		}catch(Exception ez){
			status = 0;	//"failure";
			throw new DAOException(ez.getMessage());
		}
		LOGGER.info("<<<<<< ManageSeviceDAOImpl.updateFacilitatorDtlToHonor");
		return status;
	}
	/**
	 * update payment dtls after honored
	 * @param detailId
	 * @param createdBy
	 * @return status
	 * **/
	@Transactional(propagation=Propagation.REQUIRED)
	public String updatePaymentDtlToHonor(int detailId,	String createdBy) throws DAOException {
		LOGGER.info(">>>>>> ManageSeviceDAOImpl.updatePaymentDtlToHonor");
		String status = "failure";
		try{
			sessionFactory.getCurrentSession().getNamedQuery("updatePaymentDtlAfterHonor")
				.setParameter("detailId", detailId)
				.setParameter("createdBy",createdBy)
				.setParameter("currentDate",new Date()).executeUpdate();
			status = "success";
		}catch(Exception ez){
			status = "failure";
		}
		LOGGER.info("<<<<<< ManageSeviceDAOImpl.updatePaymentDtlToHonor");
		return status;
	}
	/**
	 * update payment dtls after disHonored
	 * @param detailId
	 * @param createdBy
	 * @return status
	 * **/
	@Transactional(propagation=Propagation.REQUIRED)
	public String updatePaymentDtlToDishonor(int detailId, String createdBy) throws DAOException {
		LOGGER.info(">>>>>> ManageSeviceDAOImpl.updatePaymentDtlToHonor");
		String status = "failure";
		try{
			sessionFactory.getCurrentSession().getNamedQuery("updatePaymentDtlForDishonor")
				.setParameter("detailId", detailId)
				.setParameter("createdBy",createdBy)
				.setParameter("currentDate",new Date()).executeUpdate();
			status = "success";
		}catch(Exception ez){
			status = "failure";
		}
		LOGGER.info("<<<<<< ManageSeviceDAOImpl.updatePaymentDtlToHonor");
		return status;
	}
	
	@Transactional(propagation=Propagation.REQUIRED)
	public String updatePaymentHdrToDishonor(int headerId, String createdBy) throws DAOException {
		LOGGER.info(">>>>>> ManageSeviceDAOImpl.updatePaymentHdrToDishonor");
		String status = "failure";
		try{			
			JctPaymentHeader obj = (JctPaymentHeader) sessionFactory.getCurrentSession().getNamedQuery("fetchPaymentHeaderPk")					
					.setParameter("headerId", headerId).uniqueResult();	
			obj.setJctPmtHdrSoftDelete(1);
			obj.setJctPmtHdrModifiedBy(createdBy);
			obj.setJctPmtHdrModifiedTs(new Date());			
			update(obj);
			status = "success";
		} catch(Exception ez) {
			status = "failure";
		}
		LOGGER.info("<<<<<< ManageSeviceDAOImpl.updatePaymentHdrToDishonor");
		return status;
	}
	

	@Transactional(propagation=Propagation.REQUIRED)
	public int updateCheckPaymentDetails(int jctUserId) throws DAOException {
		LOGGER.info(">>>>>> ManageSeviceDAOImpl.updateCheckPaymentDetails");
		int status = 0; //	failure
		try{
			sessionFactory.getCurrentSession().getNamedQuery("updateCheckPaymentDetailsToHonor")
				.setParameter("jctUserId", jctUserId).executeUpdate();
			status = 1 ;//"success";
		}catch(Exception ez){
			status = 0; //"failure";
		}
		LOGGER.info("<<<<<< ManageSeviceDAOImpl.updateCheckPaymentDetails");
		return status;
	}

	@Transactional(propagation=Propagation.REQUIRED)
	public String updateCheckPaymentDetailsToDishonor(int pk) throws DAOException {
		LOGGER.info(">>>>>> ManageSeviceDAOImpl.updateCheckPaymentDetailsToDishonor");
		String status = "failure";
		try{
			sessionFactory.getCurrentSession().getNamedQuery("updateCheckPaymentDetailsToDishonor")
				.setParameter("pk", pk).executeUpdate();
			status = "success";
		}catch(Exception ez){
			status = "failure";
		}
		LOGGER.info("<<<<<< ManageSeviceDAOImpl.updateCheckPaymentDetailsToDishonor");
		return status;
	}
	
	@Transactional(propagation=Propagation.REQUIRED)
	public int updateAmountInPaymentHdr(String registeredPaymentHdrId, double costPerUser) throws DAOException {
		LOGGER.info(">>>>>> ManageSeviceDAOImpl.updateAmountInPaymentHdr");	
		int count = 0;
		String[] splitHdrId = registeredPaymentHdrId.toString().split("#");
		for(String hdrId: splitHdrId) {
				count += sessionFactory.getCurrentSession().getNamedQuery("updateAmountByHdrId")
					.setParameter("hdrId", Integer.parseInt(hdrId))
					.setParameter("updatedAmount", costPerUser).executeUpdate();
		}		
		LOGGER.info("<<<<<< ManageSeviceDAOImpl.updateAmountInPaymentHdr");
		return count;		
	}

	@Transactional(propagation=Propagation.REQUIRED)
	public UserProfileDTO getDefaultProfileObj(String userProfileName)
			throws DAOException {
		LOGGER.info(">>>>>> ManageSeviceDAOImpl.getDefaultProfileObj");
		return (UserProfileDTO) sessionFactory.getCurrentSession().getNamedQuery("fetchDefaultProfileObj")
				.setParameter("userProfileName", userProfileName)
				.list().get(0);
	}

	@Transactional(propagation=Propagation.REQUIRED)
	public UserGroupDTO getDefaultGroupObj(String userGroupName,
			String customerId, int role) throws DAOException {
		LOGGER.info(">>>>>> ManageSeviceDAOImpl.getDefaultGroupObj");
		return (UserGroupDTO) sessionFactory.getCurrentSession().getNamedQuery("fetchDefaultGroupObj")
				.setParameter("userGroupName", userGroupName)
				.setParameter("customerId", customerId)
				.setParameter("role", role)
				.list().get(0);
	}

	@Transactional(propagation=Propagation.REQUIRED)
	public JctUser checkIndividualForRenew(UserVO userVO, int roleId)
			throws DAOException {
		LOGGER.info(">>>>>> ManageSeviceDAOImpl.checkIndividualForRenew");
		return (JctUser) sessionFactory.getCurrentSession().getNamedQuery("fetchIndividualForRenew")
				.setString("emailId", userVO.getEmail())
				.setParameter("roleId", roleId).list().get(0);
	}

	@Transactional(propagation=Propagation.REQUIRED)
	public List<String> getInactiveUserListByEmail(List<String> emailIdList)
			throws DAOException {
		LOGGER.info(">>>>>> ManageSeviceDAOImpl.getInactiveUserListByEmail");
		return sessionFactory.getCurrentSession().getNamedQuery("fetchInactiveUserListByEmail")
				.setParameterList("emailIdList", emailIdList).list();
	}

	@Transactional(propagation=Propagation.REQUIRED)
	public JctUser getPasswordObject(String emailId, int role)
			throws DAOException {
		LOGGER.info(">>>>>> ManageSeviceDAOImpl.getPasswordObject");
		return (JctUser) sessionFactory.getCurrentSession().
					getNamedQuery("getExistingUserByEmailAndRoleId")
					.setParameter("email", emailId)
					.setParameter("roleId", role).uniqueResult();
	}

	@Transactional(propagation=Propagation.REQUIRED)
	public ExistingUserDTO getUserExpirationDtls(String emailId, int role)
			throws JCTException {
		LOGGER.info(">>>>>> ManageSeviceDAOImpl.getUserExpirationDtls");
		ExistingUserDTO list = null;
		list = (ExistingUserDTO) sessionFactory.getCurrentSession().getNamedQuery("getUserExpDtlsByEmailAndRole")
				.setParameter("emailId", emailId)
				.setParameter("role", role).uniqueResult();
		
		LOGGER.info("<<<<<< ManageSeviceDAOImpl.getUserExpirationDtls");
		return list;
	}
	@Transactional(propagation=Propagation.REQUIRED)
	public FacilitatorDetailDTO getFaciExpirationDtls(String customerId)
			throws JCTException {
		LOGGER.info(">>>>>> ManageSeviceDAOImpl.getUserExpirationDtls");
		FacilitatorDetailDTO list = null;
		list = (FacilitatorDetailDTO) sessionFactory.getCurrentSession().getNamedQuery("fetchFacilitatorSubRenewDtls")
				.setParameter("customerId", customerId).uniqueResult();	
		
		LOGGER.info("<<<<<< ManageSeviceDAOImpl.getUserExpirationDtls");
		return list;
	}

	@Transactional(propagation=Propagation.REQUIRED)
	public String getFirstName(String emailId, int role) throws JCTException {
		LOGGER.info(">>>>>> ManageSeviceDAOImpl.getFirstName");
		String fName = "";
		fName = (String) sessionFactory.getCurrentSession().getNamedQuery("fetchUserFirstName")
				.setParameter("userName", emailId)
				.setParameter("role", role)
				.uniqueResult();
		
		LOGGER.info("<<<<<< ManageSeviceDAOImpl.getFirstName");
		return fName;
	}

	@Transactional(propagation=Propagation.REQUIRED)
	public JctUser getFullName(String emailId) throws JCTException {
		// TODO Auto-generated method stub
		LOGGER.info(">>>>>> ManageSeviceDAOImpl.getFullName");
		JctUser userObj = null;		
		userObj = (JctUser) sessionFactory.getCurrentSession().getNamedQuery("fetchFacilitatorByEmail")
				.setParameter("facilitatorEmail", emailId)
			//	.setParameter("role", role)
				.uniqueResult();
		LOGGER.info("<<<<<< ManageSeviceDAOImpl.getFullName");
		return userObj;
	}
	
}