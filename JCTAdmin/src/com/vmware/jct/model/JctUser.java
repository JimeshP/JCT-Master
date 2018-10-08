package com.vmware.jct.model;

import java.io.Serializable;

import javax.persistence.*;

import java.util.Date;


/**
 * 
 * <p><b>Class name:</b> JctUser.java</p>
 * <p><b>Author:</b> InterraIT</p>
 * <p><b>Description:</b>  This class basically object reference of jct_user table </p>
 * <p><b>Copyrights:</b> 	All rights reserved by InterraIT and should only be used for its internal application development.</p>
 * <p><b>Revision History:</b>
 * 	<li>InterraIT - 31/Jan/2014 - Implement comments, introduce Named query </li>
 * </p>
 */

@NamedQueries({
	@NamedQuery(name = "fetchJctUser",
				query = "from JctUser where jctUserEmail = :email and jctPassword = :password " +
						"and (jctActiveYn = :active or jctActiveYn = :inactive or jctActiveYn = :chequeProb or jctActiveYn = :justCreated or jctActiveYn = :resetPassword) " +
						"and jctUserSoftDelete = 0 and jctUserRole IN (2,3) and jctUserDisabled = 0"),	
	@NamedQuery(name="fetchJctUserByEmailId",
				query="from JctUser where jctUserEmail = :emailId"),
	@NamedQuery(name="fetchJctUserByEmailIdList",
				query="from JctUser where jctUserEmail IN :emailIdList"),
	@NamedQuery(name="fetchJctUserByEmailIdPassword",
				query="from JctUser where jctUserEmail = :email and jctPassword = :password"),
	@NamedQuery(name="fetchUserList",
				query="from JctUser where jctUserEmail = :email"),
	@NamedQuery(name = "fetchAllUserList",
				query = "select new com.vmware.jct.dao.dto.ExistingUserDTO(user.jctUserEmail , user.jctUserDetails.jctUserDetailsGroupName, " +
						"user.jctUserSoftDelete, user.jctActiveYn, user.jctAccountExpirationDate, user.jctUserRole.jctRoleId, " +
						"user.jctUserDetails.jctUserDetailsLastName, user.jctUserDetails.jctUserDetailsFirstName, user.jctCreatedTs, " +
						"user.jctUserCustomerId, user.jctUserDetails.jctUserDetailsAdminId, user.jctUserRole.jctRoleId, user.jctCreatedBy) " +
						"from JctUser user where user.jctUserRole.jctRoleId IN(1, 3) and user.jctUserDispIdentifier = 0 and user.jctUserDisabled = 0"),
	@NamedQuery(name = "fetchAllUserListOnlyGroup",
				query = "select new com.vmware.jct.dao.dto.ExistingUserDTO(user.jctUserEmail , user.jctUserDetails.jctUserDetailsGroupName, " +
						"user.jctUserSoftDelete, user.jctActiveYn, user.jctAccountExpirationDate, user.jctUserRole.jctRoleId, " +
						"user.jctUserDetails.jctUserDetailsLastName, user.jctUserDetails.jctUserDetailsFirstName, user.jctCreatedTs, " +
						"user.jctUserCustomerId, user.jctUserDetails.jctUserDetailsAdminId, user.jctUserRole.jctRoleId, user.jctCreatedBy) " +
						"from JctUser user where user.jctUserRole.jctRoleId IN(1, 3)  and user.jctUserDetails.jctUserDetailsGroupId = :uGrpId " +
						"and user.jctUserDispIdentifier = 0 and user.jctUserDisabled = 0"),
	@NamedQuery(name = "fetchAllUserListOnlyType",
				query = "select new com.vmware.jct.dao.dto.ExistingUserDTO(user.jctUserEmail , user.jctUserDetails.jctUserDetailsGroupName, " +
						"user.jctUserSoftDelete, user.jctActiveYn, user.jctAccountExpirationDate, user.jctUserRole.jctRoleId, " +
						"user.jctUserDetails.jctUserDetailsLastName, user.jctUserDetails.jctUserDetailsFirstName, user.jctCreatedTs, " +
						"user.jctUserCustomerId, user.jctUserDetails.jctUserDetailsAdminId, user.jctUserRole.jctRoleId, user.jctCreatedBy) " +
						"from JctUser user where user.jctUserRole.jctRoleId IN(1, 3)  and user.jctUserRole.jctRoleId = :uType " +
						"and user.jctUserDispIdentifier = 0 and user.jctUserDisabled = 0"),
	@NamedQuery(name = "fetchAllUserListGroupTypeSelc",
				query = "select new com.vmware.jct.dao.dto.ExistingUserDTO(user.jctUserEmail , user.jctUserDetails.jctUserDetailsGroupName, " +
						"user.jctUserSoftDelete, user.jctActiveYn, user.jctAccountExpirationDate, user.jctUserRole.jctRoleId, " +
						"user.jctUserDetails.jctUserDetailsLastName, user.jctUserDetails.jctUserDetailsFirstName, user.jctCreatedTs, " +
						"user.jctUserCustomerId, user.jctUserDetails.jctUserDetailsAdminId, user.jctUserRole.jctRoleId, user.jctCreatedBy) " +
						"from JctUser user where user.jctUserRole.jctRoleId IN(1, 3) and user.jctUserDetails.jctUserDetailsGroupId = :uGrpId " +
						"and user.jctUserRole.jctRoleId = :uType and user.jctUserDispIdentifier = 0 and user.jctUserDisabled = 0"),
	@NamedQuery(name = "fetchUserByEmail", 
				query = "from JctUser user where user.jctUserRole = 2 and user.jctUserSoftDelete = 0 and user.jctUserDisabled = 0"),
	@NamedQuery(name = "fetchUserByPk", 
				query = "from JctUser where jctUserId = :pk"),
	@NamedQuery(name = "updateSoftDeleteStatus", 
				query = "Update JctUser user set user.jctUserSoftDelete = :softDelete where user.jctUserEmail = :emailId"),
	@NamedQuery(name = "fetchAllUserListByUserGroup",
				query = "select new com.vmware.jct.dao.dto.ExistingUserDTO(user.jctUserEmail) from JctUser user " +
						"where user.jctUserDetails.jctUserDetailsProfileId > 1 and user.jctUserDetails.jctUserDetailsGroupName = :grpName " +
						"and user.jctActiveYn = 1 and user.jctUserRole = 1 and user.jctUserDispIdentifier = 0 and user.jctUserDisabled = 0"),
	@NamedQuery(name = "updateSoftDeleteStatusInBatch", 
				query = "Update JctUser user set user.jctUserSoftDelete = :softDelete where user.jctUserEmail IN  ( :emailIdList ) " +
						"and user.jctUserRole.jctRoleId = :roleId"),
	@NamedQuery(name = "fetchAllUserListByUserGroupAndSoftDelete",
				query = "select new com.vmware.jct.dao.dto.ExistingUserDTO(user.jctUserEmail) from JctUser user " +
						"where user.jctUserDetails.jctUserDetailsProfileId > 1 and user.jctUserDetails.jctUserDetailsGroupName = :grpName " +
						"and user.jctUserSoftDelete = :softDelete and user.jctUserRole.jctRoleId IN(1, 3) and user.jctUserDisabled = 0"),
	@NamedQuery(name = "fetchAllUserListByUserGroupAndSoftDeleteAll",
				query = "select new com.vmware.jct.dao.dto.ExistingUserDTO(user.jctUserEmail) from JctUser user " +
						"where user.jctUserDetails.jctUserDetailsProfileId > 1 and user.jctUserDetails.jctUserDetailsGroupName = :grpName " +
						"and user.jctUserSoftDelete = :softDelete and user.jctUserRole.jctRoleId IN(1, 3) and user.jctUserRole.jctRoleId = :userRoleId and user.jctUserDisabled = 0"),
	@NamedQuery(name = "fetchAllUserListByUserGroupAndSoftDeleteForAdmin",
				query = "select new com.vmware.jct.dao.dto.ExistingUserDTO(user.jctUserEmail) from JctUser user " +
						"where user.jctUserDetails.jctUserDetailsProfileId > 1 and user.jctUserDetails.jctUserDetailsGroupName = :grpName " +
						"and user.jctUserSoftDelete = :softDelete and user.jctCreatedBy = :createdBy and user.jctUserRole.jctRoleId = :roleId and user.jctUserDisabled = 0"),
	@NamedQuery(name = "fetchAllUserListByUserGroupForAdmin",
				query = "select new com.vmware.jct.dao.dto.ExistingUserDTO(user.jctUserEmail, user.jctActiveYn) from JctUser user where " +
						"user.jctUserDetails.jctUserDetailsProfileId > 1 and user.jctUserDetails.jctUserDetailsGroupName = :grpName and " +
						"user.jctUserRole.jctRoleId = :roleId and user.jctUserDispIdentifier = 0 and user.jctUserDisabled = 0"),
	@NamedQuery(name = "getUsrByUsrGrp",
				query = "from JctUser user where user.jctUserDetails.jctUserDetailsGroupName = :userGroupName and user.jctUserDisabled = 0"),
	@NamedQuery(name = "fetchFacilitatorByEmail", 
				query = "from JctUser user where user.jctUserEmail = :facilitatorEmail and user.jctUserRole = 3 and user.jctUserSoftDelete = 0 and user.jctUserDisabled = 0"),
	@NamedQuery(name = "fetchFacilitatorID", 
				query = "select user.jctUserId from JctUser user where user.jctUserEmail = :emailAddress and user.jctUserRole.jctRoleId = :role " +
						"and user.jctUserSoftDelete = 0 and user.jctUserDisabled = 0"),
	@NamedQuery(name="fetchUserByFacilitatorId",
				query="select new com.vmware.jct.dao.dto.ExistingUserDTO(user.jctUserEmail , user.jctUserDetails.jctUserDetailsGroupName, " +
						"user.jctUserSoftDelete, user.jctActiveYn, user.jctAccountExpirationDate, " +
						"user.jctUserId, user.jctUserDetails.jctUserDetailsLastName, user.jctUserDetails.jctUserDetailsFirstName, user.jctCreatedTs) " +
						"from JctUser user where user.jctUserCustomerId = :customerId and user.jctUserRole = 1 and user.jctUserDispIdentifier = 0 and user.jctUserDisabled = 0"),
	@NamedQuery(name = "fetchActivateInactivateUserForFacilitator",
				query = "select new com.vmware.jct.dao.dto.ExistingUserDTO(user.jctUserEmail) from JctUser user where " +
						"user.jctUserCustomerId = :customerId and user.jctUserRole.jctRoleId = 1 " +
						"and user.jctUserSoftDelete = :softDelete and user.jctUserDetails.jctUserDetailsGroupName = :userGroup and user.jctUserDisabled = 0"),				
	@NamedQuery(name = "fetchUserToResetPasswordFacilitator",
				query = "select new com.vmware.jct.dao.dto.ExistingUserDTO(user.jctUserEmail) from JctUser user " +
						"where user.jctUserCustomerId = :customerId and  user.jctUserSoftDelete = 0 and user.jctActiveYn = 1 " +
						"and user.jctUserRole.jctRoleId = 1 and user.jctUserDetails.jctUserDetailsGroupName = :userGroup and user.jctUserDisabled = 0"),											
	@NamedQuery(name = "updateSoftDeleteStatusInBatchFacilitator", 
				query = "Update JctUser user set user.jctUserSoftDelete = :softDelete where user.jctUserEmail IN  ( :emailIdList ) " +
						"and user.jctUserCustomerId = :customerId and user.jctUserRole.jctRoleId = 1 and user.jctUserDisabled = 0"),		
	@NamedQuery(name="userToResetPasswordFacilitator",
				query="from JctUser user where user.jctUserEmail = :emailId and user.jctUserDetails.jctUserDetailsFacilitatorId = :facilitatorId and user.jctUserDisabled = 0"),
	@NamedQuery(name="fetchJctUserByEmailIdAndRoleId",
				query="from JctUser user where user.jctUserEmail = :emailId and user.jctUserRole.jctRoleId = :roleId and user.jctUserDisabled = 0"),
	@NamedQuery(name="getAllActiveFacilitatorWhoCreatedUsers",
				query="SELECT DISTINCT (user.jctUserDetails.jctUserDetailsFacilitatorId) from JctUser user where " +
						"user.jctUserDetails.jctUserDetailsAdminId = 0 and user.jctUserDetails.jctUserDetailsFacilitatorId > 0 " +
						"and user.jctUserRole.jctRoleId = 1 and user.jctUserDisabled = 0"),
	@NamedQuery(name = "getAllToBeExpiredUsersRegByFacilitator",
				query = "select new com.vmware.jct.dao.dto.AccountExpiryDTO(user.jctUserEmail , user.jctAccountExpirationDate) from JctUser user " +
						"where user.jctUserDetails.jctUserDetailsFacilitatorId = :facilitatorId and user.jctUserSoftDelete = 0 and user.jctUserDisabled = 0 " +
						"and user.jctAccountExpirationDate BETWEEN :startDate AND :endDate "),
	@NamedQuery(name = "getEmailIdByUserId", 
				query = "select user.jctUserEmail from JctUser user where user.jctUserId = :jctUserId and user.jctUserRole.jctRoleId = :role " +
						"and user.jctUserSoftDelete = 0 and user.jctUserDisabled = 0"),
	@NamedQuery(name = "fetchExpiryDateToRenew", 
				query = "select user.jctAccountExpirationDate from JctUser user where user.jctUserId = :userId " +
						"and user.jctUserRole.jctRoleId = :role and user.jctUserSoftDelete = 0 and user.jctUserDisabled = 0"),
	@NamedQuery(name = "renewUserByFacilitator", 
				query = "Update JctUser user set user.jctAccountExpirationDate = :expiryDate, user.jctLastmodifiedBy = :modifiedBy, " +
						"user.lastmodifiedTs = :modifiedTs  where user.jctUserId = :userId"),
	@NamedQuery(name = "fetchAllDistinctUser",
				query = "select distinct new com.vmware.jct.dao.dto.UserDTO(user.jctUserId, user.jctUserEmail) from JctUser user " +
						"where user.jctUserSoftDelete = 0 and user.jctUserRole.jctRoleId = 1 and " +
						"user.jctUserCustomerId = :customerId and user.jctUserEmail != :facilitatorEmail and user.jctUserDisabled = 0"),
	@NamedQuery(name = "fetchAllFacilitatorList",
				query = "select new com.vmware.jct.dao.dto.ExistingUserDTO(user.jctUserEmail , user.jctUserDetails.jctUserDetailsGroupName, " +
						"user.jctUserSoftDelete, user.jctActiveYn) from JctUser user " +
						"where user.jctUserRole.jctRoleId = 3 and user.jctUserDispIdentifier = 0 and user.jctActiveYn !=  6 " +
						"and user.jctUserDisabled = 0 and user.jctUserCustomerId = :customerId"),
	@NamedQuery(name="fetchFacilitatorByEmailId",
				query="from JctUser where jctUserEmail = :emailId and jctUserRole.jctRoleId = 3 and " +
					   "jctUserSoftDelete = 0 and jctUserCustomerId = :customerId and jctUserDisabled = 0"),
	@NamedQuery(name="getExistingUserByEmailAndRoleId",
				query="from JctUser where jctUserEmail = :email and jctUserRole.jctRoleId = :roleId and jctUserSoftDelete = 0 and jctUserDisabled = 0"),
	@NamedQuery(name="getExistingUserFacilitatorByEmail",
				query="from JctUser where jctUserEmail = :email and jctUserRole.jctRoleId IN (1,3) and jctUserSoftDelete = 0 and jctUserDisabled = 0 ORDER BY jctUserRole.jctRoleId"),
	@NamedQuery(name = "fetchMaxDate",
				query = "select max(YEAR(jctCreatedTs)) from JctUser"),
	@NamedQuery(name="getExistingFacilitatorByID",
				query="from JctUser where jctUserCustomerId = :facilitatorID and jctUserSoftDelete = 0 and jctUserDisabled = 0"),
	@NamedQuery(name="getFacilitatorEmailByID",
				query="select jctUserId, jctUserEmail from JctUser " +
						"where jctUserCustomerId = :facilitatorID and jctUserSoftDelete = 0 and jctUserRole.jctRoleId = 3"),
	@NamedQuery(name="getFacilitatorByMailId",
				query="select jctUserId, jctUserEmail, jctUserCustomerId from JctUser " +
					"where jctUserName = :facilitatorID and jctUserSoftDelete = 0 and jctUserRole.jctRoleId = 3"),
	@NamedQuery(name="getCustId",
				query="select user.jctUserCustomerId,user.jctUserName from JctUser user where user.jctActiveYn = 6" +
						"and user.jctUserSoftDelete = 0 and user.jctUserDisabled = 0"),	
	@NamedQuery(name = "updateNewUserToHonorForSingleUser", 
				query = "Update JctUser user set user.jctActiveYn = 1, user.jctAccountExpirationDate = :expiryDate " +
						"where user.jctUserId = :jctUserId and user.jctUserDisabled = 0"),			
	@NamedQuery(name = "updateNewFacilitatorToHonorForSingleUser", 
				query = "Update JctUser user set user.jctActiveYn = 10, user.jctAccountExpirationDate = :expiryDate " +
						"where user.jctUserId = :jctUserId and user.jctUserDisabled = 0"),
	@NamedQuery(name = "updateNewUserToHonorForFaciUser",
				query = "Update JctUser set jctActiveYn = 1, jctAccountExpirationDate = :expiryDate " +
						"where jctUserCustomerId = (Select c.jctCheckPaymentCustomerId from JctCheckPaymentUserDetails c where c.jctUserId = :jctUserId) " +
						"AND jctUserRole = 1"),
	@NamedQuery(name = "getFaciUserObjForDishonor",
				query = "from JctUser where jctUserCustomerId = (Select c.jctCheckPaymentCustomerId from JctCheckPaymentUserDetails c where c.jctUserId = :jctUserId) " +
						"AND jctUserRole = 1"),
	@NamedQuery(name = "getUserIdByUsernameCustIdRole",
				query = "from JctUser " +
						"where jctUserEmail = :userName and jctUserRole.jctRoleId IN (1,3) and jctUserSoftDelete = 0 and jctUserDisabled = 0"),
	@NamedQuery(name = "getUserIdByUsernameCustIdRoleForSingleUser",
				query = "from JctUser " +
						"where jctUserEmail = :userName and jctUserCustomerId = :customerId and jctUserRole.jctRoleId = :role and jctUserSoftDelete = 0 " +
						"and jctUserDisabled = 0"), 					
	@NamedQuery(name = "disableUserForRefundRequestByUserId",
				query = "update JctUser set jctUserSoftDelete = 1, jctUserDisabled = 1 " +
						"where jctUserId = :userId"),
	@NamedQuery(name="fetchIndiUsersByEmailList",
				query="select user.jctUserName from JctUser user where user.jctUserEmail IN :emailIdList and user.jctUserRole.jctRoleId = 1 " +
						"and user.jctUserSoftDelete = 0 and jctUserDisabled = 0"),
	@NamedQuery(name="fetchUsersForFacilitatorEmailList",
				query="select user.jctUserName from JctUser user where user.jctUserEmail IN :emailIdList and user.jctUserRole.jctRoleId = 1 " +
						"and user.jctUserCustomerId = :customerId and user.jctUserSoftDelete = 0 and jctUserDisabled = 0"),
	@NamedQuery(name="fetchRenewedUserByFacilitatorId",
				query="select new com.vmware.jct.dao.dto.ExistingUserDTO(user.jctUserEmail , user.jctUserDetails.jctUserDetailsGroupName, " +
						"user.jctUserSoftDelete, user.jctActiveYn, user.jctAccountExpirationDate, user.jctUserId, user.jctUserCustomerId) " +
						"from JctUser user where user.jctUserRole = 1 and user.jctUserDispIdentifier = 0 " +
						"and user.jctUserDisabled = 0 and user.jctUserEmail IN :validEmailIds"),
	@NamedQuery(name="fetchIndividualForRenew",
				query="from JctUser user where user.jctUserEmail = :emailId and user.jctUserRole.jctRoleId = :roleId " +
						"and user.jctUserDisabled = 0 and user.jctUserSoftDelete = 0 and user.jctActiveYn = 3"),
	@NamedQuery(name="fetchInactiveUserListByEmail",
				query="select user.jctUserName from JctUser user where user.jctUserEmail IN :emailIdList and user.jctUserRole.jctRoleId = 1 " +
						"and user.jctUserDisabled = 0 and user.jctActiveYn != 3"),
	@NamedQuery(name="isIndividualUserExist" ,
				query = "SELECT user.jctUserName FROM JctUser user where user.jctUserName = :emailId AND user.jctUserSoftDelete = 0 AND user.jctUserRole.jctRoleId = 1"),
	@NamedQuery(name="isIndividualActiveUserExist" ,
				query = "SELECT user.jctUserName FROM JctUser user where user.jctUserName = :emailId AND user.jctUserSoftDelete = 0 AND user.jctUserRole.jctRoleId = 1"
						+ " AND user.jctActiveYn = 3"),
	@NamedQuery(name="isIndividualInactiveUserExist" ,
				query = "SELECT user.jctUserName FROM JctUser user where user.jctUserName = :emailId AND user.jctUserSoftDelete = 0 AND user.jctUserRole.jctRoleId = 1"
						+ " AND user.jctActiveYn != 3"),
	@NamedQuery(name="getAccExpDateByEmailId" ,
				query = "SELECT new com.vmware.jct.dao.dto.UserAccountDTO (user.jctAccountExpirationDate, user.jctUserCustomerId) FROM JctUser user where user.jctUserName = :emailId AND user.jctUserSoftDelete = 0 AND user.jctUserRole.jctRoleId = 1"),
	@NamedQuery(name="shopifyUserCheckerQry" ,
				query = "SELECT user.jctUserName FROM JctUser user where user.jctUserName = :emailId AND user.jctUserSoftDelete = 0 AND user.jctUserRole.jctRoleId = :userType"),
	
	@NamedQuery(name = "getUserExpDtlsByEmailAndRole",
				query = "select new com.vmware.jct.dao.dto.ExistingUserDTO(user.jctUserEmail , user.jctUserDetails.jctUserDetailsGroupName, " +
			"user.jctUserSoftDelete, user.jctActiveYn, user.jctAccountExpirationDate, user.jctUserDetails.jctUserDetailsProfileId, user.jctUserCustomerId) " +
			"from JctUser user where user.jctUserRole.jctRoleId = :role AND user.jctUserEmail = :emailId " +
			"AND user.jctUserDisabled = 0 AND user.jctUserSoftDelete = 0"),
	@NamedQuery(name = "activateFaciAndResetPwd",
				query = "Update JctUser user set user.jctPassword = :password , user.jctActiveYn = :activeFlag where user.jctUserEmail = :emailId and user.jctUserRole.jctRoleId = :role"),
	@NamedQuery(name = "fetchFacilitatorByEmailForActivation", 
				query = "from JctUser user where user.jctUserEmail = :email and user.jctUserRole.jctRoleId = 3 and user.jctUserSoftDelete = 0 and user.jctUserDisabled = 0 and user.jctActiveYn IN (10, 11)")
})

@Entity
@Table(name="jct_user")
public class JctUser implements Serializable {
	private static final long serialVersionUID = 1L;

	@Id
	@Column(name="jct_user_id")
	private int jctUserId;

	@Column(name="jct_user_name")
	private String jctUserName;
	
	@Column(name="jct_password")
	private String jctPassword;
	
	@Column(name="jct_user_email")
	private String jctUserEmail;

	@Column(name="jct_active_yn")
	private int jctActiveYn;
	
	@Column(name="jct_version")
	private int jctVersion;
	
	@Temporal(TemporalType.TIMESTAMP)
	@Column(name="jct_created_ts")
	private Date jctCreatedTs;
	
	@Column(name="jct_lastmodified_by")
	private String jctLastmodifiedBy;

	@Column(name="jct_created_by")
	private String jctCreatedBy;

	@Temporal(TemporalType.TIMESTAMP)
	@Column(name="lastmodified_ts")
	private Date lastmodifiedTs;
	
	@OneToOne
	@JoinColumn(name="jct_role_id")
	private JctUserRole jctUserRole;
	
	@Column(name = "jct_user_soft_delete")
	private int jctUserSoftDelete;
	
	@Column(name = "jct_user_disabled")
	private int jctUserDisabled;
	
	
	@Temporal(TemporalType.TIMESTAMP)
	@Column(name="jct_account_expiration_date")
	private Date jctAccountExpirationDate;
	
	@Column(name="jct_user_customer_id")
	private String jctUserCustomerId;
	
	@OneToOne(cascade=CascadeType.ALL, mappedBy ="jctUser")
	private JctUserDetails jctUserDetails;
	
	@Column(name="jct_user_disp_identifier")
	private int jctUserDispIdentifier;

	public int getJctUserId() {
		return jctUserId;
	}

	public void setJctUserId(int jctUserId) {
		this.jctUserId = jctUserId;
	}

	public String getJctUserName() {
		return jctUserName;
	}

	public void setJctUserName(String jctUserName) {
		this.jctUserName = jctUserName;
	}

	public String getJctPassword() {
		return jctPassword;
	}

	public void setJctPassword(String jctPassword) {
		this.jctPassword = jctPassword;
	}

	public String getJctUserEmail() {
		return jctUserEmail;
	}

	public void setJctUserEmail(String jctUserEmail) {
		this.jctUserEmail = jctUserEmail;
	}

	public int getJctActiveYn() {
		return jctActiveYn;
	}

	public void setJctActiveYn(int jctActiveYn) {
		this.jctActiveYn = jctActiveYn;
	}

	public int getJctVersion() {
		return jctVersion;
	}

	public void setJctVersion(int jctVersion) {
		this.jctVersion = jctVersion;
	}

	public Date getJctCreatedTs() {
		return jctCreatedTs;
	}

	public void setJctCreatedTs(Date jctCreatedTs) {
		this.jctCreatedTs = jctCreatedTs;
	}

	public String getJctLastmodifiedBy() {
		return jctLastmodifiedBy;
	}

	public void setJctLastmodifiedBy(String jctLastmodifiedBy) {
		this.jctLastmodifiedBy = jctLastmodifiedBy;
	}

	public String getJctCreatedBy() {
		return jctCreatedBy;
	}

	public void setJctCreatedBy(String jctCreatedBy) {
		this.jctCreatedBy = jctCreatedBy;
	}

	public Date getLastmodifiedTs() {
		return lastmodifiedTs;
	}

	public void setLastmodifiedTs(Date lastmodifiedTs) {
		this.lastmodifiedTs = lastmodifiedTs;
	}

	public JctUserRole getJctUserRole() {
		return jctUserRole;
	}

	public void setJctUserRole(JctUserRole jctUserRole) {
		this.jctUserRole = jctUserRole;
	}

	public int getJctUserSoftDelete() {
		return jctUserSoftDelete;
	}

	public void setJctUserSoftDelete(int jctUserSoftDelete) {
		this.jctUserSoftDelete = jctUserSoftDelete;
	}

	public Date getJctAccountExpirationDate() {
		return jctAccountExpirationDate;
	}

	public void setJctAccountExpirationDate(Date jctAccountExpirationDate) {
		this.jctAccountExpirationDate = jctAccountExpirationDate;
	}

	public String getJctUserCustomerId() {
		return jctUserCustomerId;
	}

	public void setJctUserCustomerId(String jctUserCustomerId) {
		this.jctUserCustomerId = jctUserCustomerId;
	}

	public JctUserDetails getJctUserDetails() {
		return jctUserDetails;
	}

	public void setJctUserDetails(JctUserDetails jctUserDetails) {
		this.jctUserDetails = jctUserDetails;
	}

	public int getJctUserDispIdentifier() {
		return jctUserDispIdentifier;
	}

	public void setJctUserDispIdentifier(int jctUserDispIdentifier) {
		this.jctUserDispIdentifier = jctUserDispIdentifier;
	}
	public int getJctUserDisabled() {
		return jctUserDisabled;
	}

	public void setJctUserDisabled(int jctUserDisabled) {
		this.jctUserDisabled = jctUserDisabled;
	}
}