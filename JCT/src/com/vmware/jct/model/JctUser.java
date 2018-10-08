package com.vmware.jct.model;

import java.io.Serializable;

import javax.persistence.*;

import java.util.Date;


/**
 * 
 * <p><b>Class name:</b> JctUser.java</p>
 * <p><b>Author:</b> InterraIT</p>
 * <p><b>Description:</b>  This class basically object reference of jct_user table </p>
 * <p><b>Copyrights:</b> 	All rights reserved by Interra IT and should only be used for its internal application development.</p>
 * <p><b>Revision History:</b>
 * <li>InterraIT - 31/Jan/2014 - Implement comments, introduce Named query </li>
 * <li>InterraIT - 28/Oct/2014 - changed the structure</li>
 * </p>
 */

@NamedQueries({
	@NamedQuery(name = "fetchJctUser",
				query = "from JctUser where jctUserEmail = :email and jctPassword = :password " +
						"and (jctActiveYn = :active or jctActiveYn = :inactive or jctActiveYn = :complete or jctActiveYn = :chequeProb) " +
						"and jctUserRole.jctRoleId = 1"),
	@NamedQuery(name="fetchJctUserByEmailId",
				query="from JctUser where jctUserEmail = :emailId and jctUserRole = 1 and jctUserSoftDelete = 0 and jctUserDisabled = 0"),
	@NamedQuery(name="fetchJctFacilitatorUserByEmailId",
				query="from JctUser where jctUserEmail = :emailId and jctUserRole = 3"),
	@NamedQuery(name="fetchJctUserByEmailIdPassword",
				query="from JctUser where jctUserEmail = :email and jctPassword = :password"),
	@NamedQuery(name="fetchUserList",
				query="from JctUser where jctUserEmail = :email and jctUserRole.jctRoleId = :roleId  and jctUserDisabled = 0"),
	@NamedQuery(name="fetchUserListAll",
				query="from JctUser where jctUserEmail = :email"),
	@NamedQuery(name="fetchProfileByEmailId",
				query="Select user.jctUserDetails.jctUserDetailsProfileId from JctUser user where jctUserEmail = :email"),
	@NamedQuery(name = "fetchDemographicByEmailId",
				query = "Select new com.vmware.jct.dao.dto.DemographicDTO(jctUser.jctUserDetails.jctUserOnetOccupation) " +
						"from JctUser jctUser where jctUser.jctUserEmail = :emailId and jctUser.jctUserRole.jctRoleId = 1"),
	@NamedQuery(name = "fetchInactiveJctUserByEmailId", 
				query = "from JctUser usr where usr.jctUserEmail = :emailId AND usr.jctUserRole.jctRoleId = 1 " +
						"AND (usr.jctActiveYn = 1 OR usr.jctActiveYn = 2)"),
	@NamedQuery(name = "fetchActiveJctUserByEmailId", 
				query = "from JctUser usr where usr.jctUserEmail = :emailId AND usr.jctUserRole.jctRoleId = 1 " +
						"AND usr.jctActiveYn = 3"),
	@NamedQuery(name = "getEmailIdByUserId", 
				query = "Select usr.jctUserEmail from JctUser usr where usr.jctUserId = :userId"),
	@NamedQuery(name = "getMailedPassword", 
				query = "Select usr.jctPassword from JctUser usr where usr.jctUserEmail = :jctEmail and usr.jctUserSoftDelete = 0 " +
						"and usr.jctUserRole = 1 and jctUserDisabled = 0"),				
	@NamedQuery(name = "getInitialPassword", 
				query = "Select usr.jctPassword from JctUser usr where usr.jctUserEmail = :jctEmail and usr.jctUserSoftDelete = 0 and usr.jctUserRole = 1"),
	@NamedQuery(name = "getFacilitatorDetails", 
				query = "Select new com.vmware.jct.service.vo.MyAccountVO(usr.jctUserDetails.jctUserDetailsFirstName, usr.jctUserDetails.jctUserDetailsLastName , usr.jctUserEmail) from JctUser usr where usr.jctUserCustomerId = :custId and usr.jctUserSoftDelete = 0 and usr.jctUserRole = 3")
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
	
	@Column(name="jct_user_disp_identifier")
	private int jctUserDispIdentifier;
	
	@OneToOne(cascade=CascadeType.ALL, mappedBy ="jctUser")
	private JctUserDetails jctUserDetails;

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
	
	public int getJctUserDisabled() {
		return jctUserDisabled;
	}

	public void setJctUserDisabled(int jctUserDisabled) {
		this.jctUserDisabled = jctUserDisabled;
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
}