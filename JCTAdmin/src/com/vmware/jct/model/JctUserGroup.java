package com.vmware.jct.model;

import java.io.Serializable;

import java.util.Date;

import javax.persistence.Column;
import javax.persistence.Entity;
import javax.persistence.Id;
import javax.persistence.JoinColumn;
import javax.persistence.NamedQueries;
import javax.persistence.NamedQuery;
import javax.persistence.OneToOne;
import javax.persistence.Table;
import javax.persistence.Temporal;
import javax.persistence.TemporalType;
import javax.persistence.Version;
/**
 * 
 * <p><b>Class name:</b> JctUserGroup.java</p>
 * <p><b>Author:</b> InterraIT</p>
 * <p><b>Description:</b>  This class basically object reference of jct_user_group table </p>
 * <p><b>Copyrights:</b> 	All rights reserved by InterraIT and should only be used for its internal application development.</p>
 * <p><b>Revision History:</b>
 * 	<li>InterraIT - 31/Jan/2014 - Implement comments, introduce Named query </li>
 * </p>
 */
@NamedQueries({
	@NamedQuery(name = "fetchAllUserGroup",
				query = "select new com.vmware.jct.dao.dto.UserGroupDTO(userGroup.jctUserGroupDesc, userGroup.jctUserProfileDesc, " +
						"userGroup.jctActiveStatus, userGroup.jctUserGroup, userGroup.jctUserProfile.jctUserProfile) " +
						"from JctUserGroup userGroup " +
						"where userGroup.jctSoftDelete = 0 and userGroup.jctUserRoleId = 2"),
	@NamedQuery(name = "fetchAllDistinctUserGroup",
				query = "select distinct new com.vmware.jct.dao.dto.UserGroupDTO(userGroup.jctUserGroupDesc, userGroup.jctUserGroup) " +
						"from JctUserGroup userGroup where userGroup.jctSoftDelete = 0 and userGroup.jctUserRoleId = 2"),
	@NamedQuery(name = "fetchUserGroup",
				query = "from JctUserGroup userGroup where userGroup.jctUserGroup = :pkId"),
	@NamedQuery(name = "fetchUserGroupByDesc",
				query = "from JctUserGroup userGroup where userGroup.jctSoftDelete = 0 and UPPER(userGroup.jctUserGroupDesc)= :userGroup " +
						"and userGroup.jctUserProfile.jctUserProfile = :userProfile"),
	@NamedQuery(name = "fetchUserGroupByProfileId",
				query = "select new com.vmware.jct.dao.dto.UserGroupDTO(userGroup.jctUserGroupDesc, userGroup.jctUserProfileDesc, " +
						"userGroup.jctActiveStatus, userGroup.jctUserGroup, userGroup.jctUserProfile.jctUserProfile) " +
						"from JctUserGroup userGroup " +
						"where userGroup.jctSoftDelete = 0 and userGroup.jctUserProfile.jctUserProfile = :profileId and userGroup.jctUserRoleId = 2"),
	@NamedQuery(name = "fetchUserGroupId",
				query = "from JctUserGroup userGroup " +
						"where userGroup.jctUserGroup = :userGrpId and userGroup.jctSoftDelete = 0 and userGroup.jctUserCustomerId= :customerId"),
	@NamedQuery(name = "fetchUserGroupForFacilitator",
				query = "select new com.vmware.jct.dao.dto.UserGroupDTO(userGroup.jctUserGroupDesc, userGroup.jctUserProfileDesc, " +
						"userGroup.jctActiveStatus, userGroup.jctUserGroup, userGroup.jctUserProfile.jctUserProfile) " +
						"from JctUserGroup userGroup " +
						"where userGroup.jctSoftDelete = 0 and userGroup.jctUserProfile.jctUserProfile = :profileId and userGroup.jctUserCustomerId= :customerId"),
	@NamedQuery(name = "fetchUserGroupByDescForFacilitator",
				query = "from JctUserGroup userGroup where userGroup.jctSoftDelete = 0 and UPPER(userGroup.jctUserGroupDesc)= :userGroup " +
						"and userGroup.jctUserProfile.jctUserProfile = :userProfile and userGroup.jctUserCustomerId = :customerId"),
	@NamedQuery(name = "fetchUserGroupByCustId",
				query = "select distinct new com.vmware.jct.dao.dto.UserGroupDTO(userGroup.jctUserGroupDesc, userGroup.jctUserGroup) " +
						"from JctUserGroup userGroup " +
						"where userGroup.jctSoftDelete = 0 and userGroup.jctUserCustomerId= :customerId"),				
	@NamedQuery(name = "fetchDefaultGroupObj",
				query = "select distinct new com.vmware.jct.dao.dto.UserGroupDTO(userGroup.jctUserGroupDesc, userGroup.jctUserGroup) " +
						"from JctUserGroup userGroup " +
						"where userGroup.jctSoftDelete = 0 and userGroup.jctUserGroupDesc = :userGroupName " +
						"and userGroup.jctUserCustomerId = :customerId and userGroup.jctUserRoleId = :role"),
	@NamedQuery(name = "fetchUserGroupByGrpName",
				query = "from JctUserGroup userGroup " +
						"where userGroup.jctUserGroupDesc = :userGrpDesc and userGroup.jctSoftDelete = 0 and userGroup.jctUserCustomerId= :customerId"),
	@NamedQuery(name = "fetchUserGroupNameById",
			query = "select jctUserGroupDesc from JctUserGroup userGroup " +
					"where userGroup.jctUserGroup = :userGrpId and userGroup.jctSoftDelete = 0")
})
@Entity
@Table(name = "jct_user_group")
public class JctUserGroup implements Serializable{
	private static final long serialVersionUID = 1L;
	
	@Id
	@Column(name = "jct_user_group")
	private int jctUserGroup;
	
	@Column(name = "jct_user_group_desc")
	private String jctUserGroupDesc;
	
	@Column(name = "jct_soft_delete")
	private int jctSoftDelete;
	
	@Version
	@Column(name = "version")
	private int version;
	
	@Temporal(TemporalType.TIMESTAMP)
	@Column(name = "jct_created_ts")
	private Date jctCreatedTs;
	
	@Temporal(TemporalType.TIMESTAMP)
	@Column(name = "jct_lastmodified_ts")
	private Date jctLastmodifiedTs;
	
	@Column(name = "jct_lastmodified_by")
	private String jctLastmodifiedBy;
	
	@Column(name = "jct_created_by")
	private String jctCreatedBy;
	
	@OneToOne
	@JoinColumn(name="jct_user_profile")
	private JctUserProfile jctUserProfile;
	
	@Column(name = "jct_user_profile_desc")
	private String jctUserProfileDesc;
	
	@Column(name = "jct_active_status")
	private int jctActiveStatus;
	
	@Column(name = "jct_user_customer_id")
	private String jctUserCustomerId;

	@Column(name = "jct_user_role_id")
	private int jctUserRoleId;
	
	/**
	 * @return the jctUserGroup
	 */
	public int getJctUserGroup() {
		return jctUserGroup;
	}

	/**
	 * @param jctUserGroup the jctUserGroup to set
	 */
	public void setJctUserGroup(int jctUserGroup) {
		this.jctUserGroup = jctUserGroup;
	}

	/**
	 * @return the jctUserGroupDesc
	 */
	public String getJctUserGroupDesc() {
		return jctUserGroupDesc;
	}

	/**
	 * @param jctUserGroupDesc the jctUserGroupDesc to set
	 */
	public void setJctUserGroupDesc(String jctUserGroupDesc) {
		this.jctUserGroupDesc = jctUserGroupDesc;
	}

	/**
	 * @return the jctSoftDelete
	 */
	public int getJctSoftDelete() {
		return jctSoftDelete;
	}

	/**
	 * @param jctSoftDelete the jctSoftDelete to set
	 */
	public void setJctSoftDelete(int jctSoftDelete) {
		this.jctSoftDelete = jctSoftDelete;
	}

	/**
	 * @return the version
	 */
	public int getVersion() {
		return version;
	}

	/**
	 * @param version the version to set
	 */
	public void setVersion(int version) {
		this.version = version;
	}

	/**
	 * @return the jctCreatedTs
	 */
	public Date getJctCreatedTs() {
		return jctCreatedTs;
	}

	/**
	 * @param jctCreatedTs the jctCreatedTs to set
	 */
	public void setJctCreatedTs(Date jctCreatedTs) {
		this.jctCreatedTs = jctCreatedTs;
	}

	/**
	 * @return the jctLastmodifiedTs
	 */
	public Date getJctLastmodifiedTs() {
		return jctLastmodifiedTs;
	}

	/**
	 * @param jctLastmodifiedTs the jctLastmodifiedTs to set
	 */
	public void setJctLastmodifiedTs(Date jctLastmodifiedTs) {
		this.jctLastmodifiedTs = jctLastmodifiedTs;
	}

	/**
	 * @return the jctLastmodifiedBy
	 */
	public String getJctLastmodifiedBy() {
		return jctLastmodifiedBy;
	}

	/**
	 * @param jctLastmodifiedBy the jctLastmodifiedBy to set
	 */
	public void setJctLastmodifiedBy(String jctLastmodifiedBy) {
		this.jctLastmodifiedBy = jctLastmodifiedBy;
	}

	/**
	 * @return the jctCreatedBy
	 */
	public String getJctCreatedBy() {
		return jctCreatedBy;
	}

	/**
	 * @param jctCreatedBy the jctCreatedBy to set
	 */
	public void setJctCreatedBy(String jctCreatedBy) {
		this.jctCreatedBy = jctCreatedBy;
	}

	/**
	 * @return the jctUserProfile
	 */
	public JctUserProfile getJctUserProfile() {
		return jctUserProfile;
	}

	/**
	 * @param jctUserProfile the jctUserProfile to set
	 */
	public void setJctUserProfile(JctUserProfile jctUserProfile) {
		this.jctUserProfile = jctUserProfile;
	}

	/**
	 * @return the jctUserProfileDesc
	 */
	public String getJctUserProfileDesc() {
		return jctUserProfileDesc;
	}

	/**
	 * @param jctUserProfileDesc the jctUserProfileDesc to set
	 */
	public void setJctUserProfileDesc(String jctUserProfileDesc) {
		this.jctUserProfileDesc = jctUserProfileDesc;
	}

	/**
	 * @return the jctActiveStatus
	 */
	public int getJctActiveStatus() {
		return jctActiveStatus;
	}

	/**
	 * @param jctActiveStatus the jctActiveStatus to set
	 */
	public void setJctActiveStatus(int jctActiveStatus) {
		this.jctActiveStatus = jctActiveStatus;
	}

	public String getJctUserCustomerId() {
		return jctUserCustomerId;
	}

	public void setJctUserCustomerId(String jctUserCustomerId) {
		this.jctUserCustomerId = jctUserCustomerId;
	}

	public int getJctUserRoleId() {
		return jctUserRoleId;
	}

	public void setJctUserRoleId(int jctUserRoleId) {
		this.jctUserRoleId = jctUserRoleId;
	}
	
}
