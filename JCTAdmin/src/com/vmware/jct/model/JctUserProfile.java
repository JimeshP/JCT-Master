package com.vmware.jct.model;

import java.io.Serializable;
import java.util.Date;

import javax.persistence.Column;
import javax.persistence.Entity;
import javax.persistence.Id;
import javax.persistence.NamedQueries;
import javax.persistence.NamedQuery;
import javax.persistence.Table;
import javax.persistence.Temporal;
import javax.persistence.TemporalType;
import javax.persistence.Version;
/**
 * 
 * <p><b>Class name:</b> JctUserProfile.java</p>
 * <p><b>Author:</b> InterraIT</p>
 * <p><b>Description:</b>  This class basically object reference of jct_user_profile table </p>
 * <p><b>Copyrights:</b> 	All rights reserved by InterraIT and should only be used for its internal application development.</p>
 * <p><b>Revision History:</b>
 * 	<li>InterraIT - 31/Jan/2014 - Implement comments, introduce Named query </li>
 * </p>
 */
@NamedQueries({
	@NamedQuery(name = "fetchAllUserProfile",
				query = "select new com.vmware.jct.dao.dto.UserProfileDTO(userProfile.jctUserProfileDesc) from JctUserProfile userProfile " +
						"where userProfile.jctSoftDelete = 0 and jctUserProfile > 1"),
	@NamedQuery(name = "fetchAllUserProfileWithId",
				query = "select new com.vmware.jct.dao.dto.UserProfileDTO(userProfile.jctUserProfile, userProfile.jctUserProfileDesc) " +
						"from JctUserProfile userProfile where userProfile.jctSoftDelete = 0 and jctUserProfile > 1"),
	@NamedQuery(name = "fetchUserProfileByDesc",
				query = "from JctUserProfile userProfile where userProfile.jctSoftDelete = 0 and " +
						"UPPER(userProfile.jctUserProfileDesc)= :userProfile"),
	@NamedQuery(name = "fetchUserProfile",
				query = "select new com.vmware.jct.dao.dto.UserProfileDTO(userProfile.jctUserProfile, userProfile.jctUserProfileDesc) " +
						"from JctUserProfile userProfile where userProfile.jctSoftDelete = 0"),
	@NamedQuery(name = "fetchDefaultProfileId",
				query = "select userProfile.jctUserProfile from JctUserProfile userProfile " +
						"where userProfile.jctUserProfileDesc = :profileName and userProfile.jctSoftDelete = 0"),					
	@NamedQuery(name = "fetchDefaultProfileObj",
				query = "select new com.vmware.jct.dao.dto.UserProfileDTO(userProfile.jctUserProfile, userProfile.jctUserProfileDesc) " +
						"from JctUserProfile userProfile where userProfile.jctUserProfileDesc = :userProfileName and userProfile.jctSoftDelete = 0"),
	@NamedQuery(name = "fetchDefaultProfileObjByPk",
				query = "from JctUserProfile where jctUserProfile = :pk and jctSoftDelete = 0"),
	@NamedQuery(name = "fetchProfileIdByPk",
				query = "from JctUserProfile where jctUserProfile = :userProfileId and jctSoftDelete = 0")
})
@Entity
@Table(name = "jct_user_profile")
public class JctUserProfile implements Serializable{
	private static final long serialVersionUID = 1L;
	
	@Id
	@Column(name = "jct_user_profile")
	private int jctUserProfile;
	
	@Column(name = "jct_user_profile_desc")
	private String jctUserProfileDesc;
	
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

	/**
	 * @return the jctUserProfile
	 */
	public int getJctUserProfile() {
		return jctUserProfile;
	}

	/**
	 * @param jctUserProfile the jctUserProfile to set
	 */
	public void setJctUserProfile(int jctUserProfile) {
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
	 * @return the serialversionuid
	 */
	public static long getSerialversionuid() {
		return serialVersionUID;
	}
	
}
