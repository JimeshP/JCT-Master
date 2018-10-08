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
	@NamedQuery(name = "fetchTcByProfileId",
				query = "Select jctTcDescription from JctTermsAndConditions " +
						"where jctUserProfile.jctUserProfile = :userProfile and jctTcUserType = :userType and jctTcSoftDelete = 0")
})

@Entity
@Table(name="jct_terms_and_condition")
public class JctTermsAndConditions implements Serializable {
	private static final long serialVersionUID = 1L;

	@Id
	@Column(name="jct_tc_id")
	private int jctTcId;

	@Column(name="jct_tc_description")
	private String jctTcDescription;
	
	@Column(name="jct_tc_created_by")
	private String jctTcCreated_by;
	
	@Temporal(TemporalType.TIMESTAMP)
	@Column(name="jct_tc_created_on")
	private Date jctTcCreatedOn;

	@Column(name="jct_tc_modified_by")
	private String jctTcModifiedBy;
	
	@Temporal(TemporalType.TIMESTAMP)
	@Column(name="jct_tc_modified_on")
	private Date jctTcModifiedOn;	
	
	@OneToOne
    @JoinColumn(name="jct_user_profile")
	private JctUserProfile jctUserProfile;
	
	@Column(name="jct_tc_soft_delete")
	private int jctTcSoftDelete;
	
	@Column(name="jct_tc_user_type")
	private int jctTcUserType;

	public int getJctTcId() {
		return jctTcId;
	}

	public void setJctTcId(int jctTcId) {
		this.jctTcId = jctTcId;
	}

	public String getJctTcDescription() {
		return jctTcDescription;
	}

	public void setJctTcDescription(String jctTcDescription) {
		this.jctTcDescription = jctTcDescription;
	}

	public String getJctTcCreated_by() {
		return jctTcCreated_by;
	}

	public void setJctTcCreated_by(String jctTcCreated_by) {
		this.jctTcCreated_by = jctTcCreated_by;
	}

	public String getJctTcModifiedBy() {
		return jctTcModifiedBy;
	}

	public void setJctTcModifiedBy(String createdBy) {
		this.jctTcModifiedBy = createdBy;
	}

	public int getJctTcSoftDelete() {
		return jctTcSoftDelete;
	}

	public void setJctTcSoftDelete(int jctTcSoftDelete) {
		this.jctTcSoftDelete = jctTcSoftDelete;
	}

	public Date getJctTcCreatedOn() {
		return jctTcCreatedOn;
	}

	public void setJctTcCreatedOn(Date jctTcCreatedOn) {
		this.jctTcCreatedOn = jctTcCreatedOn;
	}

	public Date getJctTcModifiedOn() {
		return jctTcModifiedOn;
	}

	public void setJctTcModifiedOn(Date jctTcModifiedOn) {
		this.jctTcModifiedOn = jctTcModifiedOn;
	}

	public JctUserProfile getJctUserProfile() {
		return jctUserProfile;
	}

	public void setJctUserProfile(JctUserProfile jctUserProfile) {
		this.jctUserProfile = jctUserProfile;
	}

	public int getJctTcUserType() {
		return jctTcUserType;
	}

	public void setJctTcUserType(int jctTcUserType) {
		this.jctTcUserType = jctTcUserType;
	}	
}