package com.vmware.jct.model;

import java.io.Serializable;

import javax.persistence.*;

import java.util.Date;
/**
 * 
 * <p><b>Class name:</b> JctJobAttribute.java</p>
 * <p><b>Author:</b> InterraIt</p>
 * <p><b>Description:</b>  This class basically object reference of jct_job_attribute table </p>
 * <p><b>Copyrights:</b> 	All rights reserved by Interra IT and should only be used for its internal application development.</p>
 * <p><b>Revision History:</b>
 * 	<li>InterraIt - 31/Jan/2014 - Implement comments, introduce Named query </li>
 * </p>
 */
@NamedQueries({
	@NamedQuery(name = "fetchAllJobAttribute",
				query = "select new com.vmware.jct.dao.dto.JobAttributeDTO(jctAttribute.jctJobAttributeId," +
						"jctAttribute.jctJobAttributeCode, jctAttribute.jctJobAttributeName, " +
						"jctAttribute.jctJobAttributeDesc, jctAttribute.jctProfilesDesc, jctAttribute.jctJobAttributeOrder) " +
						"from JctJobAttribute jctAttribute where jctAttribute.jctUserProfile.jctUserProfile =:profileId " +
						"and jctAttribute.jctJobAttributeSoftDelete = 0")
	
})
@Entity
@Table(name="jct_job_attribute")
public class JctJobAttribute implements Serializable {
	private static final long serialVersionUID = 1L;

	@Id
	@Column(name="jct_job_attribute_id")
	private int jctJobAttributeId;

	@Column(name="jct_job_attribute_code")
	private String jctJobAttributeCode;

	@Column(name="jct_job_attribute_created_by")
	private String jctJobAttributeCreatedBy;

	@Temporal(TemporalType.TIMESTAMP)
	@Column(name="jct_job_attribute_created_ts")
	private Date jctJobAttributeCreatedTs;

	@Column(name="jct_job_attribute_desc")
	private String jctJobAttributeDesc;

	@Column(name="jct_job_attribute_lastmodified_by")
	private String jctJobAttributeLastmodifiedBy;

	@Temporal(TemporalType.TIMESTAMP)
	@Column(name="jct_job_attribute_lastmodified_ts")
	private Date jctJobAttributeLastmodifiedTs;

	@Column(name="jct_job_attribute_soft_delete")
	private byte jctJobAttributeSoftDelete;
	
	@Column(name="jct_job_attribute_name")
	private String jctJobAttributeName;

	private int version;

	@Column(name="jct_job_attribute_order")
	private int jctJobAttributeOrder;
	
	/*//bi-directional many-to-one association to JctStatus
	@ManyToOne
	@JoinColumn(name="jct_job_attribute_status_id")
	private JctStatus jctStatus;*/
	
	@OneToOne
	@JoinColumn(name="jct_profile_id")
	private JctUserProfile jctUserProfile;
	
	@Column(name="jct_profile_desc")
	private String jctProfilesDesc;

	public JctJobAttribute() {
	}

	public int getJctJobAttributeId() {
		return this.jctJobAttributeId;
	}

	public void setJctJobAttributeId(int jctJobAttributeId) {
		this.jctJobAttributeId = jctJobAttributeId;
	}

	public String getJctJobAttributeCode() {
		return this.jctJobAttributeCode;
	}

	public void setJctJobAttributeCode(String jctJobAttributeCode) {
		this.jctJobAttributeCode = jctJobAttributeCode;
	}

	public String getJctJobAttributeCreatedBy() {
		return this.jctJobAttributeCreatedBy;
	}

	public void setJctJobAttributeCreatedBy(String jctJobAttributeCreatedBy) {
		this.jctJobAttributeCreatedBy = jctJobAttributeCreatedBy;
	}

	public Date getJctJobAttributeCreatedTs() {
		return this.jctJobAttributeCreatedTs;
	}

	public void setJctJobAttributeCreatedTs(Date jctJobAttributeCreatedTs) {
		this.jctJobAttributeCreatedTs = jctJobAttributeCreatedTs;
	}

	public String getJctJobAttributeDesc() {
		return this.jctJobAttributeDesc;
	}

	public void setJctJobAttributeDesc(String jctJobAttributeDesc) {
		this.jctJobAttributeDesc = jctJobAttributeDesc;
	}

	public String getJctJobAttributeLastmodifiedBy() {
		return this.jctJobAttributeLastmodifiedBy;
	}

	public void setJctJobAttributeLastmodifiedBy(String jctJobAttributeLastmodifiedBy) {
		this.jctJobAttributeLastmodifiedBy = jctJobAttributeLastmodifiedBy;
	}

	public Date getJctJobAttributeLastmodifiedTs() {
		return this.jctJobAttributeLastmodifiedTs;
	}

	public void setJctJobAttributeLastmodifiedTs(Date jctJobAttributeLastmodifiedTs) {
		this.jctJobAttributeLastmodifiedTs = jctJobAttributeLastmodifiedTs;
	}

	public byte getJctJobAttributeSoftDelete() {
		return this.jctJobAttributeSoftDelete;
	}

	public void setJctJobAttributeSoftDelete(byte jctJobAttributeSoftDelete) {
		this.jctJobAttributeSoftDelete = jctJobAttributeSoftDelete;
	}

	public int getVersion() {
		return this.version;
	}

	public void setVersion(int version) {
		this.version = version;
	}

	/*public JctStatus getJctStatus() {
		return this.jctStatus;
	}

	public void setJctStatus(JctStatus jctStatus) {
		this.jctStatus = jctStatus;
	}*/
	
	public String getJctProfilesDesc() {
		return jctProfilesDesc;
	}

	public void setJctProfilesDesc(String jctProfilesDesc) {
		this.jctProfilesDesc = jctProfilesDesc;
	}

	public JctUserProfile getJctUserProfile() {
		return jctUserProfile;
	}

	public void setJctUserProfile(JctUserProfile jctUserProfile) {
		this.jctUserProfile = jctUserProfile;
	}

	public String getJctJobAttributeName() {
		return jctJobAttributeName;
	}

	public void setJctJobAttributeName(String jctJobAttributeName) {
		this.jctJobAttributeName = jctJobAttributeName;
	}

	public int getJctJobAttributeOrder() {
		return jctJobAttributeOrder;
	}

	public void setJctJobAttributeOrder(int jctJobAttributeOrder) {
		this.jctJobAttributeOrder = jctJobAttributeOrder;
	}

}