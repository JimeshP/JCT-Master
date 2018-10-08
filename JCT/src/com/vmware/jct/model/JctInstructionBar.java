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
	@NamedQuery(name = "fetchInstructionById",
				query = "from JctInstructionBar instruction where instruction.jctUserProfile.jctUserProfile = :profileId " +
						"and instruction.jctPageDetails = :relatedPage and instruction.jctInstructionBarSoftDelete = 0"),
	@NamedQuery(name = "fetchInstructionByPk", 
				query = "from JctInstructionBar instruction where instruction.jctInstructionBarId = :pk")
})
@Entity
@Table(name="jct_instruction_bar")
public class JctInstructionBar implements Serializable {
	private static final long serialVersionUID = 1L;

	@Id
	@Column(name="jct_instruction_bar_id")
	private int jctInstructionBarId;

	@Column(name="jct_page_details")
	private String jctPageDetails;

	@Column(name="jct_instruction_bar_desc")
	private String jctInstructionBarDesc;
	
	@Column(name="jct_instruction_bar_created_by")
	private String jctInstructionBarCreatedBy;

	@Temporal(TemporalType.TIMESTAMP)
	@Column(name="jct_instruction_bar_created_ts")
	private Date jctInstructionBarCreatedTs;


	@Column(name="jct_instruction_bar_lastmodified_by")
	private String jctInstructionBarLastmodifiedBy;

	@Temporal(TemporalType.TIMESTAMP)
	@Column(name="jct_instruction_bar_lastmodified_ts")
	private Date jctInstructionBarLastmodifiedTs;

	@Column(name="jct_instruction_bar_soft_delete")
	private int jctInstructionBarSoftDelete;
	
	@Column(name="jct_instruction_type")
	private String jctInstructionType;
	
	@Column(name="jct_video_path")
	private String jctVideoPath;

	private int version;	

	@OneToOne
	@JoinColumn(name="jct_profile_id")
	private JctUserProfile jctUserProfile;
	
	@Column(name="jct_profile_desc")
	private String jctProfilesDesc;
	
	
	public JctInstructionBar() {
	}


	public int getJctInstructionBarId() {
		return jctInstructionBarId;
	}


	public void setJctInstructionBarId(int jctInstructionBarId) {
		this.jctInstructionBarId = jctInstructionBarId;
	}


	public String getJctPageDetails() {
		return jctPageDetails;
	}


	public void setJctPageDetails(String jctPageDetails) {
		this.jctPageDetails = jctPageDetails;
	}


	public String getJctInstructionBarDesc() {
		return jctInstructionBarDesc;
	}


	public void setJctInstructionBarDesc(String jctInstructionBarDesc) {
		this.jctInstructionBarDesc = jctInstructionBarDesc;
	}


	public String getJctInstructionBarCreatedBy() {
		return jctInstructionBarCreatedBy;
	}


	public void setJctInstructionBarCreatedBy(String jctInstructionBarCreatedBy) {
		this.jctInstructionBarCreatedBy = jctInstructionBarCreatedBy;
	}


	public Date getJctInstructionBarCreatedTs() {
		return jctInstructionBarCreatedTs;
	}


	public void setJctInstructionBarCreatedTs(Date jctInstructionBarCreatedTs) {
		this.jctInstructionBarCreatedTs = jctInstructionBarCreatedTs;
	}


	public String getJctInstructionBarLastmodifiedBy() {
		return jctInstructionBarLastmodifiedBy;
	}


	public void setJctInstructionBarLastmodifiedBy(
			String jctInstructionBarLastmodifiedBy) {
		this.jctInstructionBarLastmodifiedBy = jctInstructionBarLastmodifiedBy;
	}


	public Date getJctInstructionBarLastmodifiedTs() {
		return jctInstructionBarLastmodifiedTs;
	}


	public void setJctInstructionBarLastmodifiedTs(
			Date jctInstructionBarLastmodifiedTs) {
		this.jctInstructionBarLastmodifiedTs = jctInstructionBarLastmodifiedTs;
	}


	public int getJctInstructionBarSoftDelete() {
		return jctInstructionBarSoftDelete;
	}


	public void setJctInstructionBarSoftDelete(int jctInstructionBarSoftDelete) {
		this.jctInstructionBarSoftDelete = jctInstructionBarSoftDelete;
	}


	public int getVersion() {
		return version;
	}


	public void setVersion(int version) {
		this.version = version;
	}


	public JctUserProfile getJctUserProfile() {
		return jctUserProfile;
	}


	public void setJctUserProfile(JctUserProfile jctUserProfile) {
		this.jctUserProfile = jctUserProfile;
	}


	public String getJctProfilesDesc() {
		return jctProfilesDesc;
	}


	public void setJctProfilesDesc(String jctProfilesDesc) {
		this.jctProfilesDesc = jctProfilesDesc;
	}


	public String getJctInstructionType() {
		return jctInstructionType;
	}


	public void setJctInstructionType(String jctInstructionType) {
		this.jctInstructionType = jctInstructionType;
	}


	public String getJctVideoPath() {
		return jctVideoPath;
	}


	public void setJctVideoPath(String jctVideoPath) {
		this.jctVideoPath = jctVideoPath;
	}

	
}