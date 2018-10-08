package com.vmware.jct.model;

import java.io.Serializable;
import javax.persistence.*;

import java.sql.Timestamp;


/**
 * 
 * <p><b>Class name:</b> JctAdditionalInfo.java</p>
 * <p><b>Author:</b> InterraIt</p>
 * <p><b>Description:</b>  This class basically object reference of jct_additional_info table. Store all additional information in this table </p>
 * <p><b>Copyrights:</b> 	All rights reserved by Interra IT and should only be used for its internal application development.</p>
 * <p><b>Revision History:</b>
 * 	<li>InterraIt - 31/Jan/2014 - Implement comments </li>
 * </p>
 */
@Entity
@Table(name="jct_additional_info")
public class JctAdditionalInfo implements Serializable {
	private static final long serialVersionUID = 1L;

	@Id
	@GeneratedValue(strategy=GenerationType.IDENTITY)
	@Column(name="jct_additional_info_id")
	private int jctAdditionalInfoId;

	@Column(name="jct_as_created_by")
	private String jctAsCreatedBy;

	@Column(name="jct_as_created_ts")
	private Timestamp jctAsCreatedTs;

	@Column(name="jct_as_lastmodified_by")
	private String jctAsLastmodifiedBy;

	@Column(name="jct_as_lastmodified_ts")
	private Timestamp jctAsLastmodifiedTs;

	@Column(name="jct_current_industry")
	private String jctCurrentIndustry;

	@Column(name="jct_current_org_desc")
	private String jctCurrentOrgDesc;

	@Column(name="jct_email")
	private String jctEmail;

	@Column(name="jct_gender")
	private String jctGender;

	@Column(name="jct_higest_edu_level")
	private String jctHigestEduLevel;

	@Column(name="jct_job_desc")
	private String jctJobDesc;

	@Column(name="jct_jobref_no")
	private String jctJobrefNo;

	@Column(name="jct_nationality")
	private String jctNationality;

	@Column(name="jct_shared_disgram")
	private byte jctSharedDisgram;

	@Column(name="jct_year_of_birth")
	private String jctYearOfBirth;

	@Column(name="jct_yr_current_job")
	private String jctYrCurrentJob;

	@Column(name="jct_yr_current_org")
	private String jctYrCurrentOrg;

	private int version;

	//bi-directional many-to-one association to JctStatus
	@ManyToOne
	@JoinColumn(name="jct_status_id")
	private JctStatus jctStatus;

	public JctAdditionalInfo() {
	}

	public int getJctAdditionalInfoId() {
		return this.jctAdditionalInfoId;
	}

	public void setJctAdditionalInfoId(int jctAdditionalInfoId) {
		this.jctAdditionalInfoId = jctAdditionalInfoId;
	}

	public String getJctAsCreatedBy() {
		return this.jctAsCreatedBy;
	}

	public void setJctAsCreatedBy(String jctAsCreatedBy) {
		this.jctAsCreatedBy = jctAsCreatedBy;
	}

	public Timestamp getJctAsCreatedTs() {
		return this.jctAsCreatedTs;
	}

	public void setJctAsCreatedTs(Timestamp jctAsCreatedTs) {
		this.jctAsCreatedTs = jctAsCreatedTs;
	}

	public String getJctAsLastmodifiedBy() {
		return this.jctAsLastmodifiedBy;
	}

	public void setJctAsLastmodifiedBy(String jctAsLastmodifiedBy) {
		this.jctAsLastmodifiedBy = jctAsLastmodifiedBy;
	}

	public Timestamp getJctAsLastmodifiedTs() {
		return this.jctAsLastmodifiedTs;
	}

	public void setJctAsLastmodifiedTs(Timestamp jctAsLastmodifiedTs) {
		this.jctAsLastmodifiedTs = jctAsLastmodifiedTs;
	}

	public String getJctCurrentIndustry() {
		return this.jctCurrentIndustry;
	}

	public void setJctCurrentIndustry(String jctCurrentIndustry) {
		this.jctCurrentIndustry = jctCurrentIndustry;
	}

	public String getJctCurrentOrgDesc() {
		return this.jctCurrentOrgDesc;
	}

	public void setJctCurrentOrgDesc(String jctCurrentOrgDesc) {
		this.jctCurrentOrgDesc = jctCurrentOrgDesc;
	}

	public String getJctEmail() {
		return this.jctEmail;
	}

	public void setJctEmail(String jctEmail) {
		this.jctEmail = jctEmail;
	}

	public String getJctGender() {
		return this.jctGender;
	}

	public void setJctGender(String jctGender) {
		this.jctGender = jctGender;
	}

	public String getJctHigestEduLevel() {
		return this.jctHigestEduLevel;
	}

	public void setJctHigestEduLevel(String jctHigestEduLevel) {
		this.jctHigestEduLevel = jctHigestEduLevel;
	}

	public String getJctJobDesc() {
		return this.jctJobDesc;
	}

	public void setJctJobDesc(String jctJobDesc) {
		this.jctJobDesc = jctJobDesc;
	}

	public String getJctJobrefNo() {
		return this.jctJobrefNo;
	}

	public void setJctJobrefNo(String jctJobrefNo) {
		this.jctJobrefNo = jctJobrefNo;
	}

	public String getJctNationality() {
		return this.jctNationality;
	}

	public void setJctNationality(String jctNationality) {
		this.jctNationality = jctNationality;
	}

	public byte getJctSharedDisgram() {
		return this.jctSharedDisgram;
	}

	public void setJctSharedDisgram(byte jctSharedDisgram) {
		this.jctSharedDisgram = jctSharedDisgram;
	}

	public String getJctYearOfBirth() {
		return this.jctYearOfBirth;
	}

	public void setJctYearOfBirth(String jctYearOfBirth) {
		this.jctYearOfBirth = jctYearOfBirth;
	}

	public String getJctYrCurrentJob() {
		return this.jctYrCurrentJob;
	}

	public void setJctYrCurrentJob(String jctYrCurrentJob) {
		this.jctYrCurrentJob = jctYrCurrentJob;
	}

	public String getJctYrCurrentOrg() {
		return this.jctYrCurrentOrg;
	}

	public void setJctYrCurrentOrg(String jctYrCurrentOrg) {
		this.jctYrCurrentOrg = jctYrCurrentOrg;
	}

	public int getVersion() {
		return this.version;
	}

	public void setVersion(int version) {
		this.version = version;
	}

	public JctStatus getJctStatus() {
		return this.jctStatus;
	}

	public void setJctStatus(JctStatus jctStatus) {
		this.jctStatus = jctStatus;
	}

}