package com.vmware.jct.service.vo;

import java.util.Date;


import com.vmware.jct.model.JctUserProfile;

public class TermsAndConditionsVO {	
	private int jctTcId;
	private String jctTcDescription;
	private String jctTcCreated_by;
	private Date jctTcCreatedOn;
	private int jctTcModifiedBy;
	private Date jctTcModifiedOn;
	private JctUserProfile jctUserProfile;
	private String jctTcSoftDelete;
	
	private int StatusCode;
	private String statusDesc;
	
	
	
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
	public Date getJctTcCreatedOn() {
		return jctTcCreatedOn;
	}
	public void setJctTcCreatedOn(Date jctTcCreatedOn) {
		this.jctTcCreatedOn = jctTcCreatedOn;
	}
	public int getJctTcModifiedBy() {
		return jctTcModifiedBy;
	}
	public void setJctTcModifiedBy(int jctTcModifiedBy) {
		this.jctTcModifiedBy = jctTcModifiedBy;
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
	public String getJctTcSoftDelete() {
		return jctTcSoftDelete;
	}
	public void setJctTcSoftDelete(String jctTcSoftDelete) {
		this.jctTcSoftDelete = jctTcSoftDelete;
	}
	public int getStatusCode() {
		return StatusCode;
	}
	public void setStatusCode(int statusCode) {
		StatusCode = statusCode;
	}
	public String getStatusDesc() {
		return statusDesc;
	}
	public void setStatusDesc(String statusDesc) {
		this.statusDesc = statusDesc;
	}
	
	
}