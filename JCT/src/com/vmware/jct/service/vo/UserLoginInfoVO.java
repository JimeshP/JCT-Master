package com.vmware.jct.service.vo;

import java.util.Date;
/**
 * 
 * <p><b>Class name:</b> UserLoginInfoVO.java</p>
 * <p><b>Author:</b> InterraIT</p>
 * <p><b>Purpose:</b> This class acts as a data transfer object.
 * <p><b>Description:</b> This class acts as a data transfer object.. </p>
 * <p><b>Copyrights:</b> 	All rights reserved by Interra IT and should only be used for its internal application development.</p>
 * <p><b>Created Date:</b> 20/Feb/2014</p>
 * <p><b>Revision History:</b>
 * 	<li>6/May/2014: Changed the saving of action plan to dynamic.</li>
 * </p>
 */
public class UserLoginInfoVO {
	private String jobReferenceNo;
	private int userId;
	private Date startTime;
	private Date endTime;
	private int statusId;
	private String pageInfo;
	private int version;
	private Date createdTs;
	private Date modifiedTs;
	private String modifiedBy;
	public int getIsCompleted() {
		return isCompleted;
	}
	public void setIsCompleted(int isCompleted) {
		this.isCompleted = isCompleted;
	}
	private String createdBy;
	private int isCompleted;
	public String getJobReferenceNo() {
		return jobReferenceNo;
	}
	public void setJobReferenceNo(String jobReferenceNo) {
		this.jobReferenceNo = jobReferenceNo;
	}
	public int getUserId() {
		return userId;
	}
	public void setUserId(int userId) {
		this.userId = userId;
	}
	public Date getStartTime() {
		return startTime;
	}
	public void setStartTime(Date startTime) {
		this.startTime = startTime;
	}
	public Date getEndTime() {
		return endTime;
	}
	public void setEndTime(Date endTime) {
		this.endTime = endTime;
	}
	public int getStatusId() {
		return statusId;
	}
	public void setStatusId(int statusId) {
		this.statusId = statusId;
	}
	public String getPageInfo() {
		return pageInfo;
	}
	public void setPageInfo(String pageInfo) {
		this.pageInfo = pageInfo;
	}
	public int getVersion() {
		return version;
	}
	public void setVersion(int version) {
		this.version = version;
	}
	public Date getCreatedTs() {
		return createdTs;
	}
	public void setCreatedTs(Date createdTs) {
		this.createdTs = createdTs;
	}
	public Date getModifiedTs() {
		return modifiedTs;
	}
	public void setModifiedTs(Date modifiedTs) {
		this.modifiedTs = modifiedTs;
	}
	public String getModifiedBy() {
		return modifiedBy;
	}
	public void setModifiedBy(String modifiedBy) {
		this.modifiedBy = modifiedBy;
	}
	public String getCreatedBy() {
		return createdBy;
	}
	public void setCreatedBy(String createdBy) {
		this.createdBy = createdBy;
	}
}