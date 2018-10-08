package com.vmware.jct.service.vo;
/**
 * 
 * <p><b>Class name:</b> JctBeforeSketchHeaderVO.java</p>
 * <p><b>Author:</b> InterraIT</p>
 * <p><b>Purpose:</b> This class acts as a data transfer object.
 * <p><b>Description:</b> This class acts as a data transfer object.. </p>
 * <p><b>Copyrights:</b> 	All rights reserved by Interra IT and should only be used for its internal application development.</p>
 * <p><b>Created Date:</b> 20/Feb/2014</p>
 * <p><b>Revision History:</b>
 * 	<li>6/May/2014: Changed the saving of action plan to dynamic.</li>
 * </p>
 */
public class JctBeforeSketchHeaderVO {
	private String jobReferenceNumber;
	private String totalJson;
	private byte[] snapShot;
	private String jobTitle;
	private String initialJsonVal;
	private String createdBy;
	private String imageBase64String;
	private int timeTaken;
	private int jctUserId;
	
	public int getTimeTaken() {
		return timeTaken;
	}
	public void setTimeTaken(int timeTaken) {
		this.timeTaken = timeTaken;
	}
	public String getImageBase64String() {
		return imageBase64String;
	}
	public void setImageBase64String(String imageBase64String) {
		this.imageBase64String = imageBase64String;
	}
	public String getCreatedBy() {
		return createdBy;
	}
	public void setCreatedBy(String createdBy) {
		this.createdBy = createdBy;
	}
	public String getJobTitle() {
		return jobTitle;
	}
	public void setJobTitle(String jobTitle) {
		this.jobTitle = jobTitle;
	}
	public String getJobReferenceNumber() {
		return jobReferenceNumber;
	}
	public void setJobReferenceNumber(String jobReferenceNumber) {
		this.jobReferenceNumber = jobReferenceNumber;
	}
	public String getTotalJson() {
		return totalJson;
	}
	public void setTotalJson(String totalJson) {
		this.totalJson = totalJson;
	}
	
	public byte[] getSnapShot() {
		return snapShot;
	}
	public void setSnapShot(byte[] snapShot) {
		this.snapShot = snapShot;
	}
	public String getInitialJsonVal() {
		return initialJsonVal;
	}
	public void setInitialJsonVal(String initialJsonVal) {
		this.initialJsonVal = initialJsonVal;
	}
	public int getJctUserId() {
		return jctUserId;
	}
	public void setJctUserId(int jctUserId) {
		this.jctUserId = jctUserId;
	}
	
}
