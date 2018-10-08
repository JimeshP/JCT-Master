package com.vmware.jct.service.vo;
/**
 * 
 * <p><b>Class name:</b> InstructionVO.java</p>
 * <p><b>Author:</b> InterraIT</p>
 * <p><b>Purpose:</b> This class acts as a value object. Transfer data from presentation layer to controller and vice versa
 * <p><b>Description:</b>This class acts as a value object. Transfer data from presentation layer to controller and vice versa </p>
 * <p><b>Copyrights:</b> 	All rights reserved by Interra IT and should only be used for its internal application development.</p>
 * <p><b>Created Date:</b> 20/Feb/2014</p>
 * <p><b>Revision History:</b>
 * 	<li>6/May/2014: Changed the saving of action plan to dynamic.</li>
 * </p>
 */
public class InstructionVO {
	public String jctInstructionDesc;
	public String jctRelatedPageDesc;
	private Integer userProfileId;
	private String userProfileName;
	private String createdBy;
	private int primaryKeyVal;
	private int activeStatus;
	private byte[] videoFile;
	
	public Integer getUserProfileId() {
		return userProfileId;
	}
	public void setUserProfileId(Integer userProfileId) {
		this.userProfileId = userProfileId;
	}
	/**
	 * @return the userProfileName
	 */
	public String getUserProfileName() {
		return userProfileName;
	}
	/**
	 * @param userProfileName the userProfileName to set
	 */
	public void setUserProfileName(String userProfileName) {
		this.userProfileName = userProfileName;
	}
	
	/**
	 * @return the createdBy
	 */
	public String getCreatedBy() {
		return createdBy;
	}
	/**
	 * @param createdBy the createdBy to set
	 */
	public void setCreatedBy(String createdBy) {
		this.createdBy = createdBy;
	}
	/**
	 * @return the primaryKeyVal
	 */
	public int getPrimaryKeyVal() {
		return primaryKeyVal;
	}
	/**
	 * @param primaryKeyVal the primaryKeyVal to set
	 */
	public void setPrimaryKeyVal(int primaryKeyVal) {
		this.primaryKeyVal = primaryKeyVal;
	}
	/**
	 * @return the activeStatus
	 */
	public int getActiveStatus() {
		return activeStatus;
	}
	/**
	 * @param activeStatus the activeStatus to set
	 */
	public void setActiveStatus(int activeStatus) {
		this.activeStatus = activeStatus;
	}
	/**
	 * 
	 * @return
	 */
	public String getJctInstructionDesc() {
		return jctInstructionDesc;
	}
	/**
	 * 
	 * @param jctInstructionDesc
	 */
	public void setJctInstructionDesc(String jctInstructionDesc) {
		this.jctInstructionDesc = jctInstructionDesc;
	}
	/**
	 * 
	 * @return
	 */
	public String getJctRelatedPageDesc() {
		return jctRelatedPageDesc;
	}
	/**
	 * 
	 * @param jctRelatedPageDesc
	 */
	public void setJctRelatedPageDesc(String jctRelatedPageDesc) {
		this.jctRelatedPageDesc = jctRelatedPageDesc;
	}
	public byte[] getVideoFile() {
		return videoFile;
	}
	public void setVideoFile(byte[] videoFile) {
		this.videoFile = videoFile;
	}
	
}
