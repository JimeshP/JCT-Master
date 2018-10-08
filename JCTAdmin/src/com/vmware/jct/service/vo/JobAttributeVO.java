package com.vmware.jct.service.vo;
/**
 * 
 * <p><b>Class name:</b> JobAttributeVO.java</p>
 * <p><b>Author:</b> InterraIT</p>
 * <p><b>Purpose:</b> This class acts as a value object. Transfer data from presentation layer to controller and vice versa
 * <p><b>Description:</b>This class acts as a value object. Transfer data from presentation layer to controller and vice versa </p>
 * <p><b>Copyrights:</b> 	All rights reserved by Interra IT and should only be used for its internal application development.</p>
 * <p><b>Created Date:</b> 20/Feb/2014</p>
 * <p><b>Revision History:</b>
 * 	<li>6/May/2014: Changed the saving of action plan to dynamic.</li>
 * </p>
 */
public class JobAttributeVO {
	private String jctJobAttributeCode;
	private String jctJobAttributeName;
	private String jctJobAttributeDesc;
	private Integer userProfileId;
	private String userProfileName;
	private String createdBy;
	private int primaryKeyVal;
	private int activeStatus;
	private int jctJobAttributeOrder;
	
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
	public String getJctJobAttributeCode() {
		return jctJobAttributeCode;
	}
	/**
	 * 
	 * @param jctJobAttributeCode
	 */
	public void setJctJobAttributeCode(String jctJobAttributeCode) {
		this.jctJobAttributeCode = jctJobAttributeCode;
	}
	/**
	 * 
	 * @return
	 */
	public String getJctJobAttributeDesc() {
		return jctJobAttributeDesc;
	}
	/**
	 * 
	 * @param jctJobAttributeDesc
	 */
	public void setJctJobAttributeDesc(String jctJobAttributeDesc) {
		this.jctJobAttributeDesc = jctJobAttributeDesc;
	}
	/**
	 * 
	 * @return
	 */
	public String getJctJobAttributeName() {
		return jctJobAttributeName;
	}
	/**
	 * 
	 * @param jctJobAttributeName
	 */
	public void setJctJobAttributeName(String jctJobAttributeName) {
		this.jctJobAttributeName = jctJobAttributeName;
	}
	/**
	 * 
	 * @return
	 */
	public int getJctJobAttributeOrder() {
		return jctJobAttributeOrder;
	}
	/**
	 * 
	 * @param jctJobAttributeOrder
	 */
	public void setJctJobAttributeOrder(int jctJobAttributeOrder) {
		this.jctJobAttributeOrder = jctJobAttributeOrder;
	}
	
}
