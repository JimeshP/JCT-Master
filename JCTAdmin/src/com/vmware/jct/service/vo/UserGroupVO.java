package com.vmware.jct.service.vo;
/**
 * 
 * <p><b>Class name:</b> UserGroupVO.java</p>
 * <p><b>Author:</b> InterraIT</p>
 * <p><b>Purpose:</b> This class acts as a value object. Transfer data from presentation layer to controller and vice versa
 * <p><b>Description:</b>This class acts as a value object. Transfer data from presentation layer to controller and vice versa </p>
 * <p><b>Copyrights:</b> 	All rights reserved by Interra IT and should only be used for its internal application development.</p>
 * <p><b>Created Date:</b> 20/Feb/2014</p>
 * <p><b>Revision History:</b>
 * 	<li>6/May/2014: Changed the saving of action plan to dynamic.</li>
 * </p>
 */
public class UserGroupVO {
	private String userGroupName;
	private String userProfileName;
	private Integer userProfileId;
	private String createdBy;
	private int primaryKeyVal;
	private int activeStatus;
	private String customerId;
	private int roleId;
	/**
	 * @return the userGroupName
	 */
	public String getUserGroupName() {
		return userGroupName;
	}
	/**
	 * @param userGroupName the userGroupName to set
	 */
	public void setUserGroupName(String userGroupName) {
		this.userGroupName = userGroupName;
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
	 * @return the userProfileId
	 */
	public Integer getUserProfileId() {
		return userProfileId;
	}
	/**
	 * @param userProfileId the userProfileId to set
	 */
	public void setUserProfileId(Integer userProfileId) {
		this.userProfileId = userProfileId;
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
	public String getCustomerId() {
		return customerId;
	}
	public void setCustomerId(String customerId) {
		this.customerId = customerId;
	}
	public int getRoleId() {
		return roleId;
	}
	public void setRoleId(int roleId) {
		this.roleId = roleId;
	}
}
