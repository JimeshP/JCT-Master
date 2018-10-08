package com.vmware.jct.dao.dto;

import java.io.Serializable;
/**
 * 
 * <p><b>Class name:</b>UserGroupDTO.java</p>
 * <p><b>Author:</b>  InterrIT</p>
 * <p><b>Purpose:</b>UserGroupDTO provides access to user group data which are fetched into them.</p>
 * <p><b>Description:</b></p>
 * <pre>
 * 		<b>Revision History:</b>
 * 		
 * </pre>
 * <p><b></b></p>
 */
public class UserGroupDTO implements Serializable {
	private static final long serialVersionUID = 1L;
	
	private String userGroupName;
	private String userProfileDesc;
	private int activeStatus;
	private int jctUserGroupId;
	private int userProfileId;
	
	public UserGroupDTO() {}
	
	public UserGroupDTO(String userGroupName, int jctUserGroupId) {
		this.userGroupName = userGroupName;
		this.jctUserGroupId = jctUserGroupId;
	}
	
	
	public UserGroupDTO(String userGroupName, String userProfileDesc, int activeStatus, int jctUserGroupId, int userProfileId) {
		this.userGroupName = userGroupName;
		this.userProfileDesc = userProfileDesc;
		this.activeStatus = activeStatus;
		this.jctUserGroupId = jctUserGroupId;
		this.userProfileId = userProfileId;
	}
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
	 * @return the userProfileDesc
	 */
	public String getUserProfileDesc() {
		return userProfileDesc;
	}
	/**
	 * @param userProfileDesc the userProfileDesc to set
	 */
	public void setUserProfileDesc(String userProfileDesc) {
		this.userProfileDesc = userProfileDesc;
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
	 * @return the jctUserGroupId
	 */
	public int getJctUserGroupId() {
		return jctUserGroupId;
	}

	/**
	 * @param jctUserGroupId the jctUserGroupId to set
	 */
	public void setJctUserGroupId(int jctUserGroupId) {
		this.jctUserGroupId = jctUserGroupId;
	}
	
	/**
	 * @return the userProfileId
	 */
	public int getUserProfileId() {
		return userProfileId;
	}

	/**
	 * @param userProfileId the userProfileId to set
	 */
	public void setUserProfileId(int userProfileId) {
		this.userProfileId = userProfileId;
	}
	
}
