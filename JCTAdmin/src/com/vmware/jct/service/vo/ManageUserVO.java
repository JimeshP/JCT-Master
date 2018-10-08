package com.vmware.jct.service.vo;

import java.util.List;
import java.util.Map;

import com.vmware.jct.dao.dto.ExistingUserDTO;
/**
 * 
 * <p><b>Class name:</b> ManageUserVO.java</p>
 * <p><b>Author:</b> InterraIT</p>
 * <p><b>Purpose:</b> This class acts as a value object. Transfer data from presentation layer to controller and vice versa
 * <p><b>Description:</b>This class acts as a value object. Transfer data from presentation layer to controller and vice versa </p>
 * <p><b>Copyrights:</b> 	All rights reserved by Interra IT and should only be used for its internal application development.</p>
 * <p><b>Created Date:</b> 20/Feb/2014</p>
 * <p><b>Revision History:</b>
 * 	<li>6/May/2014: Changed the saving of action plan to dynamic.</li>
 * </p>
 */
public class ManageUserVO {
	private String existingUserProfileList;
	private String existingUserGroupList;
	private int statusCode;
	private String statusDesc;
	private Map<Integer, String> userProfileMap;
	private List<ExistingUserDTO> existingUsers;
	private String existing;
	private String subscribedUsersList;
	private String userEmail;
	private String facilitatorEmail;
	private int deafultProfileId;
	
	private List<String> existingEmailIds;
	
	private List<String> validEmailIds;
	private List<String> invalidEmailIds;
	private List<String> nonExistingEmailIds;
	private List<String> inactiveEmailIds;
	

	public List<String> getValidEmailIds() {
		return validEmailIds;
	}

	public void setValidEmailIds(List<String> validEmailIds) {
		this.validEmailIds = validEmailIds;
	}

	public List<String> getInvalidEmailIds() {
		return invalidEmailIds;
	}

	public void setInvalidEmailIds(List<String> invalidEmailIds) {
		this.invalidEmailIds = invalidEmailIds;
	}

	public List<String> getExistingEmailIds() {
		return existingEmailIds;
	}

	public void setExistingEmailIds(List<String> existingEmailIds) {
		this.existingEmailIds = existingEmailIds;
	}

	/**
	 * @return the existingUserProfileList
	 */
	public String getExistingUserProfileList() {
		return existingUserProfileList;
	}

	/**
	 * @param existingUserProfileList the existingUserProfileList to set
	 */
	public void setExistingUserProfileList(String existingUserProfileList) {
		this.existingUserProfileList = existingUserProfileList;
	}
	
	
	/**
	 * @return the existingUserGroupList
	 */
	public String getExistingUserGroupList() {
		return existingUserGroupList;
	}

	/**
	 * @param existingUserGroupList the existingUserGroupList to set
	 */
	public void setExistingUserGroupList(String existingUserGroupList) {
		this.existingUserGroupList = existingUserGroupList;
	}

	/**
	 * @return the satusCode
	 */
	public int getStatusCode() {
		return statusCode;
	}

	/**
	 * @param satusCode the satusCode to set
	 */
	public void setStatusCode(int statusCode) {
		this.statusCode = statusCode;
	}

	/**
	 * @return the statusDesc
	 */
	public String getStatusDesc() {
		return statusDesc;
	}

	/**
	 * @param statusDesc the statusDesc to set
	 */
	public void setStatusDesc(String statusDesc) {
		this.statusDesc = statusDesc;
	}

	/**
	 * @return the userProfileMap
	 */
	public Map<Integer, String> getUserProfileMap() {
		return userProfileMap;
	}

	/**
	 * @param userProfileMap the userProfileMap to set
	 */
	public void setUserProfileMap(Map<Integer, String> userProfileMap) {
		this.userProfileMap = userProfileMap;
	}

	/**
	 * @return the existingUsers
	 */
	public List<ExistingUserDTO> getExistingUsers() {
		return existingUsers;
	}

	/**
	 * @param existingUsers the existingUsers to set
	 */
	public void setExistingUsers(List<ExistingUserDTO> existingUsers) {
		this.existingUsers = existingUsers;
	}

	/**
	 * 
	 * @return
	 */
	public String getExisting() {
		return existing;
	}

	/**
	 * 
	 * @param existing
	 */
	public void setExisting(String existing) {
		this.existing = existing;
	}

	/**** THIS IS ADDED FOR FACILITATOR****/
	/**
	 * 
	 * @return
	 */
	public String getSubscribedUsersList() {
		return subscribedUsersList;
	}
	/**
	 * 
	 * @param subscribedUsersList
	 */
	public void setSubscribedUsersList(String subscribedUsersList) {
		this.subscribedUsersList = subscribedUsersList;
	}

	public String getUserEmail() {
		return userEmail;
	}

	public void setUserEmail(String userEmail) {
		this.userEmail = userEmail;
	}

	public String getFacilitatorEmail() {
		return facilitatorEmail;
	}

	public void setFacilitatorEmail(String facilitatorEmail) {
		this.facilitatorEmail = facilitatorEmail;
	}

	public int getDeafultProfileId() {
		return deafultProfileId;
	}

	public void setDeafultProfileId(int deafultProfileId) {
		this.deafultProfileId = deafultProfileId;
	}

	public List<String> getNonExistingEmailIds() {
		return nonExistingEmailIds;
	}

	public void setNonExistingEmailIds(List<String> nonExistingEmailIds) {
		this.nonExistingEmailIds = nonExistingEmailIds;
	}

	public List<String> getInactiveEmailIds() {
		return inactiveEmailIds;
	}

	public void setInactiveEmailIds(List<String> inactiveEmailIds) {
		this.inactiveEmailIds = inactiveEmailIds;
	}
	
}
