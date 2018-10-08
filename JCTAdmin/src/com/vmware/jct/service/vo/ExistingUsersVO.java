package com.vmware.jct.service.vo;

import java.util.List;

import com.vmware.jct.dao.dto.AccountExpiryDTO;
import com.vmware.jct.dao.dto.ExistingUserDTO;
/**
 * 
 * <p><b>Class name:</b> ExistingUsersVO.java</p>
 * <p><b>Author:</b> InterraIT</p>
 * <p><b>Purpose:</b> This class acts as a value object. Transfer data from presentation layer to controller and vice versa
 * <p><b>Description:</b>This class acts as a value object. Transfer data from presentation layer to controller and vice versa </p>
 * <p><b>Copyrights:</b> 	All rights reserved by Interra IT and should only be used for its internal application development.</p>
 * <p><b>Created Date:</b> 20/Feb/2014</p>
 * <p><b>Revision History:</b>
 * 	<li>6/May/2014: Changed the saving of action plan to dynamic.</li>
 * </p>
 */
public class ExistingUsersVO {
	private List<ExistingUserDTO> existingUsers;
	private List<AccountExpiryDTO> userForRenewal;
	private int statusCode;
	private String statusDesc;
	private int userType;
	private String userGroup;
	private List<ExistingUserDTO> existingFacilitators;
	
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
	 * @return the statusCode
	 */
	public int getStatusCode() {
		return statusCode;
	}

	/**
	 * @param statusCode the statusCode to set
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

	public int getUserType() {
		return userType;
	}

	public void setUserType(int userType) {
		this.userType = userType;
	}

	/**
	 * @return the userForRenewal
	 */
	public List<AccountExpiryDTO> getUserForRenewal() {
		return userForRenewal;
	}
	/**
	 * @param statusDesc the userForRenewal to set
	 */
	public void setUserForRenewal(List<AccountExpiryDTO> userForRenewal) {
		this.userForRenewal = userForRenewal;
	}

	/**
	 * @return the userGroup
	 */
	public String getUserGroup() {
		return userGroup;
	}

	/**
	 * @param userGroup the userGroup to set
	 */
	public void setUserGroup(String userGroup) {
		this.userGroup = userGroup;
	}

	/**
	 * @return the existingFacilitators
	 */
	public List<ExistingUserDTO> getExistingFacilitators() {
		return existingFacilitators;
	}

	/**
	 * @param userGroup the existingFacilitators to set
	 */
	public void setExistingFacilitators(List<ExistingUserDTO> existingFacilitators) {
		this.existingFacilitators = existingFacilitators;
	}
	
}
