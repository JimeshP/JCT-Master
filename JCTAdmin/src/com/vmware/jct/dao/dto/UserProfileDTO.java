package com.vmware.jct.dao.dto;

import java.io.Serializable;
/**
 * 
 * <p><b>Class name:</b>UserProfileDTO.java</p>
 * <p><b>Author:</b>  InterrIT</p>
 * <p><b>Purpose:</b>UserProfileDTO provides access to user profile data which are fetched into them.</p>
 * <p><b>Description:</b></p>
 * <pre>
 * 		<b>Revision History:</b>
 * 		
 * </pre>
 * <p><b></b></p>
 */
public class UserProfileDTO implements Serializable {
	private static final long serialVersionUID = 1L;
	
	private Integer userProfileId;
	private String userProfileDesc;
	/**
	 * Default Constructor
	 */
	public UserProfileDTO(){ }
	
	public UserProfileDTO(String userProfileDesc) {
		this.userProfileDesc = userProfileDesc;
	}
	
	public UserProfileDTO(Integer userProfileId, String userProfileDesc) {
		this.userProfileDesc = userProfileDesc;
		this.userProfileId = userProfileId;
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
}
