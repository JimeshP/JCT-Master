package com.vmware.jct.dao.dto;

import java.io.Serializable;
import java.util.List;
/**
 * 
 * <p><b>Class name:</b>NewUserDTO.java</p>
 * <p><b>Author:</b>  InterrIT</p>
 * <p><b>Purpose:</b>NewUserDTO provides access to new user data which are fetched into them.</p>
 * <p><b>Description:</b></p>
 * <pre>
 * 		<b>Revision History:</b>
 * 		
 * </pre>
 * <p><b></b></p>
 */
public class NewUserDTO implements Serializable {

	private static final long serialVersionUID = 1L;
	private String emailId;
	private String firstName;
	private String password;
	private String groupName;
	private List<NewUserDTO> newRegistration;
	private List<NewUserDTO> alreadyRegistered;
	/**
	 * @return the emailId
	 */
	public String getEmailId() {
		return emailId;
	}
	/**
	 * @param emailId the emailId to set
	 */
	public void setEmailId(String emailId) {
		this.emailId = emailId;
	}
	/**
	 * @return the password
	 */
	public String getPassword() {
		return password;
	}
	/**
	 * @param password the password to set
	 */
	public void setPassword(String password) {
		this.password = password;
	}
	/**
	 * @return the registeredUsers
	 */
	public List<NewUserDTO> getNewRegistration() {
		return newRegistration;
	}
	/**
	 * @param registeredUsers the registeredUsers to set
	 */
	public void setNewRegistration(List<NewUserDTO> newRegistration) {
		this.newRegistration = newRegistration;
	}
	/**
	 * @return the alreadyRegistered
	 */
	public List<NewUserDTO> getAlreadyRegistered() {
		return alreadyRegistered;
	}
	/**
	 * @param alreadyRegistered the alreadyRegistered to set
	 */
	public void setAlreadyRegistered(List<NewUserDTO> alreadyRegistered) {
		this.alreadyRegistered = alreadyRegistered;
	}
	/**
	 * @return the groupName
	 */
	public String getGroupName() {
		return groupName;
	}
	/**
	 * @param groupName the groupName to set
	 */
	public void setGroupName(String groupName) {
		this.groupName = groupName;
	}
	public String getFirstName() {
		return firstName;
	}
	public void setFirstName(String firstName) {
		this.firstName = firstName;
	}
	
	
	
}
