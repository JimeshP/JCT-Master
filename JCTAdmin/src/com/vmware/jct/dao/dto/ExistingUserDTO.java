package com.vmware.jct.dao.dto;

import java.io.Serializable;
import java.util.Date;
/**
 * 
 * <p><b>Class name:</b>ExistingUserDTO.java</p>
 * <p><b>Author:</b>  InterrIT</p>
 * <p><b>Purpose:</b>ExistingUserDTO provides access to user data which are fetched into them.</p>
 * <p><b>Description:</b></p>
 * <pre>
 * 		<b>Revision History:</b>
 * 		
 * </pre>
 * <p><b></b></p>
 */
public class ExistingUserDTO implements Serializable {
	private static final long serialVersionUID = 1L;
	private String email;
	private String userGroup;
	private int softDelete;
	private int activeInactiveStatus;
	private Date jctAccountExpirationDate;
	private int jctProfileId;
	private String customerId;
	private String lastName;
	private String firstName;
	private Date userRegistrationDate;
	private int jctUserDetailAdminId;
	private int roleId;
	private String createdBy;
	
	public ExistingUserDTO() {}
	
	public ExistingUserDTO(String email) {
		this.email = email;
	}
	
	public ExistingUserDTO(String email, int activeInactiveStatus) {
		this.email = email;
		this.activeInactiveStatus = activeInactiveStatus;
	}
	
	public ExistingUserDTO(String email, String userGroup, int softDelete, int activeInactiveStatus) {
		this.email = email;
		this.userGroup = userGroup;
		this.softDelete = softDelete;
		this.activeInactiveStatus = activeInactiveStatus;
	}
	
	public ExistingUserDTO(String email, String userGroup, int softDelete, int activeInactiveStatus, Date jctAccountExpirationDate) {
		this.email = email;
		this.userGroup = userGroup;
		this.softDelete = softDelete;
		this.activeInactiveStatus = activeInactiveStatus;
		this.jctAccountExpirationDate = jctAccountExpirationDate;
	}
	
	public ExistingUserDTO(String email, String userGroup, int softDelete, int activeInactiveStatus, Date jctAccountExpirationDate, int jctProfileId) {
		this.email = email;
		this.userGroup = userGroup;
		this.softDelete = softDelete;
		this.activeInactiveStatus = activeInactiveStatus;
		this.jctAccountExpirationDate = jctAccountExpirationDate;
		this.jctProfileId = jctProfileId;
	}
	
	public ExistingUserDTO(String email, String userGroup, int softDelete, int activeInactiveStatus, Date jctAccountExpirationDate, int jctProfileId, String customerId) {
		this.email = email;
		this.userGroup = userGroup;
		this.softDelete = softDelete;
		this.activeInactiveStatus = activeInactiveStatus;
		this.jctAccountExpirationDate = jctAccountExpirationDate;
		this.jctProfileId = jctProfileId;
		this.customerId = customerId;
	}
	
	public ExistingUserDTO(String email, String userGroup, int softDelete, int activeInactiveStatus, Date jctAccountExpirationDate, 
			int jctProfileId, String lastName, String firstName, Date userRegistrationDate) {
		this.email = email;
		this.userGroup = userGroup;
		this.softDelete = softDelete;
		this.activeInactiveStatus = activeInactiveStatus;
		this.jctAccountExpirationDate = jctAccountExpirationDate;
		this.jctProfileId = jctProfileId;
		this.lastName = lastName;
		this.firstName = firstName;
		this.userRegistrationDate = userRegistrationDate;
	}

	public ExistingUserDTO(String email, String userGroup, int softDelete, int activeInactiveStatus, Date jctAccountExpirationDate, 
			int jctProfileId, String lastName, String firstName, Date userRegistrationDate, String customerId, int jctUserDetailAdminId, int roleId, String createdBy) {
		this.email = email;
		this.userGroup = userGroup;
		this.softDelete = softDelete;
		this.activeInactiveStatus = activeInactiveStatus;
		this.jctAccountExpirationDate = jctAccountExpirationDate;
		this.jctProfileId = jctProfileId;
		this.lastName = lastName;
		this.firstName = firstName;
		this.userRegistrationDate = userRegistrationDate;
		this.customerId = customerId;
		this.jctUserDetailAdminId = jctUserDetailAdminId;
		this.roleId = roleId;
		this.createdBy = createdBy;
	}

	/**
	 * @return the email
	 */
	public String getEmail() {
		return email;
	}

	/**
	 * @param email the email to set
	 */
	public void setEmail(String email) {
		this.email = email;
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
	 * @return the softDelete
	 */
	public int getSoftDelete() {
		return softDelete;
	}

	/**
	 * @param softDelete the softDelete to set
	 */
	public void setSoftDelete(int softDelete) {
		this.softDelete = softDelete;
	}

	/**
	 * @return the activeInactiveStatus
	 */
	public int getActiveInactiveStatus() {
		return activeInactiveStatus;
	}

	/**
	 * @param activeInactiveStatus the activeInactiveStatus to set
	 */
	public void setActiveInactiveStatus(int activeInactiveStatus) {
		this.activeInactiveStatus = activeInactiveStatus;
	}

	/**
	 * @return the jctAccountExpirationDate
	 */
	public Date getJctAccountExpirationDate() {
		return jctAccountExpirationDate;
	}
	/**
	 * @param jctAccountExpirationDate the jctAccountExpirationDate to set
	 */
	public void setJctAccountExpirationDate(Date jctAccountExpirationDate) {
		this.jctAccountExpirationDate = jctAccountExpirationDate;
	}

	public int getJctProfileId() {
		return jctProfileId;
	}

	public void setJctProfileId(int jctProfileId) {
		this.jctProfileId = jctProfileId;
	}

	public String getCustomerId() {
		return customerId;
	}

	public void setCustomerId(String customerId) {
		this.customerId = customerId;
	}

	public String getLastName() {
		return lastName;
	}

	public void setLastName(String lastName) {
		this.lastName = lastName;
	}

	public String getFirstName() {
		return firstName;
	}

	public void setFirstName(String firstName) {
		this.firstName = firstName;
	}

	public Date getUserRegistrationDate() {
		return userRegistrationDate;
	}

	public void setUserRegistrationDate(Date userRegistrationDate) {
		this.userRegistrationDate = userRegistrationDate;
	}

	public int getJctUserDetailAdminId() {
		return jctUserDetailAdminId;
	}

	public void setJctUserDetailAdminId(int jctUserDetailAdminId) {
		this.jctUserDetailAdminId = jctUserDetailAdminId;
	}

	public int getRoleId() {
		return roleId;
	}

	public void setRoleId(int roleId) {
		this.roleId = roleId;
	}

	public String getCreatedBy() {
		return createdBy;
	}

	public void setCreatedBy(String createdBy) {
		this.createdBy = createdBy;
	}
	
}
