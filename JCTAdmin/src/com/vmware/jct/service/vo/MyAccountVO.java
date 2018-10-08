package com.vmware.jct.service.vo;
/**
 * 
 * <p><b>Class name:</b> MyAccountVO.java</p>
 * <p><b>Author:</b> InterraIT</p>
 * <p><b>Purpose:</b> This class acts as a value object. Transfer data from presentation layer to controller and vice versa
 * <p><b>Description:</b>This class acts as a value object. Transfer data from presentation layer to controller and vice versa </p>
 * <p><b>Copyrights:</b> 	All rights reserved by Interra IT and should only be used for its internal application development.</p>
 * <p><b>Created Date:</b> 20/Feb/2014</p>
 * <p><b>Revision History:</b>
 * 	<li>6/May/2014: Changed the saving of action plan to dynamic.</li>
 * </p>
 */
public class MyAccountVO {
	private String userEmailId;
	private String existingUserDataList;
	private int statusCode;
	private String statusDesc;

	private String userFirstName;
	private String userLastName;
	private int primaryKey;
	private String userPassword;
	private String userName;
	
	/**
	 * 04.02.2015
	 */
	private String customerId;
	public String getCustomerId() {
		return customerId;
	}

	public void setCustomerId(String customerId) {
		this.customerId = customerId;
	}

	/**
	 * @return the satusCode
	 */
	public int getStatusCode() {
		return this.statusCode;
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
		return this.statusDesc;
	}
	
	/**
	 * @param statusDesc the statusDesc to set
	 */
	public void setStatusDesc(String statusDesc) {
		this.statusDesc = statusDesc;
	}

	/**
	 * 
	 * @return
	 */
	public String getExistingUserDataList() {
		return existingUserDataList;
	}

	/**
	 * 
	 * @param existingUserDataList
	 */
	public void setExistingUserDataList(String existingUserDataList) {
		this.existingUserDataList = existingUserDataList;
	}

	/**
	 * 
	 * @return
	 */
	public String getUserEmailId() {
		return userEmailId;
	}

	/**
	 * 
	 * @param userEmailId
	 */
	public void setUserEmailId(String userEmailId) {
		this.userEmailId = userEmailId;
	}

	/**
	 * 
	 * @return
	 */
	public String getUserFirstName() {
		return userFirstName;
	}

	/**
	 * 
	 * @param userFirstName
	 */
	public void setUserFirstName(String userFirstName) {
		this.userFirstName = userFirstName;
	}

	/**
	 * 
	 * @return
	 */
	public String getUserLastName() {
		return userLastName;
	}

	/**
	 * 
	 * @param userLastName
	 */
	public void setUserLastName(String userLastName) {
		this.userLastName = userLastName;
	}

	/**
	 * 
	 * @return
	 */
	public int getPrimaryKey() {
		return primaryKey;
	}

	/**
	 * 
	 * @param primaryKey
	 */
	public void setPrimaryKey(int primaryKey) {
		this.primaryKey = primaryKey;
	}

	/**
	 * 
	 * @return
	 */
	public String getUserPassword() {
		return userPassword;
	}

	/**
	 * 
	 * @param userPassword
	 */
	public void setUserPassword(String userPassword) {
		this.userPassword = userPassword;
	}
	/***** THIS IS ADDED FOR PUBLIC VERSION****/
	/**
	 * 
	 * @return
	 */
	public String getUserName() {
		return userName;
	}
	/**
	 * 
	 * @param userName
	 */
	public void setUserName(String userName) {
		this.userName = userName;
	}
	
}
