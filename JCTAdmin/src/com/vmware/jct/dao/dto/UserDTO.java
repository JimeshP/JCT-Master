package com.vmware.jct.dao.dto;

/**
 * 
 * <p><b>Class name:</b>UserDTO.java</p>
 * <p><b>Author:</b>  InterrIT</p>
 * <p><b>Purpose:</b>UserDTO provides access to user data which are fetched into them.</p>
 * <p><b>Description:</b></p>
 * <pre>
 * 		<b>Revision History:</b>
 * 		
 * </pre>
 * <p><b></b></p>
 */
public class UserDTO {

	private int userId;
	private String firstName;
	private String lastName;
	private String password;
	private String userEmail;
	private int activeStatus;
	private String name;
	private String custId;
	
	public UserDTO(int userId, String firstName, String lastName, String password, String userEmail, int activeStatus){
		this.userId=userId;
		this.firstName=firstName;
		this.lastName=lastName;
		this.password=password;
		this.userEmail=userEmail;
		this.activeStatus=activeStatus;
	}
	
	public UserDTO(int userId, String userEmail){
		this.userId=userId;		
		this.userEmail=userEmail;

	}
	public UserDTO(String name,String custId){
		this.name=name;
		this.custId=custId;
	}
	public String getName() {
		return name;
	}

	public void setName(String name) {
		this.name = name;
	}

	public String getCustId() {
		return custId;
	}

	public void setCustId(String custId) {
		this.custId = custId;
	}

	/**
	 * 
	 * @return
	 */
	public int getUserId() {
		return userId;
	}
	/**
	 * 
	 * @param userId
	 */
	public void setUserId(int userId) {
		this.userId = userId;
	}
	/**
	 * 
	 * @return
	 */
	public String getFirstName() {
		return firstName;
	}
	/**
	 * 
	 * @param firstName
	 */
	public void setFirstName(String firstName) {
		this.firstName = firstName;
	}
	/**
	 * 
	 * @return
	 */
	public String getLastName() {
		return lastName;
	}
	/**
	 * 
	 * @param lastName
	 */
	public void setLastName(String lastName) {
		this.lastName = lastName;
	}
	/**
	 * 
	 * @return
	 */
	public String getPassword() {
		return password;
	}
	/**
	 * 
	 * @param password
	 */
	public void setPassword(String password) {
		this.password = password;
	}
	/**
	 * 
	 * @return
	 */
	public String getUserEmail() {
		return userEmail;
	}
	/**
	 * 
	 * @param userEmail
	 */
	public void setUserEmail(String userEmail) {
		this.userEmail = userEmail;
	}
	/**
	 * 
	 * @return
	 */
	public int getActiveStatus() {
		return activeStatus;
	}
	/**
	 * 
	 * @param activeStatus
	 */
	public void setActiveStatus(int activeStatus) {
		this.activeStatus = activeStatus;
	}
	
}
