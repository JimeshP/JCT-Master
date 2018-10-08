package com.vmware.jct.dao.dto;

import java.io.Serializable;


/**
 * 
 * <p><b>Class name:</b>JobAttributeDTO.java</p>
 * <p><b>Author:</b>  InterrIT</p>
 * <p><b>Purpose:</b>JobAttributeDTO provides access to user data which are fetched into them.</p>
 * <p><b>Description:</b></p>
 * <pre>
 * 		<b>Revision History:</b>
 * 		
 * </pre>
 * <p><b></b></p>
 */
public class JobAttributeDTO implements Serializable{

	private static final long serialVersionUID = -2699569190148629648L;
	
	private int jctJobAttributeId;
	private String jctJobAttributeCode;
	private String jctJobAttributeCreatedBy;
	private String jctJobAttributeName;
	private String jctJobAttributeDesc;
	private String jctJobAttributeLastmodifiedBy;		
	private byte jctJobAttributeSoftDelete;
	private String userProfileDesc;
	private int jctJobAttributeOrder;
	private int userProfileID;

	
	public JobAttributeDTO() {
		
	}
	
	public JobAttributeDTO(int jctJobAttributeId){
		this.jctJobAttributeId = jctJobAttributeId;	
	}
	
	public JobAttributeDTO(int jctJobAttributeId,String jctJobAttributeCode,String jctJobAttributeName){
		this.jctJobAttributeId = jctJobAttributeId;
		this.jctJobAttributeCode = jctJobAttributeCode;
		this.jctJobAttributeName = jctJobAttributeName;
	}
	
	public JobAttributeDTO (int jctJobAttributeId,String jctJobAttributeCode,String jctJobAttributeName, 
			String userProfileDesc, String jctJobAttributeDesc, int jctJobAttributeOrder, int userProfileID) {
		this.jctJobAttributeId = jctJobAttributeId;
		this.jctJobAttributeCode = jctJobAttributeCode;
		this.jctJobAttributeName = jctJobAttributeName;
		this.jctJobAttributeDesc = jctJobAttributeDesc;
		this.userProfileDesc = userProfileDesc;
		this.jctJobAttributeOrder = jctJobAttributeOrder;
		this.userProfileID=userProfileID;
	}

	public JobAttributeDTO(int jctJobAttributeId,String jctJobAttributeName, int jctJobAttributeOrder){
		this.jctJobAttributeId = jctJobAttributeId;		
		this.jctJobAttributeName = jctJobAttributeName;
		this.jctJobAttributeOrder = jctJobAttributeOrder;
	}
	
	/**
	 * @return the jctJobAttributeId
	 */
	public int getJctJobAttributeId() {
		return jctJobAttributeId;
	}

	/**
	 * @param jctJobAttributeId the jctJobAttributeId to set
	 */
	public void setJctJobAttributeId(int jctJobAttributeId) {
		this.jctJobAttributeId = jctJobAttributeId;
	}

	/**
	 * @return the jctJobAttributeCode
	 */
	public String getJctJobAttributeCode() {
		return jctJobAttributeCode;
	}

	/**
	 * @param jctJobAttributeCode the jctJobAttributeCode to set
	 */
	public void setJctJobAttributeCode(String jctJobAttributeCode) {
		this.jctJobAttributeCode = jctJobAttributeCode;
	}

	/**
	 * @return the jctJobAttributeCreatedBy
	 */
	public String getJctJobAttributeCreatedBy() {
		return jctJobAttributeCreatedBy;
	}

	/**
	 * @param jctJobAttributeCreatedBy the jctJobAttributeCreatedBy to set
	 */
	public void setJctJobAttributeCreatedBy(String jctJobAttributeCreatedBy) {
		this.jctJobAttributeCreatedBy = jctJobAttributeCreatedBy;
	}

	/**
	 * @return the jctJobAttributeDesc
	 */
	public String getJctJobAttributeDesc() {
		return jctJobAttributeDesc;
	}

	/**
	 * @param jctJobAttributeDesc the jctJobAttributeDesc to set
	 */
	public void setJctJobAttributeDesc(String jctJobAttributeDesc) {
		this.jctJobAttributeDesc = jctJobAttributeDesc;
	}

	/**
	 * @return the jctJobAttributeLastmodifiedBy
	 */
	public String getJctJobAttributeLastmodifiedBy() {
		return jctJobAttributeLastmodifiedBy;
	}

	/**
	 * @param jctJobAttributeLastmodifiedBy the jctJobAttributeLastmodifiedBy to set
	 */
	public void setJctJobAttributeLastmodifiedBy(
			String jctJobAttributeLastmodifiedBy) {
		this.jctJobAttributeLastmodifiedBy = jctJobAttributeLastmodifiedBy;
	}

	/**
	 * @return the jctJobAttributeSoftDelete
	 */
	public byte getJctJobAttributeSoftDelete() {
		return jctJobAttributeSoftDelete;
	}

	/**
	 * @param jctJobAttributeSoftDelete the jctJobAttributeSoftDelete to set
	 */
	public void setJctJobAttributeSoftDelete(byte jctJobAttributeSoftDelete) {
		this.jctJobAttributeSoftDelete = jctJobAttributeSoftDelete;
	}
	/**
	 * 
	 * @return
	 */
	public String getUserProfileDesc() {
		return userProfileDesc;
	}
	/**
	 * 
	 * @param userProfileDesc
	 */
	public void setUserProfileDesc(String userProfileDesc) {
		this.userProfileDesc = userProfileDesc;
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
	/**
	 * 
	 * @return
	 */
	public int getUserProfileID() {
		return userProfileID;
	}
	/**
	 * 
	 * @param userProfileID
	 */
	public void setUserProfileID(int userProfileID) {
		this.userProfileID = userProfileID;
	}

}
