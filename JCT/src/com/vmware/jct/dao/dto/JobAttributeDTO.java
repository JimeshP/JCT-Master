/**
 * 
 */
package com.vmware.jct.dao.dto;

import java.io.Serializable;


/**
 * 
 * <p><b>Class name:</b> JobAttributeDTO.java</p>
 * <p><b>Author:</b> InterraIT</p>
 * <p><b>Purpose:</b> This class acts as a data transfer object.
 * <p><b>Description:</b> This class acts as a data transfer object.. </p>
 * <p><b>Copyrights:</b> 	All rights reserved by Interra IT and should only be used for its internal application development.</p>
 * <p><b>Created Date:</b> 20/Feb/2014</p>
 * <p><b>Revision History:</b>
 * 	<li>6/May/2014: Changed the saving of action plan to dynamic.</li>
 * </p>
 */
public class JobAttributeDTO implements Serializable{

	/**
	 * 
	 */
	private static final long serialVersionUID = -2699569190148629648L;
	
	
	public int jctJobAttributeId;
	public String jctJobAttributeCode;
	public String jctJobAttributeCreatedBy;
	public String jctJobAttributeDesc;
	public String jctJobAttributeLastmodifiedBy;
	public byte jctJobAttributeSoftDelete;
	public String jctJobAttributeName;
	public String jctProfileDesc;
	public int jctJobAttributeOrder;
	
	public JobAttributeDTO(){
		
	}
	public JobAttributeDTO(int jctJobAttributeId,String jctJobAttributeCode,String jctJobAttributeName, 
			String jctJobAttributeDesc, String jctProfileDesc, int jctJobAttributeOrder){
		this.jctJobAttributeId    = jctJobAttributeId;
		this.jctJobAttributeCode  = jctJobAttributeCode;
		this.jctJobAttributeName  = jctJobAttributeName;
		this.jctJobAttributeDesc  = jctJobAttributeDesc;
		this.jctProfileDesc 	  = jctProfileDesc;
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
	public String getJctJobAttributeName() {
		return jctJobAttributeName;
	}
	public void setJctJobAttributeName(String jctJobAttributeName) {
		this.jctJobAttributeName = jctJobAttributeName;
	}
	public String getJctProfileDesc() {
		return jctProfileDesc;
	}
	public void setJctProfileDesc(String jctProfileDesc) {
		this.jctProfileDesc = jctProfileDesc;
	}
	public int getJctJobAttributeOrder() {
		return jctJobAttributeOrder;
	}
	public void setJctJobAttributeOrder(int jctJobAttributeOrder) {
		this.jctJobAttributeOrder = jctJobAttributeOrder;
	}
	
	

}
