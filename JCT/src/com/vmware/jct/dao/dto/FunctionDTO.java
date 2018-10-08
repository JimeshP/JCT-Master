/**
 * 
 */
package com.vmware.jct.dao.dto;

import java.io.Serializable;
/**
 * 
 * <p><b>Class name:</b> FunctionDTO.java</p>
 * <p><b>Author:</b> InterraIT</p>
 * <p><b>Purpose:</b> This class acts as a data transfer object.
 * <p><b>Description:</b> This class acts as a data transfer object.. </p>
 * <p><b>Copyrights:</b> 	All rights reserved by Interra IT and should only be used for its internal application development.</p>
 * <p><b>Created Date:</b> 20/Feb/2014</p>
 * <p><b>Revision History:</b>
 * 	<li>6/May/2014: Changed the saving of action plan to dynamic.</li>
 * </p>
 */
public class FunctionDTO implements Serializable{
	
	private static final long serialVersionUID = 1L;

	public int jctFunctionId;

	public String jctFunctionName;
	
	
	public FunctionDTO(){
		
	}
	public FunctionDTO(int jctFunctionId,String jctFunctionName){
		this.jctFunctionId=jctFunctionId;
		this.jctFunctionName=jctFunctionName;
		
	}
	/**
	 * @return the jctFunctionId
	 */
	public int getJctFunctionId() {
		return jctFunctionId;
	}
	/**
	 * @param jctFunctionId the jctFunctionId to set
	 */
	public void setJctFunctionId(int jctFunctionId) {
		this.jctFunctionId = jctFunctionId;
	}
	/**
	 * @return the jctFunctionName
	 */
	public String getJctFunctionName() {
		return jctFunctionName;
	}
	/**
	 * @param jctFunctionName the jctFunctionName to set
	 */
	public void setJctFunctionName(String jctFunctionName) {
		this.jctFunctionName = jctFunctionName;
	}
	
	

}
