/**
 * 
 */
package com.vmware.jct.dao.dto;

import java.io.Serializable;

/**
 * 
 * <p><b>Class name:</b>FunctionDTO.java</p>
 * <p><b>Author:</b>  InterrIT</p>
 * <p><b>Purpose:</b>FunctionDTO provides access to function data which are fetched into them.</p>
 * <p><b>Description:</b></p>
 * <pre>
 * 		<b>Revision History:</b>
 * 		
 * </pre>
 * <p><b></b></p>
 */
public class FunctionDTO implements Serializable{
	
	/**
	 * 
	 */
	private static final long serialVersionUID = 1L;

	public int jctFunctionId;

	public String jctFunctionName;
	public int jctFunctionsOrder;
	
	
	public FunctionDTO(){
		
	}
	public FunctionDTO(int jctFunctionId,String jctFunctionName, int jctFunctionsOrder){
		this.jctFunctionId=jctFunctionId;
		this.jctFunctionName=jctFunctionName;
		this.jctFunctionsOrder=jctFunctionsOrder;
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
	/**
	 * @return the jctFunctionsOrder
	 */
	public int getJctFunctionsOrder() {
		return jctFunctionsOrder;
	}
	/**
	 * @param jctFunctionsOrder the jctFunctionsOrder to set
	 */
	public void setJctFunctionsOrder(int jctFunctionsOrder) {
		this.jctFunctionsOrder = jctFunctionsOrder;
	}
}
