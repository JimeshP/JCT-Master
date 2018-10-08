/**
 * 
 */
package com.vmware.jct.dao.dto;

import java.io.Serializable;

/**
 * 
 * <p><b>Class name:</b> LevelDTO.java</p>
 * <p><b>Author:</b> InterraIT</p>
 * <p><b>Purpose:</b> This class acts as a data transfer object.
 * <p><b>Description:</b> This class acts as a data transfer object.. </p>
 * <p><b>Copyrights:</b> 	All rights reserved by Interra IT and should only be used for its internal application development.</p>
 * <p><b>Created Date:</b> 20/Feb/2014</p>
 * <p><b>Revision History:</b>
 * 	<li>6/May/2014: Changed the saving of action plan to dynamic.</li>
 * </p>
 */
public class LevelDTO implements Serializable{
	
	private static final long serialVersionUID = 1L;

	public int jctLevelId;

	public String jctLevelName;
	
	
	public LevelDTO(){
		
	}
	public LevelDTO(int jctLevelId,String jctLevelName){
		this.jctLevelId=jctLevelId;
		this.jctLevelName=jctLevelName;
		
	}
	/**
	 * @return the jctLevelId
	 */
	public int getJctLevelId() {
		return jctLevelId;
	}
	/**
	 * @param jctLevelId the jctLevelId to set
	 */
	public void setJctLevelId(int jctLevelId) {
		this.jctLevelId = jctLevelId;
	}
	/**
	 * @return the jctLevelName
	 */
	public String getJctLevelName() {
		return jctLevelName;
	}
	/**
	 * @param jctLevelName the jctLevelName to set
	 */
	public void setJctLevelName(String jctLevelName) {
		this.jctLevelName = jctLevelName;
	}
	
	

}
