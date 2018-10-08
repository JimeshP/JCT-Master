package com.vmware.jct.dao.dto;

import java.io.Serializable;

/**
 * 
 * <p><b>Class name:</b>LevelDTO.java</p>
 * <p><b>Author:</b>  InterrIT</p>
 * <p><b>Purpose:</b>LevelDTO provides access to user level data which are fetched into them.</p>
 * <p><b>Description:</b></p>
 * <pre>
 * 		<b>Revision History:</b>
 * 		
 * </pre>
 * <p><b></b></p>
 */
public class LevelDTO implements Serializable{
	/**
	 * 
	 */
	private static final long serialVersionUID = 1L;

	public int jctLevelId;

	public String jctLevelName;
	public int jctLevelOrder;
	
	
	public LevelDTO(){
		
	}
	public LevelDTO(int jctLevelId,String jctLevelName, int jctLevelOrder){
		this.jctLevelId=jctLevelId;
		this.jctLevelName=jctLevelName;
		this.jctLevelOrder=jctLevelOrder;
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
	/**
	 * @return the jctLevelOrder
	 */
	public int getJctLevelOrder() {
		return jctLevelOrder;
	}
	/**
	 * @param jctLevelOrder the jctLevelOrder to set
	 */
	public void setJctLevelOrder(int jctLevelOrder) {
		this.jctLevelOrder = jctLevelOrder;
	}
}
