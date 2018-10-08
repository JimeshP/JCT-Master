/**
 * 
 */
package com.vmware.jct.dao.dto;

import java.io.Serializable;

/**
 * 
 * <p><b>Class name:</b> RegionDTO.java</p>
 * <p><b>Author:</b> InterraIT</p>
 * <p><b>Purpose:</b> This class acts as a data transfer object.
 * <p><b>Description:</b> This class acts as a data transfer object.. </p>
 * <p><b>Copyrights:</b> 	All rights reserved by Interra IT and should only be used for its internal application development.</p>
 * <p><b>Created Date:</b> 20/Feb/2014</p>
 * <p><b>Revision History:</b>
 * 	<li>6/May/2014: Changed the saving of action plan to dynamic.</li>
 * </p>
 */
public class RegionDTO implements Serializable{
	
	/**
	 * 
	 */
	private static final long serialVersionUID = -2595979524767735082L;

	public int jctRegionId;

	public String jctRegionName;
	
	
	public RegionDTO(){
		
	}
	public RegionDTO(int jctRegionId,String jctRegionName){
		this.jctRegionId=jctRegionId;
		this.jctRegionName=jctRegionName;
		
	}
	/**
	 * @return the jctRegionId
	 */
	public int getJctRegionId() {
		return jctRegionId;
	}
	/**
	 * @param jctRegionId the jctRegionId to set
	 */
	public void setJctRegionId(int jctRegionId) {
		this.jctRegionId = jctRegionId;
	}
	/**
	 * @return the jctRegionName
	 */
	public String getJctRegionName() {
		return jctRegionName;
	}
	/**
	 * @param jctRegionName the jctRegionName to set
	 */
	public void setJctRegionName(String jctRegionName) {
		this.jctRegionName = jctRegionName;
	}

	


}
