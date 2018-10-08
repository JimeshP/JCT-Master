package com.vmware.jct.dao.dto;

import java.io.Serializable;

/**
 * 
 * <p><b>Class name:</b>RegionDTO.java</p>
 * <p><b>Author:</b>  InterrIT</p>
 * <p><b>Purpose:</b>RegionDTO provides access to region data which are fetched into them.</p>
 * <p><b>Description:</b></p>
 * <pre>
 * 		<b>Revision History:</b>
 * 		
 * </pre>
 * <p><b></b></p>
 */
public class RegionDTO implements Serializable{
	
	private static final long serialVersionUID = -2595979524767735082L;

	public int jctRegionId;

	public String jctRegionName;
	
	
	public RegionDTO(){
		
	}
	public RegionDTO(int jctRegionId,String jctRegionName){
		this.jctRegionId=jctRegionId;
		this.jctRegionName=jctRegionName;
		
	}
	
	public RegionDTO(int jctRegionId){
		this.jctRegionId=jctRegionId;		
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
