package com.vmware.jct.dao.dto;
/**
 * 
 * <p><b>Class name:</b>OnetOccupationDTO.java</p>
 * <p><b>Author:</b>  InterrIT</p>
 * <p><b>Purpose:</b>OnetOccupationDTO provides access to onet data which are fetched into them.</p>
 * <p><b>Description:</b></p>
 * <pre>
 * 		<b>Revision History:</b>
 * 		
 * </pre>
 * <p><b></b></p>
 */
public class OnetOccupationDTO {

	private int jctOnetOccupationId;	
	private String jctOnetOccupationCode;
	private String jctOnetOccupationTitle;
	private String jctOnetOccupationDesc;
	
	public OnetOccupationDTO(){
		
	}
	public OnetOccupationDTO(int jctOnetOccupationId, String jctOnetOccupationCode, String jctOnetOccupationTitle, String jctOnetOccupationDesc){
		this.jctOnetOccupationId = jctOnetOccupationId;
		this.jctOnetOccupationCode = jctOnetOccupationCode;
		this.jctOnetOccupationTitle = jctOnetOccupationTitle;
		this.jctOnetOccupationDesc = jctOnetOccupationDesc;
	}
		
	
	public int getJctOnetOccupationId() {
		return jctOnetOccupationId;
	}
	public void setJctOnetOccupationId(int jctOnetOccupationId) {
		this.jctOnetOccupationId = jctOnetOccupationId;
	}
	public String getJctOnetOccupationCode() {
		return jctOnetOccupationCode;
	}
	public void setJctOnetOccupationCode(String jctOnetOccupationCode) {
		this.jctOnetOccupationCode = jctOnetOccupationCode;
	}
	public String getJctOnetOccupationTitle() {
		return jctOnetOccupationTitle;
	}
	public void setJctOnetOccupationTitle(String jctOnetOccupationTitle) {
		this.jctOnetOccupationTitle = jctOnetOccupationTitle;
	}
	public String getJctOnetOccupationDesc() {
		return jctOnetOccupationDesc;
	}
	public void setJctOnetOccupationDesc(String jctOnetOccupationDesc) {
		this.jctOnetOccupationDesc = jctOnetOccupationDesc;
	}
}
