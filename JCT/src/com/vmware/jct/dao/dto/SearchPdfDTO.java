package com.vmware.jct.dao.dto;

import java.io.Serializable;
import java.util.Date;

/**
 * 
 * <p><b>Class name:</b> SearchPdfDTO.java</p>
 * <p><b>Author:</b> InterraIT</p>
 * <p><b>Purpose:</b> This class acts as a data transfer object.
 * <p><b>Description:</b> This class acts as a data transfer object.. </p>
 * <p><b>Copyrights:</b> 	All rights reserved by Interra IT and should only be used for its internal application development.</p>
 * <p><b>Created Date:</b> 01/Oct/2015</p>
 * <p><b>Revision History:</b>
 * 	<li>01/Oct/2015: Changed the saving of action plan to dynamic.</li>
 * </p>
 */
public class SearchPdfDTO implements Serializable{
	
	private static final long serialVersionUID = -2595979524767735082L;	

	String jctFileLocation;
	String jctFileName;
	Date jctCreatedTimestamp;	
	
	public SearchPdfDTO(String jctFileLocation, String jctFileName, Date jctCreatedTimestamp) {
		this.jctFileLocation = jctFileLocation;
		this.jctFileName = jctFileName;
		this.jctCreatedTimestamp = jctCreatedTimestamp;
	}
	
	public String getJctFileLocation() {
		return jctFileLocation;
	}
	public void setJctFileLocation(String jctFileLocation) {
		this.jctFileLocation = jctFileLocation;
	}
	public String getJctFileName() {
		return jctFileName;
	}
	public void setJctFileName(String jctFileName) {
		this.jctFileName = jctFileName;
	}
	public Date getJctCreatedTimestamp() {
		return jctCreatedTimestamp;
	}
	public void setJctCreatedTimestamp(Date jctCreatedTimestamp) {
		this.jctCreatedTimestamp = jctCreatedTimestamp;
	}
}