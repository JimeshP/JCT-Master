package com.vmware.jct.dao.dto;

import java.io.Serializable;
/**
 * 
 * <p><b>Class name:</b> StatusSearchDTO.java</p>
 * <p><b>Author:</b> InterraIT</p>
 * <p><b>Purpose:</b> This class acts as a data transfer object.
 * <p><b>Description:</b> This class acts as a data transfer object.. </p>
 * <p><b>Copyrights:</b> 	All rights reserved by Interra IT and should only be used for its internal application development.</p>
 * <p><b>Created Date:</b> 25/Jun/2014</p>
 * <p><b>Revision History:</b>
 * <li></li>
 * </p>
 */
public class StatusSearchDTO implements Serializable {
	
	private static final long serialVersionUID = 1L;
	
	public String jobReferenceNos;
	public String imageString;
	public String occupation;

	public StatusSearchDTO(String jobReferenceNos,String imageString){
		this.jobReferenceNos=jobReferenceNos;
		this.imageString=imageString;
	}
	
	public StatusSearchDTO(String jobReferenceNos,String imageString,String occupation){
		this.jobReferenceNos=jobReferenceNos;
		this.imageString=imageString;
		this.occupation=occupation;
	}

	public String getJobReferenceNos() {
		return jobReferenceNos;
	}

	public void setJobReferenceNos(String jobReferenceNos) {
		this.jobReferenceNos = jobReferenceNos;
	}

	public String getImageString() {
		return imageString;
	}

	public void setImageString(String imageString) {
		this.imageString = imageString;
	}

	public String getOccupation() {
		return occupation;
	}

	public void setOccupation(String occupation) {
		this.occupation = occupation;
	}
	
}
