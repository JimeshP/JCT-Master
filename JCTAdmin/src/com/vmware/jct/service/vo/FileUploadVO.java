package com.vmware.jct.service.vo;

/**
 * 
 * <p><b>Class name:</b> FileUploadVO.java</p>
 * <p><b>Author:</b> InterraIT</p>
 * <p><b>Purpose:</b> This class acts as a value object. Transfer data from presentation layer to controller and vice versa
 * <p><b>Description:</b>This class acts as a value object. Transfer data from presentation layer to controller and vice versa </p>
 * <p><b>Copyrights:</b> 	All rights reserved by Interra IT and should only be used for its internal application development.</p>
 * <p><b>Created Date:</b> 10/Oct/2014</p>
 * <p><b>Revision History:</b>
 * </p>
 */

public class FileUploadVO {
	
	private int statusCode;
	private String statusDesc;
	
	
	public int getStatusCode() {
		return statusCode;
	}
	public void setStatusCode(int statusCode) {
		this.statusCode = statusCode;
	}
	public String getStatusDesc() {
		return statusDesc;
	}
	public void setStatusDesc(String statusDesc) {
		this.statusDesc = statusDesc;
	}
	
	
}
