package com.vmware.jct.service.vo;

/**
 * 
 * <p><b>Class name:</b> ExistingUsersVO.java</p>
 * <p><b>Author:</b> InterraIT</p>
 * <p><b>Purpose:</b> This class acts as a value object. Transfer data from presentation layer to controller and vice versa
 * <p><b>Description:</b>This class acts as a value object. Transfer data from presentation layer to controller and vice versa </p>
 * <p><b>Copyrights:</b> 	All rights reserved by Interra IT and should only be used for its internal application development.</p>
 * <p><b>Created Date:</b> 20/Feb/2014</p>
 * <p><b>Revision History:</b>
 * 	<li>6/May/2014: Changed the saving of action plan to dynamic.</li>
 * </p>
 */
public class ExistingFacilitatorVO {
	private String faciUserName;
	private int statusCode;
	private String statusDesc;
	private long jctFacTotalLimit;
	private long jctFacSubscribeLimit;
	
	public String getFaciUserName() {
		return faciUserName;
	}
	public void setFaciUserName(String faciUserName) {
		this.faciUserName = faciUserName;
	}
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
	public long getJctFacTotalLimit() {
		return jctFacTotalLimit;
	}
	public void setJctFacTotalLimit(long jctFacTotalLimit) {
		this.jctFacTotalLimit = jctFacTotalLimit;
	}
	public long getJctFacSubscribeLimit() {
		return jctFacSubscribeLimit;
	}
	public void setJctFacSubscribeLimit(long jctFacSubscribeLimit) {
		this.jctFacSubscribeLimit = jctFacSubscribeLimit;
	}
	
}
