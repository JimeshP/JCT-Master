package com.vmware.jct.service.vo;
/**
 * 
 * <p><b>Class name:</b> JctAfterSketchPageOneVO.java</p>
 * <p><b>Author:</b> InterraIT</p>
 * <p><b>Purpose:</b> This class acts as a data transfer object.
 * <p><b>Description:</b> This class acts as a data transfer object.. </p>
 * <p><b>Copyrights:</b> 	All rights reserved by Interra IT and should only be used for its internal application development.</p>
 * <p><b>Created Date:</b> 20/Feb/2014</p>
 * <p><b>Revision History:</b>
 * 	<li>6/May/2014: Changed the saving of action plan to dynamic.</li>
 * </p>
 */
public class JctAfterSketchPageOneVO {
	private String jobReferenceNo;
	private String elementCode;
	private int elementId;
	private String position;
	private int softDelete;
	private int passionCount;
	private int valueCount;
	private int strengthCount;
	public int getPassionCount() {
		return passionCount;
	}
	public void setPassionCount(int passionCount) {
		this.passionCount = passionCount;
	}
	public int getValueCount() {
		return valueCount;
	}
	public void setValueCount(int valueCount) {
		this.valueCount = valueCount;
	}
	public int getStrengthCount() {
		return strengthCount;
	}
	public void setStrengthCount(int strengthCount) {
		this.strengthCount = strengthCount;
	}
	public String getJobReferenceNo() {
		return jobReferenceNo;
	}
	public void setJobReferenceNo(String jobReferenceNo) {
		this.jobReferenceNo = jobReferenceNo;
	}
	public String getElementCode() {
		return elementCode;
	}
	public void setElementCode(String elementCode) {
		this.elementCode = elementCode;
	}
	public int getElementId() {
		return elementId;
	}
	public void setElementId(int elementId) {
		this.elementId = elementId;
	}
	public String getPosition() {
		return position;
	}
	public void setPosition(String position) {
		this.position = position;
	}
	public int getSoftDelete() {
		return softDelete;
	}
	public void setSoftDelete(int softDelete) {
		this.softDelete = softDelete;
	}
}
