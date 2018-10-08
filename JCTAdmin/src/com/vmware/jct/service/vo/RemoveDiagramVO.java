package com.vmware.jct.service.vo;

import java.util.Date;

public class RemoveDiagramVO {
	private int rowId;
	private String beforeSketchBaseString;
	private String aftereSketchBaseString;
	private int statusCode;
	private String message;
	private String userName;
	private long totalDiagramCount;
	private Date jctCreatedTs;
	
	public int getRowId() {
		return rowId;
	}
	public void setRowId(int rowId) {
		this.rowId = rowId;
	}
	public String getBeforeSketchBaseString() {
		return beforeSketchBaseString;
	}
	public void setBeforeSketchBaseString(String beforeSketchBaseString) {
		this.beforeSketchBaseString = beforeSketchBaseString;
	}
	public String getAftereSketchBaseString() {
		return aftereSketchBaseString;
	}
	public void setAftereSketchBaseString(String aftereSketchBaseString) {
		this.aftereSketchBaseString = aftereSketchBaseString;
	}
	public int getStatusCode() {
		return statusCode;
	}
	public void setStatusCode(int statusCode) {
		this.statusCode = statusCode;
	}
	public String getMessage() {
		return message;
	}
	public void setMessage(String message) {
		this.message = message;
	}
	public String getUserName() {
		return userName;
	}
	public void setUserName(String userName) {
		this.userName = userName;
	}
	public long getTotalDiagramCount() {
		return totalDiagramCount;
	}
	public void setTotalDiagramCount(long totalDiagramCount) {
		this.totalDiagramCount = totalDiagramCount;
	}
	public Date getJctCreatedTs() {
		return jctCreatedTs;
	}
	public void setJctCreatedTs(Date jctCreatedTs) {
		this.jctCreatedTs = jctCreatedTs;
	}
}
