package com.vmware.jct.service.vo;

import java.util.List;

public class RemoveDiagramVOList {
	private List<RemoveDiagramVO> diagramList;
	private int statusCode;
	private String statusMessage;
	private int rowId;
	private String emailId;
	private String rowIdSetString;

	public List<RemoveDiagramVO> getDiagramList() {
		return diagramList;
	}

	public void setDiagramList(List<RemoveDiagramVO> diagramList) {
		this.diagramList = diagramList;
	}

	public int getStatusCode() {
		return statusCode;
	}

	public void setStatusCode(int statusCode) {
		this.statusCode = statusCode;
	}

	public int getRowId() {
		return rowId;
	}

	public void setRowId(int rowId) {
		this.rowId = rowId;
	}

	public String getEmailId() {
		return emailId;
	}

	public void setEmailId(String emailId) {
		this.emailId = emailId;
	}

	public String getStatusMessage() {
		return statusMessage;
	}

	public void setStatusMessage(String statusMessage) {
		this.statusMessage = statusMessage;
	}

	public String getRowIdSetString() {
		return rowIdSetString;
	}

	public void setRowIdSetString(String rowIdSetString) {
		this.rowIdSetString = rowIdSetString;
	}

}
