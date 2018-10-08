package com.vmware.jct.service.vo;

public class FacilitatorIndividualReportVO {
	// General
	private String reportList;
	private int statusCode;
	private String statusMessage;
	public int getStatusCode() {
		return statusCode;
	}
	public void setStatusCode(int statusCode) {
		this.statusCode = statusCode;
	}
	public String getStatusMessage() {
		return statusMessage;
	}
	public void setStatusMessage(String statusMessage) {
		this.statusMessage = statusMessage;
	}
	public String getReportList() {
		return reportList;
	}
	public void setReportList(String reportList) {
		this.reportList = reportList;
	}
	
	
}
