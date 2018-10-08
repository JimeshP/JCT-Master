package com.vmware.jct.service.vo;

import java.util.List;

public class BeforeSketchReportVO {
	// Non Repeating
	private String flName;
	private String emailId;
	private String functionGroup;
	private String jobLevel;
	private String totalTime;
	private List<BeforeSketchReportListVO> list;
	
	public String getName() {
		return flName;
	}
	public void setName(String flName) {
		this.flName = flName;
	}
	public String getEmailId() {
		return emailId;
	}
	public void setEmailId(String emailId) {
		this.emailId = emailId;
	}
	public String getFunctionGroup() {
		return functionGroup;
	}
	public void setFunctionGroup(String functionGroup) {
		this.functionGroup = functionGroup;
	}
	public String getJobLevel() {
		return jobLevel;
	}
	public void setJobLevel(String jobLevel) {
		this.jobLevel = jobLevel;
	}
	public String getTotalTime() {
		return totalTime;
	}
	public void setTotalTime(String totalTime) {
		this.totalTime = totalTime;
	}
	public List<BeforeSketchReportListVO> getList() {
		return list;
	}
	public void setList(List<BeforeSketchReportListVO> list) {
		this.list = list;
	}
}
