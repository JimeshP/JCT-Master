package com.vmware.jct.service.vo;

import java.util.List;

public class BeforeSketchToAfterSketchReportVO {
	// Non Repeating
	private String flName;
	private String emailId;
	private String functionGroup;
	private String jobLevel;
	private List<BeforeSketchToAfterSketchReportListVO> list;
	
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
	public List<BeforeSketchToAfterSketchReportListVO> getList() {
		return list;
	}
	public void setList(List<BeforeSketchToAfterSketchReportListVO> list) {
		this.list = list;
	}	
}
