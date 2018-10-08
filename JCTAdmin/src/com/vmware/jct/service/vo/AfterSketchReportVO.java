package com.vmware.jct.service.vo;

import java.util.ArrayList;
import java.util.List;

public class AfterSketchReportVO {
	// Non repeating
	private String flName;
	private String emailId;
	private String timeSpent;
	private String functionGroup;
	private String jobLevel;
	private String totalTaskCount;
	private String totalMappingCount;
	private String totalRoleCount;
	private List<AfterSketchReportListVO> repeatingList = new ArrayList<AfterSketchReportListVO>();
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
	public String getTimeSpent() {
		return timeSpent;
	}
	public void setTimeSpent(String timeSpent) {
		this.timeSpent = timeSpent;
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
	public String getTotalTaskCount() {
		return totalTaskCount;
	}
	public void setTotalTaskCount(String totalTaskCount) {
		this.totalTaskCount = totalTaskCount;
	}
	public String getTotalMappingCount() {
		return totalMappingCount;
	}
	public void setTotalMappingCount(String totalMappingCount) {
		this.totalMappingCount = totalMappingCount;
	}
	public String getTotalRoleCount() {
		return totalRoleCount;
	}
	public void setTotalRoleCount(String totalRoleCount) {
		this.totalRoleCount = totalRoleCount;
	}
	public List<AfterSketchReportListVO> getRepeatingList() {
		return repeatingList;
	}
	public void setRepeatingList(List<AfterSketchReportListVO> repeatingList) {
		this.repeatingList = repeatingList;
	}
}
