package com.vmware.jct.service.vo;

public class AfterSketchReportListVO {
	//Repeating
	private String taskDesc;
	private String taskAllocation;
	private String phaka;	//	mapping
	private String role;
	public String getTaskDesc() {
		return taskDesc;
	}
	public void setTaskDesc(String taskDesc) {
		this.taskDesc = taskDesc;
	}
	public String getTaskAllocation() {
		return taskAllocation;
	}
	public void setTaskAllocation(String taskAllocation) {
		this.taskAllocation = taskAllocation;
	}
	public String getPhaka() {
		return phaka;
	}
	public void setPhaka(String phaka) {
		this.phaka = phaka;
	}
	public String getRole() {
		return role;
	}
	public void setRole(String role) {
		this.role = role;
	}
}
