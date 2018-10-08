package com.vmware.jct.service.vo;

public class BeforeSketchToAfterSketchReportListVO {
    
	private String taskStatus;
	private String beforeSketchTaskDesc;
	private String afterSketchTaskDesc;
	private String taskDiff;
	private String bsTime;
	private String asTime;
	private String timeDiff;
	
	public String getTaskStatus() {
		return taskStatus;
	}
	public void setTaskStatus(String taskStatus) {
		this.taskStatus = taskStatus;
	}
	public String getBeforeSketchTaskDesc() {
		return beforeSketchTaskDesc;
	}
	public void setBeforeSketchTaskDesc(String beforeSketchTaskDesc) {
		this.beforeSketchTaskDesc = beforeSketchTaskDesc;
	}
	public String getAfterSketchTaskDesc() {
		return afterSketchTaskDesc;
	}
	public void setAfterSketchTaskDesc(String afterSketchTaskDesc) {
		this.afterSketchTaskDesc = afterSketchTaskDesc;
	}
	public String getTaskDiff() {
		return taskDiff;
	}
	public void setTaskDiff(String taskDiff) {
		this.taskDiff = taskDiff;
	}
	public String getBsTime() {
		return bsTime;
	}
	public void setBsTime(String bsTime) {
		this.bsTime = bsTime;
	}
	public String getAsTime() {
		return asTime;
	}
	public void setAsTime(String asTime) {
		this.asTime = asTime;
	}
	public String getTimeDiff() {
		return timeDiff;
	}
	public void setTimeDiff(String timeDiff) {
		this.timeDiff = timeDiff;
	}	
}
