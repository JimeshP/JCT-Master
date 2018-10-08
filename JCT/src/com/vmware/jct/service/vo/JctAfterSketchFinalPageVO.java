package com.vmware.jct.service.vo;
/**
 * 
 * <p><b>Class name:</b> JctAfterSketchFinalPageVO.java</p>
 * <p><b>Author:</b> InterraIT</p>
 * <p><b>Purpose:</b> This class acts as a value object.
 * <p><b>Description:</b> This class acts as a value object .. </p>
 * <p><b>Copyrights:</b> 	All rights reserved by Interra IT and should only be used for its internal application development.</p>
 * <p><b>Created Date:</b> 20/Feb/2014</p>
 * <p><b>Revision History:</b>
 * 	<li>6/May/2014: Changed the saving of action plan to dynamic.</li>
 * </p>
 */
public class JctAfterSketchFinalPageVO {
	private String jobRefNo;
	private String elementCode;
	private String elementDesc;
	private String position;
	private String roleDescription;
	private int taskEnergy;
	private int taskId;
	private String taskDescription;
	private String createdBy;
	private String additionalDesc;
	
	public String getJobRefNo() {
		return jobRefNo;
	}
	public void setJobRefNo(String jobRefNo) {
		this.jobRefNo = jobRefNo;
	}
	public String getElementCode() {
		return elementCode;
	}
	public void setElementCode(String elementCode) {
		this.elementCode = elementCode;
	}
	public String getElementDesc() {
		return elementDesc;
	}
	public void setElementDesc(String elementDesc) {
		this.elementDesc = elementDesc;
	}
	public String getPosition() {
		return position;
	}
	public void setPosition(String position) {
		this.position = position;
	}
	public String getRoleDescription() {
		return roleDescription;
	}
	public void setRoleDescription(String roleDescription) {
		this.roleDescription = roleDescription;
	}
	public int getTaskEnergy() {
		return taskEnergy;
	}
	public void setTaskEnergy(int taskEnergy) {
		this.taskEnergy = taskEnergy;
	}
	public int getTaskId() {
		return taskId;
	}
	public void setTaskId(int taskId) {
		this.taskId = taskId;
	}
	public String getCreatedBy() {
		return createdBy;
	}
	public void setCreatedBy(String createdBy) {
		this.createdBy = createdBy;
	}
	public String getAdditionalDesc() {
		return additionalDesc;
	}
	public void setAdditionalDesc(String additionalDesc) {
		this.additionalDesc = additionalDesc;
	}
	public String getTaskDescription() {
		return taskDescription;
	}
	public void setTaskDescription(String taskDescription) {
		this.taskDescription = taskDescription;
	}
	
	public boolean equals(Object o) {
		JctAfterSketchFinalPageVO vo = (JctAfterSketchFinalPageVO) o;
		boolean result =  this.jobRefNo.equals(vo.getJobRefNo()) &&  
				this.elementCode.equals(vo.getElementCode()) &&
				this.elementDesc.equals(vo.getElementDesc()) &&
				/*this.position.equals(vo.getPosition()) &&*/
				this.roleDescription.equals(vo.getRoleDescription()) && 
				this.taskDescription.equals(vo.getTaskDescription()) && 
				/*this.createdBy.equals(vo.getCreatedBy()) && */
				this.additionalDesc.equals(vo.getAdditionalDesc()) && 
				this.taskEnergy == vo.getTaskEnergy();
		
		return result;
	}
}