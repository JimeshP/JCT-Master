package com.vmware.jct.dao.dto;

import java.io.Serializable;
/**
 * 
 * <p><b>Class name:</b>BeforeSketchDTO.java</p>
 * <p><b>Author:</b>  InterrIT</p>
 * <p><b>Purpose:</b>BeforeSketchDTO provides access to Before Sketch data which are fetched into them.</p>
 * <p><b>Description:</b></p>
 * <pre>
 * 		<b>Revision History:</b>
 * 		
 * </pre>
 * <p><b></b></p>
 */
public class BeforeSketchDTO implements Serializable{
	private static final long serialVersionUID = 1L;
	private String jobReferenceNo;
	private String functionGroup;
	private String jobLevel;
	private String taskDesc;
	private int energy;
	private int totalTime;
	
	public BeforeSketchDTO(){}
	
	public BeforeSketchDTO(String jobReferenceNo, String functionGroup, String jobLevel,
						   String taskDesc, int energy, int totalTime){
		this.jobReferenceNo = jobReferenceNo;
		this.functionGroup = functionGroup;
		this.jobLevel = jobLevel;
		this.taskDesc = taskDesc;
		this.energy = energy;
		this.totalTime = totalTime;
	}
	/**
	 * 
	 * @return
	 */
	public String getJobReferenceNo() {
		return jobReferenceNo;
	}
	/**
	 * 
	 * @param jobReferenceNo
	 */
	public void setJobReferenceNo(String jobReferenceNo) {
		this.jobReferenceNo = jobReferenceNo;
	}
	/**
	 * 
	 * @return
	 */
	public String getFunctionGroup() {
		return functionGroup;
	}
	/**
	 * 
	 * @param functionGroup
	 */
	public void setFunctionGroup(String functionGroup) {
		this.functionGroup = functionGroup;
	}
	/**
	 * 
	 * @return
	 */
	public String getJobLevel() {
		return jobLevel;
	}
	/**
	 * 
	 * @param jobLevel
	 */
	public void setJobLevel(String jobLevel) {
		this.jobLevel = jobLevel;
	}
	/**
	 * 
	 * @return
	 */
	public String getTaskDesc() {
		return taskDesc;
	}
	/**
	 * 
	 * @param taskDesc
	 */
	public void setTaskDesc(String taskDesc) {
		this.taskDesc = taskDesc;
	}
	/**
	 * 
	 * @return
	 */
	public int getEnergy() {
		return energy;
	}
	/**
	 * 
	 * @param energy
	 */
	public void setEnergy(int energy) {
		this.energy = energy;
	}
	/**
	 * 
	 * @return
	 */
	public int getTotalTime() {
		return totalTime;
	}
	/**
	 * 
	 * @param totalTime
	 */
	public void setTotalTime(int totalTime) {
		this.totalTime = totalTime;
	}
}
