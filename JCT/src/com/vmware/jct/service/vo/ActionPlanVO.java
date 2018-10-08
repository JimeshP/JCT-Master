package com.vmware.jct.service.vo;
/**
 * 
 * <p><b>Class name:</b> ActionPlanVO.java</p>
 * <p><b>Author:</b> InterraIT</p>
 * <p><b>Purpose:</b> This class acts as a value object.
 * <p><b>Description:</b> This class acts as a value object .. </p>
 * <p><b>Copyrights:</b> 	All rights reserved by Interra IT and should only be used for its internal application development.</p>
 * <p><b>Created Date:</b> 20/Feb/2014</p>
 * <p><b>Revision History:</b>
 * 	<li>6/May/2014: Changed the saving of action plan to dynamic.</li>
 * </p>
 */
public class ActionPlanVO {
	private String jobRefNo;
	private String questionDesc;
	private String questionSubDesc;
	private String answerDesc;
	private String userModified;
	private int jctUserId;
	//New way
	private String myPlan;
	public String getJobRefNo() {
		return jobRefNo;
	}
	public void setJobRefNo(String jobRefNo) {
		this.jobRefNo = jobRefNo;
	}
	public String getQuestionDesc() {
		return questionDesc;
	}
	public void setQuestionDesc(String questionDesc) {
		this.questionDesc = questionDesc;
	}
	public String getQuestionSubDesc() {
		return questionSubDesc;
	}
	public void setQuestionSubDesc(String questionSubDesc) {
		this.questionSubDesc = questionSubDesc;
	}
	public String getAnswerDesc() {
		return answerDesc;
	}
	public void setAnswerDesc(String answerDesc) {
		this.answerDesc = answerDesc;
	}
	public String getUserModified() {
		return userModified;
	}
	public void setUserModified(String userModified) {
		this.userModified = userModified;
	}
	public String getMyPlan() {
		return myPlan;
	}
	public void setMyPlan(String myPlan) {
		this.myPlan = myPlan;
	}
	public int getJctUserId() {
		return jctUserId;
	}
	public void setJctUserId(int jctUserId) {
		this.jctUserId = jctUserId;
	}
}
