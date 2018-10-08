package com.vmware.jct.service.vo;
/**
 * 
 * <p><b>Class name:</b> QuestionnaireVO.java</p>
 * <p><b>Author:</b> InterraIT</p>
 * <p><b>Purpose:</b> This class acts as a data transfer object.
 * <p><b>Description:</b> This class acts as a data transfer object.. </p>
 * <p><b>Copyrights:</b> 	All rights reserved by Interra IT and should only be used for its internal application development.</p>
 * <p><b>Created Date:</b> 20/Feb/2014</p>
 * <p><b>Revision History:</b>
 * 	<li>6/May/2014: Changed the saving of action plan to dynamic.</li>
 * </p>
 */
public class QuestionnaireVO {
	private String jobReferenceNumber;
	private String answerDescription;
	private String questionDescription;
	private String createdBy;
	private String qtnSubDesc;
	/** THIS IS FOR PUBLIC VERSION **/
	private int jctUserId;
	/********************************/
	
	public String getQtnSubDesc() {
		return qtnSubDesc;
	}
	public void setQtnSubDesc(String qtnSubDesc) {
		this.qtnSubDesc = qtnSubDesc;
	}
	public String getJobReferenceNumber() {
		return jobReferenceNumber;
	}
	public void setJobReferenceNumber(String jobReferenceNumber) {
		this.jobReferenceNumber = jobReferenceNumber;
	}
	public String getAnswerDescription() {
		return answerDescription;
	}
	public void setAnswerDescription(String answerDescription) {
		this.answerDescription = answerDescription;
	}
	public String getQuestionDescription() {
		return questionDescription;
	}
	public void setQuestionDescription(String questionDescription) {
		this.questionDescription = questionDescription;
	}
	public String getCreatedBy() {
		return createdBy;
	}
	public void setCreatedBy(String createdBy) {
		this.createdBy = createdBy;
	}
	/** THIS IS FOR PUBLIC VERSION **/
	public int getJctUserId() {
		return jctUserId;
	}
	public void setJctUserId(int jctUserId) {
		this.jctUserId = jctUserId;
	}
	/********************************/
}
