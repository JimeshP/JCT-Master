package com.vmware.jct.dao.dto;
/**
 * 
 * <p><b>Class name:</b>SurveyQuestionDTO.java</p>
 * <p><b>Author:</b>  InterrIT</p>
 * <p><b>Purpose:</b>SurveyQuestionDTO provides access to survey question data which are fetched into them.</p>
 * <p><b>Description:</b></p>
 * <pre>
 * 		<b>Revision History:</b>
 * 		
 * </pre>
 * <p><b></b></p>
 */
public class SurveyQuestionDTO {
	private int jctUserType;
	private int answerType;
	private String mainQuestion;
	private String subQtns;
	private String isMandatory;
	private int surveyQtnUserId;
	private long surveyQtnCount;
	private String profileName;
	
	public SurveyQuestionDTO(long surveyQtnCount, int surveyQtnUserId) {
		this.surveyQtnUserId = surveyQtnUserId;
		this.surveyQtnCount = surveyQtnCount;
	}


	public SurveyQuestionDTO(int jctUserType, int answerType, String mainQuestion) {
		this.jctUserType = jctUserType;
		this.answerType = answerType;
		this.mainQuestion = mainQuestion;
	}


	public SurveyQuestionDTO(int jctUserType, int answerType, String mainQuestion, String isMandatory) {
		this.jctUserType = jctUserType;
		this.answerType = answerType;
		this.mainQuestion = mainQuestion;
		this.isMandatory = isMandatory;
	}
	
	public SurveyQuestionDTO(int jctUserType, int answerType, String mainQuestion, String isMandatory, String profileName) {
		this.jctUserType = jctUserType;
		this.answerType = answerType;
		this.mainQuestion = mainQuestion;
		this.isMandatory = isMandatory;
		this.profileName = profileName;
	}
	
	public String getIsMandatory() {
		return isMandatory;
	}

	public void setIsMandatory(String isMandatory) {
		this.isMandatory = isMandatory;
	}

	public SurveyQuestionDTO(String subQtns) {
		this.subQtns = subQtns;
	}

	public int getJctUserType() {
		return jctUserType;
	}

	public void setJctUserType(int jctUserType) {
		this.jctUserType = jctUserType;
	}

	public int getAnswerType() {
		return answerType;
	}

	public void setAnswerType(int answerType) {
		this.answerType = answerType;
	}

	public String getMainQuestion() {
		return mainQuestion;
	}

	public void setMainQuestion(String mainQuestion) {
		this.mainQuestion = mainQuestion;
	}

	public String getSubQtns() {
		return subQtns;
	}

	public void setSubQtns(String subQtns) {
		this.subQtns = subQtns;
	}
	
	public int getSurveyQtnUserId() {
		return surveyQtnUserId;
	}

	public void setSurveyQtnUserId(int surveyQtnUserId) {
		this.surveyQtnUserId = surveyQtnUserId;
	}


	public long getSurveyQtnCount() {
		return surveyQtnCount;
	}


	public void setSurveyQtnCount(long surveyQtnCount) {
		this.surveyQtnCount = surveyQtnCount;
	}


	public String getProfileName() {
		return profileName;
	}


	public void setProfileName(String profileName) {
		this.profileName = profileName;
	}

	
}
