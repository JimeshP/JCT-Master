package com.vmware.jct.service.vo;

public class SurveyQuestionsVO {
	// General
	private String createdBy;
	private int userType;
	private int statusCode;
	private String message;
	private String isMandatory;
	private String userProfile;
	
	public String getIsMandatory() {
		return isMandatory;
	}

	public void setIsMandatory(String isMandatory) {
		this.isMandatory = isMandatory;
	}

	// For Free Text
	private String surveyTextQuestion;
	
	// For others
	private String mainQuestion;
	private String subQuestionsString;
	
	// For result list
	private String resultList;
	

	public String getCreatedBy() {
		return createdBy;
	}

	public void setCreatedBy(String createdBy) {
		this.createdBy = createdBy;
	}

	public int getUserType() {
		return userType;
	}

	public void setUserType(int userType) {
		this.userType = userType;
	}

	public String getSurveyTextQuestion() {
		return surveyTextQuestion;
	}

	public void setSurveyTextQuestion(String surveyTextQuestion) {
		this.surveyTextQuestion = surveyTextQuestion;
	}

	public int getStatusCode() {
		return statusCode;
	}

	public void setStatusCode(int statusCode) {
		this.statusCode = statusCode;
	}

	public String getMessage() {
		return message;
	}

	public void setMessage(String message) {
		this.message = message;
	}

	public String getMainQuestion() {
		return mainQuestion;
	}

	public void setMainQuestion(String mainQuestion) {
		this.mainQuestion = mainQuestion;
	}

	public String getSubQuestionsString() {
		return subQuestionsString;
	}

	public void setSubQuestionsString(String subQuestionsString) {
		this.subQuestionsString = subQuestionsString;
	}

	public String getResultList() {
		return resultList;
	}

	public void setResultList(String resultList) {
		this.resultList = resultList;
	}

	public String getUserProfile() {
		return userProfile;
	}

	public void setUserProfile(String userProfile) {
		this.userProfile = userProfile;
	}
}
