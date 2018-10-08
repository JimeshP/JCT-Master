package com.vmware.jct.model;

import java.io.Serializable;
import java.util.Date;

import javax.persistence.Column;
import javax.persistence.Entity;
import javax.persistence.GeneratedValue;
import javax.persistence.GenerationType;
import javax.persistence.Id;
import javax.persistence.NamedQueries;
import javax.persistence.NamedQuery;
import javax.persistence.Table;
import javax.persistence.Temporal;
import javax.persistence.TemporalType;
/**
 * 
 * <p><b>Class name:</b> JctSurveyQuestions.java</p>
 * <p><b>Author:</b> InterraIT</p>
 * <p><b>Description:</b>  This class basically object reference of jct_survey_questions table </p>
 * <p><b>Copyrights:</b> 	All rights reserved by InterraIT and should only be used for its internal application development.</p>
 * <p><b>Revision History:</b>
 * 	<li>InterraIT - 14/Nov/2014 - Implemented the class </li>
 * </p>
 */
@NamedQueries({
	@NamedQuery(name = "answerdQtns",
				query = "select obj from JctSurveyQuestions obj where obj.jctSurveyTakenBy = :email and jctSurveyUserType = 1")
	
})
@Entity
@Table(name="jct_survey_questions")
public class JctSurveyQuestions implements Serializable {
	private static final long serialVersionUID = 1L;
		
	@Id
	@GeneratedValue(strategy=GenerationType.IDENTITY)
	@Column(name="jct_survey_question_id")
	private int jctSurveyQuestionId;
	
	@Column(name="jct_survey_user_type")
	private int jctSurveyUserType;
	
	@Column(name="jct_survey_taken_by")
	private String jctSurveyTakenBy;
	
	@Column(name="jct_survey_taken_by_user_id")
	private int jctSurveyTakenByUserId;
	
	@Column(name="jct_survey_answer_type")
	private int jctSurveyAnswerType;
	
	@Column(name="jct_survey_main_qtn")
	private String jctSurveyMainQtn;
	
	@Column(name="jct_survey_sub_qtn")
	private String jctSurveySubQtn;
	
	@Column(name="jct_survey_answer")
	private String jctSurveyAnswer;
	
	@Temporal(TemporalType.TIMESTAMP)
	@Column(name="jct_survey_created_ts")
	private Date jctSurveyCreatedTs;
	
	@Column(name="jct_survey_soft_delete")
	private int jctSurveySoftDelete;

	public int getJctSurveyQuestionId() {
		return jctSurveyQuestionId;
	}

	public void setJctSurveyQuestionId(int jctSurveyQuestionId) {
		this.jctSurveyQuestionId = jctSurveyQuestionId;
	}

	public int getJctSurveyUserType() {
		return jctSurveyUserType;
	}

	public void setJctSurveyUserType(int jctSurveyUserType) {
		this.jctSurveyUserType = jctSurveyUserType;
	}

	public String getJctSurveyTakenBy() {
		return jctSurveyTakenBy;
	}

	public void setJctSurveyTakenBy(String jctSurveyTakenBy) {
		this.jctSurveyTakenBy = jctSurveyTakenBy;
	}

	public int getJctSurveyTakenByUserId() {
		return jctSurveyTakenByUserId;
	}

	public void setJctSurveyTakenByUserId(int jctSurveyTakenByUserId) {
		this.jctSurveyTakenByUserId = jctSurveyTakenByUserId;
	}

	public int getJctSurveyAnswerType() {
		return jctSurveyAnswerType;
	}

	public void setJctSurveyAnswerType(int jctSurveyAnswerType) {
		this.jctSurveyAnswerType = jctSurveyAnswerType;
	}

	public String getJctSurveyMainQtn() {
		return jctSurveyMainQtn;
	}

	public void setJctSurveyMainQtn(String jctSurveyMainQtn) {
		this.jctSurveyMainQtn = jctSurveyMainQtn;
	}

	public String getJctSurveySubQtn() {
		return jctSurveySubQtn;
	}

	public void setJctSurveySubQtn(String jctSurveySubQtn) {
		this.jctSurveySubQtn = jctSurveySubQtn;
	}

	public String getJctSurveyAnswer() {
		return jctSurveyAnswer;
	}

	public void setJctSurveyAnswer(String jctSurveyAnswer) {
		this.jctSurveyAnswer = jctSurveyAnswer;
	}

	public Date getJctSurveyCreatedTs() {
		return jctSurveyCreatedTs;
	}

	public void setJctSurveyCreatedTs(Date jctSurveyCreatedTs) {
		this.jctSurveyCreatedTs = jctSurveyCreatedTs;
	}

	public int getJctSurveySoftDelete() {
		return jctSurveySoftDelete;
	}

	public void setJctSurveySoftDelete(int jctSurveySoftDelete) {
		this.jctSurveySoftDelete = jctSurveySoftDelete;
	}
}
