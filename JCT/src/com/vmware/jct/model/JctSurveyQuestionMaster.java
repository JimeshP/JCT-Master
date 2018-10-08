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
 * <p><b>Class name:</b> JctSurveyQuestionMaster.java</p>
 * <p><b>Author:</b> InterraIT</p>
 * <p><b>Description:</b>  This class basically object reference of jct_survey_question_master table </p>
 * <p><b>Copyrights:</b> 	All rights reserved by InterraIT and should only be used for its internal application development.</p>
 * <p><b>Revision History:</b>
 * 	<li>InterraIT - 21/Oct/2014 - Implemented the class </li>
 * </p>
 */
@NamedQueries({
	@NamedQuery(name = "getTextSurveyQuestion",
				query = "select DISTINCT(survey.jctSurveyQtnMasterMainQtn), survey.jctSurveyQtnMasterMandatory from JctSurveyQuestionMaster survey WHERE "
						+ "survey.jctSurveyQtnMasterSoftDel = 0 AND survey.jctSurveyQtnMasterAnsType = 1 " +
						"AND survey.jctSurveyQtnMasterUserType = 1 ORDER BY survey.jctSurveyQtnMasterOrder"),
	@NamedQuery(name = "getMainQuestionForMultiple",
				query = "select DISTINCT(survey.jctSurveyQtnMasterMainQtn), survey.jctSurveyQtnMasterMandatory from JctSurveyQuestionMaster survey WHERE "
						+ "survey.jctSurveyQtnMasterSoftDel = 0 AND survey.jctSurveyQtnMasterAnsType = :ansType " +
						"AND survey.jctSurveyQtnMasterUserType = 1 ORDER BY survey.jctSurveyQtnMasterOrder"),
	@NamedQuery(name = "getSubQuestions",
				query = "select survey.jctSurveyQtnMasterSubQtn from JctSurveyQuestionMaster survey WHERE "
						+ "survey.jctSurveyQtnMasterSoftDel = 0 AND jctSurveyQtnMasterAnsType = :ansType AND jctSurveyQtnMasterMainQtn = :mainQtn AND jctSurveyQtnMasterUserType = 1"),
	
	@NamedQuery(name = "getAllSurveyQuestions",
				query = "select DISTINCT(survey.jctSurveyQtnMasterMainQtn), survey.jctSurveyQtnMasterMandatory, survey.jctSurveyQtnMasterAnsType from JctSurveyQuestionMaster survey WHERE "
						+ "survey.jctSurveyQtnMasterSoftDel = 0 " +
						"AND survey.jctSurveyQtnMasterUserType = 1 ORDER BY survey.jctSurveyQtnMasterOrder")
	
})
@Entity
@Table(name="jct_survey_question_master")
public class JctSurveyQuestionMaster implements Serializable {
	private static final long serialVersionUID = 1L;
	
	@Id
	@GeneratedValue(strategy=GenerationType.IDENTITY)
	@Column(name="jct_survey_qtn_master_id")
	private int jctSurveyQtnMasterId;
	
	@Column(name="jct_survey_qtn_master_user_type")
	private int jctSurveyQtnMasterUserType;
	
	@Column(name="jct_survey_qtn_master_ans_type")
	private int jctSurveyQtnMasterAnsType;
	
	@Column(name="jct_survey_qtn_master_main_qtn")
	private String jctSurveyQtnMasterMainQtn;
	
	@Column(name="jct_survey_qtn_master_sub_qtn")
	private String jctSurveyQtnMasterSubQtn;
	
	@Column(name="jct_survey_qtn_master_soft_del")
	private int jctSurveyQtnMasterSoftDel;
	
	@Column(name="jct_survey_qtn_master_created_by")
	private String jctSurveyQtnMasterCreatedBy;
	
	
	@Column(name="jct_survey_qtn_master_mandatory")
	private String jctSurveyQtnMasterMandatory;
	
	
	public String getJctSurveyQtnMasterMandatory() {
		return jctSurveyQtnMasterMandatory;
	}

	public void setJctSurveyQtnMasterMandatory(String jctSurveyQtnMasterMandatory) {
		this.jctSurveyQtnMasterMandatory = jctSurveyQtnMasterMandatory;
	}

	@Temporal(TemporalType.TIMESTAMP)
	@Column(name="jct_survey_qtn_master_created_ts")
	private Date jctSurveyQtnMasterCreatedTs;
	
	@Column(name="jct_survey_qtn_master_modified_by")
	private String jctSurveyQtnMasterModifiedBy;
	
	@Temporal(TemporalType.TIMESTAMP)
	@Column(name="jct_survey_qtn_master_modified_ts")
	private Date jctSurveyQtnMasterModifiedTs;
	
	@Column(name="jct_survey_qtn_master_profile_id")
	private int jctSurveyQtnMasterProfileId;
	
	@Column(name="jct_survey_qtn_master_profile_name")
	private String jctSurveyQtnMasterProfileName;
	
	@Column(name="jct_survey_qtn_master_order")
	private int jctSurveyQtnMasterOrder;

	public int getJctSurveyQtnMasterId() {
		return jctSurveyQtnMasterId;
	}

	public void setJctSurveyQtnMasterId(int jctSurveyQtnMasterId) {
		this.jctSurveyQtnMasterId = jctSurveyQtnMasterId;
	}

	public int getJctSurveyQtnMasterUserType() {
		return jctSurveyQtnMasterUserType;
	}

	public void setJctSurveyQtnMasterUserType(int jctSurveyQtnMasterUserType) {
		this.jctSurveyQtnMasterUserType = jctSurveyQtnMasterUserType;
	}

	public int getJctSurveyQtnMasterAnsType() {
		return jctSurveyQtnMasterAnsType;
	}

	public void setJctSurveyQtnMasterAnsType(int jctSurveyQtnMasterAnsType) {
		this.jctSurveyQtnMasterAnsType = jctSurveyQtnMasterAnsType;
	}

	public String getJctSurveyQtnMasterMainQtn() {
		return jctSurveyQtnMasterMainQtn;
	}

	public void setJctSurveyQtnMasterMainQtn(String jctSurveyQtnMasterMainQtn) {
		this.jctSurveyQtnMasterMainQtn = jctSurveyQtnMasterMainQtn;
	}

	public String getJctSurveyQtnMasterSubQtn() {
		return jctSurveyQtnMasterSubQtn;
	}

	public void setJctSurveyQtnMasterSubQtn(String jctSurveyQtnMasterSubQtn) {
		this.jctSurveyQtnMasterSubQtn = jctSurveyQtnMasterSubQtn;
	}

	public int getJctSurveyQtnMasterSoftDel() {
		return jctSurveyQtnMasterSoftDel;
	}

	public void setJctSurveyQtnMasterSoftDel(int jctSurveyQtnMasterSoftDel) {
		this.jctSurveyQtnMasterSoftDel = jctSurveyQtnMasterSoftDel;
	}

	public String getJctSurveyQtnMasterCreatedBy() {
		return jctSurveyQtnMasterCreatedBy;
	}

	public void setJctSurveyQtnMasterCreatedBy(String jctSurveyQtnMasterCreatedBy) {
		this.jctSurveyQtnMasterCreatedBy = jctSurveyQtnMasterCreatedBy;
	}

	public Date getJctSurveyQtnMasterCreatedTs() {
		return jctSurveyQtnMasterCreatedTs;
	}

	public void setJctSurveyQtnMasterCreatedTs(Date jctSurveyQtnMasterCreatedTs) {
		this.jctSurveyQtnMasterCreatedTs = jctSurveyQtnMasterCreatedTs;
	}

	public String getJctSurveyQtnMasterModifiedBy() {
		return jctSurveyQtnMasterModifiedBy;
	}

	public void setJctSurveyQtnMasterModifiedBy(String jctSurveyQtnMasterModifiedBy) {
		this.jctSurveyQtnMasterModifiedBy = jctSurveyQtnMasterModifiedBy;
	}

	public Date getJctSurveyQtnMasterModifiedTs() {
		return jctSurveyQtnMasterModifiedTs;
	}

	public void setJctSurveyQtnMasterModifiedTs(Date jctSurveyQtnMasterModifiedTs) {
		this.jctSurveyQtnMasterModifiedTs = jctSurveyQtnMasterModifiedTs;
	}

	public int getJctSurveyQtnMasterProfileId() {
		return jctSurveyQtnMasterProfileId;
	}

	public void setJctSurveyQtnMasterProfileId(int jctSurveyQtnMasterProfileId) {
		this.jctSurveyQtnMasterProfileId = jctSurveyQtnMasterProfileId;
	}

	public String getJctSurveyQtnMasterProfileName() {
		return jctSurveyQtnMasterProfileName;
	}

	public void setJctSurveyQtnMasterProfileName(
			String jctSurveyQtnMasterProfileName) {
		this.jctSurveyQtnMasterProfileName = jctSurveyQtnMasterProfileName;
	}

	public int getJctSurveyQtnMasterOrder() {
		return jctSurveyQtnMasterOrder;
	}

	public void setJctSurveyQtnMasterOrder(int jctSurveyQtnMasterOrder) {
		this.jctSurveyQtnMasterOrder = jctSurveyQtnMasterOrder;
	}
	
}
