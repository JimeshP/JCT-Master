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
	@NamedQuery(name = "getAllSurveyQuestion",
				query = "select new com.vmware.jct.dao.dto.SurveyQuestionDTO( "
						+ "survey.jctSurveyQtnMasterUserType, survey.jctSurveyQtnMasterAnsType, survey.jctSurveyQtnMasterMainQtn, survey.jctSurveyQtnMasterMandatory, survey.jctSurveyQtnMasterProfileName) "
						+ "from JctSurveyQuestionMaster survey WHERE survey.jctSurveyQtnMasterSoftDel = 0 "
						+ "GROUP BY survey.jctSurveyQtnMasterMainQtn, survey.jctSurveyQtnMasterUserType ORDER BY survey.jctSurveyQtnMasterOrder"),
	@NamedQuery(name = "getAllSurveyQuestionByUserType",
				query = "select new com.vmware.jct.dao.dto.SurveyQuestionDTO( "
						+ "survey.jctSurveyQtnMasterUserType, survey.jctSurveyQtnMasterAnsType, survey.jctSurveyQtnMasterMainQtn, survey.jctSurveyQtnMasterMandatory, survey.jctSurveyQtnMasterProfileName) "
						+ "from JctSurveyQuestionMaster survey WHERE survey.jctSurveyQtnMasterSoftDel = 0 "
						+ "AND survey.jctSurveyQtnMasterUserType = :userType GROUP BY survey.jctSurveyQtnMasterMainQtn ORDER BY survey.jctSurveyQtnMasterOrder"),	
	@NamedQuery(name = "getAllSurveyQuestionByUserTypeAndQstnType",
				query = "select new com.vmware.jct.dao.dto.SurveyQuestionDTO( "
						+ "survey.jctSurveyQtnMasterUserType, survey.jctSurveyQtnMasterAnsType, survey.jctSurveyQtnMasterMainQtn, survey.jctSurveyQtnMasterMandatory, survey.jctSurveyQtnMasterProfileName) "
						+ "from JctSurveyQuestionMaster survey WHERE survey.jctSurveyQtnMasterSoftDel = 0 "
						+ "AND survey.jctSurveyQtnMasterUserType = :userType AND survey.jctSurveyQtnMasterAnsType = :qtnType GROUP BY survey.jctSurveyQtnMasterMainQtn ORDER BY survey.jctSurveyQtnMasterOrder"),
	@NamedQuery(name = "getAllSurveySubQuestionsFromMainQuestion",
				query = "select new com.vmware.jct.dao.dto.SurveyQuestionDTO(survey.jctSurveyQtnMasterSubQtn) "
						+ "from JctSurveyQuestionMaster survey WHERE survey.jctSurveyQtnMasterSoftDel = 0 "
						+ "AND survey.jctSurveyQtnMasterAnsType = :answerType AND survey.jctSurveyQtnMasterUserType = :userType "
						+ "AND survey.jctSurveyQtnMasterMainQtn = :mainQtn"),
	@NamedQuery(name = "getNosOfSubQtns",
				query = "select COUNT(survey.jctSurveyQtnMasterSubQtn) "
						+ "from JctSurveyQuestionMaster survey WHERE survey.jctSurveyQtnMasterSoftDel = 0 "
						+ "AND survey.jctSurveyQtnMasterAnsType = :answerType AND survey.jctSurveyQtnMasterUserType = :userType "
						+ "AND survey.jctSurveyQtnMasterMainQtn = :mainQtn"),
	@NamedQuery(name = "softDeleteTheSurveyQuestion",
				query = "UPDATE JctSurveyQuestionMaster survey SET survey.jctSurveyQtnMasterSoftDel = 1 WHERE "
						+ "survey.jctSurveyQtnMasterUserType = :userType AND survey.jctSurveyQtnMasterAnsType = :answerType "
						+ "AND survey.jctSurveyQtnMasterMainQtn = :qtnDescription"),
	@NamedQuery(name = "fetchTextSurveyQtnsForEdit",
				query = "from JctSurveyQuestionMaster survey WHERE survey.jctSurveyQtnMasterSoftDel = 0 AND "
						+ "survey.jctSurveyQtnMasterUserType = :userType AND survey.jctSurveyQtnMasterAnsType = :answerType "
						+ "AND survey.jctSurveyQtnMasterMainQtn = :originalQuestion"),
	@NamedQuery(name = "getAllSurveyQuestionByQstnType",
				query = "select new com.vmware.jct.dao.dto.SurveyQuestionDTO( "
						+ "survey.jctSurveyQtnMasterUserType, survey.jctSurveyQtnMasterAnsType, survey.jctSurveyQtnMasterMainQtn, survey.jctSurveyQtnMasterMandatory, survey.jctSurveyQtnMasterProfileName) "
						+ "from JctSurveyQuestionMaster survey WHERE survey.jctSurveyQtnMasterSoftDel = 0 "
						+ "AND survey.jctSurveyQtnMasterAnsType = :qtnType GROUP BY survey.jctSurveyQtnMasterMainQtn ORDER BY survey.jctSurveyQtnMasterOrder"),
	@NamedQuery(name = "getAllSurveySubQuestionsFromMainQuestionOnlyAnsType",
				query = "select new com.vmware.jct.dao.dto.SurveyQuestionDTO(survey.jctSurveyQtnMasterSubQtn) "
						+ "from JctSurveyQuestionMaster survey WHERE survey.jctSurveyQtnMasterSoftDel = 0 "
						+ "AND survey.jctSurveyQtnMasterAnsType = :answerType "
						+ "AND survey.jctSurveyQtnMasterMainQtn = :mainQtn"),
	@NamedQuery(name = "getNosOfSubQtnsByAnsType",
				query = "select COUNT(survey.jctSurveyQtnMasterSubQtn) "
						+ "from JctSurveyQuestionMaster survey WHERE survey.jctSurveyQtnMasterSoftDel = 0 "
						+ "AND survey.jctSurveyQtnMasterAnsType = :answerType "
						+ "AND survey.jctSurveyQtnMasterMainQtn = :mainQtn"),
	@NamedQuery(name = "getTextSurveyQuestion",
				query = "select DISTINCT(survey.jctSurveyQtnMasterMainQtn), survey.jctSurveyQtnMasterMandatory " +
						"from JctSurveyQuestionMaster survey WHERE "
						+ "survey.jctSurveyQtnMasterSoftDel = 0 AND jctSurveyQtnMasterAnsType = 1 " +
						"AND jctSurveyQtnMasterUserType = 3"),
	@NamedQuery(name = "getMainQuestionForMultiple",
				query = "select DISTINCT(survey.jctSurveyQtnMasterMainQtn), survey.jctSurveyQtnMasterMandatory " +
						"from JctSurveyQuestionMaster survey WHERE "
						+ "survey.jctSurveyQtnMasterSoftDel = 0 AND jctSurveyQtnMasterAnsType = :ansType " +
						"AND jctSurveyQtnMasterUserType = 3"),
	@NamedQuery(name = "getSubQuestions",
				query = "select survey.jctSurveyQtnMasterSubQtn " +
						"from JctSurveyQuestionMaster survey WHERE "
						+ "survey.jctSurveyQtnMasterSoftDel = 0 AND jctSurveyQtnMasterAnsType = :ansType " +
						"AND jctSurveyQtnMasterMainQtn = :mainQtn AND jctSurveyQtnMasterUserType = 3"),
	@NamedQuery(name = "fetchgSurveyQtnsByAllParams",
				query = "select new com.vmware.jct.dao.dto.SurveyQuestionDTO( "
						+ "survey.jctSurveyQtnMasterUserType, survey.jctSurveyQtnMasterAnsType, survey.jctSurveyQtnMasterMainQtn, survey.jctSurveyQtnMasterMandatory, survey.jctSurveyQtnMasterProfileName) "
						+ "from JctSurveyQuestionMaster survey WHERE survey.jctSurveyQtnMasterSoftDel = 0 "
						+ "AND survey.jctSurveyQtnMasterUserType = :userType AND survey.jctSurveyQtnMasterAnsType = :qtnType AND survey.jctSurveyQtnMasterProfileId = :profileId GROUP BY survey.jctSurveyQtnMasterMainQtn ORDER BY survey.jctSurveyQtnMasterOrder"),
	@NamedQuery(name = "getAllSurveyQuestionByQstnTypeAndProfileId",
				query = "select new com.vmware.jct.dao.dto.SurveyQuestionDTO( "
						+ "survey.jctSurveyQtnMasterUserType, survey.jctSurveyQtnMasterAnsType, survey.jctSurveyQtnMasterMainQtn, survey.jctSurveyQtnMasterMandatory, survey.jctSurveyQtnMasterProfileName) "
						+ "from JctSurveyQuestionMaster survey WHERE survey.jctSurveyQtnMasterSoftDel = 0 "
						+ "AND survey.jctSurveyQtnMasterAnsType = :qtnType AND survey.jctSurveyQtnMasterProfileId = :profileId GROUP BY survey.jctSurveyQtnMasterMainQtn ORDER BY survey.jctSurveyQtnMasterOrder"),
	@NamedQuery(name = "getSrvQtnByUserTypeAndProfileId",
				query = "select new com.vmware.jct.dao.dto.SurveyQuestionDTO( "
						+ "survey.jctSurveyQtnMasterUserType, survey.jctSurveyQtnMasterAnsType, survey.jctSurveyQtnMasterMainQtn, survey.jctSurveyQtnMasterMandatory, survey.jctSurveyQtnMasterProfileName) "
						+ "from JctSurveyQuestionMaster survey WHERE survey.jctSurveyQtnMasterSoftDel = 0 "
						+ "AND survey.jctSurveyQtnMasterUserType = :userType AND survey.jctSurveyQtnMasterProfileId = :profileId GROUP BY survey.jctSurveyQtnMasterMainQtn ORDER BY survey.jctSurveyQtnMasterOrder"),
	@NamedQuery(name = "getAllSubQtnsByUserTypeAndProfile",
				query = "select new com.vmware.jct.dao.dto.SurveyQuestionDTO(survey.jctSurveyQtnMasterSubQtn) "
						+ "from JctSurveyQuestionMaster survey WHERE survey.jctSurveyQtnMasterSoftDel = 0 "
						+ "AND survey.jctSurveyQtnMasterAnsType = :answerType AND survey.jctSurveyQtnMasterUserType = :userType "
						+ "AND survey.jctSurveyQtnMasterMainQtn = :mainQtn AND survey.jctSurveyQtnMasterProfileId = :profileId"),
	@NamedQuery(name = "getNosOfSubQtnsProfileId",
				query = "select COUNT(survey.jctSurveyQtnMasterSubQtn) "
						+ "from JctSurveyQuestionMaster survey WHERE survey.jctSurveyQtnMasterSoftDel = 0 "
						+ "AND survey.jctSurveyQtnMasterAnsType = :answerType AND survey.jctSurveyQtnMasterUserType = :userType AND survey.jctSurveyQtnMasterProfileId = :profileId "
						+ "AND survey.jctSurveyQtnMasterMainQtn = :mainQtn"),
	@NamedQuery(name = "getAllSurveySubQuestionsFromMainQuestionProfileId",
				query = "select new com.vmware.jct.dao.dto.SurveyQuestionDTO(survey.jctSurveyQtnMasterSubQtn) "
						+ "from JctSurveyQuestionMaster survey WHERE survey.jctSurveyQtnMasterSoftDel = 0 "
						+ "AND survey.jctSurveyQtnMasterAnsType = :answerType AND survey.jctSurveyQtnMasterUserType = :userType "
						+ "AND survey.jctSurveyQtnMasterMainQtn = :mainQtn AND survey.jctSurveyQtnMasterProfileId = :profileId"),
	@NamedQuery(name = "fetchToUpdateSurveyQtn",
				query = "from JctSurveyQuestionMaster obj where obj.jctSurveyQtnMasterProfileId = :profileId " +
						"and obj.jctSurveyQtnMasterMainQtn = :mainQtn and obj.jctSurveyQtnMasterUserType = :userType " +
						"and obj.jctSurveyQtnMasterAnsType = :ansType " +
						"and obj.jctSurveyQtnMasterSoftDel = 0"),
	@NamedQuery(name = "fetchOrderSurveyQtn",
				query = "select max(obj.jctSurveyQtnMasterOrder) from JctSurveyQuestionMaster obj " +
						"where obj.jctSurveyQtnMasterSoftDel = 0 and obj.jctSurveyQtnMasterProfileId = :profileId " +
						"and obj.jctSurveyQtnMasterUserType = :userType"),
	@NamedQuery(name = "getAllSurveyFaciQuestions",
				query = "select DISTINCT(survey.jctSurveyQtnMasterMainQtn), survey.jctSurveyQtnMasterMandatory, survey.jctSurveyQtnMasterAnsType from JctSurveyQuestionMaster survey WHERE "
						+ "survey.jctSurveyQtnMasterSoftDel = 0 " +
						"AND survey.jctSurveyQtnMasterUserType = 3 ORDER BY survey.jctSurveyQtnMasterOrder")
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
	
	@Column(name="jct_survey_qtn_master_mandatory")
	private String jctSurveyQtnMasterMandatory;	

	@Column(name="jct_survey_qtn_master_created_by")
	private String jctSurveyQtnMasterCreatedBy;
	
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
	public String getJctSurveyQtnMasterMandatory() {
		return jctSurveyQtnMasterMandatory;
	}

	public void setJctSurveyQtnMasterMandatory(String jctSurveyQtnMasterMandatory) {
		this.jctSurveyQtnMasterMandatory = jctSurveyQtnMasterMandatory;
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
