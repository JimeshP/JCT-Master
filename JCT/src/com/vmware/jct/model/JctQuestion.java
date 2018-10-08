package com.vmware.jct.model;

import java.io.Serializable;

import javax.persistence.*;

import java.util.Date;
/**
 * 
 * <p><b>Class name:</b> JctQuestion.java</p>
 * <p><b>Author:</b> InterraIt</p>
 * <p><b>Description:</b>  This class basically object reference of jct_questions table </p>
 * <p><b>Copyrights:</b> 	All rights reserved by Interra IT and should only be used for its internal application development.</p>
 * <p><b>Revision History:</b>
 * 	<li>InterraIt - 31/Jan/2014 - Implement comments, introduce Named query </li>
 * </p>
 */

@NamedQueries({
	@NamedQuery(name = "fetchAllQuestionByTitle",
				query = "Select jctQuestion.jctQuestionsDesc from JctQuestion jctQuestion " +
						"where jctQuestion.jctUserProfile.jctUserProfile = :profileId and jctQuestion.jctQuestionBsas=:bsas " +
						"and jctQuestion.jctQuestionsSoftDelete = 0 order by jctQuestion.jctQuestionsId"),
	@NamedQuery(name = "fetchAllActionPlanQuestionByTitle",
				query = "select new com.vmware.jct.dao.dto.QuestionearDTO(jctQuestion.jctQuestionsDesc,jctQuestion.jctQuestionSubDesc) " +
						"from JctQuestion jctQuestion where jctQuestion.jctUserProfile.jctUserProfile = :profileId " +
						"and jctQuestion.jctQuestionBsas=:bsas and jctQuestion.jctQuestionsSoftDelete = 0 order by jctQuestion.jctQuestionsId"),
    @NamedQuery(name = "getActionPlanMainQtnListByProfile",
    			query = "select distinct (jctQuestion.jctQuestionsDesc) from JctQuestion jctQuestion " +
    					"where jctQuestion.jctUserProfile.jctUserProfile = :profileId and jctQuestion.jctQuestionBsas= 'AS' " +
    					"and jctQuestion.jctQuestionsSoftDelete = 0 order by jctQuestion.jctQuestionsId"),
    @NamedQuery(name = "getActionPlanSubQtnList",
    			query = "select distinct (jctQuestion.jctQuestionSubDesc) from JctQuestion jctQuestion " +
    					"where jctQuestion.jctUserProfile.jctUserProfile = :profileId and jctQuestion.jctQuestionBsas= 'AS' " +
    					"and jctQuestion.jctQuestionsDesc = :mainQtn and jctQuestion.jctQuestionsSoftDelete = 0 " +
    					"order by jctQuestion.jctQuestionsId"),
	@NamedQuery(name = "getQestnrMainQtnListByProfile",
				query = "select distinct (jctQuestion.jctQuestionsDesc) from JctQuestion jctQuestion " +
						"where jctQuestion.jctUserProfile.jctUserProfile = :profileId and jctQuestion.jctQuestionBsas= 'BS' " +
						"and jctQuestion.jctQuestionsSoftDelete = 0 order by jctQuestion.jctQuestionsOrder"),
    @NamedQuery(name = "getQtnreSubQtnList",
    			query = "select jctQuestion.jctQuestionSubDesc from JctQuestion jctQuestion " +
    					"where jctQuestion.jctUserProfile.jctUserProfile = :profileId and jctQuestion.jctQuestionBsas= 'BS' " +
    					"and jctQuestion.jctQuestionsDesc = :mainQtn and jctQuestion.jctQuestionsSoftDelete = 0 " +
    					"order by jctQuestion.jctSubQuestionsOrder")
})
@Entity
@Table(name="jct_questions")
public class JctQuestion implements Serializable {
	private static final long serialVersionUID = 1L;

	@Id
	@Column(name="jct_questions_id")
	private int jctQuestionsId;

	@Column(name="jct_bs_created_by")
	private String jctBsCreatedBy;

	@Column(name="jct_bs_lastmodified_by")
	private String jctBsLastmodifiedBy;

	@Temporal(TemporalType.TIMESTAMP)
	@Column(name="jct_bs_lastmodified_ts")
	private Date jctBsLastmodifiedTs;

	@Temporal(TemporalType.TIMESTAMP)
	@Column(name="jct_ds_created_ts")
	private Date jctDsCreatedTs;

	/*@Column(name="jct_levels_desc")
	private String jctLevelsDesc;*/

	@Column(name="jct_questions_desc")
	private String jctQuestionsDesc;

	@Column(name="jct_questions_soft_delete")
	private int jctQuestionsSoftDelete;

	private int version;
	
	@Column(name="jct_question_sub_desc")
	private String jctQuestionSubDesc;
	
	@Column(name="jct_question_bsas ")
	private String jctQuestionBsas ;

	@Column(name="jct_question_order")
	private int jctQuestionsOrder;
	
	@Column(name="jct_sub_question_order")
	private int jctSubQuestionsOrder;
	
	/*//bi-directional many-to-one association to JctLevel
	@ManyToOne
	@JoinColumn(name="jct_levels_id")
	private JctLevel jctLevel;*/
	
	@OneToOne
	@JoinColumn(name="jct_profile_id")
	private JctUserProfile jctUserProfile;
	
	@Column(name="jct_profile_desc")
	private String jctProfilesDesc;
	

	public JctQuestion() {
	}

	public int getJctQuestionsId() {
		return this.jctQuestionsId;
	}

	public void setJctQuestionsId(int jctQuestionsId) {
		this.jctQuestionsId = jctQuestionsId;
	}

	public String getJctBsCreatedBy() {
		return this.jctBsCreatedBy;
	}

	public void setJctBsCreatedBy(String jctBsCreatedBy) {
		this.jctBsCreatedBy = jctBsCreatedBy;
	}

	public String getJctBsLastmodifiedBy() {
		return this.jctBsLastmodifiedBy;
	}

	public void setJctBsLastmodifiedBy(String jctBsLastmodifiedBy) {
		this.jctBsLastmodifiedBy = jctBsLastmodifiedBy;
	}

	public Date getJctBsLastmodifiedTs() {
		return this.jctBsLastmodifiedTs;
	}

	public void setJctBsLastmodifiedTs(Date jctBsLastmodifiedTs) {
		this.jctBsLastmodifiedTs = jctBsLastmodifiedTs;
	}

	public Date getJctDsCreatedTs() {
		return this.jctDsCreatedTs;
	}

	public void setJctDsCreatedTs(Date jctDsCreatedTs) {
		this.jctDsCreatedTs = jctDsCreatedTs;
	}

	

	public String getJctQuestionsDesc() {
		return this.jctQuestionsDesc;
	}

	public void setJctQuestionsDesc(String jctQuestionsDesc) {
		this.jctQuestionsDesc = jctQuestionsDesc;
	}

	public int getJctQuestionsSoftDelete() {
		return this.jctQuestionsSoftDelete;
	}

	public void setJctQuestionsSoftDelete(int jctQuestionsSoftDelete) {
		this.jctQuestionsSoftDelete = jctQuestionsSoftDelete;
	}

	public int getVersion() {
		return this.version;
	}

	public void setVersion(int version) {
		this.version = version;
	}

	/**
	 * @return the jctLevelsDesc
	 *//*
	public String getJctLevelsDesc() {
		return jctLevelsDesc;
	}

	*//**
	 * @param jctLevelsDesc the jctLevelsDesc to set
	 *//*
	public void setJctLevelsDesc(String jctLevelsDesc) {
		this.jctLevelsDesc = jctLevelsDesc;
	}*/

	/**
	 * @return the jctLevel
	 *//*
	public JctLevel getJctLevel() {
		return jctLevel;
	}

	*//**
	 * @param jctLevel the jctLevel to set
	 *//*
	public void setJctLevel(JctLevel jctLevel) {
		this.jctLevel = jctLevel;
	}*/

	/**
	 * @return the jctQuestionSubDesc
	 */
	public String getJctQuestionSubDesc() {
		return jctQuestionSubDesc;
	}

	/**
	 * @param jctQuestionSubDesc the jctQuestionSubDesc to set
	 */
	public void setJctQuestionSubDesc(String jctQuestionSubDesc) {
		this.jctQuestionSubDesc = jctQuestionSubDesc;
	}

	/**
	 * @return the jctQuestionBsas
	 */
	public String getJctQuestionBsas() {
		return jctQuestionBsas;
	}

	/**
	 * @param jctQuestionBsas the jctQuestionBsas to set
	 */
	public void setJctQuestionBsas(String jctQuestionBsas) {
		this.jctQuestionBsas = jctQuestionBsas;
	}
	
	public int getJctQuestionsOrder() {
		return jctQuestionsOrder;
	}

	public void setJctQuestionsOrder(int jctQuestionsOrder) {
		this.jctQuestionsOrder = jctQuestionsOrder;
	}
	
	public int getJctSubQuestionsOrder() {
		return jctSubQuestionsOrder;
	}

	public void setJctSubQuestionsOrder(int jctSubQuestionsOrder) {
		this.jctSubQuestionsOrder = jctSubQuestionsOrder;
	}
	

}