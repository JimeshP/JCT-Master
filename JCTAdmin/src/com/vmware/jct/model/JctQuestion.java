package com.vmware.jct.model;

import java.io.Serializable;
import javax.persistence.*;

import java.util.Date;
/**
 * 
 * <p><b>Class name:</b> JctQuestion.java</p>
 * <p><b>Author:</b> InterraIT</p>
 * <p><b>Description:</b>  This class basically object reference of jct_questions table </p>
 * <p><b>Copyrights:</b>All rights reserved by InterraIT and should only be used for its 
 * internal application development.</p>
 * <p><b>Revision History:</b>
 * 	<li>InterraIT - 31/Jan/2014 - Implement comments, introduce Named query </li>
 * </p>
 */
@NamedQueries({
        @NamedQuery(name = "fetchAllRefQtn",
        			query = "select new com.vmware.jct.dao.dto.QuestionearDTO(questions.jctQuestionsDesc, questions.jctQuestionSubDesc, " +
        					"questions.jctProfilesDesc, questions.jctQuestionBsas, questions.jctQuestionsId, " +
        					"questions.jctQuestionsSoftDelete, questions.jctUserProfile.jctUserProfile, " +
        					"questions.jctQuestionsOrder, questions.jctSubQuestionsOrder) from JctQuestion questions " +
        					"where questions.jctQuestionsSoftDelete = 0 and questions.jctQuestionBsas = :relatedPage " +
        					"order by questions.jctUserProfile.jctUserProfile, questions.jctQuestionsOrder, " +
        					"questions.jctSubQuestionsOrder, questions.jctBsLastmodifiedTs desc"),
		@NamedQuery(name = "fetchAllRefQtnByProfileId",
					query = "select new com.vmware.jct.dao.dto.QuestionearDTO(questions.jctQuestionsDesc, questions.jctQuestionSubDesc, " +
							"questions.jctProfilesDesc, questions.jctQuestionBsas, questions.jctQuestionsId, " +
							"questions.jctQuestionsSoftDelete, questions.jctUserProfile.jctUserProfile, questions.jctQuestionsOrder, " +
							"questions.jctSubQuestionsOrder) from JctQuestion questions where questions.jctQuestionsSoftDelete = 0 " +
							"and questions.jctUserProfile.jctUserProfile = :profileId and questions.jctQuestionBsas = :relatedPage " +
							"order by questions.jctQuestionsOrder, questions.jctSubQuestionsOrder, questions.jctBsLastmodifiedTs desc"),
    	@NamedQuery(name = "fetchRefQtn",
    				query = "from JctQuestion questions where  questions.jctQuestionsId = :pkId and jctQuestionsSoftDelete = 0"),
    	@NamedQuery(name = "fetchRefQtnByOrder",
    				query = "from JctQuestion questions where questions.jctUserProfile.jctUserProfile = :profileId " +
    						"and questions.jctSubQuestionsOrder= :qtnOrder and questions.jctQuestionBsas=:page " +
    						"and questions.jctQuestionsSoftDelete = 0"),
    	@NamedQuery(name = "fetchRefQtnByDesc",
    				query = "from JctQuestion questions where questions.jctQuestionsSoftDelete = 0 and " +
    						"UPPER(questions.jctQuestionsDesc)= :refQtnDesc and questions.jctUserProfile.jctUserProfile = :userProfileId"),
    	@NamedQuery(name = "fetchActionPlanByDesc",
    				query = "from JctQuestion questions where questions.jctQuestionsSoftDelete = 0 and " +
    						"UPPER(questions.jctQuestionsDesc)= :refQtnDesc and UPPER(questions.jctQuestionSubDesc)= :subQtnDesc " +
    						"and questions.jctUserProfile.jctUserProfile = :userProfileId and questions.jctQuestionBsas=:code"),
        @NamedQuery(name = "fetchQuestionBySubOrder",
        			query = "select new com.vmware.jct.dao.dto.QuestionearDTO(questions.jctQuestionsId) from JctQuestion questions " +
        					"where questions.jctSubQuestionsOrder > :subQtnOrder and questions.jctQuestionsOrder = :mainQtnOrder " +
        					"and questions.jctQuestionsDesc = :mainQtnDesc and questions.jctQuestionsSoftDelete = 0 and " +
        					"questions.jctUserProfile.jctUserProfile = :profileId and questions.jctQuestionBsas = :relatedPage"),
    	@NamedQuery(name = "fetchOrderQtn",
    				query = "select max(questions.jctQuestionsOrder) from JctQuestion questions " +
    						"where questions.jctQuestionsSoftDelete = 0 and questions.jctUserProfile.jctUserProfile = :profileId " +
    						"and questions.jctQuestionBsas= :qtnCode and questions.jctQuestionsDesc <> :qtnDesc"),
    	@NamedQuery(name = "fetchMainQtnList",
    				query = "select DISTINCT new com.vmware.jct.dao.dto.QuestionearDTO( questions.jctQuestionsDesc, questions.jctQuestionsOrder) " +
    						"from JctQuestion questions where questions.jctQuestionsSoftDelete = 0 and " +
    						"questions.jctUserProfile.jctUserProfile = :profileId and questions.jctQuestionBsas = :relatedPage " +
    						"order by questions.jctQuestionsOrder"),
    	@NamedQuery(name = "fetchSubQtnList",
    				query = "select new com.vmware.jct.dao.dto.QuestionearDTO( questions.jctQuestionSubDesc, questions.jctSubQuestionsOrder) " +
    						"from JctQuestion questions where questions.jctQuestionsSoftDelete = 0 and " +
    						"questions.jctUserProfile.jctUserProfile = :profileId and questions.jctQuestionBsas = :relatedPage " +
    						"and questions.jctQuestionsDesc =:mainQtn order by questions.jctSubQuestionsOrder"),
    	@NamedQuery(name = "fetchToUpdateMainOrder",
    				query = "from JctQuestion questions where questions.jctUserProfile.jctUserProfile = :profileId " +
    						"and questions.jctQuestionsDesc = :mainQtnDesc and questions.jctQuestionBsas = :relatedPage " +
    						"and questions.jctQuestionsSoftDelete = 0"),
		@NamedQuery(name = "mainQtnList",
					query = "select new com.vmware.jct.dao.dto.QuestionearDTO(questions.jctQuestionsId) " +
							"from JctQuestion questions where questions.jctQuestionsDesc = :mainQtnDesc and questions.jctQuestionsSoftDelete = 0 " +
							"and questions.jctUserProfile.jctUserProfile = :profileId and questions.jctQuestionBsas = :relatedPage"),
		@NamedQuery(name = "fetchToUpdateSubOrder",
					query = "from JctQuestion questions where questions.jctUserProfile.jctUserProfile = :profileId and " +
							"questions.jctQuestionsDesc = :mainQtn and questions.jctQuestionSubDesc = :subQtnDesc " +
							"and questions.jctQuestionBsas = :relatedPage and questions.jctQuestionsSoftDelete = 0"),
		@NamedQuery(name = "fetchTotalSubQtn",
					query = "select new com.vmware.jct.dao.dto.QuestionearDTO(questions.jctQuestionsId) from JctQuestion questions " +
							"where questions.jctQuestionsDesc = :mainQtnDesc and questions.jctQuestionsSoftDelete = 0 " +
							"and questions.jctUserProfile.jctUserProfile = :profileId and questions.jctQuestionBsas = :relatedPage"),
	    @NamedQuery(name = "fetchMainQtnByOrder",
	    			query = "select new com.vmware.jct.dao.dto.QuestionearDTO(questions.jctQuestionsId) from JctQuestion questions " +
	    					"where questions.jctQuestionsOrder > :mainQtnOrder and questions.jctQuestionsSoftDelete = 0 and " +
	    					"questions.jctUserProfile.jctUserProfile = :profileId and questions.jctQuestionBsas = :relatedPage"),
    	@NamedQuery(name = "fetchAllExistingQtn",
    				query = "select DISTINCT new com.vmware.jct.dao.dto.QuestionearDTO( questions.jctQuestionsDesc) from JctQuestion questions " +
    						"where questions.jctQuestionsSoftDelete = 0 and questions.jctQuestionBsas =:code and " +
    						"questions.jctUserProfile.jctUserProfile = :userProfileId")
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

	@Column(name="jct_questions_desc")
	private String jctQuestionsDesc;

	@Column(name="jct_questions_soft_delete")
	private int jctQuestionsSoftDelete;

	private int version;
	
	@Column(name="jct_question_sub_desc")
	private String jctQuestionSubDesc;
	
	@Column(name="jct_question_bsas ")
	private String jctQuestionBsas ;

	
	//bi-directional many-to-one association to JctLevel
	/*@ManyToOne
	@JoinColumn(name="jct_profile_id")
	private JctLevel jctProfile;*/
	
	@OneToOne
	@JoinColumn(name="jct_profile_id")
	private JctUserProfile jctUserProfile;
	
	@Column(name="jct_profile_desc")
	private String jctProfilesDesc;

	@Column(name="jct_question_order")
	private int jctQuestionsOrder;
	
	@Column(name="jct_sub_question_order")
	private int jctSubQuestionsOrder;
	
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

	public String getJctProfilesDesc() {
		return jctProfilesDesc;
	}

	public void setJctProfilesDesc(String jctProfilesDesc) {
		this.jctProfilesDesc = jctProfilesDesc;
	}

	public JctUserProfile getJctUserProfile() {
		return jctUserProfile;
	}

	public void setJctUserProfile(JctUserProfile jctUserProfile) {
		this.jctUserProfile = jctUserProfile;
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