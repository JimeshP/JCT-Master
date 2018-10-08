package com.vmware.jct.model;

import java.io.Serializable;
import java.util.Date;

import javax.persistence.Column;
import javax.persistence.Entity;
import javax.persistence.GeneratedValue;
import javax.persistence.GenerationType;
import javax.persistence.Id;
import javax.persistence.JoinColumn;
import javax.persistence.ManyToOne;
import javax.persistence.NamedQueries;
import javax.persistence.NamedQuery;
import javax.persistence.Table;
import javax.persistence.Temporal;
import javax.persistence.TemporalType;
import javax.persistence.Version;
/**
 * 
 * <p><b>Class name:</b> JctBeforeSketchQuestion.java</p>
 * <p><b>Author:</b> InterraIt</p>
 * <p><b>Description:</b>  This class basically object reference of jct_before_sketch_question table </p>
 * <p><b>Copyrights:</b> 	All rights reserved by Interra IT and should only be used for its internal application development.</p>
 * <p><b>Revision History:</b>
 * 	<li>InterraIt - 31/Jan/2014 - Implement comments, introduce Named query </li>
 * </p>
 */


@NamedQueries({
	@NamedQuery(name = "fetchAllQuestion",
				query = "from JctBeforeSketchQuestion where jctBsJobrefNo = :refNo " +
						"and jctBsSoftDelete = 0 and jctStatus = 4"),
	@NamedQuery(name = "fetchAllQuestionByRefNo",
				query = "from JctBeforeSketchQuestion where jctBsJobrefNo = :jobRefNo and jctBsSoftDelete = :deleted"),
	@NamedQuery(name = "getQuestionnaireMainQtnListByJrNo",
				query = "select distinct (qtn.jctBsQuestionDesc) from JctBeforeSketchQuestion qtn " +
						"where jctBsJobrefNo = :refNo and jctBsSoftDelete = 0"),
    @NamedQuery(name = "getQuestionnaireSubQtnListByJrNo",
    			query = "select subqtn.jctBsSubQuestion from JctBeforeSketchQuestion subqtn " +
    					"where jctBsJobrefNo = :refNo and subqtn.jctBsSoftDelete = 0 and subqtn.jctBsQuestionDesc = :mainQuestion"),
    @NamedQuery(name = "getQuestionnaireAnswerListByJrNo",
    			query = "select ans.jctBsAnswarDesc from JctBeforeSketchQuestion ans " +
    					"where ans.jctBsJobrefNo = :refNo and ans.jctBsSoftDelete = 0 and " +
    					"ans.jctBsQuestionDesc = :mainQuestion and ans.jctBsSubQuestion = :subQuestion"),
    @NamedQuery(name = "getFrozenQtnCount", 
    			query = "SELECT COUNT (jctBsJobrefNo) FROM JctBeforeSketchQuestion qtn WHERE " +
    					"qtn.jctBsJobrefNo = :jrNo AND qtn.jctBsCreatedBy = :user AND " +
    					"qtn.jctUser.jctUserId = :userId AND qtn.jctBsSoftDelete = 0"),
    @NamedQuery(name="getDummyQuestionnaireData", 
    			query = "from JctBeforeSketchQuestion where jctBsJobrefNo = :jrNo"),
    @NamedQuery(name = "deleteDummyQtnData", 
    			query = "DELETE FROM JctBeforeSketchQuestion qtn where qtn.jctBsJobrefNo = :jrNo")
})

@Entity
@Table(name="jct_before_sketch_question")
public class JctBeforeSketchQuestion implements Serializable {
	private static final long serialVersionUID = 1L;

	@Id
	@GeneratedValue(strategy=GenerationType.IDENTITY)
	@Column(name="jct_bs_question_id")
	private int jctBsQuestionId;

	@Column(name="jct_bs_created_by")
	private String jctBsCreatedBy;

	@Column(name="jct_bs_jobref_no")
	private String jctBsJobrefNo;

	@Column(name="jct_bs_lastmodified_by")
	private String jctBsLastmodifiedBy;

	@Temporal(TemporalType.TIMESTAMP)
	@Column(name="jct_bs_lastmodified_ts")
	private Date jctBsLastmodifiedTs;

	@Column(name="jct_bs_question_desc")
	private String jctBsQuestionDesc;

	@Temporal(TemporalType.TIMESTAMP)
	@Column(name="jct_ds_created_ts")
	private Date jctDsCreatedTs;

	@Version
	private int version;
	
	@Column(name="jct_bs_answar_desc")
	private String jctBsAnswarDesc;
	
	@Column(name="jct_bs_soft_delete")
	private int jctBsSoftDelete;

	//bi-directional many-to-one association to JctStatus
	@ManyToOne
	@JoinColumn(name="jct_bs_status_id")
	private JctStatus jctStatus;
	
	/** THIS IS FOR PUBLIC VERSION **/
	@ManyToOne
	@JoinColumn(name="jct_bs_user_id")
	private JctUser jctUser;
	/********************************/
	
	@Column(name="jct_bs_header_id")
	private int jctBsHeaderId;
	
	@Column(name="jct_bs_sub_question")
	private String jctBsSubQuestion;

	public int getJctBsHeaderId() {
		return jctBsHeaderId;
	}

	public void setJctBsHeaderId(int jctBsHeaderId) {
		this.jctBsHeaderId = jctBsHeaderId;
	}

	public JctBeforeSketchQuestion() {
	}

	public int getJctBsQuestionId() {
		return this.jctBsQuestionId;
	}

	public void setJctBsQuestionId(int jctBsQuestionId) {
		this.jctBsQuestionId = jctBsQuestionId;
	}

	public String getJctBsCreatedBy() {
		return this.jctBsCreatedBy;
	}

	public void setJctBsCreatedBy(String jctBsCreatedBy) {
		this.jctBsCreatedBy = jctBsCreatedBy;
	}

	public String getJctBsJobrefNo() {
		return this.jctBsJobrefNo;
	}

	public void setJctBsJobrefNo(String jctBsJobrefNo) {
		this.jctBsJobrefNo = jctBsJobrefNo;
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

	public String getJctBsQuestionDesc() {
		return this.jctBsQuestionDesc;
	}

	public void setJctBsQuestionDesc(String jctBsQuestionDesc) {
		this.jctBsQuestionDesc = jctBsQuestionDesc;
	}

	public Date getJctDsCreatedTs() {
		return this.jctDsCreatedTs;
	}

	public void setJctDsCreatedTs(Date jctDsCreatedTs) {
		this.jctDsCreatedTs = jctDsCreatedTs;
	}

	public int getVersion() {
		return this.version;
	}

	public void setVersion(int version) {
		this.version = version;
	}

	public JctStatus getJctStatus() {
		return this.jctStatus;
	}

	public void setJctStatus(JctStatus jctStatus) {
		this.jctStatus = jctStatus;
	}
	
	public String getJctBsAnswarDesc() {
		return jctBsAnswarDesc;
	}

	public void setJctBsAnswarDesc(String jctBsAnswarDesc) {
		this.jctBsAnswarDesc = jctBsAnswarDesc;
	}
	
	public int getJctBsSoftDelete() {
		return jctBsSoftDelete;
	}

	public void setJctBsSoftDelete(int jctBsSoftDelete) {
		this.jctBsSoftDelete = jctBsSoftDelete;
	}

	public String getJctBsSubQuestion() {
		return jctBsSubQuestion;
	}

	public void setJctBsSubQuestion(String jctBsSubQuestion) {
		this.jctBsSubQuestion = jctBsSubQuestion;
	}

	public JctUser getJctUser() {
		return jctUser;
	}

	public void setJctUser(JctUser jctUser) {
		this.jctUser = jctUser;
	}
}