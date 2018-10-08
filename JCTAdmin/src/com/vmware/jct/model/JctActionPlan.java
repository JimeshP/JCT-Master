package com.vmware.jct.model;
import java.io.Serializable;
import javax.persistence.*;

import java.util.Date;
/**
 * 
 * <p><b>Class name:</b> JctActionPlan.java</p>
 * <p><b>Author:</b> InterraIT</p>
 * <p><b>Description:</b>  This class basically object reference of jct_action_plan table </p>
 * <p><b>Copyrights:</b> 	All rights reserved by InterraIT and should only be used for its internal application development.</p>
 * <p><b>Revision History:</b>
 * 	<li>InterraIT - 31/Jan/2014 - Implement comments, introduce Named query </li>
 * </p>
 */
@NamedQueries({
	@NamedQuery(name = "fetchListOfActionPlan",query = "from JctActionPlan where jctJobrefNo = :refNo and jctActionPlanSoftDelete=0")
})
@Entity
@Table(name="jct_action_plan")
public class JctActionPlan implements Serializable {
	private static final long serialVersionUID = 1L;

	@Id
	@Column(name="jct_action_plan_id")
	private int jctActionPlanId;

	@Column(name="jct_as_answar_desc")
	private String jctAsAnswarDesc;

	@Column(name="jct_as_created_by")
	private String jctAsCreatedBy;

	@Temporal(TemporalType.TIMESTAMP)
	@Column(name="jct_as_created_ts")
	private Date jctAsCreatedTs;

	@Column(name="jct_as_lastmodified_by")
	private String jctAsLastmodifiedBy;

	@Temporal(TemporalType.TIMESTAMP)
	@Column(name="jct_as_lastmodified_ts")
	private Date jctAsLastmodifiedTs;

	@Column(name="jct_as_question_desc")
	private String jctAsQuestionDesc;

	@Column(name="jct_as_question_sub_desc")
	private String jctAsQuestionSubDesc;

	@Column(name="jct_jobref_no")
	private String jctJobrefNo;
	
	@Column(name="jct_action_plan_soft_delete")
	private int jctActionPlanSoftDelete;
	
	@Column(name="jct_as_header_id")
	private int jctAsHeaderId;

	public int getJctAsHeaderId() {
		return jctAsHeaderId;
	}

	public void setJctAsHeaderId(int jctAsHeaderId) {
		this.jctAsHeaderId = jctAsHeaderId;
	}

	public int getJctActionPlanSoftDelete() {
		return jctActionPlanSoftDelete;
	}

	public void setJctActionPlanSoftDelete(int jctActionPlanSoftDelete) {
		this.jctActionPlanSoftDelete = jctActionPlanSoftDelete;
	}

	@Version
	private int version;

	//bi-directional many-to-one association to JctStatus
	@ManyToOne
	@JoinColumn(name="jct_status_id")
	private JctStatus jctStatus;

	public JctActionPlan() {
	}

	public int getJctActionPlanId() {
		return this.jctActionPlanId;
	}

	public void setJctActionPlanId(int jctActionPlanId) {
		this.jctActionPlanId = jctActionPlanId;
	}

	public String getJctAsAnswarDesc() {
		return this.jctAsAnswarDesc;
	}

	public void setJctAsAnswarDesc(String jctAsAnswarDesc) {
		this.jctAsAnswarDesc = jctAsAnswarDesc;
	}

	public String getJctAsCreatedBy() {
		return this.jctAsCreatedBy;
	}

	public void setJctAsCreatedBy(String jctAsCreatedBy) {
		this.jctAsCreatedBy = jctAsCreatedBy;
	}

	public Date getJctAsCreatedTs() {
		return this.jctAsCreatedTs;
	}

	public void setJctAsCreatedTs(Date jctAsCreatedTs) {
		this.jctAsCreatedTs = jctAsCreatedTs;
	}

	public String getJctAsLastmodifiedBy() {
		return this.jctAsLastmodifiedBy;
	}

	public void setJctAsLastmodifiedBy(String jctAsLastmodifiedBy) {
		this.jctAsLastmodifiedBy = jctAsLastmodifiedBy;
	}

	public Date getJctAsLastmodifiedTs() {
		return this.jctAsLastmodifiedTs;
	}

	public void setJctAsLastmodifiedTs(Date jctAsLastmodifiedTs) {
		this.jctAsLastmodifiedTs = jctAsLastmodifiedTs;
	}

	public String getJctAsQuestionDesc() {
		return this.jctAsQuestionDesc;
	}

	public void setJctAsQuestionDesc(String jctAsQuestionDesc) {
		this.jctAsQuestionDesc = jctAsQuestionDesc;
	}

	public String getJctAsQuestionSubDesc() {
		return this.jctAsQuestionSubDesc;
	}

	public void setJctAsQuestionSubDesc(String jctAsQuestionSubDesc) {
		this.jctAsQuestionSubDesc = jctAsQuestionSubDesc;
	}

	public String getJctJobrefNo() {
		return this.jctJobrefNo;
	}

	public void setJctJobrefNo(String jctJobrefNo) {
		this.jctJobrefNo = jctJobrefNo;
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


}
