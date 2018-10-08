/**
 * 
 */
package com.vmware.jct.model;
import java.io.Serializable;

import javax.persistence.*;

import java.util.Date;
/**
 * 
 * <p><b>Class name:</b> JctActionPlan.java</p>
 * <p><b>Author:</b> InterraIt</p>
 * <p><b>Description:</b>  This class basically object reference of jct_action_plan table. Store all action plan related information  in this table </p>
 * <p><b>Copyrights:</b> 	All rights reserved by Interra IT and should only be used for its internal application development.</p>
 * <p><b>Revision History:</b>
 * 	<li>InterraIt - 31/Jan/2014 - Implement comments </li>
 * </p>
 */
@NamedQueries({
	@NamedQuery(name = "fetchListOfActionPlan",
				query = "from JctActionPlan where jctJobrefNo = :refNo and jctActionPlanSoftDelete=0 and jctStatus = 4"),
	@NamedQuery(name = "fetchListOfActionPlanEdited",
				query = "from JctActionPlan where jctJobrefNo = :refNo and jctActionPlanSoftDelete=0 and jctStatus = 5"),
	@NamedQuery(name = "getActionPlanMainQtnListByJrNo",
				query = "select distinct (actionPlan.jctAsQuestionDesc) from JctActionPlan actionPlan " +
						"where jctJobrefNo = :refNo and jctActionPlanSoftDelete=0"),
    @NamedQuery(name = "getActionPlanSubQtnListByJrNo",
    			query = "select distinct (plan.jctAsQuestionSubDesc) from JctActionPlan plan " +
    					"where jctJobrefNo = :refNo and jctActionPlanSoftDelete=0 and plan.jctAsQuestionDesc = :mainQuestion"),
    @NamedQuery(name = "getAnswerListByJrNo",
    			query = "select plan.jctAsAnswarDesc from JctActionPlan plan where jctJobrefNo = :refNo " +
    					"and jctActionPlanSoftDelete=0 and plan.jctAsQuestionDesc = :mainQuestion " +
    					"and plan.jctAsQuestionSubDesc = :subQuestion"),
    @NamedQuery(name="getFrozenActionPlanCount", 
    			query="SELECT COUNT(jctJobrefNo) FROM JctActionPlan plan WHERE plan.jctJobrefNo = :jrNo " +
    					"AND plan.jctAsCreatedBy = :user AND plan.jctUser.jctUserId = :userId AND plan.jctActionPlanSoftDelete = 0"),
    @NamedQuery(name="getActionPlanDummyData", 
    			query="from JctActionPlan where jctJobrefNo = :jrNo"),
    @NamedQuery(name="deleteAllTempActionPlanData", 
    			query="DELETE FROM JctActionPlan plan WHERE plan.jctJobrefNo = :jrNo")
                	
})
@Entity
@Table(name="jct_action_plan")
public class JctActionPlan implements Serializable {
	private static final long serialVersionUID = 1L;

	@Id
	@GeneratedValue(strategy=GenerationType.IDENTITY)
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
	
	/** THIS IS FOR PUBLIC VERSION **/
	@ManyToOne
	@JoinColumn(name="jct_action_plan_user_id")
	private JctUser jctUser;
	/********************************/

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
	/** THIS IS FOR PUBLIC VERSION **/	
	public JctUser getJctUser() {
		return jctUser;
	}
	public void setJctUser(JctUser jctUser) {
		this.jctUser = jctUser;
	}
	/********************************/

}
