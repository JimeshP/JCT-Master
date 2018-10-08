package com.vmware.jct.model;

import java.io.Serializable;
import javax.persistence.*;
import java.util.Date;
import java.util.List;


/**
 * 
 * <p><b>Class name:</b> JctStatus.java</p>
 * <p><b>Author:</b> InterraIt</p>
 * <p><b>Description:</b>  This class basically object reference of jct_status table. This is a master table for status </p>
 * <p><b>Copyrights:</b> 	All rights reserved by Interra IT and should only be used for its internal application development.</p>
 * <p><b>Revision History:</b>
 * 	<li>InterraIt - 31/Jan/2014 - Implement comments, introduce Named query </li>
 * </p>
 */
@Entity
@Table(name="jct_status")
public class JctStatus implements Serializable {
	private static final long serialVersionUID = 1L;

	@Id
	@Column(name="jct_status_id")
	private int jctStatusId;

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

	@Column(name="jct_status_desc")
	private String jctStatusDesc;

	private int version;

	//bi-directional many-to-one association to JctActionPlan
	@OneToMany(mappedBy="jctStatus")
	private List<JctActionPlan> jctActionPlans;

	//bi-directional many-to-one association to JctAdditionalInfo
	@OneToMany(mappedBy="jctStatus")
	private List<JctAdditionalInfo> jctAdditionalInfos;

	//bi-directional many-to-one association to JctAfterSketch
	@OneToMany(mappedBy="jctStatus")
	private List<JctAfterSketch> jctAfterSketches;

	//bi-directional many-to-one association to JctBeforeSketchDetail
	@OneToMany(mappedBy="jctStatus")
	private List<JctBeforeSketchDetail> jctBeforeSketchDetails;

	//bi-directional many-to-one association to JctBeforeSketchHeader
	@OneToMany(mappedBy="jctStatus")
	private List<JctBeforeSketchHeader> jctBeforeSketchHeaders;

	//bi-directional many-to-one association to JctBeforeSketchQuestion
	@OneToMany(mappedBy="jctStatus")
	private List<JctBeforeSketchQuestion> jctBeforeSketchQuestions;

	//bi-directional many-to-one association to JctJobAttribute
	/*@OneToMany(mappedBy="jctStatus")
	private List<JctJobAttribute> jctJobAttributes;*/

	//bi-directional many-to-one association to JctUserLoginInfo
	@OneToMany(mappedBy="jctStatus")
	private List<JctUserLoginInfo> jctUserLoginInfos;

	public JctStatus() {
	}

	public int getJctStatusId() {
		return this.jctStatusId;
	}

	public void setJctStatusId(int jctStatusId) {
		this.jctStatusId = jctStatusId;
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

	public String getJctStatusDesc() {
		return this.jctStatusDesc;
	}

	public void setJctStatusDesc(String jctStatusDesc) {
		this.jctStatusDesc = jctStatusDesc;
	}

	public int getVersion() {
		return this.version;
	}

	public void setVersion(int version) {
		this.version = version;
	}

	public List<JctActionPlan> getJctActionPlans() {
		return this.jctActionPlans;
	}

	public void setJctActionPlans(List<JctActionPlan> jctActionPlans) {
		this.jctActionPlans = jctActionPlans;
	}

	public JctActionPlan addJctActionPlan(JctActionPlan jctActionPlan) {
		getJctActionPlans().add(jctActionPlan);
		jctActionPlan.setJctStatus(this);

		return jctActionPlan;
	}

	public JctActionPlan removeJctActionPlan(JctActionPlan jctActionPlan) {
		getJctActionPlans().remove(jctActionPlan);
		jctActionPlan.setJctStatus(null);

		return jctActionPlan;
	}

	public List<JctAdditionalInfo> getJctAdditionalInfos() {
		return this.jctAdditionalInfos;
	}

	public void setJctAdditionalInfos(List<JctAdditionalInfo> jctAdditionalInfos) {
		this.jctAdditionalInfos = jctAdditionalInfos;
	}

	public JctAdditionalInfo addJctAdditionalInfo(JctAdditionalInfo jctAdditionalInfo) {
		getJctAdditionalInfos().add(jctAdditionalInfo);
		jctAdditionalInfo.setJctStatus(this);

		return jctAdditionalInfo;
	}

	public JctAdditionalInfo removeJctAdditionalInfo(JctAdditionalInfo jctAdditionalInfo) {
		getJctAdditionalInfos().remove(jctAdditionalInfo);
		jctAdditionalInfo.setJctStatus(null);

		return jctAdditionalInfo;
	}

	public List<JctAfterSketch> getJctAfterSketches() {
		return this.jctAfterSketches;
	}

	public void setJctAfterSketches(List<JctAfterSketch> jctAfterSketches) {
		this.jctAfterSketches = jctAfterSketches;
	}

	public JctAfterSketch addJctAfterSketch(JctAfterSketch jctAfterSketch) {
		getJctAfterSketches().add(jctAfterSketch);
		jctAfterSketch.setJctStatus(this);

		return jctAfterSketch;
	}

	public JctAfterSketch removeJctAfterSketch(JctAfterSketch jctAfterSketch) {
		getJctAfterSketches().remove(jctAfterSketch);
		jctAfterSketch.setJctStatus(null);

		return jctAfterSketch;
	}

	public List<JctBeforeSketchDetail> getJctBeforeSketchDetails() {
		return this.jctBeforeSketchDetails;
	}

	public void setJctBeforeSketchDetails(List<JctBeforeSketchDetail> jctBeforeSketchDetails) {
		this.jctBeforeSketchDetails = jctBeforeSketchDetails;
	}

	public JctBeforeSketchDetail addJctBeforeSketchDetail(JctBeforeSketchDetail jctBeforeSketchDetail) {
		getJctBeforeSketchDetails().add(jctBeforeSketchDetail);
		jctBeforeSketchDetail.setJctStatus(this);

		return jctBeforeSketchDetail;
	}

	public JctBeforeSketchDetail removeJctBeforeSketchDetail(JctBeforeSketchDetail jctBeforeSketchDetail) {
		getJctBeforeSketchDetails().remove(jctBeforeSketchDetail);
		jctBeforeSketchDetail.setJctStatus(null);

		return jctBeforeSketchDetail;
	}

	public List<JctBeforeSketchHeader> getJctBeforeSketchHeaders() {
		return this.jctBeforeSketchHeaders;
	}

	public void setJctBeforeSketchHeaders(List<JctBeforeSketchHeader> jctBeforeSketchHeaders) {
		this.jctBeforeSketchHeaders = jctBeforeSketchHeaders;
	}

	public JctBeforeSketchHeader addJctBeforeSketchHeader(JctBeforeSketchHeader jctBeforeSketchHeader) {
		getJctBeforeSketchHeaders().add(jctBeforeSketchHeader);
		jctBeforeSketchHeader.setJctStatus(this);

		return jctBeforeSketchHeader;
	}

	public JctBeforeSketchHeader removeJctBeforeSketchHeader(JctBeforeSketchHeader jctBeforeSketchHeader) {
		getJctBeforeSketchHeaders().remove(jctBeforeSketchHeader);
		jctBeforeSketchHeader.setJctStatus(null);

		return jctBeforeSketchHeader;
	}

	public List<JctBeforeSketchQuestion> getJctBeforeSketchQuestions() {
		return this.jctBeforeSketchQuestions;
	}

	public void setJctBeforeSketchQuestions(List<JctBeforeSketchQuestion> jctBeforeSketchQuestions) {
		this.jctBeforeSketchQuestions = jctBeforeSketchQuestions;
	}

	public JctBeforeSketchQuestion addJctBeforeSketchQuestion(JctBeforeSketchQuestion jctBeforeSketchQuestion) {
		getJctBeforeSketchQuestions().add(jctBeforeSketchQuestion);
		jctBeforeSketchQuestion.setJctStatus(this);

		return jctBeforeSketchQuestion;
	}

	public JctBeforeSketchQuestion removeJctBeforeSketchQuestion(JctBeforeSketchQuestion jctBeforeSketchQuestion) {
		getJctBeforeSketchQuestions().remove(jctBeforeSketchQuestion);
		jctBeforeSketchQuestion.setJctStatus(null);

		return jctBeforeSketchQuestion;
	}

	/*public List<JctJobAttribute> getJctJobAttributes() {
		return this.jctJobAttributes;
	}

	public void setJctJobAttributes(List<JctJobAttribute> jctJobAttributes) {
		this.jctJobAttributes = jctJobAttributes;
	}

	public JctJobAttribute addJctJobAttribute(JctJobAttribute jctJobAttribute) {
		getJctJobAttributes().add(jctJobAttribute);
		jctJobAttribute.setJctStatus(this);

		return jctJobAttribute;
	}

	public JctJobAttribute removeJctJobAttribute(JctJobAttribute jctJobAttribute) {
		getJctJobAttributes().remove(jctJobAttribute);
		jctJobAttribute.setJctStatus(null);

		return jctJobAttribute;
	}*/

	public List<JctUserLoginInfo> getJctUserLoginInfos() {
		return this.jctUserLoginInfos;
	}

	public void setJctUserLoginInfos(List<JctUserLoginInfo> jctUserLoginInfos) {
		this.jctUserLoginInfos = jctUserLoginInfos;
	}

	public JctUserLoginInfo addJctUserLoginInfo(JctUserLoginInfo jctUserLoginInfo) {
		getJctUserLoginInfos().add(jctUserLoginInfo);
		jctUserLoginInfo.setJctStatus(this);

		return jctUserLoginInfo;
	}

	public JctUserLoginInfo removeJctUserLoginInfo(JctUserLoginInfo jctUserLoginInfo) {
		getJctUserLoginInfos().remove(jctUserLoginInfo);
		jctUserLoginInfo.setJctStatus(null);

		return jctUserLoginInfo;
	}

}