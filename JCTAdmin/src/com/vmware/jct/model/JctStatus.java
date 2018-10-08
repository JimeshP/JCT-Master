package com.vmware.jct.model;

import java.io.Serializable;
import javax.persistence.*;
import java.util.Date;
import java.util.List;


/**
 * 
 * <p><b>Class name:</b> JctStatus.java</p>
 * <p><b>Author:</b> InterraIT</p>
 * <p><b>Description:</b>  This class basically object reference of jct_status table. This is a master table for status </p>
 * <p><b>Copyrights:</b> 	All rights reserved by InterraIT and should only be used for its internal application development.</p>
 * <p><b>Revision History:</b>
 * 	<li>InterraIT - 31/Jan/2014 - Implement comments, introduce Named query </li>
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

}