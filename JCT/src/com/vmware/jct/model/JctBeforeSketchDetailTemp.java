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


/**
 * 
 * <p><b>Class name:</b> JctBeforeSketchDetailTemp.java</p>
 * <p><b>Author:</b> InterraIt</p>
 * <p><b>Description:</b>  This class basically object reference of jct_before_sketch_details_temp table </p>
 * <p><b>Copyrights:</b> 	All rights reserved by Interra IT and should only be used for its internal application development.</p>
 * <p><b>Revision History:</b>
 * 	<li>InterraIt - 16/June/2014 - Introduced the class. </li>
 * </p>
 */
@NamedQueries({
	@NamedQuery(name = "deleteDummyBSTempChildData", 
				query = "DELETE FROM JctBeforeSketchDetailTemp dtl where dtl.jctBsJobrefNo = :jrNo")
})

@Entity
@Table(name="jct_before_sketch_details_temp")
public class JctBeforeSketchDetailTemp implements Serializable {
	private static final long serialVersionUID = 1L;

	@Id
	@GeneratedValue(strategy = GenerationType.IDENTITY)
	@Column(name="jct_bs_detail_id")
	private int jctBsDetailId;
	

	@Column(name="jct_bs_task_id")
	private int jctBsTaskId;

	@Column(name="jct_bs_created_by")
	private String jctBsCreatedBy;

	@Column(name="jct_bs_energy")
	private int jctBsEnergy;

	@Column(name="jct_bs_jobref_no")
	private String jctBsJobrefNo;

	@Column(name="jct_bs_lastmodified_by")
	private String jctBsLastmodifiedBy;

	@Temporal(TemporalType.TIMESTAMP)
	@Column(name="jct_bs_lastmodified_ts")
	private Date jctBsLastmodifiedTs;

	@Column(name="jct_bs_position")
	private String jctBsPosition;

	@Column(name="jct_bs_task_desc")
	private String jctBsTaskDesc;

	@Column(name="jct_bs_time")
	private int jctBsTime;

	@Temporal(TemporalType.TIMESTAMP)
	@Column(name="jct_ds_created_ts")
	private Date jctDsCreatedTs;

	private int version;
	
	@Column(name="jct_bs_soft_delete")
	private int softDelete;

	//bi-directional many-to-one association to JctBeforeSketchHeader
	@ManyToOne
	@JoinColumn(name="jct_bs_header_id")
	private JctBeforeSketchHeaderTemp jctBeforeSketchHeader;

	//bi-directional many-to-one association to JctStatus
	@ManyToOne
	@JoinColumn(name="jct_bs_status_id")
	private JctStatus jctStatus;

	public int getJctBsDetailId() {
		return jctBsDetailId;
	}

	public void setJctBsDetailId(int jctBsDetailId) {
		this.jctBsDetailId = jctBsDetailId;
	}
	
	public JctBeforeSketchDetailTemp() {
	}

	public int getJctBsTaskId() {
		return this.jctBsTaskId;
	}

	public void setJctBsTaskId(int jctBsTaskId) {
		this.jctBsTaskId = jctBsTaskId;
	}

	public String getJctBsCreatedBy() {
		return this.jctBsCreatedBy;
	}

	public void setJctBsCreatedBy(String jctBsCreatedBy) {
		this.jctBsCreatedBy = jctBsCreatedBy;
	}

	public int getJctBsEnergy() {
		return this.jctBsEnergy;
	}

	public void setJctBsEnergy(int jctBsEnergy) {
		this.jctBsEnergy = jctBsEnergy;
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

	public String getJctBsPosition() {
		return this.jctBsPosition;
	}

	public void setJctBsPosition(String jctBsPosition) {
		this.jctBsPosition = jctBsPosition;
	}

	public String getJctBsTaskDesc() {
		return this.jctBsTaskDesc;
	}

	public void setJctBsTaskDesc(String jctBsTaskDesc) {
		this.jctBsTaskDesc = jctBsTaskDesc;
	}

	public int getJctBsTime() {
		return this.jctBsTime;
	}

	public void setJctBsTime(int jctBsTime) {
		this.jctBsTime = jctBsTime;
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

	public JctBeforeSketchHeaderTemp getJctBeforeSketchHeader() {
		return this.jctBeforeSketchHeader;
	}

	public void setJctBeforeSketchHeader(JctBeforeSketchHeaderTemp jctBeforeSketchHeader) {
		this.jctBeforeSketchHeader = jctBeforeSketchHeader;
	}

	public JctStatus getJctStatus() {
		return this.jctStatus;
	}

	public void setJctStatus(JctStatus jctStatus) {
		this.jctStatus = jctStatus;
	}
	
	public int getSoftDelete() {
		return softDelete;
	}

	public void setSoftDelete(int softDelete) {
		this.softDelete = softDelete;
	}

}