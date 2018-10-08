package com.vmware.jct.model;

import java.io.Serializable;

import javax.persistence.*;

import java.util.Date;


/**
 * 
 * <p><b>Class name:</b> JctAfterSketchFinalpageDetailTemp.java</p>
 * <p><b>Author:</b> InterraIt</p>
 * <p><b>Description:</b>  This class basically object reference of jct_after_sketch_finalpage_details table </p>
 * <p><b>Copyrights:</b> 	All rights reserved by Interra IT and should only be used for its internal application development.</p>
 * <p><b>Revision History:</b>
 * 	<li>InterraIt - 17/Jun/2014 - Introduced the class.</li>
 * </p>
 */
@NamedQueries({
	@NamedQuery(name="findASFinalPageChildrenTemp",
				query="from JctAfterSketchFinalpageDetailTemp where jctAsJobrefNo = :refNo and jctAsSoftDelete = 0"),
	@NamedQuery(name="deleteAllASPageTwoTempData", 
				query="DELETE FROM JctAfterSketchFinalpageDetailTemp dtl where dtl.jctAsJobrefNo = :jrNo")
})
@Entity
@Table(name="jct_after_sketch_finalpage_details_temp")
public class JctAfterSketchFinalpageDetailTemp implements Serializable {
	private static final long serialVersionUID = 1L;

	@Id
	@GeneratedValue(strategy=GenerationType.IDENTITY)
	@Column(name="jct_as_finalpage_details_id")
	private int jctAsFinalpageDetailsId;

	@Column(name="jct_as_created_by")
	private String jctAsCreatedBy;

	@Temporal(TemporalType.TIMESTAMP)
	@Column(name="jct_as_created_ts")
	private Date jctAsCreatedTs;

	@Column(name="jct_as_element_code")
	private String jctAsElementCode;

	@Column(name="jct_as_jobref_no")
	private String jctAsJobrefNo;

	@Column(name="jct_as_lastmodified_by")
	private String jctAsLastmodifiedBy;

	@Temporal(TemporalType.TIMESTAMP)
	@Column(name="jct_as_lastmodified_ts")
	private Date jctAsLastmodifiedTs;

	@Column(name="jct_as_position")
	private String jctAsPosition;

	@Column(name="jct_as_role_desc")
	private String jctAsRoleDesc;

	@Column(name="jct_as_soft_delete")
	private int jctAsSoftDelete;

	@Column(name="jct_as_status_id")
	private int jctAsStatusId;

	@Column(name="jct_as_task_desc")
	private String jctAsTaskDesc;

	@Column(name="jct_as_task_energy")
	private int jctAsTaskEnergy;

	@Column(name="jct_as_task_id")
	private int jctAsTaskId;
	
	@Column(name="jct_as_additional_desc")
	private String jctAsAdditionalDesc;

	private int version;

	//bi-directional many-to-one association to JctJobAttribute
	@Column(name="jct_as_element_desc")
	private String jctAsElementDesc;

	//bi-directional many-to-one association to JctAfterSketchHeader
	@ManyToOne
	@JoinColumn(name="jct_as_header_id")
	private JctAfterSketchHeaderTemp jctAfterSketchHeaderTemp;

	public JctAfterSketchFinalpageDetailTemp() {
	}

	public int getJctAsFinalpageDetailsId() {
		return this.jctAsFinalpageDetailsId;
	}

	public void setJctAsFinalpageDetailsId(int jctAsFinalpageDetailsId) {
		this.jctAsFinalpageDetailsId = jctAsFinalpageDetailsId;
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

	public String getJctAsElementCode() {
		return this.jctAsElementCode;
	}

	public void setJctAsElementCode(String jctAsElementCode) {
		this.jctAsElementCode = jctAsElementCode;
	}

	public String getJctAsJobrefNo() {
		return this.jctAsJobrefNo;
	}

	public void setJctAsJobrefNo(String jctAsJobrefNo) {
		this.jctAsJobrefNo = jctAsJobrefNo;
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

	public String getJctAsPosition() {
		return this.jctAsPosition;
	}

	public void setJctAsPosition(String jctAsPosition) {
		this.jctAsPosition = jctAsPosition;
	}

	public String getJctAsRoleDesc() {
		return this.jctAsRoleDesc;
	}

	public void setJctAsRoleDesc(String jctAsRoleDesc) {
		this.jctAsRoleDesc = jctAsRoleDesc;
	}

	public int getJctAsSoftDelete() {
		return this.jctAsSoftDelete;
	}

	public void setJctAsSoftDelete(int jctAsSoftDelete) {
		this.jctAsSoftDelete = jctAsSoftDelete;
	}

	public int getJctAsStatusId() {
		return this.jctAsStatusId;
	}

	public void setJctAsStatusId(int jctAsStatusId) {
		this.jctAsStatusId = jctAsStatusId;
	}

	public String getJctAsTaskDesc() {
		return this.jctAsTaskDesc;
	}

	public void setJctAsTaskDesc(String jctAsTaskDesc) {
		this.jctAsTaskDesc = jctAsTaskDesc;
	}

	public int getJctAsTaskEnergy() {
		return this.jctAsTaskEnergy;
	}

	public void setJctAsTaskEnergy(int jctAsTaskEnergy) {
		this.jctAsTaskEnergy = jctAsTaskEnergy;
	}

	public int getJctAsTaskId() {
		return this.jctAsTaskId;
	}

	public void setJctAsTaskId(int jctAsTaskId) {
		this.jctAsTaskId = jctAsTaskId;
	}

	public int getVersion() {
		return this.version;
	}

	public void setVersion(int version) {
		this.version = version;
	}

	
	
	public JctAfterSketchHeaderTemp getJctAfterSketchHeaderTemp() {
		return jctAfterSketchHeaderTemp;
	}

	public void setJctAfterSketchHeaderTemp(
			JctAfterSketchHeaderTemp jctAfterSketchHeaderTemp) {
		this.jctAfterSketchHeaderTemp = jctAfterSketchHeaderTemp;
	}

	public String getJctAsAdditionalDesc() {
		return jctAsAdditionalDesc;
	}

	public void setJctAsAdditionalDesc(String jctAsAdditionalDesc) {
		this.jctAsAdditionalDesc = jctAsAdditionalDesc;
	}

	public String getJctAsElementDesc() {
		return jctAsElementDesc;
	}

	public void setJctAsElementDesc(String jctAsElementDesc) {
		this.jctAsElementDesc = jctAsElementDesc;
	}
}