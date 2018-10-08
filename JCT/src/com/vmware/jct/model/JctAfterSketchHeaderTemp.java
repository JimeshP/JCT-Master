package com.vmware.jct.model;

import java.io.Serializable;
import java.util.Date;
import java.util.List;

import javax.persistence.CascadeType;
import javax.persistence.Column;
import javax.persistence.Entity;
import javax.persistence.GeneratedValue;
import javax.persistence.GenerationType;
import javax.persistence.Id;
import javax.persistence.JoinColumn;
import javax.persistence.Lob;
import javax.persistence.ManyToOne;
import javax.persistence.NamedQueries;
import javax.persistence.NamedQuery;
import javax.persistence.OneToMany;
import javax.persistence.Table;
import javax.persistence.Temporal;
import javax.persistence.TemporalType;
import javax.persistence.Version;

/**
 * 
 * <p><b>Class name:</b> JctAfterSketchHeaderTemp.java</p>
 * <p><b>Author:</b> InterraIt</p>
 * <p><b>Description:</b>  This class basically object reference of jct_after_sketch_header table </p>
 * <p><b>Copyrights:</b> 	All rights reserved by Interra IT and should only be used for its internal application development.</p>
 * <p><b>Revision History:</b>
 * 	<li>InterraIt - 17/Jun/2014 - Introduced the class. </li>
 * </p>
 */
@NamedQueries({
	@NamedQuery(name = "fetchAfterSketchHeaderTemp",
				query = "from JctAfterSketchHeaderTemp where jctAsJobrefNo = :refNo and jctAsSoftDelete=0"),
	@NamedQuery(name = "getAfterSketchDummyTempData",
				query = "from JctAfterSketchHeaderTemp where jctAsJobrefNo = :jrNo"),
	@NamedQuery(name = "deleteAllASTempData", 
				query = "DELETE FROM JctAfterSketchHeaderTemp tmp WHERE tmp.jctAsJobrefNo = :jrNo")
})

@Entity
@Table(name="jct_after_sketch_header_temp")
public class JctAfterSketchHeaderTemp implements Serializable {
	private static final long serialVersionUID = 1L;

	@Id
	@GeneratedValue(strategy=GenerationType.IDENTITY)
	@Column(name="jct_as_header_id")
	private int jctAsHeaderId;

	@Column(name="jct_as_created_by")
	private String jctAsCreatedBy;

	@Temporal(TemporalType.TIMESTAMP)
	@Column(name="jct_as_created_ts")
	private Date jctAsCreatedTs;

	@Lob
	@Column(name="jct_as_finalpage_initial_json_value")
	private String jctAsFinalpageInitialJsonValue;

	@Lob
	@Column(name="jct_as_finalpage_json")
	private String jctAsFinalpageJson;

	@Lob
	@Column(name="jct_as_finalpage_snapshot")
	private byte[] jctAsFinalpageSnapshot;

	@Lob
	@Column(name="jct_as_finalpage_snapshot_string")
	private String jctAsFinalpageSnapshotString;

	@Column(name="jct_as_finalpage_time_spent")
	private int jctAsFinalpageTimeSpent;

	@Column(name="jct_as_jobref_no")
	private String jctAsJobrefNo;

	@Column(name="jct_as_lastmodified_by")
	private String jctAsLastmodifiedBy;

	@Temporal(TemporalType.TIMESTAMP)
	@Column(name="jct_as_lastmodified_ts")
	private Date jctAsLastmodifiedTs;

	@Lob
	@Column(name="jct_as_pagene_json")
	private String jctAsPageneJson;

	@Lob
	@Column(name="jct_as_pageone_checked_value")
	private String jctAsPageoneCheckedValue;

	@Lob
	@Column(name="jct_as_pageone_snapshot")
	private byte[] jctAsPageoneSnapshot;

	@Lob
	@Column(name="jct_as_pageone_snapshot_string")
	private String jctAsPageoneSnapshotString;

	@Column(name="jct_as_pageone_time_spent")
	private int jctAsPageoneTimeSpent;

	@Column(name="jct_as_status_id")
	private int jctAsStatusId;

	@Column(name="jct_as_user_job_title")
	private String jctAsUserJobTitle;
	
	@Column(name="jct_final_page_div_height")
	private String jctFinalPageDivHeight;
	
	@Column(name="jct_final_page_div_width")
	private String jctFinalPageDivWidth;
	
	@Column(name="jct_checked_strength")
	private String jctCheckedStrength;
	
	@Column(name="jct_checked_passion")
	private String jctCheckedPassion;
	
	@Column(name="jct_checked_value")
	private String jctCheckedValue;
	
	@Column(name="jct_as_soft_delete")
	private int jctAsSoftDelete;
	
	@Column(name="jct_as_total_count")
	private int jctAsTotalCount;
	
	@Column(name="jct_as_insertion_sequence")
	private int jctAsInsertionSequence;
	
	public int getJctAsInsertionSequence() {
		return jctAsInsertionSequence;
	}

	public void setJctAsInsertionSequence(int jctAsInsertionSequence) {
		this.jctAsInsertionSequence = jctAsInsertionSequence;
	}

	public int getJctAsTotalCount() {
		return jctAsTotalCount;
	}

	public void setJctAsTotalCount(int jctAsTotalCount) {
		this.jctAsTotalCount = jctAsTotalCount;
	}

	public int getJctAsSoftDelete() {
		return jctAsSoftDelete;
	}

	public void setJctAsSoftDelete(int jctAsSoftDelete) {
		this.jctAsSoftDelete = jctAsSoftDelete;
	}

	public String getJctCheckedStrength() {
		return jctCheckedStrength;
	}

	public void setJctCheckedStrength(String jctCheckedStrength) {
		this.jctCheckedStrength = jctCheckedStrength;
	}

	public String getJctCheckedPassion() {
		return jctCheckedPassion;
	}

	public void setJctCheckedPassion(String jctCheckedPassion) {
		this.jctCheckedPassion = jctCheckedPassion;
	}

	public String getJctCheckedValue() {
		return jctCheckedValue;
	}

	public void setJctCheckedValue(String jctCheckedValue) {
		this.jctCheckedValue = jctCheckedValue;
	}

	public String getJctFinalPageDivHeight() {
		return jctFinalPageDivHeight;
	}

	public void setJctFinalPageDivHeight(String jctFinalPageDivHeight) {
		this.jctFinalPageDivHeight = jctFinalPageDivHeight;
	}

	public String getJctFinalPageDivWidth() {
		return jctFinalPageDivWidth;
	}

	public void setJctFinalPageDivWidth(String jctFinalPageDivWidth) {
		this.jctFinalPageDivWidth = jctFinalPageDivWidth;
	}

	@Version
	private int version;

	//bi-directional many-to-one association to JctAfterSketchFinalpageDetail
	@OneToMany(mappedBy="jctAfterSketchHeaderTemp", cascade=CascadeType.ALL)
	private List<JctAfterSketchFinalpageDetailTemp> jctAfterSketchFinalpageDetails;

	//bi-directional many-to-one association to JctAfterSketchPageoneDetail
	@OneToMany(mappedBy="jctAfterSketchHeaderTemp", cascade=CascadeType.ALL)
	private List<JctAfterSketchPageoneDetailTemp> jctAfterSketchPageoneDetails;

	/** THIS IS FOR PUBLIC VERSION **/
	@ManyToOne
	@JoinColumn(name="jct_as_temp_user_id")
	private JctUser jctUser;
	/********************************/
	
	public JctAfterSketchHeaderTemp() {
	}

	public int getJctAsHeaderId() {
		return this.jctAsHeaderId;
	}

	public void setJctAsHeaderId(int jctAsHeaderId) {
		this.jctAsHeaderId = jctAsHeaderId;
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

	public String getJctAsFinalpageInitialJsonValue() {
		return this.jctAsFinalpageInitialJsonValue;
	}

	public void setJctAsFinalpageInitialJsonValue(String jctAsFinalpageInitialJsonValue) {
		this.jctAsFinalpageInitialJsonValue = jctAsFinalpageInitialJsonValue;
	}

	public String getJctAsFinalpageJson() {
		return this.jctAsFinalpageJson;
	}

	public void setJctAsFinalpageJson(String jctAsFinalpageJson) {
		this.jctAsFinalpageJson = jctAsFinalpageJson;
	}

	public byte[] getJctAsFinalpageSnapshot() {
		return this.jctAsFinalpageSnapshot;
	}

	public void setJctAsFinalpageSnapshot(byte[] jctAsFinalpageSnapshot) {
		this.jctAsFinalpageSnapshot = jctAsFinalpageSnapshot;
	}

	public String getJctAsFinalpageSnapshotString() {
		return this.jctAsFinalpageSnapshotString;
	}

	public void setJctAsFinalpageSnapshotString(String jctAsFinalpageSnapshotString) {
		this.jctAsFinalpageSnapshotString = jctAsFinalpageSnapshotString;
	}

	public int getJctAsFinalpageTimeSpent() {
		return this.jctAsFinalpageTimeSpent;
	}

	public void setJctAsFinalpageTimeSpent(int jctAsFinalpageTimeSpent) {
		this.jctAsFinalpageTimeSpent = jctAsFinalpageTimeSpent;
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

	public String getJctAsPageneJson() {
		return this.jctAsPageneJson;
	}

	public void setJctAsPageneJson(String jctAsPageneJson) {
		this.jctAsPageneJson = jctAsPageneJson;
	}

	public byte[] getJctAsPageoneSnapshot() {
		return this.jctAsPageoneSnapshot;
	}

	public void setJctAsPageoneSnapshot(byte[] jctAsPageoneSnapshot) {
		this.jctAsPageoneSnapshot = jctAsPageoneSnapshot;
	}

	public String getJctAsPageoneSnapshotString() {
		return this.jctAsPageoneSnapshotString;
	}

	public void setJctAsPageoneSnapshotString(String jctAsPageoneSnapshotString) {
		this.jctAsPageoneSnapshotString = jctAsPageoneSnapshotString;
	}

	public int getJctAsPageoneTimeSpent() {
		return this.jctAsPageoneTimeSpent;
	}

	public void setJctAsPageoneTimeSpent(int jctAsPageoneTimeSpent) {
		this.jctAsPageoneTimeSpent = jctAsPageoneTimeSpent;
	}

	public int getJctAsStatusId() {
		return this.jctAsStatusId;
	}

	public void setJctAsStatusId(int jctAsStatusId) {
		this.jctAsStatusId = jctAsStatusId;
	}

	public String getJctAsUserJobTitle() {
		return this.jctAsUserJobTitle;
	}

	public void setJctAsUserJobTitle(String jctAsUserJobTitle) {
		this.jctAsUserJobTitle = jctAsUserJobTitle;
	}

	public int getVersion() {
		return this.version;
	}

	public void setVersion(int version) {
		this.version = version;
	}

	public List<JctAfterSketchFinalpageDetailTemp> getJctAfterSketchFinalpageDetails() {
		return this.jctAfterSketchFinalpageDetails;
	}

	public void setJctAfterSketchFinalpageDetails(List<JctAfterSketchFinalpageDetailTemp> jctAfterSketchFinalpageDetails) {
		this.jctAfterSketchFinalpageDetails = jctAfterSketchFinalpageDetails;
	}

	public JctAfterSketchFinalpageDetailTemp addJctAfterSketchFinalpageDetail(JctAfterSketchFinalpageDetailTemp jctAfterSketchFinalpageDetail) {
		getJctAfterSketchFinalpageDetails().add(jctAfterSketchFinalpageDetail);
		jctAfterSketchFinalpageDetail.setJctAfterSketchHeaderTemp(this);

		return jctAfterSketchFinalpageDetail;
	}

	public JctAfterSketchFinalpageDetail removeJctAfterSketchFinalpageDetail(JctAfterSketchFinalpageDetail jctAfterSketchFinalpageDetail) {
		getJctAfterSketchFinalpageDetails().remove(jctAfterSketchFinalpageDetail);
		jctAfterSketchFinalpageDetail.setJctAfterSketchHeader(null);

		return jctAfterSketchFinalpageDetail;
	}

	public List<JctAfterSketchPageoneDetailTemp> getJctAfterSketchPageoneDetails() {
		return this.jctAfterSketchPageoneDetails;
	}

	public void setJctAfterSketchPageoneDetails(List<JctAfterSketchPageoneDetailTemp> jctAfterSketchPageoneDetails) {
		this.jctAfterSketchPageoneDetails = jctAfterSketchPageoneDetails;
	}

	public JctAfterSketchPageoneDetailTemp addJctAfterSketchPageoneDetail(JctAfterSketchPageoneDetailTemp jctAfterSketchPageoneDetail) {
		getJctAfterSketchPageoneDetails().add(jctAfterSketchPageoneDetail);
		jctAfterSketchPageoneDetail.setJctAfterSketchHeaderTemp(this);

		return jctAfterSketchPageoneDetail;
	}

	public JctAfterSketchPageoneDetail removeJctAfterSketchPageoneDetail(JctAfterSketchPageoneDetail jctAfterSketchPageoneDetail) {
		getJctAfterSketchPageoneDetails().remove(jctAfterSketchPageoneDetail);
		jctAfterSketchPageoneDetail.setJctAfterSketchHeader(null);

		return jctAfterSketchPageoneDetail;
	}

	public String getJctAsPageoneCheckedValue() {
		return jctAsPageoneCheckedValue;
	}

	public void setJctAsPageoneCheckedValue(String jctAsPageoneCheckedValue) {
		this.jctAsPageoneCheckedValue = jctAsPageoneCheckedValue;
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