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
 * <p><b>Class name:</b> JctBeforeSketchHeader.java</p>
 * <p><b>Author:</b> InterraIt</p>
 * <p><b>Description:</b>  This class basically object reference of jct_before_sketch_header table </p>
 * <p><b>Copyrights:</b> 	All rights reserved by Interra IT and should only be used for its internal application development.</p>
 * <p><b>Revision History:</b>
 * 	<li>InterraIt - 31/Jan/2014 - Implement comments, introduce Named query </li>
 * </p>
 */

@NamedQueries({
	@NamedQuery(name = "fetchAllBeforeSketchHeader",
				query = "from JctBeforeSketchHeader where jctBsJobrefNo = :refNo and softDelete=0 and jctStatus = 4"),
	@NamedQuery(name = "fetchAllBeforeSketchHeaderEdtd",
				query = "from JctBeforeSketchHeader where jctBsJobrefNo = :refNo and softDelete=0 and jctStatus = 5"),
	@NamedQuery(name = "fetchBKJsonByRefNo",
				query = "Select hdr.jctBsJson from JctBeforeSketchHeader hdr where jctBsJobrefNo = :refNo " +
						"and hdr.softDelete=0 and hdr.jctStatus = 4"),
	@NamedQuery(name = "fetchBKJsonByRefNoEdited",
				query = "Select hdr.jctBsJson from JctBeforeSketchHeader hdr where jctBsJobrefNo = :refNo " +
						"and hdr.softDelete=0 and hdr.jctStatus = 5"),
	@NamedQuery(name = "fetchBKInitialJsonByRefNo",
				query = "Select hdr.initialJsonVal from JctBeforeSketchHeader hdr where jctBsJobrefNo = :refNo  " +
						"and hdr.softDelete=0 and hdr.jctStatus = 4"),
	@NamedQuery(name = "fetchBKInitialJsonByRefNoEdited",
				query = "Select hdr.initialJsonVal from JctBeforeSketchHeader hdr where jctBsJobrefNo = :refNo " +
						"and hdr.softDelete=0 and hdr.jctStatus = 5"),
	@NamedQuery(name = "getPrevBSDiagramByJRNO", 
				query = "from JctBeforeSketchHeader hdr where jctBsJobrefNo = :jctJobrefNo and " +
						"hdr.softDelete = 1 and hdr.jctStatus = 5"),
	@NamedQuery(name = "fetchAllBeforeSketchHeaderForPrevious", 
				query = "from JctBeforeSketchHeader hdr where jctBsJobrefNo = :refNo and hdr.softDelete=0 and hdr.jctStatus = 4"),
	@NamedQuery(name = "bsMaxPageSeq", 
				query = "from JctBeforeSketchHeader hdr where jctBsJobrefNo = :refNo and hdr.softDelete=1 and hdr.jctStatus = 5 " +
						"and hdr.jctBsInsertionSequence = (select max(hdr2.jctBsInsertionSequence) from JctBeforeSketchHeader hdr2 " +
						"where hdr2.jctBsInsertionSequence < :seqNos)"),
	@NamedQuery(name = "getBeforeSketchByEmailId",
				query = "from JctBeforeSketchHeader where softDelete=0 and jctStatus = 4 and jctBsCreatedBy = :createdBy"),
	@NamedQuery(name = "getDummyBeforeSketchData",
				query = "from JctBeforeSketchHeader where  jctBsJobrefNo = :jrNo"),
	@NamedQuery(name = "deleteDummyBSData", 
				query = "DELETE FROM JctBeforeSketchHeader header where header.jctBsJobrefNo = :jrNo")
})
@Entity
@Table(name="jct_before_sketch_header")
public class JctBeforeSketchHeader implements Serializable {
	private static final long serialVersionUID = 1L;

	@Id
	@GeneratedValue(strategy = GenerationType.IDENTITY)
	@Column(name="jct_bs_header_id")
	private int jctBsHeaderId;

	@Column(name="jct_bs_created_by")
	private String jctBsCreatedBy;

	@Column(name="jct_bs_jobref_no")
	private String jctBsJobrefNo;

	@Lob
	@Column(name="jct_bs_json")
	private String jctBsJson;

	@Column(name="jct_bs_lastmodified_by")
	private String jctBsLastmodifiedBy;

	@Temporal(TemporalType.TIMESTAMP)
	@Column(name="jct_bs_lastmodified_ts")
	private Date jctBsLastmodifiedTs;

	@Column(name="jct_bs_time_spent")
	private int jctBsTimeSpent;

	@Temporal(TemporalType.TIMESTAMP)
	@Column(name="jct_ds_created_ts")
	private Date jctDsCreatedTs;
	
	@Column(name="jct_bs_user_job_title")
	private String jctBsUserJobTitle;
	
	@Column(name="jct_bs_initial_json_value")
	private String initialJsonVal;
	
	@Column(name="jct_bs_snapshot_string")
	private String jctBsSnapshotString;
	
	@Column(name="jct_bs_soft_delete")
	private int softDelete;
	
	@Version
	private int version;
	
	@Lob
	@Column(name="jct_bs_snapshot")
	private byte[] jctBsSnapshot;

	//bi-directional many-to-one association to JctBeforeSketchDetail
	@OneToMany(mappedBy="jctBeforeSketchHeader", cascade=CascadeType.ALL)
	private List<JctBeforeSketchDetail> jctBeforeSketchDetails;

	//bi-directional many-to-one association to JctStatus
	@ManyToOne
	@JoinColumn(name="jct_bs_status_id")
	private JctStatus jctStatus;
	
	/** THIS IS FOR PUBLIC VERSION **/
	@ManyToOne
	@JoinColumn(name="jct_bs_user_id")
	private JctUser jctUser;
	/********************************/
	
	@Column(name="jct_bs_insertion_sequence")
	private int jctBsInsertionSequence;

	public int getJctBsInsertionSequence() {
		return jctBsInsertionSequence;
	}

	public void setJctBsInsertionSequence(int jctBsInsertionSequence) {
		this.jctBsInsertionSequence = jctBsInsertionSequence;
	}

	public JctBeforeSketchHeader() {
	}

	public int getJctBsHeaderId() {
		return this.jctBsHeaderId;
	}

	public void setJctBsHeaderId(int jctBsHeaderId) {
		this.jctBsHeaderId = jctBsHeaderId;
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

	public String getJctBsJson() {
		return this.jctBsJson;
	}

	public void setJctBsJson(String jctBsJson) {
		this.jctBsJson = jctBsJson;
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

	public int getJctBsTimeSpent() {
		return this.jctBsTimeSpent;
	}

	public void setJctBsTimeSpent(int jctBsTimeSpent) {
		this.jctBsTimeSpent = jctBsTimeSpent;
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

	public List<JctBeforeSketchDetail> getJctBeforeSketchDetails() {
		return this.jctBeforeSketchDetails;
	}

	public void setJctBeforeSketchDetails(List<JctBeforeSketchDetail> jctBeforeSketchDetails) {
		this.jctBeforeSketchDetails = jctBeforeSketchDetails;
	}

	public JctBeforeSketchDetail addJctBeforeSketchDetail(JctBeforeSketchDetail jctBeforeSketchDetail) {
		getJctBeforeSketchDetails().add(jctBeforeSketchDetail);
		jctBeforeSketchDetail.setJctBeforeSketchHeader(this);

		return jctBeforeSketchDetail;
	}

	public JctBeforeSketchDetail removeJctBeforeSketchDetail(JctBeforeSketchDetail jctBeforeSketchDetail) {
		getJctBeforeSketchDetails().remove(jctBeforeSketchDetail);
		jctBeforeSketchDetail.setJctBeforeSketchHeader(null);

		return jctBeforeSketchDetail;
	}

	public JctStatus getJctStatus() {
		return this.jctStatus;
	}

	public void setJctStatus(JctStatus jctStatus) {
		this.jctStatus = jctStatus;
	}
	
	public String getJctBsUserJobTitle() {
		return jctBsUserJobTitle;
	}

	public void setJctBsUserJobTitle(String jctBsUserJobTitle) {
		this.jctBsUserJobTitle = jctBsUserJobTitle;
	}
	
	public byte[] getJctBsSnapshot() {
		return jctBsSnapshot;
	}

	public void setJctBsSnapshot(byte[] jctBsSnapshot) {
		this.jctBsSnapshot = jctBsSnapshot;
	}

	public String getInitialJsonVal() {
		return initialJsonVal;
	}

	public void setInitialJsonVal(String initialJsonVal) {
		this.initialJsonVal = initialJsonVal;
	}
	
	public String getJctBsSnapshotString() {
		return jctBsSnapshotString;
	}

	public void setJctBsSnapshotString(String jctBsSnapshotString) {
		this.jctBsSnapshotString = jctBsSnapshotString;
	}
	
	public int getSoftDelete() {
		return softDelete;
	}

	public void setSoftDelete(int softDelete) {
		this.softDelete = softDelete;
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