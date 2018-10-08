package com.vmware.jct.model;

import java.io.Serializable;
import java.util.Date;

import javax.persistence.Column;
import javax.persistence.Entity;
import javax.persistence.Id;
import javax.persistence.JoinColumn;
import javax.persistence.Lob;
import javax.persistence.ManyToOne;
import javax.persistence.NamedQueries;
import javax.persistence.NamedQuery;
import javax.persistence.Table;
import javax.persistence.Temporal;
import javax.persistence.TemporalType;
import javax.persistence.Version;
/**
 * 
 * <p><b>Class name:</b> JctStatusSearch.java</p>
 * <p><b>Author:</b> InterraIT</p>
 * <p><b>Description:</b>  This class basically object reference of jct_status_search table </p>
 * <p><b>Copyrights:</b> 	All rights reserved by InterraIT and should only be used for its internal application development.</p>
 * <p><b>Revision History:</b>
 * 	<li>InterraIT - 31/Jan/2014 - Implement comments, introduce Named query </li>
 * </p>
 */
@NamedQueries({
	@NamedQuery(name = "searchBeforeSketchDiag",
				query = "Select status.jctBsSnapshotString from JctStatusSearch status " +
						"where status.jctUserLevels = :jobTitle and status.jctIsInactive=0"),
	@NamedQuery(name = "searchAfterSketchDiag",
				query = "Select status.jctAsSnapshotString from JctStatusSearch status where " +
						"status.jctUserLevels = :jobTitle and status.jctIsInactive=0"),
	@NamedQuery(name = "disableStatusSearch",
				query = "update JctStatusSearch set jctIsInactive=1 where jctJobrefNo = :jobRefNo"),
	@NamedQuery(name = "searchDiagramsToDiableByEmailId",
				query = "Select new com.vmware.jct.dao.dto.RemoveDiagramDTO(srch.jctStatusSearchId, srch.jctBsSnapshotString, " +
						"srch.jctAsSnapshotString, srch.jctJobrefNo) from JctStatusSearch srch where srch.jctSoftDelete = 0 and " +
						"srch.jctJobrefNo IN (select distinct (info.jctJobrefNo) from JctUserLoginInfo info where info.jctAsCreatedBy = :emailId)"),
	@NamedQuery(name = "softDeleteBsAndAsDiagramByAdmin",
				query = "UPDATE JctStatusSearch SET jctSoftDelete = :softDelStatus, jctLastmodifiedBy = :emailId, " +
						"jctLastmodifiedTs = :ts WHERE jctStatusSearchId = :rowId"),
	@NamedQuery(name = "searchAllDiagrams",
	query = "Select new com.vmware.jct.dao.dto.RemoveDiagramDTO(srch.jctStatusSearchId, srch.jctBsSnapshotString, " +
			"srch.jctAsSnapshotString, srch.jctJobrefNo, srch.jctCreatedTs) from JctStatusSearch srch where srch.jctSoftDelete = 0 and " +
			"srch.jctJobrefNo IN (select distinct (info.jctJobrefNo) from JctUserLoginInfo info) order by srch.jctCreatedTs desc"),
	@NamedQuery(name = "fetchAllDiagramCount",
	query = "Select COUNT(srch.jctStatusSearchId) from JctStatusSearch srch where srch.jctSoftDelete = 0 and " +
			"srch.jctJobrefNo IN (select distinct (info.jctJobrefNo) from JctUserLoginInfo info)")
})
@Entity
@Table(name="jct_status_search")
public class JctStatusSearch implements Serializable{
	
	private static final long serialVersionUID = 1L;
	
	@Id
	@Column(name="jct_status_search_id")
	private int jctStatusSearchId;
	
	@Column(name="jct_jobref_no")
	private String jctJobrefNo;
	
	@Column(name="jct_status_search_flg")
	private String jctStatusSearchFlg;
	
	@Lob
	@Column(name="jct_as_snapshot")
	private byte[] jctAsSnapshot;
	
	@Lob
	@Column(name="jct_bs_snapshot")
	private byte[] jctBsSnapshot;
	
	@Column(name="jct_bs_snapshot_string")
	private String jctBsSnapshotString;
	
	@Column(name="jct_as_snapshot_string")
	private String jctAsSnapshotString;
	
	@ManyToOne
	@JoinColumn(name="jct_status_id")
	private JctStatus jctStatusId;
	
	@Column(name="jct_soft_delete")
	private int jctSoftDelete;
	
	@Version
	private int version;
	
	@Temporal(TemporalType.TIMESTAMP)
	@Column(name="jct_created_ts")
	private Date jctCreatedTs;
	
	@Temporal(TemporalType.TIMESTAMP)
	@Column(name="jct_lastmodified_ts")
	private Date jctLastmodifiedTs;

	@Column(name="jct_lastmodified_by")
	private String jctLastmodifiedBy;
	
	@Column(name="jct_created_by")
	private String jctCreatedBy;
	
	@Column(name="jct_user_levels")
	private String jctUserLevels;
	
	@Column(name="jct_is_inactive")
	private int jctIsInactive;

	public int getJctIsInactive() {
		return jctIsInactive;
	}

	public void setJctIsInactive(int jctIsInactive) {
		this.jctIsInactive = jctIsInactive;
	}

	public String getJctUserLevels() {
		return jctUserLevels;
	}

	public void setJctUserLevels(String jctUserLevels) {
		this.jctUserLevels = jctUserLevels;
	}

	public int getJctStatusSearchId() {
		return jctStatusSearchId;
	}

	public void setJctStatusSearchId(int jctStatusSearchId) {
		this.jctStatusSearchId = jctStatusSearchId;
	}

	public String getJctJobrefNo() {
		return jctJobrefNo;
	}

	public void setJctJobrefNo(String jctJobrefNo) {
		this.jctJobrefNo = jctJobrefNo;
	}

	public String getJctStatusSearchFlg() {
		return jctStatusSearchFlg;
	}

	public void setJctStatusSearchFlg(String jctStatusSearchFlg) {
		this.jctStatusSearchFlg = jctStatusSearchFlg;
	}

	public byte[] getJctAsSnapshot() {
		return jctAsSnapshot;
	}

	public void setJctAsSnapshot(byte[] jctAsSnapshot) {
		this.jctAsSnapshot = jctAsSnapshot;
	}

	public byte[] getJctBsSnapshot() {
		return jctBsSnapshot;
	}

	public void setJctBsSnapshot(byte[] jctBsSnapshot) {
		this.jctBsSnapshot = jctBsSnapshot;
	}

	public String getJctBsSnapshotString() {
		return jctBsSnapshotString;
	}

	public void setJctBsSnapshotString(String jctBsSnapshotString) {
		this.jctBsSnapshotString = jctBsSnapshotString;
	}

	public String getJctAsSnapshotString() {
		return jctAsSnapshotString;
	}

	public void setJctAsSnapshotString(String jctAsSnapshotString) {
		this.jctAsSnapshotString = jctAsSnapshotString;
	}

	public JctStatus getJctStatusId() {
		return jctStatusId;
	}

	public void setJctStatusId(JctStatus jctStatusId) {
		this.jctStatusId = jctStatusId;
	}

	public int getJctSoftDelete() {
		return jctSoftDelete;
	}

	public void setJctSoftDelete(int jctSoftDelete) {
		this.jctSoftDelete = jctSoftDelete;
	}

	public int getVersion() {
		return version;
	}

	public void setVersion(int version) {
		this.version = version;
	}

	public Date getJctCreatedTs() {
		return jctCreatedTs;
	}

	public void setJctCreatedTs(Date jctCreatedTs) {
		this.jctCreatedTs = jctCreatedTs;
	}

	public Date getJctLastmodifiedTs() {
		return jctLastmodifiedTs;
	}

	public void setJctLastmodifiedTs(Date jctLastmodifiedTs) {
		this.jctLastmodifiedTs = jctLastmodifiedTs;
	}

	public String getJctLastmodifiedBy() {
		return jctLastmodifiedBy;
	}

	public void setJctLastmodifiedBy(String jctLastmodifiedBy) {
		this.jctLastmodifiedBy = jctLastmodifiedBy;
	}

	public String getJctCreatedBy() {
		return jctCreatedBy;
	}

	public void setJctCreatedBy(String jctCreatedBy) {
		this.jctCreatedBy = jctCreatedBy;
	}
	
	
}
