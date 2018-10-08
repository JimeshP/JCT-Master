package com.vmware.jct.model;

import java.io.Serializable;
import java.util.Date;

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
import javax.persistence.Table;
import javax.persistence.Temporal;
import javax.persistence.TemporalType;
import javax.persistence.Version;
/**
 * 
 * <p><b>Class name:</b> JctStatusSearch.java</p>
 * <p><b>Author:</b> InterraIt</p>
 * <p><b>Description:</b>  This class basically object reference of jct_status_search table.  </p>
 * <p><b>Copyrights:</b> 	All rights reserved by Interra IT and should only be used for its internal application development.</p>
 * <p><b>Revision History:</b>
 * 	<li>InterraIt - 31/Jan/2014 - Implement comments, introduce Named query </li>
 * </p>
 */
@NamedQueries({
	@NamedQuery(name = "searchBeforeSketchDiag",
				query = " select new com.vmware.jct.dao.dto.StatusSearchDTO(status.jctJobrefNo , status.jctBsSnapshotString) " +
						"from JctStatusSearch status where status.jctUserLevels = :jobTitle " +
						"and status.jctIsInactive=0 and status.jctSoftDelete=0"),
	@NamedQuery(name = "searchBeforeSketchJRNO",
				query = "Select status.jctJobrefNo from JctStatusSearch status where jctJobrefNo = :jrno " +
						"and status.jctUserLevels = :jobTitle and status.jctIsInactive=0 and status.jctSoftDelete=0 " +
						"and status.jctBsSnapshotString = :snpShtStr"),
	@NamedQuery(name = "searchASAvailableByJRNO",
				query = "Select status.jctJobrefNo from JctStatusSearch status where status.jctUserLevels = :jobTitle " +
						"and status.jctIsInactive=0 and status.jctSoftDelete=0 and status.jctJobrefNo = :jctJobrefNo"),
	@NamedQuery(name = "getBSDiagramByJRNO",
				query = "Select status.jctAsSnapshotString from JctStatusSearch status where status.jctIsInactive=0 " +
						"and status.jctSoftDelete=0 and status.jctJobrefNo = :jctJobrefNo"),
	@NamedQuery(name = "searchAfterSketchDiag",
				query = " select new com.vmware.jct.dao.dto.StatusSearchDTO(status.jctJobrefNo , status.jctAsSnapshotString) " +
						"from JctStatusSearch status where status.jctUserLevels = :jobTitle " +
						"and status.jctIsInactive=0 and status.jctSoftDelete=0"),
	@NamedQuery(name = "searchAfterSketchJRNO",
				query = "Select status.jctJobrefNo from JctStatusSearch status where status.jctUserLevels = :jobTitle " +
						"and status.jctIsInactive=0 and status.jctSoftDelete=0 and status.jctAsSnapshotString = :snpShtStr"),
	@NamedQuery(name = "searchBSAvailableByJRNO",
				query = "Select status.jctJobrefNo from JctStatusSearch status where status.jctUserLevels = :jobTitle " +
						"and status.jctIsInactive=0 and status.jctSoftDelete=0 and status.jctJobrefNo = :jctJobrefNo"),
	@NamedQuery(name = "getASDiagramByJRNO",
				query = "Select status.jctBsSnapshotString from JctStatusSearch status where status.jctIsInactive=0 " +
						"and status.jctSoftDelete=0 and status.jctJobrefNo = :jctJobrefNo"),
	@NamedQuery(name = "disableStatusSearch",
				query = "update JctStatusSearch set jctIsInactive=1 where jctJobrefNo = :jobRefNo"),
	@NamedQuery(name = "updateSoftDeleteForEdit",
				query = "update JctStatusSearch set jctSoftDelete=1 where jctSoftDelete=0 and jctJobrefNo = :jobRefNo"),
	//@NamedQuery(name = "getPrevBsForView",query = "Select max(status.jctBsSnapshotString) from JctStatusSearch status where status.jctSoftDelete=:softDel and status.jctJobrefNo = :jctJobrefNo"),
	//@NamedQuery(name = "getPrevAsForView",query = "Select max(status.jctAsSnapshotString) from JctStatusSearch status where status.jctSoftDelete=:softDel and status.jctJobrefNo = :jctJobrefNo"),
	@NamedQuery(name = "getPrevBsForView",
				query = "Select bsHdr.jctBsSnapshotString from JctBeforeSketchHeader bsHdr where bsHdr.softDelete=1 " +
						"and bsHdr.jctStatus=5 and bsHdr.jctBsJobrefNo = :jctJobrefNo and " +
						"bsHdr.jctBsHeaderId = (select max(bsHdr2.jctBsHeaderId) from JctBeforeSketchHeader bsHdr2 " +
						"where bsHdr2.softDelete=1 and bsHdr2.jctStatus=5 and bsHdr2.jctBsJobrefNo = :jctJobrefNo2)"),
	@NamedQuery(name = "getPrevAsForView",
				query = "Select asHdr.jctAsFinalpageSnapshotString from JctAfterSketchHeader asHdr " +
						"where asHdr.jctAsSoftDelete=1 and asHdr.jctAsStatusId=5 and asHdr.jctAsJobrefNo = :jctJobrefNo " +
						"and asHdr.jctAsHeaderId = (select max(asHdr2.jctAsHeaderId) " +
						"from JctAfterSketchHeader asHdr2 where asHdr2.jctAsSoftDelete=1 and " +
						"asHdr2.jctAsStatusId=5 and asHdr2.jctAsJobrefNo = :jctJobrefNo2)"),
	@NamedQuery(name = "getNosOfTimesShared",
				query = "select status.jctJobrefNo from JctStatusSearch status where status.jctJobrefNo = :jctJobrefNo"),
	@NamedQuery(name="getStatusSearchDummyData", 
				query = "from JctStatusSearch where jctJobrefNo = :jrNo"),
	@NamedQuery(name="deleteAllDummyStatusSearchData", 
				query="DELETE FROM JctStatusSearch srch where srch.jctJobrefNo = :jrNo"),
				
	@NamedQuery(name = "searchBeforeSketchDiagByOccupation",
			query = " select new com.vmware.jct.dao.dto.StatusSearchDTO(status.jctJobrefNo , status.jctBsSnapshotString, status.occupationCode) " +
			"from JctStatusSearch status where " +
			"status.jctIsInactive=0 and status.jctSoftDelete=0 AND status.jctJobrefNo NOT LIKE '00000000ADPRV%'"),
	@NamedQuery(name = "searchAfterSketchDiagByOccupation",
			query = " select new com.vmware.jct.dao.dto.StatusSearchDTO(status.jctJobrefNo , status.jctAsSnapshotString, status.occupationCode) " +
			"from JctStatusSearch status where " +
			"status.jctIsInactive=0 and status.jctSoftDelete=0  AND status.jctJobrefNo NOT LIKE '00000000ADPRV%'"),		
	@NamedQuery(name = "fetchAllBSDiagramCount",
			query = "Select COUNT(srch.jctStatusSearchId) from JctStatusSearch srch where srch.jctSoftDelete = 0 and " +
			"srch.jctJobrefNo IN (select distinct (info.jctJobrefNo) from JctUserLoginInfo info) and srch.jctJobrefNo NOT LIKE '00000000ADPRV%'"),
	@NamedQuery(name = "fetchAllASDiagramCount",
			query = "Select COUNT(srch.jctStatusSearchId) from JctStatusSearch srch where srch.jctSoftDelete = 0 and " +
			"srch.jctJobrefNo IN (select distinct (info.jctJobrefNo) from JctUserLoginInfo info)  and srch.jctJobrefNo NOT LIKE '00000000ADPRV%'"),
	@NamedQuery(name = "searchBeforeSketchDiagByOccupationCode",
			query = " select new com.vmware.jct.dao.dto.StatusSearchDTO(status.jctJobrefNo , status.jctBsSnapshotString, status.occupationCode) " +
			"from JctStatusSearch status where " +
			"status.jctIsInactive=0 and status.jctSoftDelete=0 and status.occupationCode = :code"),
	@NamedQuery(name = "searchAfterSketchDiagByOccupationCode",
			query = " select new com.vmware.jct.dao.dto.StatusSearchDTO(status.jctJobrefNo , status.jctAsSnapshotString, status.occupationCode) " +
			"from JctStatusSearch status where " +
			"status.jctIsInactive=0 and status.jctSoftDelete=0 and status.occupationCode = :code"),
	@NamedQuery(name = "fetchAllBSDiagramCountCode",
			query = "Select COUNT(srch.jctStatusSearchId) from JctStatusSearch srch where srch.occupationCode = :code and srch.jctSoftDelete = 0 and " +
			"srch.jctJobrefNo IN (select distinct (info.jctJobrefNo) from JctUserLoginInfo info)"),
	@NamedQuery(name = "fetchAllASDiagramCountCode",
			query = "Select COUNT(srch.jctStatusSearchId) from JctStatusSearch srch where srch.occupationCode = :code and srch.jctSoftDelete = 0 and " +
			"srch.jctJobrefNo IN (select distinct (info.jctJobrefNo) from JctUserLoginInfo info)")
	
})
@Entity
@Table(name="jct_status_search")
public class JctStatusSearch implements Serializable{
	
	private static final long serialVersionUID = 1L;
	
	@Id
	@GeneratedValue(strategy=GenerationType.IDENTITY)
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
	
	@Column(name="jct_status_decision")
	private int jctStatusDecision;
	
	@Column(name="occupation_code")
	private String occupationCode;

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

	public int getJctStatusDecision() {
		return jctStatusDecision;
	}

	public void setJctStatusDecision(int jctStatusDecision) {
		this.jctStatusDecision = jctStatusDecision;
	}

	public String getOccupationCode() {
		return occupationCode;
	}

	public void setOccupationCode(String occupationCode) {
		this.occupationCode = occupationCode;
	}
	
}
