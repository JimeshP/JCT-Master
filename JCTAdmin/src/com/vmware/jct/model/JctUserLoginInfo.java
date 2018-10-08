package com.vmware.jct.model;

import java.io.Serializable;
import javax.persistence.*;

import java.util.Date;


/**
 * 
 * <p><b>Class name:</b> JctUserLoginInfo.java</p>
 * <p><b>Author:</b> InterreIt</p>
 * <p><b>Description:</b>  This class basically object reference of jct_user_login_info table </p>
 * <p><b>Copyrights:</b> 	All rights reserved by Interra IT and should only be used for its internal application development.</p>
 * <p><b>Revision History:</b>
 * 	<li>InterreIt - 31/Jan/2014 - Implement comments, introduce Named query </li>
 * </p>
 */

@NamedQueries({
	@NamedQuery(name = "fetchCreatedByByJobRefNo",
			query = "select jctAsCreatedBy from JctUserLoginInfo where jctJobrefNo = :jobRefNo order by jctAsLastmodifiedTs"),
	@NamedQuery(name = "getObjByinfoCrtfBy",
			query = "from JctUserLoginInfo obj where obj.jctAsCreatedBy = :userName and obj.jctJobrefNo NOT LIKE '%00000000ADPRV%'")
})
@Entity
@Table(name="jct_user_login_info")
public class JctUserLoginInfo implements Serializable {
	private static final long serialVersionUID = 1L;

	@Id
	@GeneratedValue(strategy=GenerationType.IDENTITY)
	@Column(name="jct_user_login_info_id")
	private int jctUserLoginInfoId;

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

	@Temporal(TemporalType.TIMESTAMP)
	@Column(name="jct_end_time")
	private Date jctEndTime;

	@Column(name="jct_jobref_no")
	private String jctJobrefNo;

	@Column(name="jct_page_info")
	private String jctPageInfo;

	@Temporal(TemporalType.TIMESTAMP)
	@Column(name="jct_start_time")
	private Date jctStartTime;

	@Column(name="jct_user_id")
	private int jctUserId;
	
	@Column(name="jct_completed")
	private int jctCompleted;
	
	@Column(name="jct_next_page")
	private String jctNextPage;

	public String getJctNextPage() {
		return jctNextPage;
	}

	public void setJctNextPage(String jctNextPage) {
		this.jctNextPage = jctNextPage;
	}

	private int version;

	//bi-directional many-to-one association to JctStatus
	@ManyToOne
	@JoinColumn(name="jct_status_id")
	private JctStatus jctStatus;

	public JctUserLoginInfo() {
	}

	public int getJctUserLoginInfoId() {
		return this.jctUserLoginInfoId;
	}

	public void setJctUserLoginInfoId(int jctUserLoginInfoId) {
		this.jctUserLoginInfoId = jctUserLoginInfoId;
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

	public Date getJctEndTime() {
		return this.jctEndTime;
	}

	public void setJctEndTime(Date jctEndTime) {
		this.jctEndTime = jctEndTime;
	}

	public String getJctJobrefNo() {
		return this.jctJobrefNo;
	}

	public void setJctJobrefNo(String jctJobrefNo) {
		this.jctJobrefNo = jctJobrefNo;
	}

	public String getJctPageInfo() {
		return this.jctPageInfo;
	}

	public void setJctPageInfo(String jctPageInfo) {
		this.jctPageInfo = jctPageInfo;
	}

	public Date getJctStartTime() {
		return this.jctStartTime;
	}

	public void setJctStartTime(Date jctStartTime) {
		this.jctStartTime = jctStartTime;
	}

	public int getJctUserId() {
		return this.jctUserId;
	}

	public void setJctUserId(int jctUserId) {
		this.jctUserId = jctUserId;
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

	public int getJctCompleted() {
		return jctCompleted;
	}

	public void setJctCompleted(int jctCompleted) {
		this.jctCompleted = jctCompleted;
	}
}