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
 * <p><b>Class name:</b> JctPdfRecords.java</p>
 * <p><b>Author:</b> InterraIt</p>
 * <p><b>Description:</b>  This class basically object reference of jct_pdf_records table </p>
 * <p><b>Copyrights:</b> 	All rights reserved by Interra IT and should only be used for its internal application development.</p>
 * <p><b>Revision History:</b>
 * 	<li>InterraIt - 31/Jan/2014 - Implement comments, introduce Named query </li>
 * </p>
 */
@NamedQueries({
	@NamedQuery(name = "fetchLastUpdated",
				query = "select obj from JctPdfRecords obj where jctFileId = (select max(jctFileId) from JctPdfRecords where jctJobReferenceNo = :refNo)"),
	@NamedQuery(name = "fetchLatestPdf",
				query = "from JctPdfRecords obj where obj.jctJobReferenceNo = :jobRefNo  and obj.jctShowPdf = 0 "+
						"and obj.jctCreatedTimestamp = (select max(jctCreatedTimestamp) from JctPdfRecords where jctJobReferenceNo = :jobRefNo)"),
	/*@NamedQuery(name = "fetchOldPdf",
				query = "select new com.vmware.jct.dao.dto.SearchPdfDTO(jctFileLocation,jctFileName,jctCreatedTimestamp) " +
						"from JctPdfRecords where jctUser.jctUserId = :userId order by jctCreatedTimestamp desc")*/
	@NamedQuery(name = "fetchOldPdf",
				query = "select new com.vmware.jct.dao.dto.SearchPdfDTO(jctFileLocation,jctFileName,jctCreatedTimestamp) " +
						"from JctPdfRecords where jctUser.jctUserId = :userId")
})
@Entity
@Table(name="jct_pdf_records")
public class JctPdfRecords implements Serializable{
	private static final long serialVersionUID = 1L;
	
	@Id
	@GeneratedValue(strategy=GenerationType.IDENTITY)
	@Column(name="jct_file_id")
	private int jctFileId;
	
	@Column(name="jct_file_location")
	private String jctFileLocation;
	
	@Column(name="jct_job_reference_no")
	private String jctJobReferenceNo;
	
	@Column(name = "jct_file_name")
	private String jctFileName;
	
	@Temporal(TemporalType.TIMESTAMP)
	@Column(name="jct_created_timestamp")
	private Date jctCreatedTimestamp;
	
	@Column(name="jct_created_by")
	private String jctCreatedBy;
	
	/** THIS IS FOR PUBLIC VERSION **/
	@ManyToOne
	@JoinColumn(name="jct_user_id")
	private JctUser jctUser;	
	
	@Column(name = "jct_show_pdf")
	private int jctShowPdf;	
	/********************************/

	public int getJctFileId() {
		return jctFileId;
	}

	public void setJctFileId(int jctFileId) {
		this.jctFileId = jctFileId;
	}

	public String getJctFileLocation() {
		return jctFileLocation;
	}

	public void setJctFileLocation(String jctFileLocation) {
		this.jctFileLocation = jctFileLocation;
	}

	public String getJctJobReferenceNo() {
		return jctJobReferenceNo;
	}

	public void setJctJobReferenceNo(String jctJobReferenceNo) {
		this.jctJobReferenceNo = jctJobReferenceNo;
	}

	public String getJctFileName() {
		return jctFileName;
	}

	public void setJctFileName(String jctFileName) {
		this.jctFileName = jctFileName;
	}

	public Date getJctCreatedTimestamp() {
		return jctCreatedTimestamp;
	}

	public void setJctCreatedTimestamp(Date jctCreatedTimestamp) {
		this.jctCreatedTimestamp = jctCreatedTimestamp;
	}

	public String getJctCreatedBy() {
		return jctCreatedBy;
	}

	public void setJctCreatedBy(String jctCreatedBy) {
		this.jctCreatedBy = jctCreatedBy;
	}
	
	public JctUser getJctUser() {
		return jctUser;
	}

	public void setJctUser(JctUser jctUser) {
		this.jctUser = jctUser;
	}

	public int getJctShowPdf() {
		return jctShowPdf;
	}

	public void setJctShowPdf(int jctShowPdf) {
		this.jctShowPdf = jctShowPdf;
	}	
}