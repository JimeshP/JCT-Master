package com.vmware.jct.model;

import java.io.Serializable;
import java.util.Date;

import javax.persistence.Column;
import javax.persistence.Entity;
import javax.persistence.GeneratedValue;
import javax.persistence.GenerationType;
import javax.persistence.Id;
import javax.persistence.NamedQueries;
import javax.persistence.NamedQuery;
import javax.persistence.Table;
import javax.persistence.Temporal;
import javax.persistence.TemporalType;

/**
 * 
 * <p><b>Class name:</b> JctCompletionStatus.java</p>
 * <p><b>Author:</b> InterraIt</p>
 * <p><b>Description:</b>  This class basically object reference of jct_completion_status table </p>
 * <p><b>Copyrights:</b> 	All rights reserved by Interra IT and should only be used for its internal application development.</p>
 * <p><b>Revision History:</b>
 * 	<li>InterraIt - 17/June/2014 - Introduced the class </li>
 * </p>
 */
@NamedQueries({
	@NamedQuery(name = "fetchCompletionObjs",
				query = "from JctCompletionStatus where jctJobReferenceNo = :refNo"),
	@NamedQuery(name = "getNumberOfCompletion", 
				query = "select count(obj.jctJobReferenceNo) from JctCompletionStatus obj where " +
						"obj.jctJobReferenceNo = :jobReferenceNos"),
	@NamedQuery(name = "getLastRowCount", 
				query = "select obj.jctOptionSelected from JctCompletionStatus obj where " +
						"obj.jctJobReferenceNo = :jobRefNo and obj.jctCompletionId = (select max(obj2.jctCompletionId) " +
						"from JctCompletionStatus obj2 where obj2.jctJobReferenceNo = :jobRefNo2)"),
	@NamedQuery(name="deleteAllDummyCompletedData", 
				query="DELETE FROM JctCompletionStatus cmp WHERE cmp.jctJobReferenceNo = :jrNo")
})
@Entity
@Table(name="jct_completion_status")
public class JctCompletionStatus implements Serializable {
	private static final long serialVersionUID = 1L;
	
	@Id
	@GeneratedValue(strategy=GenerationType.IDENTITY)
	@Column(name="jct_completion_id")
	private int jctCompletionId;

	@Column(name="jct_job_reference_no")
	private String jctJobReferenceNo;

	@Temporal(TemporalType.TIMESTAMP)
	@Column(name="jct_start_date")
	private Date jctStartDate;
	
	@Temporal(TemporalType.TIMESTAMP)
	@Column(name="jct_end_date")
	private Date jctEndDate;
	
	@Column(name="jct_completion_counter")
	private int jctCompletionCounter;
	
	@Column(name="jct_option_selected")
	private int jctOptionSelected;

	public int getJctCompletionId() {
		return jctCompletionId;
	}

	public void setJctCompletionId(int jctCompletionId) {
		this.jctCompletionId = jctCompletionId;
	}

	public String getJctJobReferenceNo() {
		return jctJobReferenceNo;
	}

	public void setJctJobReferenceNo(String jctJobReferenceNo) {
		this.jctJobReferenceNo = jctJobReferenceNo;
	}

	public Date getJctStartDate() {
		return jctStartDate;
	}

	public void setJctStartDate(Date jctStartDate) {
		this.jctStartDate = jctStartDate;
	}

	public Date getJctEndDate() {
		return jctEndDate;
	}

	public void setJctEndDate(Date jctEndDate) {
		this.jctEndDate = jctEndDate;
	}

	public int getJctCompletionCounter() {
		return jctCompletionCounter;
	}

	public void setJctCompletionCounter(int jctCompletionCounter) {
		this.jctCompletionCounter = jctCompletionCounter;
	}

	/**
	 * @return the jctOptionSelected
	 */
	public int getJctOptionSelected() {
		return jctOptionSelected;
	}

	/**
	 * @param jctOptionSelected the jctOptionSelected to set
	 */
	public void setJctOptionSelected(int jctOptionSelected) {
		this.jctOptionSelected = jctOptionSelected;
	}
	
}
