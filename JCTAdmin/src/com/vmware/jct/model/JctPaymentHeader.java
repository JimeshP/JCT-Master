package com.vmware.jct.model;

import java.io.Serializable;
import java.util.Date;

import javax.persistence.*;
/**
 * 
 * <p><b>Class name:</b> JctPaymentHeader.java</p>
 * <p><b>Author:</b> InterraIT</p>
 * <p><b>Description:</b>  This class basically object reference of jct_payment_header table </p>
 * <p><b>Copyrights:</b> 	All rights reserved by InterraIT and should only be used for its internal application development.</p>
 * <p><b>Revision History:</b>
 * 	<li>InterraIT - 14/Oct/2014 - Implemented the class </li>
 * </p>
 */
@NamedQueries({	
	@NamedQuery(name = "fetchPaymentHeaderPk", 
			query = "from JctPaymentHeader hdr where hdr.jctPmtHdrId = :headerId"),
	@NamedQuery(name = "updateAmountByHdrId", 
			query = "update JctPaymentHeader set jctPmtHdrTotalAmt = :updatedAmount where jctPmtHdrId = :hdrId")
})

@Entity
@Table(name="jct_payment_header")
public class JctPaymentHeader implements Serializable {
	
	private static final long serialVersionUID = 1L;

	@Id
	@GeneratedValue(strategy=GenerationType.IDENTITY)
	@Column(name="jct_pmt_hdr_id")
	private int jctPmtHdrId;
	
	@Column(name="jct_pmt_hdr_email_id")
	private String jctPmtHdrEmailId;
	
	@Column(name="jct_pmt_hdr_user_id")
	private int jctPmtHdrUserId;
	
	@Column(name="jct_pmt_hdr_total_amt")
	private double jctPmtHdrTotalAmt;
	
	@Column(name="jct_pmt_hdr_created_by")
	private String jctPmtHdrCreatedBy;
	
	@Temporal(TemporalType.TIMESTAMP)
	@Column(name="jct_pmt_hdr_created_ts")
	private Date jctPmtHdrCreatedTs;
	
	@Column(name="jct_pmt_hdr_modified_by")
	private String jctPmtHdrModifiedBy;
	
	@Temporal(TemporalType.TIMESTAMP)
	@Column(name="jct_pmt_hdr_modified_ts")
	private Date jctPmtHdrModifiedTs;
	
	@Column(name="jct_pmt_hdr_soft_delete")
	private int jctPmtHdrSoftDelete;
		
	@OneToOne(cascade=CascadeType.ALL, mappedBy ="jctPaymentHeader")
	private JctFacilitatorDetails jctFacilitatorDetails;
	
	@Column(name="jct_user_customer_id")
	private String jctUserCustomerId;

	public int getJctPmtHdrId() {
		return jctPmtHdrId;
	}

	public void setJctPmtHdrId(int jctPmtHdrId) {
		this.jctPmtHdrId = jctPmtHdrId;
	}

	public String getJctPmtHdrEmailId() {
		return jctPmtHdrEmailId;
	}

	public void setJctPmtHdrEmailId(String jctPmtHdrEmailId) {
		this.jctPmtHdrEmailId = jctPmtHdrEmailId;
	}

	public int getJctPmtHdrUserId() {
		return jctPmtHdrUserId;
	}

	public void setJctPmtHdrUserId(int jctPmtHdrUserId) {
		this.jctPmtHdrUserId = jctPmtHdrUserId;
	}

	public double getJctPmtHdrTotalAmt() {
		return jctPmtHdrTotalAmt;
	}

	public void setJctPmtHdrTotalAmt(double jctPmtHdrTotalAmt) {
		this.jctPmtHdrTotalAmt = jctPmtHdrTotalAmt;
	}

	public String getJctPmtHdrCreatedBy() {
		return jctPmtHdrCreatedBy;
	}

	public void setJctPmtHdrCreatedBy(String jctPmtHdrCreatedBy) {
		this.jctPmtHdrCreatedBy = jctPmtHdrCreatedBy;
	}

	public Date getJctPmtHdrCreatedTs() {
		return jctPmtHdrCreatedTs;
	}

	public void setJctPmtHdrCreatedTs(Date jctPmtHdrCreatedTs) {
		this.jctPmtHdrCreatedTs = jctPmtHdrCreatedTs;
	}

	public String getJctPmtHdrModifiedBy() {
		return jctPmtHdrModifiedBy;
	}

	public void setJctPmtHdrModifiedBy(String jctPmtHdrModifiedBy) {
		this.jctPmtHdrModifiedBy = jctPmtHdrModifiedBy;
	}

	public Date getJctPmtHdrModifiedTs() {
		return jctPmtHdrModifiedTs;
	}

	public void setJctPmtHdrModifiedTs(Date jctPmtHdrModifiedTs) {
		this.jctPmtHdrModifiedTs = jctPmtHdrModifiedTs;
	}

	public int getJctPmtHdrSoftDelete() {
		return jctPmtHdrSoftDelete;
	}

	public void setJctPmtHdrSoftDelete(int jctPmtHdrSoftDelete) {
		this.jctPmtHdrSoftDelete = jctPmtHdrSoftDelete;
	}

	public JctFacilitatorDetails getJctFacilitatorDetails() {
		return jctFacilitatorDetails;
	}

	public void setJctFacilitatorDetails(JctFacilitatorDetails jctFacilitatorDetails) {
		this.jctFacilitatorDetails = jctFacilitatorDetails;
	}

	public String getJctUserCustomerId() {
		return jctUserCustomerId;
	}

	public void setJctUserCustomerId(String jctUserCustomerId) {
		this.jctUserCustomerId = jctUserCustomerId;
	}
	
}
