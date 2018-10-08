package com.vmware.jct.model;

import java.io.Serializable;
import java.util.Date;
import javax.persistence.*;
/**
 * 
 * <p><b>Class name:</b> JctPaymentDetails.java</p>
 * <p><b>Author:</b> InterraIT</p>
 * <p><b>Description:</b>  This class basically object reference of jct_payment_details table </p>
 * <p><b>Copyrights:</b> 	All rights reserved by InterraIT and should only be used for its internal application development.</p>
 * <p><b>Revision History:</b>
 * 	<li>InterraIT - 14/Oct/2014 - Implemented the class </li>
 * </p>
 */

@NamedQueries({											
	@NamedQuery(name = "updatePaymentDtlAfterHonor", 
				query = "Update JctPaymentDetails dtl set dtl.jctPmtDtlsModifiedBy =:createdBy,  dtl.jctPmtDtlsModifiedTs =:currentDate " +
						"where dtl.jctPmtDtlsId = :detailId"),
	@NamedQuery(name = "updatePaymentDtlForDishonor", 
				query = "Update JctPaymentDetails dtl " +
						"set dtl.jctPmtDtlsModifiedBy =:createdBy,  dtl.jctPmtDtlsModifiedTs =:currentDate, dtl.jctPmtDtlsSoftDelete = 1" +
						"where dtl.jctPmtDtlsId = :detailId")
})

@Entity
@Table(name="jct_payment_details")
public class JctPaymentDetails implements Serializable {

	private static final long serialVersionUID = 1L;

	@Id
	@GeneratedValue(strategy=GenerationType.IDENTITY)
	@Column(name="jct_pmt_dtls_id")
	private int jctPmtDtlsId;
	
	@Temporal(TemporalType.TIMESTAMP)
	@Column(name="jct_pmt_dtls_date")
	private Date jctPmtDtlsDate;
	
	@Column(name="jct_pmt_dtls_pmt_typ_id")
	private int jctPmtDtlsPmtTypId;
	
	@Column(name="jct_pmt_dtls_pmt_typ_desc")
	private String jctPmtDtlsPmtTypDesc;
	
	@Column(name="jct_pmt_dtls_pmt_trans_nos")
	private String jctPmtDtlsPmtTransNos;
	
	@Column(name="jct_pmt_dtls_cheque_nos")
	private String jctPmtDtlsChequeNos;
	
	@Column(name="jct_pmt_dtls_created_by")
	private String jctPmtDtlsCreatedBy;
	
	@Temporal(TemporalType.TIMESTAMP)
	@Column(name="jct_pmt_dtls_created_ts")
	private Date jctPmtDtlsCreatedTs;
	
	@Column(name="jct_pmt_dtls_modified_by")
	private String jctPmtDtlsModifiedBy;
	
	@Temporal(TemporalType.TIMESTAMP)
	@Column(name="jct_pmt_dtls_modified_ts")
	private Date jctPmtDtlsModifiedTs;
	
	@Column(name="jct_pmt_dtls_soft_delete")
	private int jctPmtDtlsSoftDelete;
	
	@OneToOne
    @JoinColumn(name="jct_pmt_hdr_id")
	private JctPaymentHeader jctPmtHdrId;

	public JctPaymentHeader getJctPmtHdrId() {
		return jctPmtHdrId;
	}

	public void setJctPmtHdrId(JctPaymentHeader jctPmtHdrId) {
		this.jctPmtHdrId = jctPmtHdrId;
	}

	public int getJctPmtDtlsId() {
		return jctPmtDtlsId;
	}

	public void setJctPmtDtlsId(int jctPmtDtlsId) {
		this.jctPmtDtlsId = jctPmtDtlsId;
	}

	public Date getJctPmtDtlsDate() {
		return jctPmtDtlsDate;
	}

	public void setJctPmtDtlsDate(Date jctPmtDtlsDate) {
		this.jctPmtDtlsDate = jctPmtDtlsDate;
	}

	public int getJctPmtDtlsPmtTypId() {
		return jctPmtDtlsPmtTypId;
	}

	public void setJctPmtDtlsPmtTypId(int jctPmtDtlsPmtTypId) {
		this.jctPmtDtlsPmtTypId = jctPmtDtlsPmtTypId;
	}

	public String getJctPmtDtlsPmtTypDesc() {
		return jctPmtDtlsPmtTypDesc;
	}

	public void setJctPmtDtlsPmtTypDesc(String jctPmtDtlsPmtTypDesc) {
		this.jctPmtDtlsPmtTypDesc = jctPmtDtlsPmtTypDesc;
	}

	public String getJctPmtDtlsPmtTransNos() {
		return jctPmtDtlsPmtTransNos;
	}

	public void setJctPmtDtlsPmtTransNos(String jctPmtDtlsPmtTransNos) {
		this.jctPmtDtlsPmtTransNos = jctPmtDtlsPmtTransNos;
	}

	public String getJctPmtDtlsChequeNos() {
		return jctPmtDtlsChequeNos;
	}

	public void setJctPmtDtlsChequeNos(String jctPmtDtlsChequeNos) {
		this.jctPmtDtlsChequeNos = jctPmtDtlsChequeNos;
	}

	public String getJctPmtDtlsCreatedBy() {
		return jctPmtDtlsCreatedBy;
	}

	public void setJctPmtDtlsCreatedBy(String jctPmtDtlsCreatedBy) {
		this.jctPmtDtlsCreatedBy = jctPmtDtlsCreatedBy;
	}

	public Date getJctPmtDtlsCreatedTs() {
		return jctPmtDtlsCreatedTs;
	}

	public void setJctPmtDtlsCreatedTs(Date jctPmtDtlsCreatedTs) {
		this.jctPmtDtlsCreatedTs = jctPmtDtlsCreatedTs;
	}

	public String getJctPmtDtlsModifiedBy() {
		return jctPmtDtlsModifiedBy;
	}

	public void setJctPmtDtlsModifiedBy(String jctPmtDtlsModifiedBy) {
		this.jctPmtDtlsModifiedBy = jctPmtDtlsModifiedBy;
	}

	public Date getJctPmtDtlsModifiedTs() {
		return jctPmtDtlsModifiedTs;
	}

	public void setJctPmtDtlsModifiedTs(Date jctPmtDtlsModifiedTs) {
		this.jctPmtDtlsModifiedTs = jctPmtDtlsModifiedTs;
	}

	public int getJctPmtDtlsSoftDelete() {
		return jctPmtDtlsSoftDelete;
	}

	public void setJctPmtDtlsSoftDelete(int jctPmtDtlsSoftDelete) {
		this.jctPmtDtlsSoftDelete = jctPmtDtlsSoftDelete;
	}
	
}
