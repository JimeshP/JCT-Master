package com.vmware.jct.model;

import java.io.Serializable;
import java.util.Date;

import javax.persistence.Column;
import javax.persistence.Entity;
import javax.persistence.Id;
import javax.persistence.JoinColumn;
import javax.persistence.NamedQueries;
import javax.persistence.NamedQuery;
import javax.persistence.OneToOne;
import javax.persistence.Table;
import javax.persistence.Temporal;
import javax.persistence.TemporalType;
/**
 * 
 * <p><b>Class name:</b> JctFacilitatorDetails.java</p>
 * <p><b>Author:</b> InterraIT</p>
 * <p><b>Description:</b>  This class basically object reference of jct_facilitator_details table </p>
 * <p><b>Copyrights:</b> 	All rights reserved by InterraIT and should only be used for its internal application development.</p>
 * <p><b>Revision History:</b>
 * 	<li>InterraIT - 14/Oct/2014 - Implemented the class </li>
 * </p>
 */
@NamedQueries({
	@NamedQuery(name = "fetchSubscribedUserADD",
				query = "select new com.vmware.jct.dao.dto.FacilitatorDetailDTO( SUM(facilitatorDetails.jctFacTotalLimit) , " +
						"SUM (facilitatorDetails.jctFacSubscribeLimit) ) from JctFacilitatorDetails facilitatorDetails " +
						"where facilitatorDetails.jctUserCustomerId = :customerId"),
	@NamedQuery(name = "fetchFacilitatorObj",
				query = "from JctFacilitatorDetails facilitatorDetails where facilitatorDetails.jctUserCustomerId = :customerId " +
						"and facilitatorDetails.jctFacType= :type and facilitatorDetails.jctFacTotalLimit <> facilitatorDetails.jctFacSubscribeLimit " +
						"order by facilitatorDetails.jctFacCreatedTs"),
	@NamedQuery(name = "fetchFacilitatorObj1",
				query = "from JctFacilitatorDetails facilitatorDetails where facilitatorDetails.jctUserCustomerId = :customerId " +
						"and facilitatorDetails.jctFacType= :type and facilitatorDetails.jctFacTotalLimit <> facilitatorDetails.jctFacSubscribeLimit " +
						"order by facilitatorDetails.jctFacCreatedTs"),
	@NamedQuery(name = "fetchFacilitatorSubRenewDtls",
				query = "select new com.vmware.jct.dao.dto.FacilitatorDetailDTO(jctFacUserName, (SUM(jctFacTotalLimit)), (SUM(jctFacSubscribeLimit))) from JctFacilitatorDetails " +
						"where jctFacUserName = :customerId GROUP BY jctFacUserName")
})
@Entity
@Table(name="jct_facilitator_details")
public class JctFacilitatorDetails implements Serializable {

	private static final long serialVersionUID = 1L;
	
	@Id
	@Column(name="jct_fac_mstr_id")
	private int jctFacMstrId;

	@Column(name="jct_fac_user_id")
	private int jctFacUserId;
	
	@Column(name="jct_fac_user_name")
	private String jctFacUserName;
	
	@Column(name="jct_fac_total_limit")
	private int jctFacTotalLimit;
	
	@Column(name="jct_fac_subscribe_limit")
	private int jctFacSubscribeLimit;
	
	@Column(name="jct_fac_type")
	private String jctFacType;
	
	@Temporal(TemporalType.TIMESTAMP)
	@Column(name="jct_fac_subscription_dt")
	private Date jctFacSubscriptionDt;
	
	@Column(name="jct_fac_created_by")
	private String jctFacCreatedBy;
	
	@Temporal(TemporalType.TIMESTAMP)
	@Column(name="jct_fac_created_ts")
	private Date jctFacCreatedTs;
	
	@OneToOne
    @JoinColumn(name="jct_pmt_hdr_id")
	private JctPaymentHeader jctPaymentHeader;
	
	@Column(name="jct_user_customer_id")
	private String jctUserCustomerId;
	
	public int getJctFacMstrId() {
		return jctFacMstrId;
	}

	public void setJctFacMstrId(int jctFacMstrId) {
		this.jctFacMstrId = jctFacMstrId;
	}

	public int getJctFacUserId() {
		return jctFacUserId;
	}

	public void setJctFacUserId(int jctFacUserId) {
		this.jctFacUserId = jctFacUserId;
	}

	public String getJctFacUserName() {
		return jctFacUserName;
	}

	public void setJctFacUserName(String jctFacUserName) {
		this.jctFacUserName = jctFacUserName;
	}

	public int getJctFacSubscribeLimit() {
		return jctFacSubscribeLimit;
	}

	public void setJctFacSubscribeLimit(int jctFacSubscribeLimit) {
		this.jctFacSubscribeLimit = jctFacSubscribeLimit;
	}

	public String getJctFacType() {
		return jctFacType;
	}

	public void setJctFacType(String jctFacType) {
		this.jctFacType = jctFacType;
	}

	public Date getJctFacSubscriptionDt() {
		return jctFacSubscriptionDt;
	}

	public void setJctFacSubscriptionDt(Date jctFacSubscriptionDt) {
		this.jctFacSubscriptionDt = jctFacSubscriptionDt;
	}

	public String getJctFacCreatedBy() {
		return jctFacCreatedBy;
	}

	public void setJctFacCreatedBy(String jctFacCreatedBy) {
		this.jctFacCreatedBy = jctFacCreatedBy;
	}

	public Date getJctFacCreatedTs() {
		return jctFacCreatedTs;
	}

	public void setJctFacCreatedTs(Date jctFacCreatedTs) {
		this.jctFacCreatedTs = jctFacCreatedTs;
	}

	public JctPaymentHeader getJctPaymentHeader() {
		return jctPaymentHeader;
	}

	public void setJctPaymentHeader(JctPaymentHeader jctPaymentHeader) {
		this.jctPaymentHeader = jctPaymentHeader;
	}

	public int getJctFacTotalLimit() {
		return jctFacTotalLimit;
	}

	public void setJctFacTotalLimit(int jctFacTotalLimit) {
		this.jctFacTotalLimit = jctFacTotalLimit;
	}

	public String getJctUserCustomerId() {
		return jctUserCustomerId;
	}

	public void setJctUserCustomerId(String jctUserCustomerId) {
		this.jctUserCustomerId = jctUserCustomerId;
	}
}
