package com.vmware.jct.model;

import java.io.Serializable;
import java.util.Date;

import javax.persistence.Column;
import javax.persistence.Entity;
import javax.persistence.GeneratedValue;
import javax.persistence.GenerationType;
import javax.persistence.Id;
import javax.persistence.JoinColumn;
import javax.persistence.NamedQueries;
import javax.persistence.NamedQuery;
import javax.persistence.OneToOne;
import javax.persistence.Table;
import javax.persistence.Temporal;
import javax.persistence.TemporalType;

@NamedQueries({					
	@NamedQuery(name = "fetchCheckPaymentUserUniqueDetails",
				query = "Select DISTINCT new com.vmware.jct.dao.dto.PaymentDetailsDTO(obj.jctCheckPaymentCheckNo,obj.jctCheckPaymentCheckDate,obj.jctPmtdtlsPmtTransNos) " +
						"from JctCheckPaymentUserDetails obj where obj.jctCheckPaymentIsHonored = 0 and obj.jctCheckPaymentSoftDelete = 0"),
	@NamedQuery(name = "fetchCheckPaymentUserUniqueDetailsByChk",
				query = "Select DISTINCT new com.vmware.jct.dao.dto.PaymentDetailsDTO(obj.jctCheckPaymentCheckNo,obj.jctCheckPaymentCheckDate,obj.jctPmtdtlsPmtTransNos) " +
						"from JctCheckPaymentUserDetails obj " +
						"where obj.jctCheckPaymentCheckNo = :chequeNum and obj.jctCheckPaymentIsHonored = 0 and obj.jctCheckPaymentSoftDelete = 0"),
						
						
						
	@NamedQuery(name = "CheckPymntUserUniqueDtlByChkAndUsrName",
				query = "Select DISTINCT new com.vmware.jct.dao.dto.PaymentDetailsDTO(obj.jctCheckPaymentCheckNo,obj.jctCheckPaymentCheckDate,obj.jctPmtdtlsPmtTransNos) " +
						"from JctCheckPaymentUserDetails obj " +
						"where obj.jctCheckPaymentCheckNo = :chequeNum and obj.jctCheckPaymentUserName = :userName " +
						"and obj.jctCheckPaymentIsHonored = 0 and obj.jctCheckPaymentSoftDelete = 0"),
						
	@NamedQuery(name= "fetchuserRemainingFields",
				query = "Select new com.vmware.jct.dao.dto.PaymentDetailsDTO(obj.jctCheckPaymentUserName, obj.jctCheckPaymentCustomerId, obj.jctCheckPaymentRoleId," +
						"obj.jctCheckPaymentHeaderId.jctPmtHdrId,obj.jctCheckPaymentUserId )" +
						"from JctCheckPaymentUserDetails obj " +
						"where obj.jctPmtdtlsPmtTransNos = :transId and obj.jctCheckPaymentIsHonored = 0 and obj.jctCheckPaymentSoftDelete = 0"),
						
	@NamedQuery(name = "updateCheckPaymentDetailsToHonor", 
				query = "Update JctCheckPaymentUserDetails chk set chk.jctCheckPaymentSoftDelete = 1 where jctUserId = :jctUserId"),
						
	@NamedQuery(name = "updateCheckPaymentDetailsToDishonor", 
				query = "Update JctCheckPaymentUserDetails dtl set dtl.jctCheckPaymentSoftDelete = 1, dtl.jctCheckPaymentIsHonored = 1" +
								"where dtl.jctCheckPaymentUserId = :pk"),	

	@NamedQuery(name="fetchCheckUserDtlByTranId",
				query = "from JctCheckPaymentUserDetails where jctPmtdtlsPmtTransNos = :tranId and jctCheckPaymentSoftDelete = 0"),					
						
	@NamedQuery(name = "fetchCheckUserDtlByCheckNoForSingleUser",
				query = "from JctCheckPaymentUserDetails " +
						"where jctCheckPaymentUserName = :userName and jctCheckPaymentCustomerId = :customerId and " +
						"jctCheckPaymentCheckNo = :checkNo and jctCheckPaymentHeaderId.jctPmtHdrId = :headerId and jctCheckPaymentSoftDelete = 0"),
						
	@NamedQuery(name="getExistingUserByTranIdFromJctChk",
				query = "Select new com.vmware.jct.dao.dto.PaymentDetailsDTO(user.jctUserName, user.jctPassword, chk.jctCheckPaymentUserType, chk.jctCheckPaymentRoleId ) " +
						"from JctCheckPaymentUserDetails chk, JctUser user " +
						"where chk.jctPmtdtlsPmtTransNos = :tranId and user.jctUserId = chk.jctUserId"),
	@NamedQuery(name="getChkNoByUserName",
				query = "Select jctCheckPaymentCheckNo from JctCheckPaymentUserDetails where jctCheckPaymentUserName = :userName " +
						"and jctCheckPaymentSoftDelete = 0 AND jctCheckPaymentIsHonored = 0")
})
@Entity
@Table(name="jct_check_payment_user_details")
public class JctCheckPaymentUserDetails implements Serializable{
	private static final long serialVersionUID = 1L;
	
	@Id
	@GeneratedValue(strategy=GenerationType.IDENTITY)
	@Column(name="jct_check_payment_user_id")
	private int jctCheckPaymentUserId;
	
	@Column(name="jct_check_payment_user_name")
	private String jctCheckPaymentUserName;
	
	@Column(name="jct_check_payment_customer_id")
	private String jctCheckPaymentCustomerId;
	
	@Column(name="jct_check_payment_role_id")
	private int jctCheckPaymentRoleId;
	
	@Column(name="jct_check_payment_check_no")
	private String jctCheckPaymentCheckNo;
	
	@Temporal(TemporalType.TIMESTAMP)
	@Column(name="jct_check_payment_check_date")
	private Date jctCheckPaymentCheckDate;
	
	@Column(name="jct_check_payment_is_honored")
	private int jctCheckPaymentIsHonored;
	
	@Column(name="jct_check_payment_created_by")
	private String jctCheckPaymentCreatedBy;
	
	@Temporal(TemporalType.TIMESTAMP)
	@Column(name="jct_check_payment_created_ts")
	private Date jctCheckPaymentCreatedTs;
	
	@Column(name="jct_check_payment_modified_by")
	private String jctCheckPaymentmModifiedBy;
	
	@Temporal(TemporalType.TIMESTAMP)
	@Column(name="jct_check_payment_modified_ts")
	private Date jctCheckPaymentModifiedTs;
	
	@Column(name="jct_check_payment_user_type")
	private String jctCheckPaymentUserType;
	
	@Column(name="jct_check_payment_soft_delete")
	private int jctCheckPaymentSoftDelete;
	
	@Column(name="jct_pmt_dtls_pmt_trans_nos")
	private String jctPmtdtlsPmtTransNos;	
	
	@Column(name="jct_user_id")
	private int jctUserId;	
	
	@OneToOne
    @JoinColumn(name="jct_check_payment_header_id")
	private JctPaymentHeader jctCheckPaymentHeaderId;
	
	@OneToOne
    @JoinColumn(name="jct_check_payment_details_id")
	private JctPaymentDetails jctCheckPaymentDetailsId;

	@Column(name="jct_user_expiry_duration")
	private int jctCheckPaymentExpiryDuration;
	
	public int getJctCheckPaymentUserId() {
		return jctCheckPaymentUserId;
	}

	public void setJctCheckPaymentUserId(int jctCheckPaymentUserId) {
		this.jctCheckPaymentUserId = jctCheckPaymentUserId;
	}

	public String getJctCheckPaymentUserName() {
		return jctCheckPaymentUserName;
	}

	public void setJctCheckPaymentUserName(String jctCheckPaymentUserName) {
		this.jctCheckPaymentUserName = jctCheckPaymentUserName;
	}

	public String getJctCheckPaymentCustomerId() {
		return jctCheckPaymentCustomerId;
	}

	public void setJctCheckPaymentCustomerId(String jctCheckPaymentCustomerId) {
		this.jctCheckPaymentCustomerId = jctCheckPaymentCustomerId;
	}

	public int getJctCheckPaymentRoleId() {
		return jctCheckPaymentRoleId;
	}

	public void setJctCheckPaymentRoleId(int jctCheckPaymentRoleId) {
		this.jctCheckPaymentRoleId = jctCheckPaymentRoleId;
	}

	public String getJctCheckPaymentCheckNo() {
		return jctCheckPaymentCheckNo;
	}

	public void setJctCheckPaymentCheckNo(String jctCheckPaymentCheckNo) {
		this.jctCheckPaymentCheckNo = jctCheckPaymentCheckNo;
	}

	public Date getJctCheckPaymentCheckDate() {
		return jctCheckPaymentCheckDate;
	}

	public void setJctCheckPaymentCheckDate(Date jctCheckPaymentCheckDate) {
		this.jctCheckPaymentCheckDate = jctCheckPaymentCheckDate;
	}

	public int getJctCheckPaymentIsHonored() {
		return jctCheckPaymentIsHonored;
	}

	public void setJctCheckPaymentIsHonored(int jctCheckPaymentIsHonored) {
		this.jctCheckPaymentIsHonored = jctCheckPaymentIsHonored;
	}

	public String getJctCheckPaymentCreatedBy() {
		return jctCheckPaymentCreatedBy;
	}

	public void setJctCheckPaymentCreatedBy(String jctCheckPaymentCreatedBy) {
		this.jctCheckPaymentCreatedBy = jctCheckPaymentCreatedBy;
	}

	public Date getJctCheckPaymentCreatedTs() {
		return jctCheckPaymentCreatedTs;
	}

	public void setJctCheckPaymentCreatedTs(Date jctCheckPaymentCreatedTs) {
		this.jctCheckPaymentCreatedTs = jctCheckPaymentCreatedTs;
	}

	public String getJctCheckPaymentmModifiedBy() {
		return jctCheckPaymentmModifiedBy;
	}

	public void setJctCheckPaymentmModifiedBy(String jctCheckPaymentmModifiedBy) {
		this.jctCheckPaymentmModifiedBy = jctCheckPaymentmModifiedBy;
	}

	public Date getJctCheckPaymentModifiedTs() {
		return jctCheckPaymentModifiedTs;
	}

	public void setJctCheckPaymentModifiedTs(Date jctCheckPaymentModifiedTs) {
		this.jctCheckPaymentModifiedTs = jctCheckPaymentModifiedTs;
	}

	public JctPaymentHeader getJctCheckPaymentHeaderId() {
		return jctCheckPaymentHeaderId;
	}

	public void setJctCheckPaymentHeaderId(JctPaymentHeader jctCheckPaymentHeaderId) {
		this.jctCheckPaymentHeaderId = jctCheckPaymentHeaderId;
	}

	public JctPaymentDetails getJctCheckPaymentDetailsId() {
		return jctCheckPaymentDetailsId;
	}

	public void setJctCheckPaymentDetailsId(
			JctPaymentDetails jctCheckPaymentDetailsId) {
		this.jctCheckPaymentDetailsId = jctCheckPaymentDetailsId;
	}

	public String getJctCheckPaymentUserType() {
		return jctCheckPaymentUserType;
	}

	public void setJctCheckPaymentUserType(String jctCheckPaymentUserType) {
		this.jctCheckPaymentUserType = jctCheckPaymentUserType;
	}

	public int getJctCheckPaymentSoftDelete() {
		return jctCheckPaymentSoftDelete;
	}

	public void setJctCheckPaymentSoftDelete(int jctCheckPaymentSoftDelete) {
		this.jctCheckPaymentSoftDelete = jctCheckPaymentSoftDelete;
	}

	public String getJctPmtdtlsPmtTransNos() {
		return jctPmtdtlsPmtTransNos;
	}

	public void setJctPmtdtlsPmtTransNos(String jctPmtdtlsPmtTransNos) {
		this.jctPmtdtlsPmtTransNos = jctPmtdtlsPmtTransNos;
	}

	public int getJctUserId() {
		return jctUserId;
	}

	public void setJctUserId(int jctUserId) {
		this.jctUserId = jctUserId;
	}

	public int getJctCheckPaymentExpiryDuration() {
		return jctCheckPaymentExpiryDuration;
	}

	public void setJctCheckPaymentExpiryDuration(int jctCheckPaymentExpiryDuration) {
		this.jctCheckPaymentExpiryDuration = jctCheckPaymentExpiryDuration;
	}
	
}
