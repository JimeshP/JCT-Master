package com.vmware.jct.service.vo;

import java.util.Date;
import java.util.List;

import com.vmware.jct.dao.dto.PaymentDetailsDTO;

public class PaymentVO {
	private int statusCode;
	private String statusDesc;
	private String userName;
	private Date paymentDate;
	private String chequeNumber;
	private String paymentDetails;
	private List<PaymentDetailsDTO> existingUsers;
	private String chkDetailsString;
	
	public String getPaymentDetails() {
		return paymentDetails;
	}
	public void setPaymentDetails(String paymentDetails) {
		this.paymentDetails = paymentDetails;
	}
	public int getStatusCode() {
		return statusCode;
	}
	public void setStatusCode(int statusCode) {
		this.statusCode = statusCode;
	}
	public String getStatusDesc() {
		return statusDesc;
	}
	public void setStatusDesc(String statusDesc) {
		this.statusDesc = statusDesc;
	}
	public String getUserName() {
		return userName;
	}
	public void setUserName(String userName) {
		this.userName = userName;
	}
	public Date getPaymentDate() {
		return paymentDate;
	}
	public void setPaymentDate(Date paymentDate) {
		this.paymentDate = paymentDate;
	}
	public String getChequeNumber() {
		return chequeNumber;
	}
	public void setChequeNumber(String chequeNumber) {
		this.chequeNumber = chequeNumber;
	}
	public List<PaymentDetailsDTO> getExistingUsers() {
		return existingUsers;
	}
	public void setExistingUsers(List<PaymentDetailsDTO> existingUsers) {
		this.existingUsers = existingUsers;
	}
	public String getChkDetailsString() {
		return chkDetailsString;
	}
	public void setChkDetailsString(String chkDetailsString) {
		this.chkDetailsString = chkDetailsString;
	}
	
}
