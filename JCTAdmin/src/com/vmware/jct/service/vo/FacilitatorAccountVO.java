package com.vmware.jct.service.vo;

public class FacilitatorAccountVO {
	// GENERAL
	private String createdBy;
	private String paymentType;
	// CHEQUE PAYMENT
	private String userGroup;
	//private String bankDetails;
	private String facilitatorEmail;
	private String paymentDate;
	private String numberOfUsers;
	private String chequeNos;
	private String totalAmountToBePaid;
	private String facToHaveAccount;
	private String facilitatorID;
	private String facilitatorFirstName;
	private String facilitatorLastName;
	
	public FacilitatorAccountVO(String userGroup,
			String facilitatorEmail, String paymentDate, String numberOfUsers,
			String chequeNos, String totalAmountToBePaid,
			String facToHaveAccount, String createdBy, String paymentType) {
		this.userGroup = userGroup;
		//this.bankDetails = bankDetails;
		this.facilitatorEmail = facilitatorEmail;
		this.paymentDate = paymentDate;
		this.numberOfUsers = numberOfUsers;
		this.chequeNos = chequeNos;
		this.totalAmountToBePaid = totalAmountToBePaid;
		this.facToHaveAccount = facToHaveAccount;
		this.createdBy = createdBy;
		this.paymentType = paymentType;
	}
	public FacilitatorAccountVO(String createdBy,String facilitatorID,String facilitatorEmail,String paymentDate, String numberOfUsers,
			String chequeNos, String totalAmountToBePaid, String paymentType) {
		this.createdBy = createdBy;
		this.facilitatorID = facilitatorID;  
		this.facilitatorEmail = facilitatorEmail;
		//this.bankDetails = bankDetails;		
		this.paymentDate = paymentDate;
		this.numberOfUsers = numberOfUsers;
		this.chequeNos = chequeNos;
		this.totalAmountToBePaid = totalAmountToBePaid;	
		this.paymentType = paymentType;
	}
	public FacilitatorAccountVO(String createdBy,String facilitatorID,String facilitatorEmail,String numberOfUsers,
			String totalAmountToBePaid, String paymentType) {
		this.createdBy = createdBy;
		this.facilitatorID = facilitatorID;  
		this.facilitatorEmail = facilitatorEmail;
		this.numberOfUsers = numberOfUsers;
		this.totalAmountToBePaid = totalAmountToBePaid;	
		this.paymentType = paymentType;
	}
	
	public FacilitatorAccountVO(String userGroup,
			String facilitatorEmail, String paymentDate, String numberOfUsers,
			String chequeNos, String totalAmountToBePaid,
			String facToHaveAccount, String createdBy, String paymentType, String facilitatorFirstName, String facilitatorLastName) {
		this.userGroup = userGroup;
		//this.bankDetails = bankDetails;
		this.facilitatorEmail = facilitatorEmail;
		this.paymentDate = paymentDate;
		this.numberOfUsers = numberOfUsers;
		this.chequeNos = chequeNos;
		this.totalAmountToBePaid = totalAmountToBePaid;
		this.facToHaveAccount = facToHaveAccount;
		this.createdBy = createdBy;
		this.paymentType = paymentType;
		this.facilitatorFirstName = facilitatorFirstName;
		this.facilitatorLastName = facilitatorLastName;
	}
	public String getUserGroup() {
		return userGroup;
	}
	public String getFacilitatorID() {
		return facilitatorID;
	}
	public void setFacilitatorID(String facilitatorID) {
		this.facilitatorID = facilitatorID;
	}
	public void setUserGroup(String userGroup) {
		this.userGroup = userGroup;
	}
/*	public String getBankDetails() {
		return bankDetails;
	}
	public void setBankDetails(String bankDetails) {
		this.bankDetails = bankDetails;
	}*/
	public String getFacilitatorEmail() {
		return facilitatorEmail;
	}
	public void setFacilitatorEmail(String facilitatorEmail) {
		this.facilitatorEmail = facilitatorEmail;
	}
	public String getPaymentDate() {
		return paymentDate;
	}
	public void setPaymentDate(String paymentDate) {
		this.paymentDate = paymentDate;
	}
	public String getNumberOfUsers() {
		return numberOfUsers;
	}
	public void setNumberOfUsers(String numberOfUsers) {
		this.numberOfUsers = numberOfUsers;
	}
	public String getChequeNos() {
		return chequeNos;
	}
	public void setChequeNos(String chequeNos) {
		this.chequeNos = chequeNos;
	}
	public String getTotalAmountToBePaid() {
		return totalAmountToBePaid;
	}
	public void setTotalAmountToBePaid(String totalAmountToBePaid) {
		this.totalAmountToBePaid = totalAmountToBePaid;
	}
	public String getFacToHaveAccount() {
		return facToHaveAccount;
	}
	public void setFacToHaveAccount(String facToHaveAccount) {
		this.facToHaveAccount = facToHaveAccount;
	}
	public String getCreatedBy() {
		return createdBy;
	}
	public void setCreatedBy(String createdBy) {
		this.createdBy = createdBy;
	}
	public String getPaymentType() {
		return paymentType;
	}
	public void setPaymentType(String paymentType) {
		this.paymentType = paymentType;
	}
	public String getFacilitatorFirstName() {
		return facilitatorFirstName;
	}
	public void setFacilitatorFirstName(String facilitatorFirstName) {
		this.facilitatorFirstName = facilitatorFirstName;
	}
	public String getFacilitatorLastName() {
		return facilitatorLastName;
	}
	public void setFacilitatorLastName(String facilitatorLastName) {
		this.facilitatorLastName = facilitatorLastName;
	}
}
