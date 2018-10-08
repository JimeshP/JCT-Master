package com.vmware.jct.service.vo;

import java.util.List;

public class GeneralUserAccountVO {
	// GENERAL
	private String createdBy;
	private String paymentType;
	private String expiryDuration;
	// CHEQUE PAYMENT
	private String userGroup;
	//private String bankDetails;
	private List<String> validEmailList;
	private List<String> inValidEmailList;
	private String paymentDate;
	private String numberOfUsers;
	private String chequeNos;
	private String totalAmountToBePaid;
	private List<String> emailWithExpDateList;
	
	public GeneralUserAccountVO (String createdBy, String paymentType, String expiryDuration, String userGroup, 
			List<String> validEmailList, List<String> inValidEmailList, String paymentDate, 
			String numberOfUsers, String chequeNos, String totalAmountToBePaid) {
		this.createdBy 				= createdBy;
		this.paymentType 			= paymentType;
		this.expiryDuration 		= expiryDuration;
		this.userGroup 				= userGroup;
		//this.bankDetails 			= bankDetails;
		this.validEmailList 		= validEmailList;
		this.inValidEmailList 		= inValidEmailList;
		this.paymentDate 			= paymentDate;
		this.numberOfUsers 			= numberOfUsers;
		this.chequeNos 				= chequeNos;
		this.totalAmountToBePaid 	= totalAmountToBePaid;
	}
	
	public GeneralUserAccountVO (String createdBy, String paymentType,
			List<String> emailWithExpDateList, String paymentDate, 
			String numberOfUsers, String chequeNos, String totalAmountToBePaid) {
		this.createdBy 				= createdBy;
		this.paymentType 			= paymentType;
		this.emailWithExpDateList 	= emailWithExpDateList;
		this.paymentDate 			= paymentDate;
		this.numberOfUsers 			= numberOfUsers;
		this.chequeNos 				= chequeNos;
		this.totalAmountToBePaid 	= totalAmountToBePaid;
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

	public String getExpiryDuration() {
		return expiryDuration;
	}

	public void setExpiryDuration(String expiryDuration) {
		this.expiryDuration = expiryDuration;
	}

	public String getUserGroup() {
		return userGroup;
	}

	public void setUserGroup(String userGroup) {
		this.userGroup = userGroup;
	}

/*	public String getBankDetails() {
		return bankDetails;
	}

	public void setBankDetails(String bankDetails) {
		this.bankDetails = bankDetails;
	}
*/
	public List<String> getValidEmailList() {
		return validEmailList;
	}

	public void setValidEmailList(List<String> validEmailList) {
		this.validEmailList = validEmailList;
	}

	public List<String> getInValidEmailList() {
		return inValidEmailList;
	}

	public void setInValidEmailList(List<String> inValidEmailList) {
		this.inValidEmailList = inValidEmailList;
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

	public List<String> getEmailWithExpDateList() {
		return emailWithExpDateList;
	}

	public void setEmailWithExpDateList(List<String> emailWithExpDateList) {
		this.emailWithExpDateList = emailWithExpDateList;
	}
}
