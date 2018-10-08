package com.vmware.jct.service.vo;

import java.util.List;

public class surveyQtnReportVO {
	// non repeating fields
	String flName;
	String userName;
	String profileName;
	String groupName;
	String custId;
	String userType;
	String createdBy;
	String expirationDate;
	List<surveyQtnRedundentListVO> qtnList;	// repeating Qtn-Ans fields
	public String getName() {
		return flName;
	}
	public void setName(String flName) {
		this.flName = flName;
	}
	public String getUserName() {
		return userName;
	}
	public void setUserName(String userName) {
		this.userName = userName;
	}
	public String getProfileName() {
		return profileName;
	}
	public void setProfileName(String profileName) {
		this.profileName = profileName;
	}
	public String getGroupName() {
		return groupName;
	}
	public void setGroupName(String groupName) {
		this.groupName = groupName;
	}
	public String getCustId() {
		return custId;
	}
	public void setCustId(String custId) {
		this.custId = custId;
	}
	public String getUserType() {
		return userType;
	}
	public void setUserType(String userType) {
		this.userType = userType;
	}
	public String getCreatedBy() {
		return createdBy;
	}
	public void setCreatedBy(String createdBy) {
		this.createdBy = createdBy;
	}
	public String getExpirationDate() {
		return expirationDate;
	}
	public void setExpirationDate(String expirationDate) {
		this.expirationDate = expirationDate;
	}
	public List<surveyQtnRedundentListVO> getQtnList() {
		return qtnList;
	}
	public void setQtnList(List<surveyQtnRedundentListVO> qtnList) {
		this.qtnList = qtnList;
	}
}