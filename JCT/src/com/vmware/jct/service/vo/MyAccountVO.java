package com.vmware.jct.service.vo;

import java.util.Date;

/**
 * 
 * <p><b>Class name:</b> MyAccountVO.java</p>
 * <p><b>Author:</b> InterraIT</p>
 * <p><b>Purpose:</b> This class acts as a data transfer object.
 * <p><b>Description:</b> This class acts as a data transfer object.. </p>
 * <p><b>Copyrights:</b> 	All rights reserved by Interra IT and should only be used for its internal application development.</p>
 * <p><b>Created Date:</b> 20/Feb/2014</p>
 * <p><b>Revision History:</b>
 * 	<li>6/May/2014: Changed the saving of action plan to dynamic.</li>
 * </p>
 */
public class MyAccountVO {
	private String firstName;
	private String lastName;
	private String region;
	private String sex;
	private String tenure;
	private String supervisePeople;
	private String functionGroup;
	private String jobLevel;
	private int statusCode;
	private String status;
	private Date jctAccountExpirationDate;
	private String customerId;
	private String onetDesc;
	
	private String faciFirstName;
	private String faciLastName;
	private String faciEmail;
	
	public MyAccountVO () {
		
	}
	
	public MyAccountVO (String firstName, String lastName, String region, String sex, 
			String tenure, String supervisePeople, String functionGroup, String jobLevel) {
		this.firstName = firstName;
		this.lastName = lastName;
		this.region = region;
		this.sex = sex;
		this.tenure = tenure;
		this.supervisePeople = supervisePeople;
		this.functionGroup = functionGroup;
		this.jobLevel = jobLevel;
	}
	
	public MyAccountVO (String faciFirstName, String faciLastName, String faciEmail) {
		this.faciFirstName = faciFirstName;
		this.faciLastName = faciLastName;
		this.faciEmail = faciEmail;
	}

	/**
	 * @return the firstName
	 */
	public String getFirstName() {
		return firstName;
	}

	/**
	 * @param firstName the firstName to set
	 */
	public void setFirstName(String firstName) {
		this.firstName = firstName;
	}

	/**
	 * @return the lastName
	 */
	public String getLastName() {
		return lastName;
	}

	/**
	 * @param lastName the lastName to set
	 */
	public void setLastName(String lastName) {
		this.lastName = lastName;
	}

	/**
	 * @return the region
	 */
	public String getRegion() {
		return region;
	}

	/**
	 * @param region the region to set
	 */
	public void setRegion(String region) {
		this.region = region;
	}

	/**
	 * @return the sex
	 */
	public String getSex() {
		return sex;
	}

	/**
	 * @param sex the sex to set
	 */
	public void setSex(String sex) {
		this.sex = sex;
	}

	/**
	 * @return the tenure
	 */
	public String getTenure() {
		return tenure;
	}

	/**
	 * @param tenure the tenure to set
	 */
	public void setTenure(String tenure) {
		this.tenure = tenure;
	}

	/**
	 * @return the supervisePeople
	 */
	public String getSupervisePeople() {
		return supervisePeople;
	}

	/**
	 * @param supervisePeople the supervisePeople to set
	 */
	public void setSupervisePeople(String supervisePeople) {
		this.supervisePeople = supervisePeople;
	}

	/**
	 * @return the functionGroup
	 */
	public String getFunctionGroup() {
		return functionGroup;
	}

	/**
	 * @param functionGroup the functionGroup to set
	 */
	public void setFunctionGroup(String functionGroup) {
		this.functionGroup = functionGroup;
	}

	/**
	 * @return the jobLevel
	 */
	public String getJobLevel() {
		return jobLevel;
	}

	/**
	 * @param jobLevel the jobLevel to set
	 */
	public void setJobLevel(String jobLevel) {
		this.jobLevel = jobLevel;
	}

	/**
	 * @return the statusCode
	 */
	public int getStatusCode() {
		return statusCode;
	}

	/**
	 * @param statusCode the statusCode to set
	 */
	public void setStatusCode(int statusCode) {
		this.statusCode = statusCode;
	}

	/**
	 * @return the status
	 */
	public String getStatus() {
		return status;
	}

	/**
	 * @param status the status to set
	 */
	public void setStatus(String status) {
		this.status = status;
	}

	public Date getJctAccountExpirationDate() {
		return jctAccountExpirationDate;
	}

	public void setJctAccountExpirationDate(Date jctAccountExpirationDate) {
		this.jctAccountExpirationDate = jctAccountExpirationDate;
	}

	public String getCustomerId() {
		return customerId;
	}

	public void setCustomerId(String customerId) {
		this.customerId = customerId;
	}

	public String getOnetDesc() {
		return onetDesc;
	}

	public void setOnetDesc(String onetDesc) {
		this.onetDesc = onetDesc;
	}

	public String getFaciFirstName() {
		return faciFirstName;
	}
	
	public void setFaciFirstName(String faciFirstName) {
		this.faciFirstName = faciFirstName;
	}
	
	public String getFaciLastName() {
		return faciLastName;
	}

	public void setFaciLastName(String faciLastName) {
		this.faciLastName = faciLastName;
	}

	

	public String getFaciEmail() {
		return faciEmail;
	}

	public void setFaciEmail(String faciEmail) {
		this.faciEmail = faciEmail;
	}
	
}
