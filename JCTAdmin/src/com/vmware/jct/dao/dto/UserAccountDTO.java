package com.vmware.jct.dao.dto;

import java.io.Serializable;
import java.util.Date;

public class UserAccountDTO implements Serializable {
	private static final long serialVersionUID = 1L;
	private Date jctAccountExpirationDate;
	private String customerId;
	
	public UserAccountDTO(Date jctAccountExpirationDate , String customerId) {
		this.jctAccountExpirationDate = jctAccountExpirationDate;
		this.customerId = customerId;
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
	
}
