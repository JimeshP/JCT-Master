package com.vmware.jct.dao.dto;

import java.util.Date;
/**
 * 
 * <p><b>Class name:</b>AccountExpiryDTO.java</p>
 * <p><b>Author:</b>  InterrIT</p>
 * <p><b>Purpose:</b>AccountExpiryDTO provides access to account expiry data which are fetched into them.</p>
 * <p><b>Description:</b></p>
 * <pre>
 * 		<b>Revision History:</b>
 * 		
 * </pre>
 * <p><b></b></p>
 */
public class AccountExpiryDTO {
	private String userName;
	private Date expiryDate;
	
	public AccountExpiryDTO (String userName, Date expiryDate) {
		this.userName = userName;
		this.expiryDate = expiryDate;
	}

	public String getUserName() {
		return userName;
	}

	public void setUserName(String userName) {
		this.userName = userName;
	}

	public Date getExpiryDate() {
		return expiryDate;
	}

	public void setExpiryDate(Date expiryDate) {
		this.expiryDate = expiryDate;
	}
}
