package com.vmware.jct.service;

import java.util.List;
import java.util.Map;
import com.vmware.jct.exception.JCTException;

/**
 * 
 * <p><b>Class name:</b>IAccountExpiryNotificationService.java</p>
 * <p><b>@author:</b> InterraIT </p>
 * <p><b>Purpose:</b> 		This Service interface represents a IAccountExpiryNotificationService interface
 * <p><b>Description:</b> 	This will have all methods where we implement all business logic and all service layer </p>
 * <p><b>Copyrights:</b> 	All rights reserved by InterraIT and should only be used for its internal application development.</p>
 *  <pre>
 * <b>Revision History:</b>
 * <li> InterraIT - 20/Oct/2014 - Introduced the class </li>
 * </pre>
 */
public interface IAccountExpiryNotificationService {
	/**
	 * Method fetches all the expiry details..
	 * @return Map<Email, List of all would be expiring account>
	 * @throws JCTException
	 */
	public Map<String, List<String>> getAcntExpryNtfcatnForFacilitator() throws JCTException;
}
