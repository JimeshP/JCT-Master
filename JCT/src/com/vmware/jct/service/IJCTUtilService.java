package com.vmware.jct.service;

import com.vmware.jct.exception.JCTException;
/**
 * 
 * <p><b>Class name:</b>IJCTUtilService.java</p>
 * <p><b>@author:</b> InterraIt </p>
 * <p><b>Purpose:</b> 		This Service interface represents a IJCTUtilService interface
 * <p><b>Description:</b> 	This will have all methods where we implement all business logic and all dao layer </p>
 * <p><b>Copyrights:</b> 	All rights reserved by Interra IT and should only be used for its internal application development.</p>
 *  <pre>
 * <b>Revision History:</b>
 * </pre>
 */
public interface IJCTUtilService {

	/**
	 * 	
	 * @param keyName
	 * @return Integer
	 * @throws WorkFlowException
	 */
	public Integer generateKey(String keyName) throws JCTException;
	
}
