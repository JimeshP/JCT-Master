package com.vmware.jct.service;

import com.vmware.jct.exception.JCTException;

/**
 * 
 * <p><b>Class name:</b>IAdminMenuService.java</p>
 * <p><b>@author:</b> InterraIT </p>
 * <p><b>Purpose:</b> 		This Service interface represents a IAdminMenuService interface
 * <p><b>Description:</b> 	This will have all methods where we implement all business logic and all service layer </p>
 * <p><b>Copyrights:</b> 	All rights reserved by InterraIT and should only be used for its internal application development.</p>
 *  <pre>
 * <b>Revision History:</b>
 * 
 * </pre>
 */
public interface IAdminMenuService {
	/**
	 * 
	 * @param userType
	 * @return
	 * @throws JCTException
	 */
	public String getMenuString ( int roleId ) throws JCTException;
}
