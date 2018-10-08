package com.vmware.jct.service;

import com.vmware.jct.exception.JCTException;
import com.vmware.jct.service.vo.AppSettingsVO;
/**
 * 
 * <p><b>Class name:</b>IAppSettingsService.java</p>
 * <p><b>@author:</b> InterraIT </p>
 * <p><b>Purpose:</b> 		This Service interface represents a IAppSettingsService interface
 * <p><b>Description:</b> 	This will have all methods where we implement all business logic and all service layer </p>
 * <p><b>Copyrights:</b> 	All rights reserved by InterraIT and should only be used for its internal application development.</p>
 *  <pre>
 * <b>Revision History:</b>
 * 
 * </pre>
 */
public interface IAppSettingsService {
	/**
	 * Method fetched the active application settings colors
	 * @param settingsVO
	 * @return
	 * @throws JCTException
	 */
	public AppSettingsVO fetchColor(AppSettingsVO settingsVO) throws JCTException;
	/**
	 * Method saves a new white lebel
	 * @param settingsVO
	 * @param jctEmail
	 * @return
	 * @throws JCTException
	 */
	public String saveNewWhiteLebel(AppSettingsVO settingsVO, String jctEmail) throws JCTException;
}
