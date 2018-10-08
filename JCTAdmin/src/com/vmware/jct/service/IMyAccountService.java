package com.vmware.jct.service;

import com.vmware.jct.exception.JCTException;
import com.vmware.jct.service.vo.MyAccountVO;
/**
 * 
 * <p><b>Class name:</b>IMyAccountService.java</p>
 * <p><b>@author:</b> InterraIT </p>
 * <p><b>Purpose:</b> 		This Service interface represents a IMyAccountService interface
 * <p><b>Description:</b> 	This will have all methods where we implement all business logic and all service layer </p>
 * <p><b>Copyrights:</b> 	All rights reserved by InterraIT and should only be used for its internal application development.</p>
 *  <pre>
 * <b>Revision History:</b>
 * InterraIT - 31/Jan/2014 - Implement exception through out the application
 * </pre>
 */
public interface IMyAccountService {
	/**
	 * Method populates existing user data
	 * @param userId
	 * @return MyAccountVO
	 * @throws JCTException
	 */
	public MyAccountVO populateExistingUserData() throws JCTException;
	/**
	 * Method updates the first name of the administrator only
	 * @param valueDesc
	 * @param tablePkId
	 * @return MyAccountVO
	 * @throws JCTException
	 */
	public MyAccountVO updateFirstName(String valueDesc, Integer tablePkId) throws JCTException;
	/**
	 * Method updates the last name of the administrator only
	 * @param valueDesc
	 * @param tablePkId
	 * @return MyAccountVO
	 * @throws JCTException
	 */
	public MyAccountVO updateLastName(String valueDesc, Integer tablePkId) throws JCTException;
	/**
	 * Method updates the email id of the administrator only
	 * @param valueDesc
	 * @param tablePkId
	 * @return MyAccountVO
	 * @throws JCTException
	 */
	public MyAccountVO updateEmailId(String valueDesc, Integer tablePkId) throws JCTException;
	/**
	 * Method updates the password of the administrator only
	 * @param valueDesc
	 * @param tablePkId
	 * @return MyAccountVO
	 * @throws JCTException
	 */
	public MyAccountVO updatePassword(String valueDesc, Integer tablePkId) throws JCTException;
}
