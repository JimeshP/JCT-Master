package com.vmware.jct.service;

import com.vmware.jct.exception.JCTException;
import com.vmware.jct.service.vo.MyAccountVO;
/**
 * 
 * <p><b>Class name:</b>IFacilitatorAccountService.java</p>
 * <p><b>@author:</b> InterraIT </p>
 * <p><b>Purpose:</b> 		This Service interface represents a IFacilitatorAccountService interface
 * <p><b>Description:</b> 	This will have all methods where we implement all business logic and all service layer </p>
 * <p><b>Copyrights:</b> 	All rights reserved by InterraIT and should only be used for its internal application development.</p>
 *  <pre>
 * <b>Revision History:</b>
 * InterraIT - 01/Oct/2014 - Implement exception through out the application
 * </pre>
 */
public interface IFacilitatorAccountService {
	/**
	 * Method populates facilitator data
	 * @param facilitatorEmail
	 * @return MyAccountVO
	 * @throws JCTException
	 */
	public MyAccountVO populateExistingFacilitatorData(String facilitatorEmail) throws JCTException;
	/**
	 * Method updates the user name of the facilitator
	 * @param valueDesc
	 * @param tablePkId
	 * @return MyAccountVO
	 * @throws JCTException
	 */
	public MyAccountVO updateUserNameFacilitator(String valueDesc, Integer tablePkId, String facilitatorEmail) throws JCTException;
	/**
	 * Method updates the email address of the facilitator
	 * @param valueDesc
	 * @param tablePkId
	 * @return MyAccountVO
	 * @throws JCTException
	 */
	public MyAccountVO updateEmailIdFacilitator(String valueDesc, Integer tablePkId, String facilitatorEmail) throws JCTException;
	/**
	 * Method updates the password of the facilitator
	 * @param valueDesc
	 * @param tablePkId
	 * @return MyAccountVO
	 * @throws JCTException
	 */
	public MyAccountVO updatePasswordFacilitator(String valueDesc, Integer tablePkId, String facilitatorEmail) throws JCTException;	
	/**
	 * Method search the user id is exist in database or not
	 * @param facilitatorEmail
	 * @return string
	 * @throws JCTException
	 */
	public String searchUserIdInDB(String facilitatorEmail) throws JCTException;
}
