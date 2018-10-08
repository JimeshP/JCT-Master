package com.vmware.jct.service;

import javax.servlet.http.HttpServletRequest;

import com.vmware.jct.exception.JCTException;
/**
 * 
 * <p><b>Class name:</b>IGeneratePDFService.java</p>
 * <p><b>@author:</b> InterraIt </p>
 * <p><b>Purpose:</b> 		This Service interface represents a IGeneratePDFService interface
 * <p><b>Description:</b> 	This will have all methods where we implement all business logic and all dao layer </p>
 * <p><b>Copyrights:</b> 	All rights reserved by Interra IT and should only be used for its internal application development.</p>
 *  <pre>
 * <b>Revision History:</b>
 * </pre>
 */
public interface IGeneratePDFService {
	/**
	 * 
	 * @param jobReferenceNumber
	 * @param emailId
	 * @param diskPath
	 * @return string Obj
	 * @throws JCTException
	 */
	public String generatePdf(String jobReferenceNumber, String emailId, String diskPath, HttpServletRequest request, int jctUserId) throws JCTException;
	/**
	 * 
	 * @param jobReferenceNo
	 * @return string Obj
	 * @throws JCTException
	 */
	public String getPathAndFile(String jobReferenceNo) throws JCTException;
}
