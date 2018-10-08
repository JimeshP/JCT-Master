package com.vmware.jct.service;

import com.vmware.jct.exception.JCTException;
/**
 * 
 * <p><b>Class name:</b>IConfirmationService.java</p>
 * <p><b>@author:</b> InterreIt </p>
 * <p><b>Purpose:</b> 		This Service interface represents a IConfirmationService interface
 * <p><b>Description:</b> 	This will have all methods where we implement all business logic and all dao layer </p>
 * <p><b>Copyrights:</b> 	All rights reserved by Interra IT and should only be used for its internal application development.</p>
 *  <pre>
 * <b>Revision History:</b>
 * </pre>
 */
public interface IConfirmationService {
	/**
	 * 
	 * @param jobReferenceNo
	 * @return string object
	 * @throws JCTException
	 */
	public String addDiagWithoutEdits(String jobReferenceNo) throws JCTException;
	/**
	 * 
	 * @param jobReferenceNo
	 * @return string object
	 * @throws JCTException
	 */
	public String addEdits(String jobReferenceNo) throws JCTException;
	/**
	 * 
	 * @param jobReferenceNo
	 * @return int object
	 * @throws JCTException
	 */
	public int disableStatus(String jobReferenceNo) throws JCTException;
	/**
	 * 
	 * @param jobReferenceNo
	 * @return int object
	 * @throws JCTException
	 */
	public String updatePrevious(String jobReferenceNo) throws JCTException;
	/**
	 * 
	 * @param jobReferenceNo
	 * @return
	 * @throws JCTException
	 */
	public int updateLinkStatus(String jobReferenceNo) throws JCTException;
	/**
	 * Method restarts the whole excersize all over again
	 * @param jobReferenceNo
	 * @return string object
	 * @throws JCTException
	 */
	public String restartExcersize(String jobReferenceNo, String distinction) throws JCTException;
	/**
	 * Method fetches the count of completion
	 * @param jobRefNo
	 * @return
	 * @throws JCTException
	 */
	public int getExcersizeCompletionCount (String jobRefNo) throws JCTException;
	
	public String getTotalTime(int userId) throws JCTException;
}
