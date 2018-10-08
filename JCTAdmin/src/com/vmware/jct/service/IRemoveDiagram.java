package com.vmware.jct.service;

import com.vmware.jct.exception.JCTException;
import com.vmware.jct.service.vo.RemoveDiagramVOList;
import com.vmware.jct.service.vo.UserVO;
/**
 * 
 * <p><b>Class name:</b>IRemoveDiagram.java</p>
 * <p><b>@author:</b> InterraIT </p>
 * <p><b>Purpose:</b> 		This Service interface represents a IRemoveDiagram interface
 * <p><b>Description:</b> 	This will have all methods where we implement all business logic and all service layer </p>
 * <p><b>Copyrights:</b> 	All rights reserved by InterraIT and should only be used for its internal application development.</p>
 *  <pre>
 * <b>Revision History:</b>
 * InterraIT - 06/Nov/2014 - Introduced the interface
 * </pre>
 */
public interface IRemoveDiagram {
	/**
	 * Method fetches the before sketch and after sketch diagrams based 
	 * on the email Id
	 * @param emailId
	 * @return
	 * @throws JCTException
	 */
	public RemoveDiagramVOList getDiagrams (String emailId) throws JCTException;
	/**
	 * Method soft deletes (marks 2) the before and after sketch diagram for the 
	 * corresponding row id.
	 * @param vo
	 * @return
	 * @throws JCTException
	 */
	public String removeDiagram (RemoveDiagramVOList vo) throws JCTException;
	/**
	 * Method fetches the before sketch and after sketch diagrams based 
	 * on the email Id
	 * @param emailId
	 * @return
	 * @throws JCTException
	 */
	public RemoveDiagramVOList populateAllDiagram (int recordIndex) throws JCTException;
	/**
	 * Method fetches the count of all the diagrams.
	 * @param null
	 * @return
	 * @throws JCTException
	 */
	public Long fetchAllDiagramCount() throws JCTException;
	/**
	 * Method fetches all user 
	 * on the email Id
	 * @param emailId
	 * @return
	 * @throws JCTException
	 */
	public UserVO searchUserForRefundRequest (String emailId) throws JCTException;
	
	/**
	 * Method fetches the before sketch and after sketch diagrams based 
	 * on the email Id
	 * @param emailId
	 * @return
	 * @throws JCTException
	 */
	public String disableUserForRefundRequest (int userId, String emailId) throws JCTException;
}
