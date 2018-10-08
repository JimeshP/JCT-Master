package com.vmware.jct.dao;

import java.util.List;

import com.vmware.jct.dao.dto.RemoveDiagramDTO;
import com.vmware.jct.exception.DAOException;
import com.vmware.jct.exception.JCTException;
import com.vmware.jct.model.JctUser;
import com.vmware.jct.service.vo.RemoveDiagramVOList;
/**
 * 
 * <p><b>Class name:</b>IRemoveDiagramDAO.java</p>
 * <p><b>@author:</b> InterraIT </p>
 * <p><b>Purpose:</b> 		This DAO interface represents a IRemoveDiagramDAO interface
 * <p><b>Description:</b> 	This will have all methods mapped to the database and fetching data from database to be used for report service. </p>
 * <p><b>Note:</b> 			This methods implemented by DAOImpl classes and execute the mapped queries and get the records.</p>  
 * <p><b>Copyrights:</b> 	All rights reserved by InterraIT and should only be used for its internal application development.</p>
 *  <pre>
 * <b>Revision History:</b>
 * <li></li>
 * </pre>
 */
public interface IRemoveDiagramDAO {
	/**
	 * Method searches for before sketch and after sketch shared diagrams
	 * @param emailId
	 * @return
	 * @throws DAOException
	 */
	public List<RemoveDiagramDTO> getDiagrams(String emailId) throws DAOException; 
	/**
	 * Method soft deletes (marks 2) the before and after sketch diagram for the 
	 * corresponding row id.
	 * @param vo
	 * @return
	 * @throws JCTException
	 */
	public int removeDiagram(RemoveDiagramVOList vo) throws DAOException;
	/**
	 * Method searches for all before sketch and after sketch shared diagrams
	 * @param emailId
	 * @return
	 * @throws DAOException
	 */
	public List<RemoveDiagramDTO> populateAllDiagram(int recordIndex) throws DAOException; 
	/**
	 * Method fetch user name against job ref no
	 * @param emailId
	 * @return
	 * @throws DAOException
	 */
	public String fetchUserName(String jobRefNo) throws DAOException;
	/**
	 * Method fetches the count of all the diagrams.
	 * @param null
	 * @return
	 * @throws DAOException
	 */
	public Long fetchAllDiagramCount() throws DAOException;	
	
	/**
	 * Method fetches the user object
	 * @param null
	 * @return
	 * @throws DAOException
	 */
	public JctUser searchUserForRefundRequest(String emailId) throws DAOException;
	
	/**
	 * Method disable an user using userId for refund request
	 * @param userId
	 * @return
	 * @throws DAOException
	 */
	public boolean disableUserForRefundRequestByUserId(int userId)throws DAOException;	
}
