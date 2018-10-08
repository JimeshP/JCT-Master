package com.vmware.jct.dao;

import java.util.List;

import com.vmware.jct.exception.DAOException;
import com.vmware.jct.exception.JCTException;
import com.vmware.jct.model.JctCompletionStatus;
import com.vmware.jct.model.JctStatusSearch;
import com.vmware.jct.model.JctUserLoginInfo;
/**
 * 
 * <p><b>Class name:</b>IConfirmationDAO.java</p>
 * <p><b>@author:</b> InterraIt </p>
 * <p><b>Purpose:</b> 		This DAO interface represents a IConfirmationDAO interface
 * <p><b>Description:</b> 	This will have all methods mapped to the database and fetching data from database to be used across the application </p>
 * <p><b>Copyrights:</b> 	All rights reserved by Interra IT and should only be used for its internal application development.</p>
 *  <pre>
 * <b>Revision History:</b>
 * </pre>
 */
public interface IConfirmationDAO {
	/**
	 * 
	 * @param jobReferenceNo
	 * @return
	 * @throws DAOException
	 */
	public String addDiagWithoutEdits(String jobReferenceNo) throws DAOException;
	/**
	 * 
	 * @param jobReferenceNo
	 * @return
	 * @throws DAOException
	 */
	public String addEdits(String jobReferenceNo) throws DAOException;
	/**
	 * 
	 * @param jobReferenceNo
	 * @return
	 * @throws DAOException
	 */
	public int disableStatus(String jobReferenceNo) throws DAOException;
	/**
	 * 
	 * @param jobReferenceNo
	 * @return
	 * @throws DAOException
	 */
	public int updateDeleteStatusForEdit(String jobReferenceNo) throws DAOException;
	/**
	 * 
	 * @param entity
	 * @return
	 * @throws DAOException
	 */
	public String saveOrUpdateBeforeSketch(JctCompletionStatus entity) throws DAOException;
	/**
	 * 
	 * @param entity
	 * @return
	 * @throws DAOException
	 */
	public String saveToStatusSearch(JctStatusSearch entity) throws DAOException;
	/**
	 * 
	 * @param jobReferenceNo
	 * @return
	 * @throws DAOException
	 */
	public int updateLinkStatus(String jobReferenceNo) throws DAOException;
	/**
	 * 
	 * @param jobReferenceNo
	 * @return
	 * @throws DAOException
	 */
	public int keepOriginal(String jobReferenceNo) throws DAOException;
	/**
	 * 
	 * @param jobReferenceNo
	 * @return
	 * @throws DAOException
	 */
	public int restartExcersize(String jobReferenceNo, String distinction) throws DAOException;
	/**
	 * 
	 * @param jobRefNo
	 * @return
	 * @throws DAOException
	 */
	public List<JctCompletionStatus> getExcersizeCompletionCount (String jobRefNo) throws DAOException;
	/**
	 * 
	 * @param jobRefNo
	 * @return
	 * @throws DAOException
	 */
	public void updateShowPDF(String jobRefNo) throws DAOException;
	
	public List<JctUserLoginInfo> getTotalTime(int userId) throws DAOException;
}
