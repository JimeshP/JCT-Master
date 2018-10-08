package com.vmware.jct.dao;

import java.util.List;

import com.vmware.jct.dao.dto.SearchPdfDTO;
import com.vmware.jct.dao.dto.StatusSearchDTO;
import com.vmware.jct.exception.DAOException;
import com.vmware.jct.model.JctAfterSketchHeader;
import com.vmware.jct.model.JctBeforeSketchHeader;
import com.vmware.jct.model.JctCompletionStatus;

/**
 * 
 * <p><b>Class name:</b>ISketchSearchDAO.java</p>
 * <p><b>Author:</b> InterraIt </p>
 * <p><b>Purpose:</b> 		This DAO interface represents a ISketchSearchDAO interface
 * <p><b>Description:</b> 	This will have all methods to search the sketch Diagram from database </p>
 * <p><b>Copyrights:</b> 	All rights reserved by Interra IT and should only be used for its internal application development.</p>
 * <p><b>Created Date:</b> 05/Feb/2014</p>
 * <b>Revision History:</b>
 * <li></li>
 */
public interface ISketchSearchDAO {
	/**
	 * 
	 * @param jobTitle
	 * @param firstResult
	 * @param maxResults
	 * @return
	 * @throws DAOException
	 */
	public List<StatusSearchDTO> getBeforeSketchList(String jobTitle, int firstResult, int maxResults) throws DAOException;
	/**
	 * 
	 * @param jobTitle
	 * @param firstResult
	 * @param maxResults
	 * @return
	 * @throws DAOException
	 */
	public List<StatusSearchDTO> getAfterSketchList(String jobTitle, int firstResult, int maxResults) throws DAOException;
	/**
	 * 
	 * @param snapShotString
	 * @param jobTitle
	 * @return Job Reference No
	 * @throws DAOException
	 */
	public String getBeforeSketchJrno(String jrno, String snapShotString, String jobTitle) throws DAOException;
	
	/**
	 * 
	 * @param jobReference,jobTitle
	 * @return String
	 * @throws DAOException
	 */
	
	public String isASAvailable (String jobReference, String jobTitle) throws DAOException;
	/**
	 * 
	 * @param jobReference,jobTitle
	 * @return String
	 * @throws DAOException
	 */
	
	public String getASDiagram (String jobReference) throws DAOException;
	/**
	 * 
	 * @param jobReference,jobTitle
	 * @return String
	 * @throws DAOException
	 */
	
	public List<JctAfterSketchHeader> getPrevASDiagram (String jobReference) throws DAOException;
	/**
	 * 
	 * @param jobReference
	 * @return
	 * @throws DAOException
	 */
	public String getPrevASDiagramForView(String jobReference, int softDel) throws DAOException;
	/**
	 * 
	 * @param snapShotString,jobTitle
	 * @return String
	 * @throws DAOException
	 */
	public String getAfterSketchJrno(String jrno, String snapShotString, String jobTitle) throws DAOException;
	/**
	 * 
	 * @param jobReference,jobTitle
	 * @return String
	 * @throws DAOException
	 */
	public String isBSAvailable(String jobReference, String jobTitle) throws DAOException;
	/**
	 * 
	 * @param jobReference
	 * @return String
	 * @throws DAOException
	 */
	public String getBSDiagram (String jobReference) throws DAOException;
	/**
	 * 
	 * @param jobReference
	 * @return JctBeforeSketchHeader
	 * @throws DAOException
	 */
	public List<JctBeforeSketchHeader> getPrevBSDiagram (String jobReference) throws DAOException;
	/**
	 * 
	 * @param jobReference
	 * @return
	 * @throws DAOException
	 */
	public String getPrevBSDiagramForView(String jobReference, int softDel) throws DAOException;
	/**
	 * 
	 * @param jobReference
	 * @return String
	 * @throws DAOException
	 */
	public List<JctCompletionStatus> getCompletionStatusObjs (String jobReference) throws DAOException;
	/**
	 * 
	 * @param jobRefNo
	 * @return
	 * @throws DAOException
	 */
	public List<String> fetchNosOfTimesShared(String jobRefNo) throws DAOException;
	/**
	 * 
	 * @param jobTitle
	 * @return
	 * @throws DAOException
	 */	
	public String getOccupationDesc (String jobTitle) throws DAOException;
	
	public List<StatusSearchDTO> getBeforeSketchList(int firstResult, int maxResults) throws DAOException;
	
	public List<StatusSearchDTO> getBeforeSketchList(int firstResult, int maxResults, String code) throws DAOException;
	
	public List<StatusSearchDTO> getAfterSketchList(int firstResult, int maxResults) throws DAOException;
	
	public List<StatusSearchDTO> getAfterSketchList(int firstResult, int maxResults, String code) throws DAOException;
	
	public Long fetchAllDiagramCount(String distinction) throws DAOException;
	
	public Long fetchAllDiagramCount(String distinction, String code) throws DAOException;
	
	public List<SearchPdfDTO> fetchOldPdf(int userId) throws DAOException;
}
