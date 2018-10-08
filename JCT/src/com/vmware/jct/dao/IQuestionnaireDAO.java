package com.vmware.jct.dao;

import java.util.Date;
import java.util.List;

import com.vmware.jct.dao.dto.JobAttributeDTO;
import com.vmware.jct.exception.DAOException;
import com.vmware.jct.model.JctBeforeSketchQuestion;
import com.vmware.jct.model.JctGlobalProfileHistory;
import com.vmware.jct.model.JctJobAttributeFrozen;
/**
 * 
 * <p><b>Class name:</b>IQuestionnaireDAO.java</p>
 * <p><b>@author:</b> InterraIt </p>
 * <p><b>Purpose:</b> 		This DAO interface represents a IQuestionnaireDAO interface
 * <p><b>Description:</b> 	This will have all methods mapped to the database and fetching data from database to be used across the application </p>
 * <p><b>Copyrights:</b> 	All rights reserved by Interra IT and should only be used for its internal application development.</p>
 *  <pre>
 * <b>Revision History:</b>
 * InterraIt - 31/Jan/2014 - Implement comments through out the application
 * </pre>
 */


public interface IQuestionnaireDAO {
	/**
	 * @param It will take JctBeforeSketchQuestion object as a input parameter
	 * @return It will returns string object
	 * @throws DAOException -it throws DAOException
	 */
	public String saveAnswers(JctBeforeSketchQuestion question) throws DAOException;
	/**
	 * @param It will take JctBeforeSketchQuestion object as a input parameter
	 * @return It will returns string object
	 * @throws DAOException -it throws DAOException
	 */
	public String mergeAnswers(JctBeforeSketchQuestion question) throws DAOException;
	/**
	 * @param It will take JctBeforeSketchQuestion object as a input parameter
	 * @return It will returns JctBeforeSketchQuestion object
	 * @throws DAOException -it throws DAOException
	 */
	public List<JctBeforeSketchQuestion> getList(String jobRefNumber) throws DAOException;
	/**
	 * @param : profileId
	 * @return It will returns JobAttributeDTO object
	 * @throws DAOException -it throws DAOException
	 */
	public List<JobAttributeDTO> getFetchJobAttribute(int profileId) throws DAOException;
	/**
	 * @param : It will take JctBeforeSketchQuestion object as a input parameter
	 * @return It will returns String object
	 * @throws DAOException -it throws DAOException
	 */
	public String remove(JctBeforeSketchQuestion question) throws DAOException;
	/**
	 * @param It will take  job reference number as a String input parameter
	 * @return It returns list of  String data from database
	 * @throws DAOException -it throws DAOException
	 */
	public List<String> getDistinctQuestionList(String jobRefNo) throws DAOException;
	/**
	 * @param It will take mainQuestionand job reference number as a String input parameter
	 * @return It returns list of  String data from database
	 * @throws DAOException -it throws DAOException
	 */
	public List<String> getDistinctSubQuestionListByJrno(String mainQuestion, String jobRefNo) throws DAOException;
	/**
	 * @param It will take mainQuestion,subQuestion and job reference number as a String input parameter
	 * @return It returns list of  String data from database
	 * @throws DAOException -it throws DAOException
	 */
	public List<String> getAnswerListByJrno(String mainQuestion, String subQuestion, String jobRefNo) throws DAOException;
	/**
	 * 
	 * @param It will take profileId   as a input parameter
	 * @return string list
	 * @throws DAOException
	 */
	public List<String> getQuestionnaireMainQtnList (int profileId) throws DAOException;
	/**
	 * 
	 * @param It will take mainQtn and profileId   as a input parameter
	 * @return string list
	 * @throws DAOException
	 */
	public List<String> getQuestionnaireSubQtnList (String mainQtn, int profileId) throws DAOException;
	/**
	 * Method fetches all the existing forzen job attributes
	 * @param userId
	 * @param createdFor
	 * @param profileId
	 * @return
	 * @throws DAOException
	 */
	public Long getExistingFrozenAttrsCount (int userId, String createdFor, int profileId) throws DAOException;
	/**
	 * Method fetches count in questionnaire table
	 * @param jrNo
	 * @param user
	 * @param userId
	 * @param profileId
	 * @return
	 * @throws DAOException
	 */
	public Long getQuestCount (String jrNo, String user, int userId, int profileId) throws DAOException;
	/**
	 * 
	 * @param headerDate
	 * @param subQtn
	 * @param mainQtn
	 * @return
	 * @throws DAOException
	 */
	public List<JctGlobalProfileHistory> getGlobalProfileChangedData(Date headerDate, String subQtn, String mainQtn, String diff) throws DAOException;
}
