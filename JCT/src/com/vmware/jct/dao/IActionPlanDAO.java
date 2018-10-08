package com.vmware.jct.dao;

import java.util.List;

import com.vmware.jct.exception.DAOException;
import com.vmware.jct.model.JctActionPlan;
/**
 * 
 * <p><b>Class name:</b>IActionPlanDAO.java</p>
 * <p><b>@author:</b> InterraIt </p>
 * <p><b>Purpose:</b> 		This DAO interface represents a IActionPlanDAO interface
 * <p><b>Description:</b> 	This will have all methods mapped to the database and fetching data from database to be used across the application </p>
 * <p><b>Note:</b> 			This methods implemented by DAOImpl classes and execute the mapped queries and get the records.</p>  
 * <p><b>Copyrights:</b> 	All rights reserved by Interra IT and should only be used for its internal application development.</p>
 *  <pre>
 * <b>Revision History:</b>
 * </pre>
 */
public interface IActionPlanDAO {
	/**
	 * @param It will take  action plan object as a input parameter
	 * @return It returns list of  String data from database
	 * @throws DAOException -it throws DAOException
	 */
	public String saveActinoPlans(JctActionPlan actionPlan) throws DAOException;
	/**
	 * @param It will take  action plan object as a input parameter
	 * @return It returns list of  String data from database
	 * @throws DAOException -it throws DAOException
	 */
	public String mergeActinoPlans(JctActionPlan actionPlan) throws DAOException;
	/**
	 * @param It will take  action plan object as a input parameter
	 * @return It returns list of  String data from database
	 * @throws DAOException -it throws DAOException
	 */
	public List<JctActionPlan> getList(String jobRefNo) throws DAOException;
	
	public List<JctActionPlan> getListEdited(String jobRefNo) throws DAOException;
	/**
	 * @param It will take  action plan object as a input parameter
	 * @return It returns list of  String data from database
	 * @throws DAOException -it throws DAOException
	 */
	public String remove(JctActionPlan object) throws DAOException;
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
	 * @param jobReferenceNos
	 * @return
	 * @throws DAOException
	 */
	public Long getNumberOfCompletion (String jobReferenceNos) throws DAOException;
	/**
	 * 
	 * @param jobReferenceNos
	 * @return
	 * @throws DAOException
	 */
	public Integer getLastCompletionStatus (String jobReferenceNos) throws DAOException;
	/**
	 * 
	 * @param jrNo
	 * @param user
	 * @param userId
	 * @param profileId
	 * @return
	 * @throws DAOException
	 */
	public Long getActionPlanCount(String jrNo, String user, int userId, int profileId) throws DAOException;
	
	public String getOnetTitle(String onetCode) throws DAOException;
}
