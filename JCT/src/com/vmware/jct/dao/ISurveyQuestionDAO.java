package com.vmware.jct.dao;

import com.vmware.jct.exception.DAOException;
import com.vmware.jct.model.JctSurveyQuestions;
/**
 * 
 * <p><b>Class name:</b>ISurveyQuestionDAO.java</p>
 * <p><b>@author:</b> InterraIt </p>
 * <p><b>Purpose:</b> 		This DAO interface represents a ISurveyQuestionDAO interface
 * <p><b>Description:</b> 	This will have all methods mapped to the database and fetching data from database to be used across the application </p>
 * <p><b>Copyrights:</b> 	All rights reserved by Interra IT and should only be used for its internal application development.</p>
 *  <pre>
 * <b>Revision History:</b>
 * InterraIt - 14/Oct/2014 - Implemented the class
 * </pre>
 */
public interface ISurveyQuestionDAO {
	/**
	 * Method saves the survey questions and answers
	 * @param survey
	 * @return
	 * @throws DAOException
	 */
	public String saveSurveyQuestion(JctSurveyQuestions survey) throws DAOException;
}
