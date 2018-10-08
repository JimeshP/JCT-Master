package com.vmware.jct.dao;

import java.util.List;

import com.vmware.jct.dao.dto.SurveyQuestionDTO;
import com.vmware.jct.exception.DAOException;
import com.vmware.jct.model.JctSurveyQuestionMaster;
import com.vmware.jct.model.JctSurveyQuestions;
import com.vmware.jct.model.JctUser;
/**
 * 
 * <p><b>Class name:</b>ISurveyQuestionDAO.java</p>
 * <p><b>@author:</b> InterraIT </p>
 * <p><b>Purpose:</b> 		This DAO interface represents a ISurveyQuestionDAO interface
 * <p><b>Description:</b> 	This will have all methods mapped to the database and fetching data from database to be used for Manage user service. </p>
 * <p><b>Note:</b> 			This methods implemented by DAOImpl classes and execute the mapped queries and get the records.</p>  
 * <p><b>Copyrights:</b> 	All rights reserved by Interra IT and should only be used for its internal application development.</p>
 *  <pre>
 * <b>Revision History:</b>
 * <li> </li>
 * </pre>
 */
public interface ISurveyQuestionDAO {
	/**
	 * Method saves the free text
	 * @param surveyMaster
	 * @return
	 * @throws DAOException
	 */
	public String saveSurveyQuestion (JctSurveyQuestionMaster surveyMaster) throws DAOException;
	/**
	 * Method fetches all the survey questions
	 * @return
	 * @throws DAOException
	 */
	public List<SurveyQuestionDTO> getExistingAllSurveyQuestion() throws DAOException;
	/**
	 * Method fetches all the survey questions by user type
	 * @param userType
	 * @return
	 * @throws DAOException
	 */
	public List<SurveyQuestionDTO> fetchAllExistingSurveyQtnsByUserType(int userType) throws DAOException;
	/**
	 * Method fetches all the survey questions by user type and question type
	 * @param userType
	 * @param qtnType
	 * @return
	 * @throws DAOException
	 */
	public List<SurveyQuestionDTO> fetchAllExistingSurveyQtnsByUserTypeAndQtnType(int userType, int qtnType) throws DAOException;
	/**
	 * Method fetches all the survey questions by question type
	 * @param qtnType
	 * @return
	 * @throws DAOException
	 */
	public List<SurveyQuestionDTO> fetchAllExistingSurveyQtnsByQtnType(int qtnType) throws DAOException;
	/**
	 * Method fetches the list of all sub questions
	 * @param answerType
	 * @param userType
	 * @param mainQtns
	 * @return
	 * @throws DAOException
	 */
	public List<SurveyQuestionDTO> getAllSubQtns (int answerType, int userType, String mainQtns) throws DAOException;
	/**
	 * Method fetches the list of all sub questions
	 * @param answerType
	 * @param mainQtns
	 * @return
	 * @throws DAOException
	 */
	public List<SurveyQuestionDTO> getAllSubQtnsByAnsType (int answerType, String mainQtns) throws DAOException;
	/**
	 * Method gets the total number of sub questions
	 * @param answerType
	 * @param userType
	 * @param mainQtns
	 * @return
	 * @throws DAOException
	 */
	public Long getNumberOfSubQuestions (int answerType, int userType, String mainQtns) throws DAOException;
	/**
	 * Method gets the total number of sub questions
	 * @param answerType
	 * @param mainQtns
	 * @return
	 * @throws DAOException
	 */
	public Long getNumberOfSubQuestionsByAnsType (int answerType, String mainQtns) throws DAOException;
	/**
	 * Method soft deletes the entry
	 * @param userType
	 * @param answerType
	 * @param qtnDescription
	 * @return int
	 * @throws DAOException
	 */
	public int deleteSurveyQtn(int userType, int answerType, String qtnDescription) throws DAOException; 
	/**
	 * 
	 * @param userType
	 * @param answerType
	 * @param originalQuestion
	 * @return
	 * @throws DAOException
	 */
	public List<JctSurveyQuestionMaster> getSurveyQuestionList(int userType, int answerType, String originalQuestion) throws DAOException; 
	/**
	 * 
	 * @param surveyMaster
	 * @return
	 * @throws DAOException
	 */
	public String updateSurveyQuestion(JctSurveyQuestionMaster surveyMaster) throws DAOException;
	/**
	 * Method saves the survey questions and answers
	 * @param survey
	 * @return
	 * @throws DAOException
	 */
	public String saveSurveyQuestion(JctSurveyQuestions survey) throws DAOException;
	/**
	 * Method populates existing user data based on primary key
	 * @param primaryKey
	 * @return JctUser entity
	 * @throws DAOException
	 */
	public JctUser fetchUserData(Integer userId) throws DAOException;
	/**
	 * Method updates the first name
	 * @param user
	 * @return status
	 * @throws DAOException
	 */
	public String updateUser(JctUser user) throws DAOException;
	/**
	 * 
	 * @param userType
	 * @param qtnType
	 * @param profileId
	 * @return
	 * @throws DAOException
	 */
	public  List<SurveyQuestionDTO> fetchgSurveyQtnsByAllParams (int userType, int qtnType, int profileId) throws DAOException;
	/**
	 * 
	 * @param qtnType
	 * @param profileId
	 * @return
	 * @throws DAOException
	 */
	public List<SurveyQuestionDTO> fetchAllExistingSurveyQtnsOnlyUserProfile (int qtnType, int profileId) throws DAOException;
	/**
	 * 
	 * @param userType
	 * @param profileId
	 * @return
	 * @throws DAOException
	 */
	public List<SurveyQuestionDTO> fetchSurveyQtnsByProfileAndUserType (int userType, int profileId) throws DAOException;
	/**
	 * 
	 * @param answerType
	 * @param userType
	 * @param profileId
	 * @param mainQtns
	 * @return
	 * @throws DAOException
	 */
	public List<SurveyQuestionDTO> getAllSubQtnsByUserTypeAndProfile (int answerType, int userType, int profileId, String mainQtns) throws DAOException;
	/**
	 * 
	 * @param answerType
	 * @param userType
	 * @param mainQtns
	 * @param profileId
	 * @return
	 * @throws DAOException
	 */
	public Long getNumberOfSubQuestionsProfileId (int answerType, int userType, String mainQtns, int profileId) throws DAOException;
	/**
	 * 
	 * @param builder
	 * @param userProfile
	 * @param userType
	 * @return
	 * @throws DAOException
	 */
	public String saveSurvryQtnOrder(String builder, Integer userProfile, Integer userType, String builderAnsType) throws DAOException;
	/**
	 * 
	 * @param i
	 * @param userType
	 * @return
	 * @throws DAOException
	 */
	public int generateOrderSurvQtn(int profileId, int userType) throws DAOException;
	
}
