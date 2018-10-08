package com.vmware.jct.service;

import com.vmware.jct.exception.JCTException;
import com.vmware.jct.service.vo.SurveyQuestionsVO;
/**
 * 
 * <p><b>Class name:</b>ISurveyQuestionService.java</p>
 * <p><b>@author:</b> InterraIT </p>
 * <p><b>Purpose:</b> 		This Service interface represents a IMyISurveyQuestionServiceAccountService interface
 * <p><b>Description:</b> 	This will have all methods where we implement all business logic and all service layer </p>
 * <p><b>Copyrights:</b> 	All rights reserved by InterraIT and should only be used for its internal application development.</p>
 *  <pre>
 * <b>Revision History:</b>
 * <li></li>
 * </pre>
 */
public interface ISurveyQuestionService {
	/**
	 * Method saves the free text
	 * @param surveyQtnsVO
	 * @return
	 * @throws JCTException
	 */
	public SurveyQuestionsVO saveFreeText (SurveyQuestionsVO surveyQtnsVO) throws JCTException;
	/**
	 * Method saves radio related survey question
	 * @param surveyQtnsVO
	 * @return
	 * @throws JCTException
	 */
	public SurveyQuestionsVO saveRadio (SurveyQuestionsVO surveyQtnsVO) throws JCTException;
	/**
	 * Method saves drop down related survey question
	 * @param surveyQtnsVO
	 * @return
	 * @throws JCTException
	 */
	public SurveyQuestionsVO saveDropDown (SurveyQuestionsVO surveyQtnsVO) throws JCTException;
	/**
	 * Method saves check box related survey question
	 * @param surveyQtnsVO
	 * @return
	 * @throws JCTException
	 */
	public SurveyQuestionsVO saveCheckBox (SurveyQuestionsVO surveyQtnsVO) throws JCTException;
	/**
	 * Method populates all the survey questions
	 * @return
	 * @throws JCTException
	 */
	public String getExistingAllSurveyQuestion () throws JCTException;
	/**
	 * Method populates all the survey questions by user type
	 * @param userType
	 * @return
	 * @throws JCTException
	 */
	public String fetchAllExistingSurveyQtnsByUserType (int userType) throws JCTException;
	/**
	 * Method populates all the survey questions by user type and question type
	 * @param userType
	 * @param qtnType
	 * @return
	 * @throws JCTException
	 */
	public String fetchAllExistingSurveyQtnsByUserTypeAndQtnType (int userType, int qtnType) throws JCTException;
	/**
	 * Method populates all the survey questions by question type
	 * @param qtnType
	 * @return
	 * @throws JCTException
	 */
	public String fetchAllExistingSurveyQtnsByQtnType (int qtnType) throws JCTException;
	/**
	 * Method soft deletes the entry
	 * @param userType
	 * @param answerType
	 * @param qtnDescription
	 * @return
	 * @throws JCTException
	 */
	public String deleteSurveyQtn (int userType, int answerType, String qtnDescription) throws JCTException;
	/**
	 * Method updates the main question of the text
	 * @param userType
	 * @param answerType
	 * @param originalQuestion
	 * @param updatedQuestion
	 * @return
	 * @throws JCTException
	 */
	public String updateTextSurveyQtn (int userType, int answerType, String originalQuestion, String updatedQuestion, String isMandatory) throws JCTException;
	/**
	 * Method updates the main question of the radio
	 * @param userType
	 * @param answerType
	 * @param originalQuestion
	 * @param updatedMainQuestion
	 * @param subQuestion
	 * @param createdBy
	 * @return
	 * @throws JCTException
	 */
	public String updateTextRadioQtn (int userType, int answerType, String originalQuestion, 
			String updatedMainQuestion, String subQuestion, String createdBy,
			String isMandatory) throws JCTException;
	/**
	 * Method updates the main question of the drop down
	 * @param userType
	 * @param answerType
	 * @param originalQuestion
	 * @param updatedMainQuestion
	 * @param subQuestion
	 * @param createdBy
	 * @return
	 * @throws JCTException
	 */
	public String updateDropDownSurveyQtn (int userType, int answerType, String originalQuestion, 
			String updatedMainQuestion, String subQuestion, String createdBy, String isMandatory) throws JCTException;
	/**
	 * Method updates the main question of the check box
	 * @param userType
	 * @param answerType
	 * @param originalQuestion
	 * @param updatedMainQuestion
	 * @param subQuestion
	 * @param createdBy
	 * @return
	 * @throws JCTException
	 */
	public String updateCheckboxSurveyQtn (int userType, int answerType, String originalQuestion, 
			String updatedMainQuestion, String subQuestion, String createdBy, String isMandatory) throws JCTException;
	
	/**
	 * Method saves all the survey questions.
	 * @param text
	 * @param radio
	 * @param drop
	 * @param check
	 * @param emailId
	 * @param userId
	 * @return
	 * @throws JCTException
	 */
	public String saveSurveyQuestions(String text, String radio, String drop, 
			String check, String emailId, int userId) throws JCTException;
	/**
	 * Method update active status for facilitator.
	 * @param emailId
	 * @param userId
	 * @return
	 * @throws JCTException
	 */
	public String updateActiveStatusFacilitator(String emailId, int userId) throws JCTException;
	/**
	 * 
	 * @param userType
	 * @param qtnType
	 * @param profileId
	 * @return
	 * @throws JCTException
	 */
	public String fetchgSurveyQtnsByAllParams (int userType, int qtnType, int profileId) throws JCTException;
	/**
	 * 
	 * @param qtnType
	 * @param profileId
	 * @return
	 * @throws JCTException
	 */
	public String fetchAllExistingSurveyQtnsOnlyUserProfile (int qtnType, int profileId) throws JCTException;
	/**
	 * 
	 * @param userType
	 * @param profileId
	 * @return
	 * @throws JCTException
	 */
	public String fetchSurveyQtnsByProfileAndUserType (int userType, int profileId) throws JCTException;
	/**
	 * 
	 * @param builder
	 * @param userProfile
	 * @param userType
	 * @return
	 * @throws JCTException
	 */
	public SurveyQuestionsVO saveSurvryQtnOrder(String builder, Integer userProfile, Integer userType, String builderAnsType) throws JCTException;
}
