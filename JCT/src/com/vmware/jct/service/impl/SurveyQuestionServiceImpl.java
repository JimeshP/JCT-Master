package com.vmware.jct.service.impl;

import java.util.Date;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;
import org.springframework.transaction.annotation.Propagation;
import org.springframework.transaction.annotation.Transactional;

import com.vmware.jct.common.utility.CommonConstants;
import com.vmware.jct.dao.ISurveyQuestionDAO;
import com.vmware.jct.exception.DAOException;
import com.vmware.jct.exception.JCTException;
import com.vmware.jct.model.JctSurveyQuestions;
import com.vmware.jct.service.ISurveyQuestionService;
/**
 * 
 * <p><b>Class name:</b> SurveyQuestionServiceImpl.java</p>
 * <p><b>Author:</b> InterreIt</p>
 * <p><b>Purpose:</b> This is a SurveyQuestionServiceImpl class. This artifact is Business layer artifact.
 * SurveyQuestionServiceImpl implement ISurveyQuestionService interface and override the following  methods.
 * -saveOrUpdateAnswers(List<QuestionnaireVO> vo, String jobRefNumber)
 * -fetchJobAttribute()
 * </p>
 * <p><b>Description:</b> This class used to perform insert, modify ,fetch data from db and delete operation. Save, update, delete operations are performed in DataAccessObject class </p>
 * <p><b>Revision History:</b>
 * 	<li>InterreIt - 19/Jan/2014 - Implement Exception </li>
 * <li>InterreIt - InterraIt - 14/Oct/2014 - Implemented the class </li>
 * </p>
 */
@Service
public class SurveyQuestionServiceImpl implements ISurveyQuestionService {

	@Autowired
	private ISurveyQuestionDAO surveyQuestionDAO;
	
	private static final Logger LOGGER = LoggerFactory.getLogger(SurveyQuestionServiceImpl.class);
	
	/**
	 * Method saves all the survey questions.
	 * @param text
	 * @param radio
	 * @param drop
	 * @param check
	 * @return
	 * @throws JCTException
	 */
	@Transactional(propagation=Propagation.REQUIRED)
	public String saveSurveyQuestions(String text, String radio, String drop,
			String check, String emailId, int userId) throws JCTException {
		LOGGER.info(">>>>>>>> SurveyQuestionServiceImpl.saveSurveyQuestions");
		String result = "failure";
		try {
			if (text.trim().length() > 0) {
				saveTextSurveyAns(text, emailId, userId);
			}
			if (radio.trim().length() > 0) {
				saveRadioSurveyAns(radio, emailId, userId);
			}
			if (drop.trim().length() > 0) {
				saveDropSurveyAns(drop, emailId, userId);
			}
			if (check.trim().length() > 0) {
				saveCheckSurveyAns(check, emailId, userId);
			}
			result = "success";
		} catch (DAOException ex) {
			LOGGER.error("UNABLE TO SAVE TO SAVE SURVEY QUESTIONS: "+ex.getLocalizedMessage());
		}
		LOGGER.info("<<<<<<<< SurveyQuestionServiceImpl.saveSurveyQuestions");
		return result;
	}
	/**
	 * Method saves only the free text surveys.
	 * @param text
	 * @return
	 * @throws DAOException
	 */
	@Transactional(propagation=Propagation.REQUIRED)
	private String saveTextSurveyAns (String text, String emailId, int userId) throws DAOException {
		LOGGER.info(">>>>>>>> SurveyQuestionServiceImpl.saveTextSurveyAns");
		//What do you think of the tool?`ONE~~How asked you to join?`NOTANSWERED~~HEHE`TWO~~HOHO`NOTANSWERED~~HIHI`THREE~~HUHU`NOTANSWERED~~Test`NOTANSWERED~~
		String status = "failure";
		String[] qtnSplit = text.split("~~");
		for (int index = 0; index < qtnSplit.length; index++) {
			String[] innerSplit = qtnSplit[index].split("`"); // Left = qtn | Right = ans
			String mainQtn = innerSplit[0];
			String answer = innerSplit[1].trim();
			if (!answer.equals("NOTANSWERED")) {
				JctSurveyQuestions surveyQuestions = new JctSurveyQuestions();
				surveyQuestions.setJctSurveyAnswer(answer);
				surveyQuestions.setJctSurveyAnswerType(CommonConstants.FREE_TEXT);
				surveyQuestions.setJctSurveyCreatedTs(new Date());
				surveyQuestions.setJctSurveyMainQtn(mainQtn);
				surveyQuestions.setJctSurveySubQtn("NA");
				surveyQuestions.setJctSurveyTakenBy(emailId);
				surveyQuestions.setJctSurveyTakenByUserId(userId);
				surveyQuestions.setJctSurveyUserType(CommonConstants.GENERAL_USER);
				surveyQuestionDAO.saveSurveyQuestion(surveyQuestions);
			}
			status = "success";
		}
		LOGGER.info("<<<<<<<< SurveyQuestionServiceImpl.saveTextSurveyAns");
		return status;
	}
	/**
	 * Method saves only the radio surveys.
	 * @param text
	 * @return
	 * @throws DAOException
	 */
	@Transactional(propagation=Propagation.REQUIRED)
	private String saveRadioSurveyAns (String radio, String emailId, int userId) throws DAOException {
		LOGGER.info(">>>>>>>> SurveyQuestionServiceImpl.saveRadioSurveyAns");
		// RADIO WITH 5 OPTIONS`OPTION 1~~Are you happy?`Yes~~
		String status = "failure";
		String[] qtnSplit = radio.split("~~");
		for (int index = 0; index < qtnSplit.length; index++) {
			try {
			String[] innerSplit = qtnSplit[index].split("`"); // Left = qtn | Right = ans
			String ans = (String) innerSplit[1].trim();
			if (!ans.equals("NOTANSWERED")) {
				if (ans.trim().length() > 0) {
					JctSurveyQuestions surveyQuestions = new JctSurveyQuestions();
					surveyQuestions.setJctSurveyAnswer((String) innerSplit[1]);
					surveyQuestions.setJctSurveyAnswerType(CommonConstants.RADIO_BTN);
					surveyQuestions.setJctSurveyCreatedTs(new Date());
					surveyQuestions.setJctSurveyMainQtn((String) innerSplit[0]);
					surveyQuestions.setJctSurveySubQtn((String) innerSplit[0]);
					surveyQuestions.setJctSurveyTakenBy(emailId);
					surveyQuestions.setJctSurveyTakenByUserId(userId);
					surveyQuestions.setJctSurveyUserType(CommonConstants.GENERAL_USER);
					surveyQuestionDAO.saveSurveyQuestion(surveyQuestions);
				}
			}
			} catch (Exception e) {
				LOGGER.error("NO OPTION SELECTED FOR RADIO");
			}
			status = "success";
		}
		LOGGER.info("<<<<<<<< SurveyQuestionServiceImpl.saveRadioSurveyAns");
		return status;
	}
	/**
	 * Method saves only the drop down surveys.
	 * @param text
	 * @return
	 * @throws DAOException
	 */
	@Transactional(propagation=Propagation.REQUIRED)
	private String saveDropSurveyAns (String drop, String emailId, int userId) throws DAOException {
		LOGGER.info(">>>>>>>> SurveyQuestionServiceImpl.saveDropSurveyAns");
		// What is your job profile?`HR~~AGAIN A DROP DOWN`VERY FUNNY~~ 
		String status = "failure";
		String[] qtnSplit = drop.split("~~");
		for (int index = 0; index < qtnSplit.length; index++) {
			String[] innerSplit = qtnSplit[index].split("`"); // Left = qtn | Right = ans
			String ans = (String) innerSplit[1].trim();
			if (!ans.equals("NOTANSWERED")) {
				if (ans.trim().length() > 0) {
					JctSurveyQuestions surveyQuestions = new JctSurveyQuestions();
					surveyQuestions.setJctSurveyAnswer((String) innerSplit[1]);
					surveyQuestions.setJctSurveyAnswerType(CommonConstants.DROP_DOWN);
					surveyQuestions.setJctSurveyCreatedTs(new Date());
					surveyQuestions.setJctSurveyMainQtn((String) innerSplit[0]);
					surveyQuestions.setJctSurveySubQtn((String) innerSplit[0]);
					surveyQuestions.setJctSurveyTakenBy(emailId);
					surveyQuestions.setJctSurveyTakenByUserId(userId);
					surveyQuestions.setJctSurveyUserType(CommonConstants.GENERAL_USER);
					surveyQuestionDAO.saveSurveyQuestion(surveyQuestions);
				}
				status = "success";
			}
		}
		LOGGER.info("<<<<<<<< SurveyQuestionServiceImpl.saveDropSurveyAns");
		return status;
	}
	/**
	 * Method saves only the check box surveys.
	 * @param text
	 * @return
	 * @throws DAOException
	 */
	@Transactional(propagation=Propagation.REQUIRED)
	private String saveCheckSurveyAns (String check, String emailId, int userId) throws DAOException {
		LOGGER.info(">>>>>>>> SurveyQuestionServiceImpl.saveCheckSurveyAns");
		// What all do you enjoy?`PLAYING#Reading#~~
		// CHEC K WITH 5`two#~~ ----    CHECK`~~ffff`~~hkgkhhuyi`~~
		String status = "failure";
		String[] qtnSplit = check.split("~~");
		for (int index = 0; index < qtnSplit.length; index++) {
			// What all do you enjoy?`PLAYING#Reading#
			String[] innerSplit = qtnSplit[index].split("`"); // Left = qtn | Right = answers (multiple)
			try {
				String[] answerSplit = innerSplit[1].split("#");
				for (int ansIndex = 0; ansIndex < answerSplit.length; ansIndex++) {
					if (!((String) answerSplit[ansIndex]).equals("NOTANSWERED")) {
						JctSurveyQuestions surveyQuestions = new JctSurveyQuestions();
						surveyQuestions.setJctSurveyAnswer((String) answerSplit[ansIndex]);
						surveyQuestions.setJctSurveyAnswerType(CommonConstants.CHECK_BOX);
						surveyQuestions.setJctSurveyCreatedTs(new Date());
						surveyQuestions.setJctSurveyMainQtn((String) innerSplit[0]);
						surveyQuestions.setJctSurveySubQtn((String) answerSplit[ansIndex]);
						surveyQuestions.setJctSurveyTakenBy(emailId);
						surveyQuestions.setJctSurveyTakenByUserId(userId);
						surveyQuestions.setJctSurveyUserType(CommonConstants.GENERAL_USER);
						surveyQuestionDAO.saveSurveyQuestion(surveyQuestions);
						status = "success";
					}
				}
			} catch (Exception ex) {
				LOGGER.error("NO SUB ELEMENT FOUND");
			}
		}
		LOGGER.info("<<<<<<<< SurveyQuestionServiceImpl.saveCheckSurveyAns");
		return status;
	}
}
