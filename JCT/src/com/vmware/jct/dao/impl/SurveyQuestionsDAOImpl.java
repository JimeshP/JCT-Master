package com.vmware.jct.dao.impl;

import org.hibernate.SessionFactory;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Repository;
import org.springframework.transaction.annotation.Propagation;
import org.springframework.transaction.annotation.Transactional;

import com.vmware.jct.dao.DataAccessObject;
import com.vmware.jct.dao.ISurveyQuestionDAO;
import com.vmware.jct.exception.DAOException;
import com.vmware.jct.model.JctSurveyQuestions;
/**
 * 
 * <p><b>Class name:</b> SurveyQuestionsDAOImpl.java</p>
 * <p><b>Author:</b> InterraIt</p>
 * <p><b>Purpose:</b> This is a SurveyQuestionsDAOImpl class. This artifact is Persistence Manager layer artifact.
 * QuestionnaireDAOImpl implement IQuestionnaireDAO interface and extend DataAccessObject  and override the following  methods.
 * -saveSurveyQuestion(JctSurveyQuestions survey)
 * <p><b>Description:</b>  This class used to perform insert, modify ,fetch data from db and delete operation. Save, update, delete operations are 
 * performed in DataAccessObject class </p>
 * <p><b>Copyrights:</b> 	All rights reserved by Interra IT and should only be used for its internal application development.</p>
 * <p><b>Revision History:</b>
 * 	<li>InterraIt - 14/Nov/2014 - Implemented the class </li>
 * </p>
 */
@Repository
public class SurveyQuestionsDAOImpl  extends DataAccessObject implements ISurveyQuestionDAO {

	@Autowired
	private SessionFactory sessionFactory;
	
	private static final Logger LOGGER = LoggerFactory.getLogger(SurveyQuestionsDAOImpl.class);
	/**
	 * Method saves the survey answers to the transaction database.
	 * @param JctSurveyQuestions
	 * @return status
	 * @throws DAOException
	 */
	@Transactional(propagation=Propagation.REQUIRED)
	public String saveSurveyQuestion(JctSurveyQuestions survey)
			throws DAOException {
		LOGGER.info(">>>>>>>> SurveyQuestionsDAOImpl.saveSurveyQuestion");
		return save(survey);
	}

}
