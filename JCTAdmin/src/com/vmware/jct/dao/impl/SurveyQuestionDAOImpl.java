package com.vmware.jct.dao.impl;

import java.util.List;

import org.hibernate.SessionFactory;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Repository;
import org.springframework.transaction.annotation.Propagation;
import org.springframework.transaction.annotation.Transactional;

import com.vmware.jct.dao.DataAccessObject;
import com.vmware.jct.dao.ISurveyQuestionDAO;
import com.vmware.jct.dao.dto.SurveyQuestionDTO;
import com.vmware.jct.exception.DAOException;
import com.vmware.jct.model.JctJobAttribute;
import com.vmware.jct.model.JctSurveyQuestionMaster;
import com.vmware.jct.model.JctSurveyQuestions;
import com.vmware.jct.model.JctUser;
/**
 * 
 * <p><b>Class name:</b> SurveyQuestionDAOImpl.java</p>
 * <p><b>Author:</b> InterraIT</p>
 * <p><b>Purpose:</b> This is a ReportDAOImpl class. This artifact is Persistence Manager layer artifact.
 * SurveyQuestionDAOImpl extends DataAccessObject and implements ISurveyQuestionDAO interface.
 * </p>
 * <p><b>Description:</b> This class used to perform insert, modify ,fetch data from db and delete operation. Save, update, delete operations are performed in DataAccessObject class </p>
 * <p><b>Revision History:</b>
 * 	<li>InterraIT - 21/Oct/2014 - Implemented the class</li>
 * </p>
 */
@Repository
public class SurveyQuestionDAOImpl extends DataAccessObject implements ISurveyQuestionDAO {

	@Autowired
	private SessionFactory sessionFactory;
	
	private static final Logger LOGGER = LoggerFactory.getLogger(SurveyQuestionDAOImpl.class);
	/**
	 * Method saves the survey questions
	 */
	@Transactional(propagation = Propagation.REQUIRED)
	public String saveSurveyQuestion(JctSurveyQuestionMaster surveyMaster)
			throws DAOException {
		LOGGER.info(">>>>>> SurveyQuestionDAOImpl.saveSurveyQuestion");
		return save(surveyMaster);
	}
	/**
	 * Method updates the survey questions
	 */
	@Transactional(propagation = Propagation.REQUIRED)
	public String updateSurveyQuestion(JctSurveyQuestionMaster surveyMaster)
			throws DAOException {
		LOGGER.info(">>>>>> SurveyQuestionDAOImpl.updateSurveyQuestion");
		return update(surveyMaster);
	}
	/**
	 * Method fetches all existing active survey question
	 */
	@Transactional(propagation = Propagation.REQUIRED)
	public List<SurveyQuestionDTO> getExistingAllSurveyQuestion()
			throws DAOException {
		LOGGER.info(">>>>>> SurveyQuestionDAOImpl.getExistingAllSurveyQuestion");
		return sessionFactory.getCurrentSession().getNamedQuery("getAllSurveyQuestion").list();
	}
	/**
	 * Method fetches all existing active survey sub question.
	 * @param userType
	 * @param mainQtns
	 * @return List
	 * @throws DAOException
	 */
	@Transactional(propagation = Propagation.REQUIRED)
	public List<SurveyQuestionDTO> getAllSubQtns(int answerType,
			int userType, String mainQtns) throws DAOException {
		LOGGER.info(">>>>>> SurveyQuestionDAOImpl.getAllSubQtns");
		return sessionFactory.getCurrentSession().getNamedQuery("getAllSurveySubQuestionsFromMainQuestion")
				.setParameter("answerType", answerType)
				.setParameter("userType", userType)
				.setParameter("mainQtn", mainQtns).list();
	}
	/**
	 * Method fetches all existing active survey sub question.
	 * @param userType
	 * @param mainQtns
	 * @return List
	 * @throws DAOException
	 */
	@Transactional(propagation = Propagation.REQUIRED)
	public List<SurveyQuestionDTO> getAllSubQtnsByAnsType(int answerType,String mainQtns) throws DAOException {
		LOGGER.info(">>>>>> SurveyQuestionDAOImpl.getAllSubQtns");
		return sessionFactory.getCurrentSession().getNamedQuery("getAllSurveySubQuestionsFromMainQuestionOnlyAnsType")
				.setParameter("answerType", answerType)
				.setParameter("mainQtn", mainQtns).list();
	}
	
	/**
	 * Method fetches all existing active survey sub question.
	 * @param answerType
	 * @param userType
	 * @param mainQtns
	 * @return List
	 * @throws DAOException
	 */
	@Transactional(propagation = Propagation.REQUIRED)
	public Long getNumberOfSubQuestions(int answerType, int userType,
			String mainQtns) throws DAOException {
		LOGGER.info(">>>>>> SurveyQuestionDAOImpl.getNumberOfSubQuestions");
		return (Long) sessionFactory.getCurrentSession().getNamedQuery("getNosOfSubQtns")
				.setParameter("answerType", answerType)
				.setParameter("userType", userType)
				.setParameter("mainQtn", mainQtns).uniqueResult();
	}
	/**
	 * Method fetches all existing active survey sub question.
	 * @param answerType
	 * @param userType
	 * @param mainQtns
	 * @return List
	 * @throws DAOException
	 */
	@Transactional(propagation = Propagation.REQUIRED)
	public Long getNumberOfSubQuestionsByAnsType(int answerType, String mainQtns) throws DAOException {
		LOGGER.info(">>>>>> SurveyQuestionDAOImpl.getNumberOfSubQuestionsByAnsType");
		return (Long) sessionFactory.getCurrentSession().getNamedQuery("getNosOfSubQtnsByAnsType")
				.setParameter("answerType", answerType)
				.setParameter("mainQtn", mainQtns).uniqueResult();
	}
	/**
	 * Method fetches all existing active survey question.
	 * @param userType
	 * @return List
	 * @throws DAOException
	 */
	@Transactional(propagation = Propagation.REQUIRED)
	public List<SurveyQuestionDTO> fetchAllExistingSurveyQtnsByUserType(int userType)
			throws DAOException {
		LOGGER.info(">>>>>> SurveyQuestionDAOImpl.fetchAllExistingSurveyQtnsByUserType");
		return sessionFactory.getCurrentSession().getNamedQuery("getAllSurveyQuestionByUserType")
				.setParameter("userType", userType).list();
	}
	/**
	 * Method fetches all existing active survey sub question.
	 * @param userType
	 * @param qtnType
	 * @return List
	 * @throws DAOException
	 */
	@Transactional(propagation = Propagation.REQUIRED)
	public List<SurveyQuestionDTO> fetchAllExistingSurveyQtnsByUserTypeAndQtnType(
			int userType, int qtnType) throws DAOException {
		LOGGER.info(">>>>>> SurveyQuestionDAOImpl.fetchAllExistingSurveyQtnsByUserType");
		return sessionFactory.getCurrentSession().getNamedQuery("getAllSurveyQuestionByUserTypeAndQstnType")
				.setParameter("userType", userType)
				.setParameter("qtnType", qtnType).list();
	}
	/**
	 * Method fetches all existing active survey sub question.
	 * @param userType
	 * @param qtnType
	 * @return List
	 * @throws DAOException
	 */
	@Transactional(propagation = Propagation.REQUIRED)
	public List<SurveyQuestionDTO> fetchAllExistingSurveyQtnsByQtnType(int qtnType) throws DAOException {
		LOGGER.info(">>>>>> SurveyQuestionDAOImpl.fetchAllExistingSurveyQtnsByQtnType");
		return sessionFactory.getCurrentSession().getNamedQuery("getAllSurveyQuestionByQstnType")
				.setParameter("qtnType", qtnType).list();
	}
	/**
	 * Method soft deletes all existing survey questions.
	 * @param userType
	 * @param answerType
	 * @param qtnDescription
	 * @return List
	 * @throws DAOException
	 */
	@Transactional(propagation = Propagation.REQUIRED)
	public int deleteSurveyQtn(int userType, int answerType,
			String qtnDescription) throws DAOException {
		return sessionFactory.getCurrentSession().getNamedQuery("softDeleteTheSurveyQuestion")
				.setParameter("userType", userType)
				.setParameter("answerType", answerType)
				.setParameter("qtnDescription", qtnDescription).executeUpdate();
	}
	/**
	 * Method fetches all existing active survey sub question.
	 * @param userType
	 * @param answerType
	 * @param originalQuestion
	 * @return List
	 * @throws DAOException
	 */
	@Transactional(propagation = Propagation.REQUIRED)
	public List<JctSurveyQuestionMaster> getSurveyQuestionList(int userType,
			int answerType, String originalQuestion) throws DAOException {
		return sessionFactory.getCurrentSession().getNamedQuery("fetchTextSurveyQtnsForEdit")
				.setParameter("userType", userType)
				.setParameter("answerType", answerType)
				.setParameter("originalQuestion", originalQuestion).list();
	}
	
	@Transactional(propagation=Propagation.REQUIRED)
	public String saveSurveyQuestion(JctSurveyQuestions survey)
			throws DAOException {
		LOGGER.info(">>>>>>>> SurveyQuestionsDAOImpl.saveSurveyQuestion");
		return save(survey);
	}
	
	@Transactional(propagation = Propagation.REQUIRED)
	public JctUser fetchUserData(Integer userId) throws DAOException {
		LOGGER.info(">>>>>> SurveyQuestionsDAOImpl.fetchUserData");				
		return (JctUser) sessionFactory.getCurrentSession().
				getNamedQuery("fetchUserByPk")
				.setParameter("pk", userId).list().get(0);
	}
	
	@Transactional(propagation = Propagation.REQUIRED)
	public String updateUser(JctUser user) throws DAOException {
		LOGGER.info(">>>>>> SurveyQuestionsDAOImpl.updateUser");
		return update(user);
	}
	/**
	 * Method fetches all existing active survey sub question.
	 * @param userType
	 * @param qtnType
	 * @return List
	 * @throws DAOException
	 */
	@Transactional(propagation = Propagation.REQUIRED)
	public List<SurveyQuestionDTO> fetchgSurveyQtnsByAllParams(
			int userType, int qtnType, int profileId) throws DAOException {
		LOGGER.info(">>>>>> SurveyQuestionDAOImpl.fetchgSurveyQtnsByAllParams");
		return sessionFactory.getCurrentSession().getNamedQuery("fetchgSurveyQtnsByAllParams")
				.setParameter("userType", userType)
				.setParameter("qtnType", qtnType)
				.setParameter("profileId", profileId).list();
	}
	
	@Transactional(propagation = Propagation.REQUIRED)
	public List<SurveyQuestionDTO> fetchAllExistingSurveyQtnsOnlyUserProfile(
			int qtnType, int profileId) throws DAOException {
		LOGGER.info(">>>>>> SurveyQuestionDAOImpl.fetchAllExistingSurveyQtnsByQtnType");
		return sessionFactory.getCurrentSession().getNamedQuery("getAllSurveyQuestionByQstnTypeAndProfileId")
				.setParameter("qtnType", qtnType).list();
	}
	@Transactional(propagation = Propagation.REQUIRED)
	public List<SurveyQuestionDTO> fetchSurveyQtnsByProfileAndUserType(
			int userType, int profileId) throws DAOException {
		LOGGER.info(">>>>>> SurveyQuestionDAOImpl.fetchSurveyQtnsByProfileAndUserType");
		return sessionFactory.getCurrentSession().getNamedQuery("getSrvQtnByUserTypeAndProfileId")
				.setParameter("userType", userType)
				.setParameter("profileId", profileId).list();
	}
	@Transactional(propagation = Propagation.REQUIRED)
	public List<SurveyQuestionDTO> getAllSubQtnsByUserTypeAndProfile(
			int answerType, int userType, int profileId, String mainQtns) throws DAOException {
		LOGGER.info(">>>>>> SurveyQuestionDAOImpl.getAllSubQtnsByUserTypeAndProfile");
		return sessionFactory.getCurrentSession().getNamedQuery("getAllSurveySubQuestionsFromMainQuestionProfileId")
				.setParameter("answerType", answerType)
				.setParameter("userType", userType)
				.setParameter("mainQtn", mainQtns)
				.setParameter("profileId", profileId).list();
	}
	@Transactional(propagation = Propagation.REQUIRED)
	public Long getNumberOfSubQuestionsProfileId(int answerType, int userType,
			String mainQtns, int profileId) throws DAOException {
		LOGGER.info(">>>>>> SurveyQuestionDAOImpl.getNumberOfSubQuestionsProfileId");
		return (Long) sessionFactory.getCurrentSession().getNamedQuery("getNosOfSubQtnsProfileId")
				.setParameter("answerType", answerType)
				.setParameter("userType", userType)
				.setParameter("mainQtn", mainQtns)
				.setParameter("profileId", profileId).uniqueResult();
	}

	@Transactional(propagation = Propagation.REQUIRED)
	public String saveSurvryQtnOrder(String builder, Integer userProfile,
			Integer userType, String builderAnsType) throws DAOException {
		LOGGER.info(">>>>>> SurveyQuestionDAOImpl.saveSurvryQtnOrder");
		String returnStr = "";
		List<JctSurveyQuestionMaster> existingList = null;
		String[] splitDesc = builder.split("~~");
		String[] splitAnsType = builderAnsType.split("~~");
		int sunQtnLength = splitDesc.length;
		int answerType = 0;
		for (int i = 0; i < sunQtnLength; i++) {
			if (splitAnsType[i].equals("Text")) {
				answerType = 1;
			} else if (splitAnsType[i].equals("Radio")) {
				answerType = 2;
			} else if (splitAnsType[i].equals("Dropdown")) {
				answerType = 3;
			} else if (splitAnsType[i].equals("CheckBox")) {
				answerType = 4;
			}
			 existingList = (List<JctSurveyQuestionMaster>) sessionFactory.
					getCurrentSession().getNamedQuery("fetchToUpdateSurveyQtn").
					setParameter("profileId", userProfile).
					setParameter("userType", userType).
					setParameter("mainQtn", splitDesc[i]).
					setParameter("ansType", answerType).list();
			if (existingList.size() != 0) {
				for (int j = 0; j < existingList.size(); j++) {
					existingList.get(j).setJctSurveyQtnMasterOrder(i+1);
					returnStr = update(existingList.get(j));
				}
			}
		}
		LOGGER.info("<<<<<< SurveyQuestionDAOImpl.saveSurvryQtnOrder");
		return returnStr;
	}

	@Transactional(propagation = Propagation.REQUIRED)
	public int generateOrderSurvQtn(int profileId, int userType)
			throws DAOException {
		LOGGER.info(">>>>>> SurveyQuestionDAOImpl.generateOrderSurvQtn");
		Integer order = null;
		Integer value = (Integer) sessionFactory.getCurrentSession().getNamedQuery("fetchOrderSurveyQtn")
				.setParameter("profileId", profileId).setParameter("userType", userType).uniqueResult();
		if( null == value) {
			order = 1;
		} else {
			order = value + 1;
		}
		return order;
	}
}