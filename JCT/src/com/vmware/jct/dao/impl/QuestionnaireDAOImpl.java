package com.vmware.jct.dao.impl;

import java.util.Date;
import java.util.List;

import org.hibernate.SessionFactory;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Repository;
import org.springframework.transaction.annotation.Propagation;
import org.springframework.transaction.annotation.Transactional;

import com.vmware.jct.dao.DataAccessObject;
import com.vmware.jct.dao.IQuestionnaireDAO;
import com.vmware.jct.dao.dto.JobAttributeDTO;
import com.vmware.jct.exception.DAOException;
import com.vmware.jct.model.JctBeforeSketchQuestion;
import com.vmware.jct.model.JctGlobalProfileHistory;


/**
 * 
 * <p><b>Class name:</b> AfterSketchDAOImpl.java</p>
 * <p><b>Author:</b> InterraIt</p>
 * <p><b>Purpose:</b> This is a QuestionnaireDAOImpl class. This artifact is Persistence Manager layer artifact.
 * QuestionnaireDAOImpl implement IQuestionnaireDAO interface and extend DataAccessObject  and override the following  methods.
 * -saveAnswers(JctBeforeSketchQuestion question)
 * -getList(String jobRefNumber)
 * -getFetchJobAttribute()
 * -remove(JctBeforeSketchQuestion question)
 * <p><b>Description:</b>  This class used to perform insert, modify ,fetch data from db and delete operation. Save, update, delete operations are 
 * performed in DataAccessObject class </p>
 * <p><b>Copyrights:</b> 	All rights reserved by Interra IT and should only be used for its internal application development.</p>
 * <p><b>Revision History:</b>
 * 	<li>InterraIt - 31/Jan/2014 - Implement comments, introduce Named query, Transactional attribute </li>
 * </p>
 */
@Repository
public class QuestionnaireDAOImpl extends DataAccessObject implements IQuestionnaireDAO {

	@Autowired
	private SessionFactory sessionFactory;
	
	private static final Logger logger = LoggerFactory.getLogger(QuestionnaireDAOImpl.class);
	
	
	/**
	 * <p><b>Description:</b>This method is used for storing Before Sketch Reflection question</p>
	 * @param It will take JctBeforeSketchQuestion as a input parameter
	 * @return  It returns string object
	 * @throws DAOException -it throws DAOException
	 * 
	 */
	@Transactional(propagation=Propagation.REQUIRED)
	public String saveAnswers(JctBeforeSketchQuestion question)
			throws DAOException {
		return save(question);
	}
	/**
	 * <p><b>Description:</b>This method is used for merging Before Sketch Reflection question</p>
	 * @param It will take JctBeforeSketchQuestion as a input parameter
	 * @return  It returns string object
	 * @throws DAOException -it throws DAOException
	 * 
	 */
	@Transactional(propagation=Propagation.REQUIRED)
	public String mergeAnswers(JctBeforeSketchQuestion question)
			throws DAOException {
		return update(question);
	}
	
	/**
	 * <p><b>Description:</b>   Baserd on job reference number this method fetch all Before Sketch Reflection question based on job reference number</p>
	 * @param It will take jobRefNo as a input parameter
	 * @return  It returns list of JctBeforeSketchQuestion data from database
	 * @throws DAOException -it throws DAOException
	 * 
	 */
	@Transactional(propagation=Propagation.REQUIRED,readOnly = true)
	public List<JctBeforeSketchQuestion> getList(String jobRefNumber) throws DAOException {
		return sessionFactory.getCurrentSession().getNamedQuery("fetchAllQuestion").setParameter("refNo", jobRefNumber).list();
	}
	
	
	/**
	 * <p><b>Description:</b>   Baserd on job reference number this method fetch all Before Sketch Reflection question based on profile id</p>
	 * @param It will take profileId as a input parameter
	 * @return  It returns list of JobAttributeDTO data from database
	 * @throws DAOException -it throws DAOException
	 * 
	 */
	@Transactional(propagation=Propagation.REQUIRED,readOnly = true)
	public List<JobAttributeDTO> getFetchJobAttribute(int profileId) throws DAOException {
		return sessionFactory.getCurrentSession().getNamedQuery("fetchAllJobAttribute").setParameter("profileId", profileId).list();
	}
	
	
	/**
	 * <p><b>Description:</b>This method is used for remove Before Sketch Reflection question </p>
	 * @param It will take JctBeforeSketchQuestion as a input parameter
	 * @return  It returns string object
	 * @throws DAOException -it throws DAOException
	 * 
	 */
	@Transactional(propagation=Propagation.REQUIRED)
	public String remove(JctBeforeSketchQuestion object) throws DAOException {
		logger.info(">>>> QuestionnaireDAOImpl.remove"); 
		String cal = "success";
		try{
			Delete(object);
		}catch(Exception ex){
			cal = "failure";
			logger.error(ex.getLocalizedMessage());
		}
		logger.info("<<<< QuestionnaireDAOImpl.remove"); 
		return cal;
	}
	/**
	 * Method returns list of distinct question answer
	 * @param It will take jobRefNo as a input parameter
	 * @return It returns string object
	 * @throws DAOException -it throws DAOException
	 */
	@Transactional(propagation=Propagation.REQUIRED)
	public List<String> getDistinctQuestionList(String jobRefNo)
			throws DAOException {
		return sessionFactory.getCurrentSession().getNamedQuery("getQuestionnaireMainQtnListByJrNo").setParameter("refNo", jobRefNo).list();
	}
	/**
	 * Method returns distinct sub question list based on main question and job reference nos
	 * @param It will take jobRefNo and mainQuestion as a input parameter
	 * @return It returns List of String object
	 * @throws DAOException -it throws DAOException
	 */
	@Transactional(propagation=Propagation.REQUIRED)
	public List<String> getDistinctSubQuestionListByJrno(String mainQuestion, String jobRefNo)
			throws DAOException {
		return sessionFactory.getCurrentSession().getNamedQuery("getQuestionnaireSubQtnListByJrNo")
				.setParameter("mainQuestion", mainQuestion)
				.setParameter("refNo", jobRefNo).list();
	}
	/**
	 * Method returs list of answers based on main question, sub question and job reference number
	 * @param It will take jobRefNo,subquestion and mainQuestion as a input parameter
	 * @return It returns List of String object
	 * @throws DAOException -it throws DAOException
	 */
	@Transactional(propagation=Propagation.REQUIRED)
	public List<String> getAnswerListByJrno(String mainQuestion, String subQuestion, String jobRefNo)
			throws DAOException {
		return sessionFactory.getCurrentSession().getNamedQuery("getQuestionnaireAnswerListByJrNo")
				.setParameter("mainQuestion", mainQuestion)
				.setParameter("subQuestion", subQuestion)
				.setParameter("refNo", jobRefNo).list();
	}
	/**
	 * <p><b>Description:</b>  Baserd on job reference number this method fetch all data from jct after sketch header table</p>
	 * @param It will take profileId as a input parameter
	 * @return It returns List of String
	 * @throws DAOException -it throws DAOException
	 * 
	 */
	@Transactional(propagation=Propagation.REQUIRED,readOnly = true)
	public List<String> getQuestionnaireMainQtnList(int profileId) throws DAOException {
		return sessionFactory.getCurrentSession().getNamedQuery("getQestnrMainQtnListByProfile")
				.setParameter("profileId", profileId).list();
	}
	/**
	 * <p><b>Description:</b>  Baserd on job reference number this method fetch all data from action plan table</p>
	 * @param It will take profileId and mainQtn as a input parameter
	 * @return It returns List of String
	 * @throws DAOException -it throws DAOException
	 * 
	 */
	@Transactional(propagation=Propagation.REQUIRED,readOnly = true)
	public List<String> getQuestionnaireSubQtnList(String mainQtn, int profileId) throws DAOException {
		return sessionFactory.getCurrentSession().getNamedQuery("getQtnreSubQtnList")
				.setParameter("mainQtn", mainQtn)
				.setParameter("profileId", profileId).list();
	}
	/**
	 * <p><b>Description:</b>Method fetches all the existing forzen job attributes</p>
	 * @param userId
	 * @param createdFor
	 * @param profileId
	 * @return
	 * @throws DAOException
	 */
	public Long getExistingFrozenAttrsCount(int userId,
			String createdFor, int profileId) throws DAOException {
		return (Long)sessionFactory.getCurrentSession().getNamedQuery("getExtingFrozenJobAttrs")
				.setParameter("userId", userId)
				.setParameter("createdFor", createdFor)
				.setParameter("profileId", profileId).uniqueResult();
	}
	/**
	 * <p><b>Description:</b>Method fetches all the existing forzen job attributes</p>
	 * @param jrNo
	 * @param user
	 * @param userId
	 * @param profileId
	 * @return
	 * @throws DAOException
	 */
	@Transactional(propagation=Propagation.REQUIRED)
	public Long getQuestCount(String jrNo, String user, int userId,
			int profileId) throws DAOException {
		return (Long)sessionFactory.getCurrentSession().getNamedQuery("getFrozenQtnCount")
				.setParameter("jrNo", jrNo)
				.setParameter("user", user)
				.setParameter("userId", userId).uniqueResult();
	}
	
	/**
	 * 
	 */
	@Transactional(propagation=Propagation.REQUIRED,readOnly = true)
	public List<JctGlobalProfileHistory> getGlobalProfileChangedData(Date headerDate, String subQtn, String mainQtn, String diff) throws DAOException {
		String namedQuery = "getGlobalRQChangeBySubQtn";
		if (diff.equals("AP")) {
			namedQuery = "getGlobalAPChangeBySubQtn";
		}
		JctGlobalProfileHistory obj = new JctGlobalProfileHistory();
		obj.setPresentTimeStamp(new Date());
		return sessionFactory.getCurrentSession().getNamedQuery(namedQuery)
				.setParameter("subQtn", subQtn)
				.setParameter("mainQtn", mainQtn)
				.setParameter("date1", headerDate)
				.setParameter("date2", obj.getPresentTimeStamp()).list();
	}
}
