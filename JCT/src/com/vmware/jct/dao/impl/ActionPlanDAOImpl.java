package com.vmware.jct.dao.impl;


import java.util.List;

import org.hibernate.SessionFactory;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Repository;
import org.springframework.transaction.annotation.Propagation;
import org.springframework.transaction.annotation.Transactional;

import com.vmware.jct.dao.DataAccessObject;
import com.vmware.jct.dao.IActionPlanDAO;
import com.vmware.jct.exception.DAOException;
import com.vmware.jct.model.JctActionPlan;

/**
 * 
 * <p><b>Class name:</b> ActionPlanDAOImpl.java</p>
 * <p><b>Author:</b> InterraIt</p>
 * <p><b>Purpose:</b> This is a ActionPlanDAOImpl class. This artifact is Persistence Manager layer artifact.
 * ActionPlanDAOImpl implement IActionPlanDAO interface and override the following  methods.
 * -saveActinoPlans(JctActionPlan actionPlan)
 * -getList(String jobRefNo)
 * -remove(JctActionPlan object)
 * <p><b>Description:</b> This class used to perform search functionality </p>
 * <p><b>Revision History:</b>
 * </p>
 */
@Repository
public class ActionPlanDAOImpl extends DataAccessObject implements IActionPlanDAO {

	@Autowired
	private SessionFactory sessionFactory;
	
	/**
	 * The method will save action plan data.
	 * @param It will take JctActionPlan object as a input parameter
	 * @return It returns string object
	 * @throws DAOException -it throws DAOException
	 */
	@Transactional(propagation=Propagation.REQUIRED)
	public String saveActinoPlans(JctActionPlan actionPlan) throws DAOException {
		return save(actionPlan);
	}
	/**
	 * The method will merge action plan data.
	 * @param It will take JctActionPlan object as a input parameter
	 * @return It returns string object
	 * @throws DAOException -it throws DAOException
	 */
	@Transactional(propagation=Propagation.REQUIRED)
	public String mergeActinoPlans(JctActionPlan actionPlan) throws DAOException {
		return update(actionPlan);
	}
	/**
	 * Method returns list of action plan.
	 * @param It will take jobRefNo as a input parameter
	 * @return It returns List of JctActionPlan object
	 * @throws DAOException -it throws DAOException
	 */
	@Transactional(propagation=Propagation.REQUIRED,readOnly = true)
	public List<JctActionPlan> getList(String jobRefNo) throws DAOException {
		return sessionFactory.getCurrentSession().getNamedQuery("fetchListOfActionPlan").setParameter("refNo", jobRefNo).list();
	}
	
	@Transactional(propagation=Propagation.REQUIRED,readOnly = true)
	public List<JctActionPlan> getListEdited(String jobRefNo) throws DAOException {
		return sessionFactory.getCurrentSession().getNamedQuery("fetchListOfActionPlanEdited").setParameter("refNo", jobRefNo).list();
	}
	/**
	 * Method deletes action plan object from database.
	 * @param It will take jobRefNo as a input parameter
	 * @return It returns String object
	 * @throws DAOException -it throws DAOException
	 */
	@Transactional(propagation=Propagation.REQUIRED)
	public String remove(JctActionPlan object) throws DAOException {
		String cal = "success";
		try{
			Delete(object);
		}catch(Exception ex){
			cal = "failure";
		}
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
		return sessionFactory.getCurrentSession().getNamedQuery("getActionPlanMainQtnListByJrNo").setParameter("refNo", jobRefNo).list();
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
		return sessionFactory.getCurrentSession().getNamedQuery("getActionPlanSubQtnListByJrNo")
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
		return sessionFactory.getCurrentSession().getNamedQuery("getAnswerListByJrNo")
				.setParameter("mainQuestion", mainQuestion)
				.setParameter("subQuestion", subQuestion)
				.setParameter("refNo", jobRefNo).list();
	}
	/**
	 * Method returns the count of items present in completion table
	 * @param It will take jobRefNo,subquestion and mainQuestion as a input parameter
	 * @return It returns Integer object
	 * @throws DAOException -it throws DAOException
	 */
	@Transactional(propagation=Propagation.REQUIRED)
	public Long getNumberOfCompletion (String jobReferenceNos) throws DAOException {
		return (Long) sessionFactory.getCurrentSession().getNamedQuery("getNumberOfCompletion")
				.setParameter("jobReferenceNos", jobReferenceNos).uniqueResult();
	}
	/**
	 * Method returns the last decision made
	 * @param It will take jobRefNo,subquestion and mainQuestion as a input parameter
	 * @return It returns Integer object
	 * @throws DAOException -it throws DAOException
	 */
	@Transactional(propagation=Propagation.REQUIRED)
	public Integer getLastCompletionStatus (String jobReferenceNos) throws DAOException {
		int count = 0;
		try {
			count = (Integer) sessionFactory.getCurrentSession().getNamedQuery("getLastRowCount")
					.setParameter("jobRefNo", jobReferenceNos)
					.setParameter("jobRefNo2", jobReferenceNos).uniqueResult();
		} catch (Exception ex) {
			
		}
		return count;
	}
	/**
	 * Method returns the count
	 * @param jrNo
	 * @param user
	 * @param userId
	 * @param profileId
	 * @return It returns Long object
	 * @throws DAOException -it throws DAOException
	 */
	@Transactional(propagation=Propagation.REQUIRED)
	public Long getActionPlanCount(String jrNo, String user, int userId,
			int profileId) throws DAOException {
		return (Long)sessionFactory.getCurrentSession().getNamedQuery("getFrozenActionPlanCount")
				.setParameter("jrNo", jrNo)
				.setParameter("user", user)
				.setParameter("userId", userId).uniqueResult();
	}
	@Transactional(propagation=Propagation.REQUIRED)
	public String getOnetTitle(String onetCode) throws DAOException {
		return (String) sessionFactory.getCurrentSession().getNamedQuery("getOccupationDesc")
				.setParameter("onetTitle", onetCode).uniqueResult();
	}
}
