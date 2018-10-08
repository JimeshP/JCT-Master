package com.vmware.jct.dao.impl;

import java.util.List;

import org.hibernate.SessionFactory;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Repository;
import org.springframework.transaction.annotation.Propagation;
import org.springframework.transaction.annotation.Transactional;

import com.vmware.jct.common.utility.CommonConstants;
import com.vmware.jct.dao.DataAccessObject;
import com.vmware.jct.dao.IBeforeSketchDAO;
import com.vmware.jct.exception.DAOException;
import com.vmware.jct.model.JctBeforeSketchHeader;
import com.vmware.jct.model.JctBeforeSketchHeaderTemp;
import com.vmware.jct.model.JctBeforeSketchQuestion;
import com.vmware.jct.model.JctQuestion;
import com.vmware.jct.service.vo.UserVO;


/**
 * 
 * <p><b>Class name:</b> BeforeSketchDAOImpl.java</p>
 * <p><b>Author:</b> InterraIt</p>
 * <p><b>Purpose:</b> This is a BeforeSketchDAOImpl class. This artifact is Persistence Manager layer artifact.
 * QuestionnaireDAOImpl implement IBeforeSketchDAO interface and extend DataAccessObject  and override the following  methods.
 * -saveOrUpdateBeforeSketch(JctBeforeSketchHeader entity)
 * -getList(String jobRefNo)
 * -getJson(String jobReferenceNo)
 * -getInitialJson(String jobReferenceNo)
 * -getFetchQuestion(String jobTitel)
 * -fetchQuestionAnswer(String jobRefNo)
 * <p><b>Description:</b>  This class used to perform insert, modify ,fetch data from db and delete operation. Save, update, delete operations are 
 * performed in DataAccessObject class </p>
 * <p><b>Copyrights:</b> 	All rights reserved by Interra IT and should only be used for its internal application development.</p>
 * <p><b>Revision History:</b>
 * 	<li>InterraIt - 31/Jan/2014 - Implement comments, introduce Named query, Transactional attribute </li>
 * </p>
 */
@Repository
public class BeforeSketchDAOImpl extends DataAccessObject implements IBeforeSketchDAO {

	@Autowired
	private SessionFactory sessionFactory;
	private static final Logger logger = LoggerFactory.getLogger(BeforeSketchDAOImpl.class);
	

	/**
	 * <p><b>Description:</b>  This method is used for storing Before Sketch data </p>
	 * @param It will take JctBeforeSketchHeader as a input parameter
	 * @return It returns string object
	 * @throws DAOException -it throws DAOException
	 * 
	 */
	@Transactional(propagation=Propagation.REQUIRED)
	public String saveOrUpdateBeforeSketch(JctBeforeSketchHeader entity)
			throws DAOException {
		return update(entity);
	}
	/**
	 * <p><b>Description:</b>  This method is used for storing Before Sketch data </p>
	 * @param It will take JctBeforeSketchHeaderTemp as a input parameter
	 * @return It returns string object
	 * @throws DAOException -it throws DAOException
	 * 
	 */
	@Transactional(propagation=Propagation.REQUIRED)
	public String saveOrUpdateBeforeSketchTemp(JctBeforeSketchHeaderTemp entity)
			throws DAOException {
		return save(entity);
	}
	
	/**
	 * <p><b>Description:</b>  Baserd on job reference number this method fetch all Before Sketch header</p>
	 * @param It will take jobRefNo as a input parameter
	 * @return It returns list of JctBeforeSketchHeader data from database
	 * @throws DAOException -it throws DAOException
	 * 
	 */
	@Transactional(propagation=Propagation.REQUIRED,readOnly = true)
	public List<JctBeforeSketchHeader> getList(String jobRefNo) throws DAOException {
		return sessionFactory.getCurrentSession().getNamedQuery("fetchAllBeforeSketchHeader").setParameter("refNo", jobRefNo).list();
	}
	
	@Transactional(propagation=Propagation.REQUIRED,readOnly = true)
	public List<JctBeforeSketchHeader> getListEdtd(String jobRefNo) throws DAOException {
		return sessionFactory.getCurrentSession().getNamedQuery("fetchAllBeforeSketchHeaderEdtd").setParameter("refNo", jobRefNo).list();
	}
	/**
	 * <p><b>Description:</b>  Baserd on job reference number this method fetch all Before Sketch header</p>
	 * @param It will take jobRefNo as a input parameter
	 * @return It returns list of JctBeforeSketchHeaderTemp data from database
	 * @throws DAOException -it throws DAOException
	 * 
	 */
	@Transactional(propagation=Propagation.REQUIRED,readOnly = true)
	public List<JctBeforeSketchHeaderTemp> getListTemp(String jobRefNo) throws DAOException {
		return sessionFactory.getCurrentSession().getNamedQuery("fetchAllBeforeSketchHeaderTemp").setParameter("refNo", jobRefNo).list();
	}
	
	/**
	 * <p><b>Description:</b>  Baserd on job reference number this method fetch Json</p>
	 * @param It will take jobRefNo as a input parameter
	 * @return It returns string object
	 * @throws DAOException -it throws DAOException
	 * 
	 */
	@Transactional(propagation=Propagation.REQUIRED,readOnly = true)
	public String getJson(String jobReferenceNo, String linkClicked) throws DAOException {
		String namedQuery = "fetchBKJsonByRefNo";
		/*if (linkClicked.equals("Y")) {
			namedQuery = "fetchBKJsonByRefNoEdited";
		}*/
		String value = (String)sessionFactory.getCurrentSession().getNamedQuery(namedQuery)
				.setParameter("refNo", jobReferenceNo).uniqueResult();
		if (null == value) {
			value = (String)sessionFactory.getCurrentSession().getNamedQuery("fetchBKJsonByRefNoEdited")
					.setParameter("refNo", jobReferenceNo).uniqueResult();
		}
		return value;
	}
	/**
	 * <p><b>Description:</b>  Baserd on job reference number this method fetch Json</p>
	 * @param It will take jobRefNo as a input parameter
	 * @return It returns string object
	 * @throws DAOException -it throws DAOException
	 * 
	 */
	@Transactional(propagation=Propagation.REQUIRED,readOnly = true)
	public String getJsonForEdit(String jobReferenceNo, String linkClicked) throws DAOException {
		String namedQuery = "fetchBKJsonByRefNoEdited";
		/*if (linkClicked.equals("Y")) {
			namedQuery = "fetchBKJsonByRefNoEdited";
		}*/
		return (String)sessionFactory.getCurrentSession().getNamedQuery(namedQuery)
				.setParameter("refNo", jobReferenceNo).uniqueResult();
	}
	/**
	 * <p><b>Description:</b>  Baserd on job reference number this method fetch Json</p>
	 * @param It will take jobRefNo as a input parameter
	 * @return It returns string object
	 * @throws DAOException -it throws DAOException
	 * 
	 */
	@Transactional(propagation=Propagation.REQUIRED,readOnly = true)
	public String getJsonForEditFromTemp(String jobReferenceNo, String linkClicked) throws DAOException {
		String namedQuery = "fetchBKJsonByRefNoEditedTemp";
		
		return (String)sessionFactory.getCurrentSession().getNamedQuery(namedQuery)
				.setParameter("refNo", jobReferenceNo).uniqueResult();
	}
	/**
	 * <p><b>Description:</b>   Baserd on job reference number this method fetch initila Json</p>
	 * @param It will take jobRefNo as a input parameter
	 * @return It returns string object
	 * @throws DAOException -it throws DAOException
	 * 
	 */
	@Transactional(propagation=Propagation.REQUIRED,readOnly = true)
	public String getInitialJson(String jobReferenceNo, String linkClicked) throws DAOException {
		String namedQuery = "fetchBKInitialJsonByRefNo";
		return (String)sessionFactory.getCurrentSession().getNamedQuery(namedQuery)
				.setParameter("refNo", jobReferenceNo).uniqueResult();
	}
	/**
	 * <p><b>Description:</b>   Baserd on job reference number this method fetch initila Json</p>
	 * @param It will take jobRefNo as a input parameter
	 * @return It returns string object
	 * @throws DAOException -it throws DAOException
	 * 
	 */
	@Transactional(propagation=Propagation.REQUIRED,readOnly = true)
	public String getInitialJsonFromTemp(String jobReferenceNo, String linkClicked) throws DAOException {
		String namedQuery = "fetchBKInitialJsonByRefNoTemp";
		return (String)sessionFactory.getCurrentSession().getNamedQuery(namedQuery)
				.setParameter("refNo", jobReferenceNo).uniqueResult();
	}
	/**
	 * <p><b>Description:</b>   Baserd on job reference number this method fetch initila Json</p>
	 * @param It will take jobRefNo as a input parameter
	 * @return It returns string object
	 * @throws DAOException -it throws DAOException
	 * 
	 */
	@Transactional(propagation=Propagation.REQUIRED,readOnly = true)
	public String getInitialJsonForEdit(String jobReferenceNo, String linkClicked) throws DAOException {
		String namedQuery = "fetchBKInitialJsonByRefNoEdited";
		/*if (linkClicked.equals("Y")) {
			namedQuery = "fetchBKInitialJsonByRefNoEdited";
		}*/
		return (String)sessionFactory.getCurrentSession().getNamedQuery(namedQuery)
				.setParameter("refNo", jobReferenceNo).uniqueResult();
	}
	
	/**
	 * <p><b>Description:</b>   Baserd on profileId this method fetch Before sketch question</p>
	 * @param It will take jobRefNo as a input parameter
	 * @return  It returns list of JctQuestion data from database
	 * @throws DAOException -it throws DAOException
	 * 
	 */
	@Transactional(propagation=Propagation.REQUIRED,readOnly = true)
	public List<JctQuestion> getFetchQuestion(int profileId) throws DAOException {
		return sessionFactory.getCurrentSession().getNamedQuery("fetchAllQuestionByTitle").setParameter("profileId", profileId)
				.setParameter("bsas",CommonConstants.BEFORE_SKETCH).list();
	}

	/**
	 * <p><b>Description:</b>   Baserd on job reference number this method fetch Before sketch question answar</p>
	 * @param It will take jobRefNo as a input parameter
	 * @return  It returns list of JctBeforeSketchQuestion data from database
	 * @throws DAOException -it throws DAOException
	 * 
	 */
	@Transactional(propagation=Propagation.REQUIRED,readOnly = true)
	public List<JctBeforeSketchQuestion> fetchQuestionAnswer(String jobRefNo) throws DAOException {
		return sessionFactory.getCurrentSession().getNamedQuery("fetchAllQuestionByRefNo")
		.setParameter("jobRefNo", jobRefNo)
		.setParameter("deleted", CommonConstants.ENTRY_NOT_DELETED).list();
	}

	/**
	 * <p><b>Description:</b>  This method is used for remove After Sketch header data </p>
	 * @param It will take jobRefNo as a input parameter
	 * @return  It returns string object
	 * @throws DAOException -it throws DAOException
	 * 
	 */
	@Transactional(propagation=Propagation.REQUIRED)
	public String remove(JctBeforeSketchHeader object) throws DAOException {
		logger.info(">>>> BeforeSketchDAOImpl.remove"); 
		String cal = "success";
		try{
			Delete(object);
		}catch(Exception ex){
			cal = "failure";
			logger.error(ex.getLocalizedMessage());
		}
		logger.info("<<<< BeforeSketchDAOImpl.remove"); 
		return cal;
	}
	/**
	 * <p><b>Description:</b>  This method is used for remove After Sketch header data </p>
	 * @param It will take jobRefNo as a input parameter
	 * @return  It returns string object
	 * @throws DAOException -it throws DAOException
	 * 
	 */
	@Transactional(propagation=Propagation.REQUIRED)
	public String removeTemp(JctBeforeSketchHeaderTemp object) throws DAOException {
		logger.info(">>>> BeforeSketchDAOImpl.removeTemp"); 
		String cal = "success";
		try{
			Delete(object);
		}catch(Exception ex){
			cal = "failure";
			logger.error(ex.getLocalizedMessage());
		}
		logger.info("<<<< BeforeSketchDAOImpl.removeTemp"); 
		return cal;
	}
	/**
	 * <p><b>Description:</b>  This method is used to fetch before Sketch header data for keeping previous</p>
	 * @param It will take jobRefNo as a input parameter
	 * @return  It returns string object
	 * @throws DAOException -it throws DAOException
	 * 
	 */
	@Transactional(propagation=Propagation.REQUIRED)
	public List<JctBeforeSketchHeader> getListForKeepingPrevious(String jobRefNo)
			throws DAOException {
		return sessionFactory.getCurrentSession().getNamedQuery("fetchAllBeforeSketchHeaderForPrevious").setParameter("refNo", jobRefNo).list();
	}
	/**
	 * <p><b>Description:</b>  This method is used to fetch before Sketch header data</p>
	 * @param It will take jobRefNo as a input parameter
	 * @return  It returns  List<JctBeforeSketchHeader>
	 * @throws DAOException -it throws DAOException
	 * 
	 */
	@Transactional(propagation=Propagation.REQUIRED)
	public List<JctBeforeSketchHeader> getBSMaxPageSeqHeader(String jobRefNo, int seqNos) throws DAOException {
		return sessionFactory.getCurrentSession().getNamedQuery("bsMaxPageSeq").setParameter("refNo", jobRefNo).setParameter("seqNos", seqNos).list();
	}
	/**
	 * <p><b>Description:</b>  This method is used to fetch before Sketch header data</p>
	 * @param It will take user vo as a input parameter
	 * @return  It returns  List<JctBeforeSketchHeader>
	 * @throws DAOException -it throws DAOException
	 * 
	 */
	@Transactional(propagation=Propagation.REQUIRED)
	public List<JctBeforeSketchHeader> getBeforeSketchActiveObject(UserVO userVo) throws DAOException {
		return sessionFactory.getCurrentSession().getNamedQuery("getBeforeSketchByEmailId").setParameter("createdBy", userVo.getEmail()).list();
	}
}
