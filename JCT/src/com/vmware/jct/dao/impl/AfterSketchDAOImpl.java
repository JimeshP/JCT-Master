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

import com.vmware.jct.common.utility.CommonConstants;
import com.vmware.jct.dao.DataAccessObject;
import com.vmware.jct.dao.IAfterSketchDAO;
import com.vmware.jct.dao.dto.QuestionearDTO;
import com.vmware.jct.exception.DAOException;
import com.vmware.jct.model.JctAfterSketchFinalpageDetail;
import com.vmware.jct.model.JctAfterSketchFinalpageDetailTemp;
import com.vmware.jct.model.JctAfterSketchHeader;
import com.vmware.jct.model.JctAfterSketchHeaderTemp;
import com.vmware.jct.model.JctAfterSketchPageoneDetail;
import com.vmware.jct.model.JctGlobalProfileHistory;
import com.vmware.jct.model.JctUserLoginInfo;
import com.vmware.jct.service.vo.UserVO;


/**
 * 
 * <p><b>Class name:</b> AfterSketchDAOImpl.java</p>
 * <p><b>Author:</b> InterraIt</p>
 * <p><b>Purpose:</b> This is a AfterSketchDAOImpl class. This artifact is Persistence Manager layer artifact.
 * AfterSketchDAOImpl implement IAfterSketchDAO interface and extend DataAccessObject  and override the following  methods.
 * -saveOrUpdatePageOne(JctAfterSketchHeader entity)
 * -getList(String jobRefNo)
 * -fetchCheckedElements(String jobRefNo)
 * -fetchASOneJson(String jobRefNo)
 * -fetchASOneJson(String jobRefNo)
 * <p><b>Description:</b>  This class used to perform insert, modify ,fetch data from db and delete operation. Save, update, delete operations are 
 * performed in DataAccessObject class </p>
 * <p><b>Copyrights:</b> 	All rights reserved by Interra IT and should only be used for its internal application development.</p>
 * <p><b>Revision History:</b>
 * 	<li>InterraIt - 31/Jan/2014 - Implement comments, introduce Named query, Transactional attribute </li>
 * </p>
 */
@Repository
public class AfterSketchDAOImpl extends DataAccessObject implements IAfterSketchDAO {

	@Autowired
	private SessionFactory sessionFactory;
	
	private static final Logger LOGGER = LoggerFactory.getLogger(AfterSketchDAOImpl.class);
	
	/**
	 * <p><b>Description:</b>  This method is used for storing After Sketch header data</p>
	 * @param It will take JctAfterSketchHeader object as a input parameter
	 * @return It returns list of all user data from database
	 * @throws DAOException -it throws DAOException
	 * 
	 */
	@Transactional(propagation=Propagation.REQUIRED)
	public String saveOrUpdatePageOne(JctAfterSketchHeader entity)
			throws DAOException {
		return save(entity);
	}
	/**
	 * <p><b>Description:</b>  This method is used for storing After Sketch header data</p>
	 * @param It will take JctAfterSketchHeader object as a input parameter
	 * @return It returns list of all user data from database
	 * @throws DAOException -it throws DAOException
	 * 
	 */
	@Transactional(propagation=Propagation.REQUIRED)
	public String updateHeaderPage(JctAfterSketchHeader entity)
			throws DAOException {
		return update(entity);
	}
	/**
	 * <p><b>Description:</b>  This method is used for storing After Sketch final page date</p>
	 * @param It will take JctAfterSketchHeader object as a input parameter
	 * @return It returns list of all user data from database
	 * @throws DAOException -it throws DAOException
	 * 
	 */
	@Transactional(propagation=Propagation.REQUIRED)
	public String updateFinalPageDetails(JctAfterSketchFinalpageDetail entity)
			throws DAOException {
		return update(entity);
	}
	
	/**
	 * <p><b>Description:</b> Baserd on job reference number this method fetch all data from jct after sketch header table</p>
	 * @param It will take jobRefNo as a input parameter
	 * @return It returns list of JctAfterSketchHeader data from database
	 * @throws DAOException -it throws DAOException
	 * 
	 */
	@Transactional(propagation=Propagation.REQUIRED,readOnly = true)
	public List<JctAfterSketchHeader> getList(String jobRefNo, String linkClicked)
			throws DAOException {
		LOGGER.info("ASDAO	----job reference number: "+jobRefNo);
		String namedQuery = "fetchAfterSketchHeader";
		/*if (linkClicked.equals("Y")) {
			namedQuery = "fetchAfterSketchHeaderEdited";
		}*/
		return sessionFactory.getCurrentSession().getNamedQuery(namedQuery)
				.setParameter("refNo", jobRefNo).list();
	}
	/**
	 * <p><b>Description:</b> Baserd on job reference number this method fetch all data from jct after sketch header table</p>
	 * @param It will take jobRefNo as a input parameter
	 * @return It returns list of JctAfterSketchHeader data from database
	 * @throws DAOException -it throws DAOException
	 * 
	 */
	@Transactional(propagation=Propagation.REQUIRED,readOnly = true)
	public List<JctAfterSketchHeader> getListForEdit(String jobRefNo, String linkClicked)
			throws DAOException {
		String namedQuery = "fetchAfterSketchHeaderEdited";
		return sessionFactory.getCurrentSession().getNamedQuery(namedQuery)
				.setParameter("refNo", jobRefNo).list();
	}
	/**
	 * <p><b>Description:</b> Baserd on job reference number this method fetch all data from jct after sketch header table</p>
	 * @param It will take jobRefNo as a input parameter
	 * @return It returns list of JctAfterSketchHeader data from database
	 * @throws DAOException -it throws DAOException
	 * 
	 */
	@Transactional(propagation=Propagation.REQUIRED,readOnly = true)
	public List<JctAfterSketchHeader> getListForKeepingPrevious(String jobRefNo)
			throws DAOException {
		return sessionFactory.getCurrentSession().getNamedQuery("fetchAllAfterSketchHeaderForPrevious").setParameter("refNo", jobRefNo).list();
	}
	
	/**
	 * <p><b>Description:</b> Baserd on job reference number this method fetch all data from jct after sketch header table</p>
	 * @param It will take jobRefNo as a input parameter
	 * @return It returns list of JctAfterSketchHeader data from database
	 * @throws DAOException -it throws DAOException
	 * 
	 */
	@Transactional(propagation=Propagation.REQUIRED,readOnly = true)
	public List<JctAfterSketchHeaderTemp> getListTemp(String jobRefNo)
			throws DAOException {
		return sessionFactory.getCurrentSession().getNamedQuery("fetchAfterSketchHeaderTemp").setParameter("refNo", jobRefNo).list();
	}
	
	/**
	 * <p><b>Description:</b> Baserd on job reference number this method fetch all data from jct after sketch detail table</p>
	 * @param It will take jobRefNo as a input parameter
	 * @return It returns list of JctAfterSketchFinalpageDetail data from database
	 * @throws DAOException -it throws DAOException
	 * 
	 */
	@Transactional(propagation=Propagation.REQUIRED,readOnly = true)
	public List<JctAfterSketchFinalpageDetail> getASFinalPageList(String jobRefNo)
			throws DAOException {
		return sessionFactory.getCurrentSession().getNamedQuery("findASFinalPageChildren").setParameter("refNo", jobRefNo).list();
	}
	/**
	 * <p><b>Description:</b> Baserd on job reference number this method fetch all data from jct after sketch detail table</p>
	 * @param It will take jobRefNo as a input parameter
	 * @return It returns list of JctAfterSketchFinalpageDetail data from database
	 * @throws DAOException -it throws DAOException
	 * 
	 */
	@Transactional(propagation=Propagation.REQUIRED,readOnly = true)
	public List<JctAfterSketchPageoneDetail> getASPageOneList(String jobRefNo)
			throws DAOException {
		return sessionFactory.getCurrentSession().getNamedQuery("findASPageOneChildren").setParameter("refNo", jobRefNo).list();
	}
	
	/**
	 * <p><b>Description:</b> Baserd on job reference number this method fetch all data from jct after sketch detail table</p>
	 * @param It will take jobRefNo as a input parameter
	 * @return It returns list of JctAfterSketchPageoneDetail data from database
	 * @throws DAOException -it throws DAOException
	 * 
	 */
	
	@Transactional(propagation=Propagation.REQUIRED,readOnly = true)
	public List<JctAfterSketchPageoneDetail> getASFirstPageList(String jobRefNo) throws DAOException {
		return sessionFactory.getCurrentSession().getNamedQuery("findASPageOneChildren").setParameter("refNo", jobRefNo).list();
	}
	
	/**
	 * <p><b>Description:</b> Baserd on job reference number this method fetch page one check value from jct after sketch header table</p>
	 * @param It will take jobRefNo as a input parameter
	 * @return It returnsstring object
	 * @throws DAOException -it throws DAOException
	 * 
	 */
	@Transactional(propagation=Propagation.REQUIRED,readOnly = true)
	public String fetchCheckedElements(String jobRefNo, String linkClicked) throws DAOException {
		String namedQuery = "findPageoneCheckedValue";
		/*if (linkClicked.equals("Y")) {
			namedQuery = "findPageoneCheckedValueEdited";
		}*/
		return (String) sessionFactory.getCurrentSession().getNamedQuery(namedQuery)
				.setParameter("jRno", jobRefNo).uniqueResult();
	}
	
	
	/**
	 * <p><b>Description:</b>  Based on job reference number this method fetch page one Json value from jct after sketch header table</p>
	 * @param It will take jobRefNo as a input parameter
	 * @return It returns string object
	 * @throws DAOException -it throws DAOException
	 * 
	 */
	@Transactional(propagation=Propagation.REQUIRED,readOnly = true)
	public String fetchASOneJson(String jobRefNo, String linkClicked) throws DAOException {
		String namedQuery = "findPageoneJson";
		/*if (linkClicked.equals("Y")) {
			namedQuery = "findPageoneJsonEdited";
		}*/
		String value = (String) sessionFactory.getCurrentSession().getNamedQuery(namedQuery)
				.setParameter("jRno", jobRefNo).uniqueResult();
		if (null == value) {
			value = (String) sessionFactory.getCurrentSession().getNamedQuery("findPageoneJsonEdited")
					.setParameter("jRno", jobRefNo).uniqueResult();
		}
		return value;
	}
	/**
	 * <p><b>Description:</b>  Based on job reference number this method fetch page one Json value from jct after sketch header table</p>
	 * @param It will take jobRefNo as a input parameter
	 * @return It returns string object
	 * @throws DAOException -it throws DAOException
	 * 
	 */
	@Transactional(propagation=Propagation.REQUIRED,readOnly = true)
	public String fetchASOneJsonForEdit(String jobRefNo, String linkClicked) throws DAOException {
		String namedQuery = "findPageoneJsonEdited";
		/*if (linkClicked.equals("Y")) {
			namedQuery = "findPageoneJsonEdited";
		}*/
		return (String) sessionFactory.getCurrentSession().getNamedQuery(namedQuery)
				.setParameter("jRno", jobRefNo).uniqueResult();
	}
	/**
	 * <p><b>Description:</b>  Based on job reference number this method fetch final page Json value from jct after sketch header table</p>
	 * @param It will take jobRefNo as a input parameter
	 * @return It returns string object
	 * @throws DAOException -it throws DAOException
	 * 
	 */
	@Transactional(propagation=Propagation.REQUIRED,readOnly = true)
	public String fetchASFinalPageJson(String jobRefNo, String linkClicked) throws DAOException {
		String namedQuery = "findFinalPageJson";
		String value = (String) sessionFactory.getCurrentSession().getNamedQuery(namedQuery)
		.setParameter("jRno", jobRefNo).uniqueResult();
		if (null == value) {
			value = (String) sessionFactory.getCurrentSession().getNamedQuery("findFinalPageJsonEdited")
					.setParameter("jRno", jobRefNo).uniqueResult();
		}
		return value;
	}
	/**
	 * <p><b>Description:</b>  Based on job reference number this method fetch final page Json value from jct after sketch header table</p>
	 * @param It will take jobRefNo as a input parameter
	 * @return It returns string object
	 * @throws DAOException -it throws DAOException
	 * 
	 */
	@Transactional(propagation=Propagation.REQUIRED,readOnly = true)
	public String fetchASFinalPageJsonForEdit(String jobRefNo, String linkClicked) throws DAOException {
		String namedQuery = "findFinalPageJsonEdited";
		return (String) sessionFactory.getCurrentSession().getNamedQuery(namedQuery)
				.setParameter("jRno", jobRefNo).uniqueResult();
	}
	/**
	 * <p><b>Description:</b>  This method is used for remove After Sketch header data </p>
	 * @param It will take JctAfterSketchHeader as a input parameter
	 * @return It returns string object
	 * @throws DAOException -it throws DAOException
	 * 
	 */
	@Transactional(propagation=Propagation.REQUIRED)
	public String remove(JctAfterSketchHeader object) throws DAOException {
		LOGGER.info(">>>> AfterSketchDAOImpl.remove");
		String cal = "success";
		try{
			Delete(object);
		}catch(Exception ex){
			cal = "failure";
			LOGGER.error(ex.getLocalizedMessage());
		}
		LOGGER.info("<<<< AfterSketchDAOImpl.remove");
		return cal;
	}
	/**
	 * <p><b>Description:</b>  This method is used for remove After Sketch header data </p>
	 * @param It will take JctAfterSketchHeader as a input parameter
	 * @return It returns string object
	 * @throws DAOException -it throws DAOException
	 * 
	 */
	@Transactional(propagation=Propagation.REQUIRED)
	public String removeTemp(JctAfterSketchHeaderTemp object) throws DAOException {
		LOGGER.info(">>>> AfterSketchDAOImpl.removeTemp");
		String cal = "success";
		try{
			Delete(object);
		}catch(Exception ex){
			cal = "failure";
			LOGGER.error(ex.getLocalizedMessage());
		}
		LOGGER.info("<<<< AfterSketchDAOImpl.removeTemp");
		return cal;
	}
	
	/**
	 * <p><b>Description:</b>  This method is used for remove After Sketch details data </p>
	 * @param It will take JctAfterSketchFinalpageDetail as a input parameter
	 * @return It returns string object
	 * @throws DAOException -it throws DAOException
	 * 
	 */
	@Transactional(propagation=Propagation.REQUIRED)
	public String removeDetails(JctAfterSketchFinalpageDetail object) throws DAOException {
		LOGGER.info(">>>> AfterSketchDAOImpl.removeDetails");
		String cal = "success";
		try{
			Delete(object);
		}catch(Exception ex){
			cal = "failure";
			LOGGER.error(ex.getLocalizedMessage());
		}
		LOGGER.info("<<<< AfterSketchDAOImpl.removeDetails");
		return cal;
	}
	
	/**
	 * <p><b>Description:</b>  This method is used for remove After Sketch details data </p>
	 * @param It will take JctAfterSketchFinalpageDetail as a input parameter
	 * @return It returns string object
	 * @throws DAOException -it throws DAOException
	 * 
	 */
	@Transactional(propagation=Propagation.REQUIRED)
	public String removeDetailsTemp(JctAfterSketchFinalpageDetailTemp object) throws DAOException {
		LOGGER.info(">>>> AfterSketchDAOImpl.removeDetailsTemp");
		String cal = "success";
		try{
			Delete(object);
		}catch(Exception ex){
			cal = "failure";
			LOGGER.error(ex.getLocalizedMessage());
		}
		LOGGER.info("<<<< AfterSketchDAOImpl.removeDetailsTemp");
		return cal;
	}
	
	/**
	 * <p><b>Description:</b>  This method is used for remove After Sketch page one details data </p>
	 * @param It will take JctAfterSketchPageoneDetail as a input parameter
	 * @return It returns string object
	 * @throws DAOException -it throws DAOException
	 * 
	 */
	@Transactional(propagation=Propagation.REQUIRED)
	public String removePageOneDetails(JctAfterSketchPageoneDetail object) throws DAOException {
		LOGGER.info(">>>> AfterSketchDAOImpl.removePageOneDetails");
		String cal = "success";
		try{
			Delete(object);
		}catch(Exception ex){
			cal = "failure";
			LOGGER.error(ex.getLocalizedMessage());
		}
		LOGGER.info("<<<< AfterSketchDAOImpl.removePageOneDetails");
		return cal;
	}

	
	/**
	 * <p><b>Description:</b>  Based on job reference number this method fetches fetches List of passion count</p>
	 * @param It will take jobRefNo as a input parameter
	 * @return It returns Integer object
	 * @throws DAOException -it throws DAOException
	 * 
	 */
	@Transactional(propagation=Propagation.REQUIRED,readOnly = true)
	public List<Integer> getPassionCount(String jobRefNo) throws DAOException {
		return (List<Integer>)sessionFactory.getCurrentSession().getNamedQuery("findPassionCount").setParameter("jRno", jobRefNo).list();
	}
	
	/**
	 * <p><b>Description:</b>  Based on job reference number this method fetches List of value count</p>
	 * @param It will take jobRefNo as a input parameter
	 * @return It returns Integer object
	 * @throws DAOException -it throws DAOException
	 * 
	 */
	@Transactional(propagation=Propagation.REQUIRED,readOnly = true)
	public List<Integer> getValueCount(String jobRefNo) throws DAOException {
		return (List<Integer>)sessionFactory.getCurrentSession().getNamedQuery("findValueCount").setParameter("jRno", jobRefNo).list();
	}
	/**
	 * <p><b>Description:</b>  Based on job reference number this method fetches List of strength count</p>
	 * @param It will take jobRefNo as a input parameter
	 * @return It returns Integer object
	 * @throws DAOException -it throws DAOException
	 * 
	 */
	@Transactional(propagation=Propagation.REQUIRED,readOnly = true)
	public List<Integer> getStrengthCount(String jobRefNo) throws DAOException {
		return (List<Integer>)sessionFactory.getCurrentSession().getNamedQuery("findStrengthCount").setParameter("jRno", jobRefNo).list();
	}
	
	/**
	 * <p><b>Description:</b> This method is used for update existing After Sketch header data and insert new 
	 *final page rows</p>
	 * @param It will take JctAfterSketchHeader as a input parameter
	 * @return It returns Integer object
	 * @throws DAOException -it throws DAOException
	 * 
	 */
	@Transactional(propagation=Propagation.REQUIRED)
	public String mergeAfterSketchDiag(JctAfterSketchHeader entity) throws DAOException{
		return update(entity);
	}
	
	
	/**
	 * <p><b>Description:</b> Based on job level  this method fetch after sketch action plan question</p>
	 * @param It will take profileId as a input parameter
	 * @return It returns Integer object
	 * @throws DAOException -it throws DAOException
	 * 
	 */
	@Transactional(propagation=Propagation.REQUIRED,readOnly = true)
	public List<QuestionearDTO> getFetchQuestion(int profileId) throws DAOException {
		return sessionFactory.getCurrentSession().getNamedQuery("fetchAllActionPlanQuestionByTitle").setParameter("profileId", profileId)
				.setParameter("bsas",CommonConstants.AFTER_SKETCH).list();
	}
	
	/**
	 * <p><b>Description:</b>  Based on job ref no this method calculate the difference between before sketch and after sketch task</p>
	 * @param It will take jobReferenceNos as a input parameter
	 * @return It returns Integer object
	 * @throws DAOException -it throws DAOException
	 * 
	 */
	@Transactional(propagation=Propagation.REQUIRED,readOnly = true)
	public Integer calculateDifference(String jobReferenceNos) throws DAOException {
		return (Integer) sessionFactory.getCurrentSession().createSQLQuery("SELECT jct_before_sketch_to_after_sketch('"+jobReferenceNos+"')").uniqueResult();
	}
	
	/**
	 * <p><b>Description:</b>  Based on job ref no this method changes the status to Completed</p>
	 * @param It will take jobReferenceNos as a input parameter
	 * @return It returns string object
	 * @throws DAOException -it throws DAOException
	 * 
	 */
	@Transactional(propagation=Propagation.REQUIRED)
	public String changeStatus(String jobReferenceNos) throws DAOException {
		LOGGER.info(">>>> AfterSketchDAOImpl.changeStatus");
		String successFlag = "failed";
		try{
			sessionFactory.getCurrentSession().getNamedQuery("updateCompletedStatus").setParameter("jrNo",jobReferenceNos).executeUpdate();
			successFlag = "success";
		} catch(Exception ex){
			LOGGER.error(ex.getLocalizedMessage());
			throw new DAOException();
		}
		LOGGER.info("<<<< AfterSketchDAOImpl.changeStatus");
		return successFlag;
	}
	
	/**
	 * <p><b>Description:</b>  Based on job ref no this method fetch after sketch final diagrams total task element count</p>
	 * @param It will take jobReferenceNos as a input parameter
	 * @return It returns integer object
	 * @throws DAOException -it throws DAOException
	 * 
	 */
	@Transactional(propagation=Propagation.REQUIRED)
	public Integer getAsTotalCount(String jobRefNo, String linkClicked) throws DAOException {
		String namedQuery = "findAsTotalCount";
		/*if (linkClicked.equals("Y")) {
			namedQuery = "findAsTotalCountEdited";
		}*/
		Integer value = (Integer)sessionFactory.getCurrentSession().getNamedQuery(namedQuery)
				.setParameter("jRno",jobRefNo).uniqueResult();
		if (null == value) {
			value = (Integer)sessionFactory.getCurrentSession().getNamedQuery("findAsTotalCountEdited")
					.setParameter("jRno",jobRefNo).uniqueResult();
		}
		return value;
	}
	/**
	 * <p><b>Description:</b>  Based on job ref no this method fetch after sketch final diagrams total task element count</p>
	 * @param It will take jobReferenceNos as a input parameter
	 * @return It returns integer object
	 * @throws DAOException -it throws DAOException
	 * 
	 */
	@Transactional(propagation=Propagation.REQUIRED)
	public Integer getAsTotalCountForEdit(String jobRefNo, String linkClicked) throws DAOException {
		String namedQuery = "findAsTotalCountEdited";
		return (Integer)sessionFactory.getCurrentSession().getNamedQuery(namedQuery)
				.setParameter("jRno",jobRefNo).uniqueResult();
	}
	/**
	 * <p><b>Description:</b>  Based on job ref no this method fetch after sketch final page's initial jon</p>
	 * @param It will take jobReferenceNos as a input parameter
	 * @return It returns string object
	 * @throws DAOException -it throws DAOException
	 * 
	 */
	@Transactional(propagation=Propagation.REQUIRED)
	public String getFinalPageInitialJson(String jobRefNo, String linkClicked) throws DAOException {
		String namedQuery = "fetchFinalPageInitialJson";
		/*if (linkClicked.equals("Y")) {
			namedQuery = "fetchFinalPageInitialJsonEdited";
		}*/
		String value = (String) sessionFactory.getCurrentSession().getNamedQuery(namedQuery)
				.setParameter("jRno", jobRefNo).uniqueResult();
		if (null == value) {
			value = (String) sessionFactory.getCurrentSession().getNamedQuery("fetchFinalPageInitialJsonEdited")
					.setParameter("jRno", jobRefNo).uniqueResult();
		}
		return value;
	}
	
	/**
	 * <p><b>Description:</b>  Based on job ref no this method fetch after sketch final page's initial jon</p>
	 * @param It will take jobReferenceNos as a input parameter
	 * @return It returns string object
	 * @throws DAOException -it throws DAOException
	 * 
	 */
	@Transactional(propagation=Propagation.REQUIRED)
	public String getFinalPageInitialJsonForEdit(String jobRefNo, String linkClicked) throws DAOException {
		String namedQuery = "fetchFinalPageInitialJsonEdited";
		/*if (linkClicked.equals("Y")) {
			namedQuery = "fetchFinalPageInitialJsonEdited";
		}*/
		return (String) sessionFactory.getCurrentSession().getNamedQuery(namedQuery)
				.setParameter("jRno", jobRefNo).uniqueResult();
	}

	/**
	 * <p><b>Description:</b>  Based on job ref no and task id  this method fetch before sketch and after sketch difference</p>
	 * @param It will take jobRefNo and taskId as a input parameter
	 * @return It returns string object
	 * @throws DAOException -it throws DAOException
	 * 
	 */
	@Transactional(propagation=Propagation.REQUIRED)
	public List<Object> getBsToAsDifference(String jobRefNo, int taskId)
			throws DAOException {
		StringBuilder sb = new StringBuilder("");
		sb.append("SELECT bstoas.jct_diff_energy, bstoas.jct_bs_to_as_task_desc_diff, bstoas.jct_diff_status from jct_bs_to_as bstoas where "
				+ "bstoas.jct_jobref_no =:jrno and bstoas.jct_task_id =:taskId and bstoas.jct_soft_delete = 0");
		return sessionFactory.getCurrentSession().createSQLQuery(sb.toString()).setParameter("jrno", jobRefNo).setParameter("taskId", taskId).list();
	}
	/**
	 * <p><b>Description:</b>  Based on task id  this method will fetch after sketch role frames</p>
	 * @param It will take taskId as a input parameter
	 * @return It returns string object
	 * @throws DAOException -it throws DAOException
	 * 
	 */
	@Transactional(propagation=Propagation.REQUIRED)
	public List<String> getRoleFrames(int taskId, String jobReferenceNumber) throws DAOException {
		StringBuilder sb = new StringBuilder("");
		sb.append("SELECT dtls.jct_as_role_desc from jct_after_sketch_finalpage_details dtls where "
				+ "dtls.jct_as_jobref_no =:jrno and dtls.jct_as_task_id =:taskId and dtls.jct_as_status_id=4 and dtls.jct_as_soft_delete=0");
		return sessionFactory.getCurrentSession().createSQLQuery(sb.toString()).setParameter("jrno", jobReferenceNumber).setParameter("taskId", taskId).list();
	}
	/**
	 * <p><b>Description:</b>  Based on job ref no  this method fetches tasks which are deleted</p>
	 * @param It will take jobRefNo and taskId as a input parameter
	 * @return It returns string object
	 * @throws DAOException -it throws DAOException
	 * 
	 */
	@Transactional(propagation=Propagation.REQUIRED)
	public List<Object> getBsToAsDeletedTask(String jobRefNo) throws DAOException {
		StringBuilder sb = new StringBuilder("");
		sb.append("SELECT bstoas.jct_bs_task_desc, bstoas.jct_bs_energy from jct_bs_to_as bstoas where "
				+ "bstoas.jct_jobref_no =:jrno and bstoas.jct_as_energy =0 and bstoas.jct_soft_delete = 0 ");
		return sessionFactory.getCurrentSession().createSQLQuery(sb.toString()).setParameter("jrno", jobRefNo).list();
	}
	/**
	 * <p><b>Description:</b>  Baserd on job reference number this method fetch all data from jct after sketch header table</p>
	 * @param It will take jobRefNo as a input parameter
	 * @return It returns List of JctAfterSketchFinalpageDetail data from db
	 * @throws DAOException -it throws DAOException
	 * 
	 */
	@Transactional(propagation=Propagation.REQUIRED,readOnly = true)
	public List<JctAfterSketchFinalpageDetail> getASFinalPageListDescOrder(String jobRefNo)
			throws DAOException {
		return sessionFactory.getCurrentSession().getNamedQuery("findASFinalPageChildrenDescOrder").setParameter("refNo", jobRefNo).list();
	}
	/**
	 * <p><b>Description:</b>  Baserd on job reference number this method fetch all distict role frames</p>
	 * @param It will take jobRefNo as a input parameter
	 * @return It returns List of String data from db
	 * @throws DAOException -it throws DAOException
	 * 
	 */
	@Transactional(propagation=Propagation.REQUIRED,readOnly = true)
	public List<String> getDistinctRoleFrames(String jobRefNo) throws DAOException {
		return sessionFactory.getCurrentSession().getNamedQuery("findASDistinctRoleFrames").setParameter("refNo", jobRefNo).list();
	}
	/**
	 * <p><b>Description:</b>  Baserd on job reference number this method fetch all data from jct after sketch header table</p>
	 * @param It will take profileId as a input parameter
	 * @return It returns List of String
	 * @throws DAOException -it throws DAOException
	 * 
	 */
	@Transactional(propagation=Propagation.REQUIRED,readOnly = true)
	public List<String> getActionPlanMainQtnList(int profileId) throws DAOException {
		return sessionFactory.getCurrentSession().getNamedQuery("getActionPlanMainQtnListByProfile")
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
	public List<String> getActionPlanSubQtnList(String mainQtn, int profileId) throws DAOException {
		return sessionFactory.getCurrentSession().getNamedQuery("getActionPlanSubQtnList")
				.setParameter("mainQtn", mainQtn)
				.setParameter("profileId", profileId).list();
	}
	/**
	 * <p><b>Description:</b> Baserd on job reference number this method fetch all data from jct after sketch detail table</p>
	 * @param It will take jobRefNo as a input parameter
	 * @return It returns list of JctAfterSketchFinalpageDetail data from database
	 * @throws DAOException -it throws DAOException
	 * 
	 */
	@Transactional(propagation=Propagation.REQUIRED,readOnly = true)
	public List<JctAfterSketchFinalpageDetailTemp> getASFinalPageListTemp(String jobRefNo)
			throws DAOException {
		return sessionFactory.getCurrentSession().getNamedQuery("findASFinalPageChildrenTemp").setParameter("refNo", jobRefNo).list();
	}
	/**
	 * <p><b>Description:</b>  This method is used for storing After Sketch header data</p>
	 * @param It will take JctAfterSketchHeaderTemp object as a input parameter
	 * @return It returns list of all user data from database
	 * @throws DAOException -it throws DAOException
	 * 
	 */
	@Transactional(propagation=Propagation.REQUIRED)
	public String saveOrUpdateAsHeaderTemp(JctAfterSketchHeaderTemp entity)
			throws DAOException {
		return save(entity);
	}
	/**
	 * <p><b>Description:</b>  This method is used to fetch before Sketch header data</p>
	 * @param It will take jobRefNo as a input parameter
	 * @return  It returns string object
	 * @throws DAOException -it throws DAOException
	 * 
	 */
	public List<JctAfterSketchHeader> getASMaxPageSeqHeader(String jobRefNo, int seqNos) throws DAOException {
		return sessionFactory.getCurrentSession().getNamedQuery("asMaxPageSeq").setParameter("refNo", jobRefNo).setParameter("seqNos", seqNos).list();
	}
	/**
	 * <p><b>Description:</b>  This method is used to fetch after Sketch header data</p>
	 * @param It will take user vo as a input parameter
	 * @return  It returns  List<JctBeforeSketchHeader>
	 * @throws DAOException -it throws DAOException
	 * 
	 */
	@Transactional(propagation=Propagation.REQUIRED)
	public List<JctAfterSketchHeader> getAfterSketchActiveObject(UserVO userVo) throws DAOException {
		return sessionFactory.getCurrentSession().getNamedQuery("getAfterSketchByEmailId").setParameter("createdBy", userVo.getEmail()).list();
	}
	/**
	 * Method gives distinct role name
	 */
	@Transactional(propagation=Propagation.REQUIRED)
	public List<String> getDistinctRoleName(String jobRefNo, String taskDesc)
			throws DAOException {
		return sessionFactory.getCurrentSession().getNamedQuery("getDistinctRoleName")
				.setParameter("jobRefNo", jobRefNo)
				.setParameter("taskDesc", taskDesc).list();
	}
	@Transactional(propagation=Propagation.REQUIRED)
	public String getFrozenAttributeById(int attributeId) throws DAOException {
		return (String) sessionFactory.getCurrentSession().getNamedQuery("getFrozenAttributeById")
				.setParameter("attributeId", attributeId).uniqueResult();
	}
	@Transactional(propagation=Propagation.REQUIRED)
	public List<JctUserLoginInfo> getLoginInfoList(String jobRefNo)
			throws DAOException {
		return sessionFactory.getCurrentSession().getNamedQuery("getLoginInfoListForGPAChange")
				.setParameter("jobRefNo", jobRefNo)
				.setParameter("jobRefNo2", jobRefNo).list();
	}
	@Transactional(propagation=Propagation.REQUIRED)
	public List<JctGlobalProfileHistory> getGlobalProfileChangedData(
			Date headerDate) throws DAOException {
		JctGlobalProfileHistory hist = new JctGlobalProfileHistory();
		hist.setJctGlobalProfileCreatedTimestamp(new Date());
		return sessionFactory.getCurrentSession().getNamedQuery("getGlobalAttributeChangesPageOneByTimeStamp")
				.setParameter("date1", headerDate)
				.setParameter("date2", hist.getJctGlobalProfileCreatedTimestamp()).list();
	}
}
