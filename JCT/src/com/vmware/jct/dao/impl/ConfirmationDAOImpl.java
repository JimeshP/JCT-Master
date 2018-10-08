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
import com.vmware.jct.dao.IConfirmationDAO;
import com.vmware.jct.exception.DAOException;
import com.vmware.jct.model.JctCompletionStatus;
import com.vmware.jct.model.JctPdfRecords;
import com.vmware.jct.model.JctStatusSearch;
import com.vmware.jct.model.JctUserLoginInfo;

/**
 * 
 * <p><b>Class name:</b> ConfirmationDAOImpl.java</p>
 * <p><b>Author:</b> InterraIt</p>
 * <p><b>Purpose:</b> This is a ConfirmationDAOImpl class. This artifact is Persistence Manager layer artifact.
 * ConfirmationDAOImpl implement IConfirmationDAO interface and override the following  methods.
 * -addDiagWithoutEdits(String jobReferenceNo)
 * -addEdits(String jobReferenceNo)
 * -disableStatus(String jobReferenceNo)
 * -updateDeleteStatusForEdit(String jobReferenceNo)
 * <p><b>Description:</b> This class used to perform final confirmation level job</p>
 * <p><b>Revision History:</b>
 * </p>
 */

@Repository
public class ConfirmationDAOImpl extends DataAccessObject implements IConfirmationDAO {

	@Autowired
	private SessionFactory sessionFactory;
	
	private static final Logger logger = LoggerFactory.getLogger(ConfirmationDAOImpl.class);
	
	
	/**
	 * <p><b>Description:</b>  This method call  function updateStatus  and update soft delete and status in different transaction table </p>
	 * @param It will take jobReferenceNo as a input parameter
	 * @return  It returns string object
	 * @throws DAOException -it throws DAOException
	 * 
	 */
	@Transactional(propagation=Propagation.REQUIRED)
	public String addDiagWithoutEdits(String jobReferenceNo)
			throws DAOException {
		logger.info(">>>> ConfirmationDAOImpl.addDiagWithoutEdits"); 
		String result = "success";
		Integer successFailureFlag = 0; //0: success
		try{
			char decider = 'C';
			successFailureFlag = (Integer) sessionFactory.getCurrentSession().createSQLQuery("SELECT jct_update_status('"+jobReferenceNo+"', '"+decider+"')").uniqueResult();
			compFunctions(jobReferenceNo);
			if(successFailureFlag == 1){
				result = "failure";
			}
		} catch(Exception ex){
			//ex.printStackTrace();
			logger.error(ex.getLocalizedMessage());
			result = "failure";
		}
		logger.info("<<<< ConfirmationDAOImpl.addDiagWithoutEdits"); 
		return result;
	}
	/**
	 * <p><b>Description:</b>  This method call  function updateStatus and update soft delete only</p>
	 * @param It will take jobReferenceNo as a input parameter
	 * @return  It returns string object
	 * @throws DAOException -it throws DAOException
	 * 
	 */
	@Transactional(propagation=Propagation.REQUIRED)
	public int updateLinkStatus(String jobReferenceNo) throws DAOException {
		Integer successFailureFlag = 0; //0: success
		try{
			char decider = 'Z';
			successFailureFlag = (Integer) sessionFactory.getCurrentSession().createSQLQuery("SELECT jct_update_status('"+jobReferenceNo+"', '"+decider+"')").uniqueResult();
			
		} catch(Exception ex) {
			logger.error(ex.getLocalizedMessage());
		}
		return successFailureFlag;
	}
	/**
	 * <p><b>Description:</b>  This method call  function updateStatus and update soft delete only</p>
	 * @param It will take jobReferenceNo as a input parameter
	 * @return  It returns string object
	 * @throws DAOException -it throws DAOException
	 * 
	 */
	@Transactional(propagation=Propagation.REQUIRED)
	public int keepOriginal(String jobReferenceNo) throws DAOException {
		Integer successFailureFlag = 0; //0: success
		try{
			char decider = 'K';
			successFailureFlag = (Integer) sessionFactory.getCurrentSession().createSQLQuery("SELECT jct_update_status('"+jobReferenceNo+"', '"+decider+"')").uniqueResult();
			compFunctions(jobReferenceNo);
		} catch(Exception ex) {
			logger.error(ex.getLocalizedMessage());
		}
		return successFailureFlag;
	}
	/**
	 * <p><b>Description:</b>  This method call  function updateStatus  and update soft delete and status in different transaction table </p>
	 * @param It will take jobReferenceNo as a input parameter
	 * @return  It returns string object
	 * @throws DAOException -it throws DAOException
	 * 
	 */
	@Transactional(propagation=Propagation.REQUIRED)
	public String addEdits(String jobReferenceNo)
			throws DAOException {
		logger.info(">>>> ConfirmationDAOImpl.addEdits"); 
		String result = "success";
		Integer successFailureFlag = 0; //0: success
		char decider = 'P';
		try{
			successFailureFlag = (Integer) sessionFactory.getCurrentSession().createSQLQuery("SELECT jct_update_status('"+jobReferenceNo+"', '"+decider+"')").uniqueResult();
			compFunctions(jobReferenceNo);
			if(successFailureFlag == 1){
				result = "failure";
			}
		} catch(Exception ex){
			result = "failure";
			logger.error(ex.getLocalizedMessage());
		}
		logger.info("<<<< ConfirmationDAOImpl.addEdits"); 
		return result;
	}
	/**
	 * <p><b>Description:</b>  This method call  function populateBeforeSketch  and populateAfterSketch </p>
	 * @param jobReferenceNo
	 * @return string 
	 * @throws DAOException
	 */
	private String compFunctions(String jobReferenceNo)
			throws DAOException {
		logger.info(">>>> compFunctions.compFunctions"); 
		Integer beforeSketchReturn = (Integer)sessionFactory.getCurrentSession().createSQLQuery("SELECT jct_populate_before_sketch('"+jobReferenceNo+"')").uniqueResult();
		logger.info("FUNCTION TO POPULATE BEFORE SKETCH FOR JOB REFERENCE NUMBER: "+jobReferenceNo+" RETURNED: "+beforeSketchReturn);
		Integer afterSketchReturn = (Integer)sessionFactory.getCurrentSession().createSQLQuery("SELECT jct_populate_after_sketch('"+jobReferenceNo+"')").uniqueResult();
		logger.info("FUNCTION TO POPULATE AFTER SKETCH FOR JOB REFERENCE NUMBER: "+jobReferenceNo+" RETURNED: "+afterSketchReturn);
		logger.info("<<<< compFunctions.compFunctions"); 
		return null;
	}
	/**
	 * <p><b>Description:</b>  This method update the soft delete status based on jobReferenceNo </p>
	 * @param jobReferenceNo
	 * @return int 
	 * @throws DAOException
	 */
	public int disableStatus(String jobReferenceNo) throws DAOException {
		char decider = 'P';
		int successFailureFlag = (Integer) sessionFactory.getCurrentSession().createSQLQuery("SELECT jct_update_status('"+jobReferenceNo+"', '"+decider+"')").uniqueResult();
		compFunctions(jobReferenceNo);
		return sessionFactory.getCurrentSession().getNamedQuery("disableStatusSearch")
				.setParameter("jobRefNo", jobReferenceNo).executeUpdate();
	}
	/**
	 * <p><b>Description:</b>  This method update the soft delete status based on jobReferenceNo </p>
	 * @param jobReferenceNo
	 * @return int 
	 * @throws DAOException
	 */
	@Transactional(propagation=Propagation.REQUIRED)
	public int updateDeleteStatusForEdit(String jobReferenceNo)
			throws DAOException {
		return sessionFactory.getCurrentSession().getNamedQuery("updateSoftDeleteForEdit")
				.setParameter("jobRefNo", jobReferenceNo).executeUpdate();
	}

	/**
	 * <p><b>Description:</b>  This method is used for storing completion status data data </p>
	 * @param It will take JctBeforeSketchHeader as a input parameter
	 * @return It returns string object
	 * @throws DAOException -it throws DAOException
	 * 
	 */
	@Transactional(propagation=Propagation.REQUIRED)
	public String saveOrUpdateBeforeSketch(JctCompletionStatus entity)
			throws DAOException {
		return update(entity);
	}
	/**
	 * <p><b>Description:</b>  This method is used for storing status search data </p>
	 * @param It will take JctBeforeSketchHeaderTemp as a input parameter
	 * @return It returns string object
	 * @throws DAOException -it throws DAOException
	 * 
	 */
	@Transactional(propagation=Propagation.REQUIRED)
	public String saveToStatusSearch(JctStatusSearch entity)
			throws DAOException {
		return save(entity);
	}
	/**
	 * <p><b>Description:</b>  This method restarts the excersize based on jobReferenceNo </p>
	 * @param jobReferenceNo
	 * @return int 
	 * @throws DAOException
	 */
	@Transactional(propagation=Propagation.REQUIRED)
	public int restartExcersize(String jobReferenceNo, String distinction) throws DAOException {
		int successFailureFlag = (Integer) sessionFactory.getCurrentSession().createSQLQuery("SELECT jct_new_diagram('"+jobReferenceNo+"','"+distinction+"')").uniqueResult();
		return successFailureFlag;
	}
	/**
	 * <p><b>Description:</b>  This method fetches list of completion entity </p>
	 * @param jobReferenceNo
	 * @return list 
	 * @throws DAOException
	 */
	@Transactional(propagation=Propagation.REQUIRED)
	public List<JctCompletionStatus> getExcersizeCompletionCount (String jobRefNo) throws DAOException {
		return sessionFactory.getCurrentSession().getNamedQuery("fetchCompletionObjs")
				.setString("refNo", jobRefNo).list();
	}
	
	
	@Transactional(propagation=Propagation.REQUIRED)
	public void updateShowPDF(String jobRefNo) throws DAOException{
		logger.info(">>>> ConfirmationDAOImpl.updateShowPDF"); 
		int pk = 0;
		JctPdfRecords obj = new JctPdfRecords();		
		obj = (JctPdfRecords) sessionFactory.getCurrentSession().getNamedQuery("fetchLatestPdf")
				.setString("jobRefNo", jobRefNo).uniqueResult();
		obj.setJctShowPdf(1);
		update(obj);
		logger.info("<<<< ConfirmationDAOImpl.updateShowPDF");
	}
	
	@Transactional(propagation=Propagation.REQUIRED)
	public List<JctUserLoginInfo> getTotalTime(int userId) throws DAOException {
		return sessionFactory.getCurrentSession().getNamedQuery("selectObjByEmailId").setParameter("userId", userId).list();
	}
}
