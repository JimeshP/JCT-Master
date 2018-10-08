package com.vmware.jct.dao.impl;

import java.util.List;

import org.hibernate.SessionFactory;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Repository;
import org.springframework.transaction.annotation.Propagation;
import org.springframework.transaction.annotation.Transactional;

import com.vmware.jct.dao.ISketchSearchDAO;
import com.vmware.jct.dao.dto.SearchPdfDTO;
import com.vmware.jct.dao.dto.StatusSearchDTO;
import com.vmware.jct.exception.DAOException;
import com.vmware.jct.model.JctAfterSketchHeader;
import com.vmware.jct.model.JctBeforeSketchHeader;
import com.vmware.jct.model.JctCompletionStatus;
/**
 * 
 * <p><b>Class name:</b> SketchSearchDAOImpl.java</p>
 * <p><b>Author:</b> InterraIt</p>
 * <p><b>Purpose:</b> This is a SketchSearchDAOImpl class. This artifact is Persistence Manager layer artifact.
 * SketchSearchDAOImpl implement ISketchSearchDAO interface and override the following  methods.
 * -getBeforeSketchList(String jobTitle, int firstResult,int maxResults)
 * -getAfterSketchList(String jobTitle, int firstResult,int maxResults)
 * <p><b>Description:</b> This class used to perform search functionality </p>
 * <p><b>Revision History:</b>
 * </p>
 */

@Repository
public class SketchSearchDAOImpl implements ISketchSearchDAO {

	@Autowired
	private SessionFactory sessionFactory;
	
	private static final Logger logger = LoggerFactory.getLogger(QuestionnaireDAOImpl.class);
	
	
	/**
	 * <p><b>Description:</b>Method searches before sketch completed diagrams</p>
	 * @param It will take jobTitle,firstResult and maxResults as a input parameter
	 * @return  It returns list of string object
	 * @throws DAOException -it throws DAOException
	 * 
	 */
	@Transactional(propagation = Propagation.REQUIRED)
	public List<StatusSearchDTO> getBeforeSketchList(String jobTitle, int firstResult,
			int maxResults) throws DAOException {
		return sessionFactory.getCurrentSession()
				.getNamedQuery("searchBeforeSketchDiag")
				.setParameter("jobTitle", jobTitle).list();
	}
	
	/**
	 * <p><b>Description:</b>Method searches after sketch completed diagrams</p>
	 * @param It will take jobTitle,firstResult and maxResults as a input parameter
	 * @return  It returns list of string object
	 * @throws DAOException -it throws DAOException
	 * 
	 */
	@Transactional(propagation = Propagation.REQUIRED)
	public List<StatusSearchDTO> getAfterSketchList(String jobTitle, int firstResult,
			int maxResults) throws DAOException {
		return sessionFactory.getCurrentSession()
				.getNamedQuery("searchAfterSketchDiag")
				.setParameter("jobTitle", jobTitle).list();
	}
	/**
	 * <p><b>Description:</b>This Method fetch before sketch jeson from before sketch detail table</p>
	 * @param It will take snapShotString, and jobTitle as a input parameter
	 * @return  It returns  string object
	 * @throws DAOException -it throws DAOException
	 * 
	 */
	@Transactional(propagation = Propagation.REQUIRED)
	public String getBeforeSketchJrno(String jrno, String snapShotString, String jobTitle)
			throws DAOException {
		return (String) sessionFactory.getCurrentSession()
				.getNamedQuery("searchBeforeSketchJRNO")
				.setParameter("jrno", jrno)
				.setParameter("jobTitle", jobTitle)
				.setParameter("snpShtStr", snapShotString).uniqueResult();
	}
	
	/**
	 * <p><b>Description:</b>This Method fetch after sketch jeson from before sketch detail table</p>
	 * @param It will take snapShotString, and jobTitle as a input parameter
	 * @return  It returns  string object
	 * @throws DAOException -it throws DAOException
	 * 
	 */
	
	@Transactional(propagation = Propagation.REQUIRED)
	public String getAfterSketchJrno(String jrno, String snapShotString, String jobTitle)
			throws DAOException {
		return (String) sessionFactory.getCurrentSession()
				.getNamedQuery("searchAfterSketchJRNO")
				.setParameter("jobTitle", jobTitle)
				.setParameter("snpShtStr", snapShotString).uniqueResult();
	}
	/**
	 * <p><b>Description:</b>This Method fetch status search data from after sketch table if it is avaliable</p>
	 * @param It will take jobReference, and jobTitle as a input parameter
	 * @return  It returns  string object
	 * @throws DAOException -it throws DAOException
	 * 
	 */
	@Transactional(propagation = Propagation.REQUIRED)
	public String isASAvailable(String jobReference, String jobTitle)
			throws DAOException {
		logger.info(">>>> SketchSearchDAOImpl.isASAvailable"); 
		String available = "Y";
		List asAvailableList = null;
		try {
			asAvailableList = sessionFactory.getCurrentSession()
			.getNamedQuery("searchASAvailableByJRNO")
			.setParameter("jctJobrefNo", jobReference)
			.setParameter("jobTitle", jobTitle).list();
			if (asAvailableList.size() == 0) {
				available = "N";
			}
			
		} catch (Exception ex) {
			logger.error(ex.getLocalizedMessage());
		}
		logger.info("<<<< SketchSearchDAOImpl.isASAvailable"); 
		return available;
	}
	
	/**
	 * <p><b>Description:</b>This Method fetch status search data from before sketch table if it is avaliable</p>
	 * @param It will take jobReference, and jobTitle as a input parameter
	 * @return  It returns  string object
	 * @throws DAOException -it throws DAOException
	 * 
	 */
	@Transactional(propagation = Propagation.REQUIRED)
	public String isBSAvailable(String jobReference, String jobTitle)
			throws DAOException {
		logger.info(">>>> SketchSearchDAOImpl.isBSAvailable"); 
		String available = "Y";
		List asAvailableList = null;
		try {
			asAvailableList = sessionFactory.getCurrentSession()
			.getNamedQuery("searchBSAvailableByJRNO")
			.setParameter("jctJobrefNo", jobReference)
			.setParameter("jobTitle", jobTitle).list();
			if (asAvailableList.size() == 0) {
				available = "N";
			}
			
		} catch (Exception ex) {
			logger.error(ex.getLocalizedMessage());
		}
		logger.info("<<<< SketchSearchDAOImpl.isBSAvailable"); 
		return available;
	}

	/**
	 * <p><b>Description:</b>This Method fetch after sketch snap short from after sketch header table</p>
	 * @param It will take jobReference as a input parameter
	 * @return  It returns  string object
	 * @throws DAOException -it throws DAOException
	 * 
	 */
	@Transactional(propagation = Propagation.REQUIRED)
	public String getASDiagram(String jobReference) throws DAOException {
		return (String) sessionFactory.getCurrentSession()
				.getNamedQuery("getBSDiagramByJRNO")
				.setParameter("jctJobrefNo", jobReference).uniqueResult();
	}
	/**
	 * <p><b>Description:</b>This Method fetch before sketch snap short from after sketch header table</p>
	 * @param It will take jobReference as a input parameter
	 * @return  It returns  string object
	 * @throws DAOException -it throws DAOException
	 * 
	 */
	@Transactional(propagation = Propagation.REQUIRED)
	public List<JctAfterSketchHeader> getPrevASDiagram(String jobReference) throws DAOException {
		return sessionFactory.getCurrentSession()
				.getNamedQuery("getPrevASDiagramByJRNO")
				.setParameter("jctJobrefNo", jobReference).list();
	}
	/**
	 * <p><b>Description:</b>This Method fetch before sketch snap short from after sketch header table</p>
	 * @param It will take jobReference as a input parameter
	 * @return  It returns  string object
	 * @throws DAOException -it throws DAOException
	 * 
	 */
	@Transactional(propagation = Propagation.REQUIRED)
	public String getPrevASDiagramForView(String jobReference, int softDel) throws DAOException {
		return (String) sessionFactory.getCurrentSession()
				.getNamedQuery("getPrevAsForView")
				.setParameter("jctJobrefNo", jobReference)
				.setParameter("jctJobrefNo2", jobReference).uniqueResult();
	}
	/**
	 * <p><b>Description:</b>This Method fetch before sketch snap short from after sketch header table</p>
	 * @param It will take jobReference as a input parameter
	 * @return  It returns  string object
	 * @throws DAOException -it throws DAOException
	 * 
	 */
	@Transactional(propagation = Propagation.REQUIRED)
	public String getBSDiagram(String jobReference) throws DAOException {
		return (String) sessionFactory.getCurrentSession()
				.getNamedQuery("getASDiagramByJRNO")
				.setParameter("jctJobrefNo", jobReference).uniqueResult();
	}
	/**
	 * <p><b>Description:</b>This Method fetch before sketch snap short from after sketch header table</p>
	 * @param It will take jobReference as a input parameter
	 * @return  It returns  string object
	 * @throws DAOException -it throws DAOException
	 * 
	 */
	@Transactional(propagation = Propagation.REQUIRED)
	public List<JctBeforeSketchHeader> getPrevBSDiagram(String jobReference) throws DAOException {
		return sessionFactory.getCurrentSession()
				.getNamedQuery("getPrevBSDiagramByJRNO")
				.setParameter("jctJobrefNo", jobReference).list();
	}
	/**
	 * <p><b>Description:</b>This Method fetch before sketch snap short from after sketch header table</p>
	 * @param It will take jobReference as a input parameter
	 * @return  It returns  string object
	 * @throws DAOException -it throws DAOException
	 * 
	 */
	@Transactional(propagation = Propagation.REQUIRED)
	public String getPrevBSDiagramForView(String jobReference, int softDel) throws DAOException {
		return (String) sessionFactory.getCurrentSession()
				.getNamedQuery("getPrevBsForView")
				.setParameter("jctJobrefNo", jobReference)
				.setParameter("jctJobrefNo2", jobReference).uniqueResult();
	}
	/**
	 * <p><b>Description:</b>This Method fetch completion status details</p>
	 * @param It will take jobReference as a input parameter
	 * @return  It returns  list of completions status objects
	 * @throws DAOException -it throws DAOException
	 * 
	 */
	@Transactional(propagation = Propagation.REQUIRED)
	public List<JctCompletionStatus> getCompletionStatusObjs(String jobReference)
			throws DAOException {
		return sessionFactory.getCurrentSession()
				.getNamedQuery("fetchCompletionObjs")
				.setParameter("refNo", jobReference).list();
	}
	/**
	 * <p><b>Description:</b>Method searches list of</p>
	 * @param It will take jobRefNo
	 * @return  It returns list of JctStatus object
	 * @throws DAOException -it throws DAOException
	 * 
	 */
	@Transactional(propagation = Propagation.REQUIRED)
	public List<String> fetchNosOfTimesShared(String jobRefNo) throws DAOException {
		return sessionFactory.getCurrentSession()
				.getNamedQuery("getNosOfTimesShared")
				.setParameter("jctJobrefNo", jobRefNo).list();
	}

	@Transactional(propagation = Propagation.REQUIRED)
	public String getOccupationDesc(String onetTitle) throws DAOException {
		logger.info(">>>> SketchSearchDAOImpl.getOccupationDesc"); 
		return (String) sessionFactory.getCurrentSession()
				.getNamedQuery("getOccupationDesc")
				.setParameter("onetTitle", onetTitle).uniqueResult();
	}
	
	/**
	 * <p><b>Description:</b>Method searches before sketch completed diagrams</p>
	 * @param It will take jobTitle,firstResult and maxResults as a input parameter
	 * @return  It returns list of string object
	 * @throws DAOException -it throws DAOException
	 * 
	 */
	@Transactional(propagation = Propagation.REQUIRED)
	public List<StatusSearchDTO> getBeforeSketchList(int firstResult,
			int maxResults) throws DAOException {
		return sessionFactory.getCurrentSession()
				.getNamedQuery("searchBeforeSketchDiagByOccupation")
				.setFirstResult(firstResult).setMaxResults(9).list();
	}
	
	/**
	 * <p><b>Description:</b>Method searches before sketch completed diagrams</p>
	 * @param It will take jobTitle,firstResult and maxResults as a input parameter
	 * @return  It returns list of string object
	 * @throws DAOException -it throws DAOException
	 * 
	 */
	@Transactional(propagation = Propagation.REQUIRED)
	public List<StatusSearchDTO> getAfterSketchList(int firstResult,
			int maxResults) throws DAOException {
		return sessionFactory.getCurrentSession()
				.getNamedQuery("searchAfterSketchDiagByOccupation")
				.setFirstResult(firstResult).setMaxResults(9).list();
	}
	
	@Transactional(propagation = Propagation.REQUIRED)
	public Long fetchAllDiagramCount(String distinction) throws DAOException {
		Long count = 0l;
		if (distinction.trim().equals("BS")) {
			count = (Long) sessionFactory.getCurrentSession().getNamedQuery("fetchAllBSDiagramCount").uniqueResult();
		} else {
			count = (Long) sessionFactory.getCurrentSession().getNamedQuery("fetchAllASDiagramCount").uniqueResult();
		}
		return count;
	}
	
	@Transactional(propagation = Propagation.REQUIRED)
	public Long fetchAllDiagramCount(String distinction, String code) throws DAOException {
		Long count = 0l;
		if (distinction.trim().equals("BS")) {
			count = (Long) sessionFactory.getCurrentSession().getNamedQuery("fetchAllBSDiagramCountCode").setParameter("code", code).uniqueResult();
		} else {
			count = (Long) sessionFactory.getCurrentSession().getNamedQuery("fetchAllASDiagramCountCode").setParameter("code", code).uniqueResult();
		}
		return count;
	}
	
	@Transactional(propagation = Propagation.REQUIRED)
	public List<StatusSearchDTO> getBeforeSketchList(int firstResult,
			int maxResults, String code) throws DAOException {
		return sessionFactory.getCurrentSession()
				.getNamedQuery("searchBeforeSketchDiagByOccupationCode")
				.setParameter("code", code)
				.setFirstResult(firstResult).setMaxResults(9).list();
	}
	
	/**
	 * <p><b>Description:</b>Method searches before sketch completed diagrams</p>
	 * @param It will take jobTitle,firstResult and maxResults as a input parameter
	 * @return  It returns list of string object
	 * @throws DAOException -it throws DAOException
	 * 
	 */
	@Transactional(propagation = Propagation.REQUIRED)
	public List<StatusSearchDTO> getAfterSketchList(int firstResult,
			int maxResults, String code) throws DAOException {
		return sessionFactory.getCurrentSession()
				.getNamedQuery("searchAfterSketchDiagByOccupationCode")
				.setParameter("code", code)
				.setFirstResult(firstResult).setMaxResults(9).list();
	}

	@Override
	@Transactional(propagation = Propagation.REQUIRED)
	public List<SearchPdfDTO> fetchOldPdf(int userId)
			throws DAOException {		
		return sessionFactory.getCurrentSession().getNamedQuery("fetchOldPdf")
				.setParameter("userId", userId)
				.list();
	}
}
