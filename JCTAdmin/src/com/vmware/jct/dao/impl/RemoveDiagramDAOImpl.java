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
import com.vmware.jct.dao.IRemoveDiagramDAO;
import com.vmware.jct.dao.dto.RemoveDiagramDTO;
import com.vmware.jct.exception.DAOException;
import com.vmware.jct.model.JctUser;
import com.vmware.jct.service.vo.RemoveDiagramVOList;
/**
 * 
 * <p><b>Class name:</b> RemoveDiagramDAOImpl.java</p>
 * <p><b>Author:</b> InterraIT</p>
 * <p><b>Purpose:</b> This is a IRemoveDiagramDAO class. This artifact is Persistence Manager layer artifact.
 * RemoveDiagramDAOImpl implement IMyAccountServiceDAO interface and override the following  methods.
 * -getDiagrams(String emailId)
 * </p>
 * <p><b>Description:</b> This class used to perform insert, modify ,fetch data from db and delete operation. Save, update, delete operations are performed in DataAccessObject class </p>
 * <p><b>Revision History:</b>
 * 	<li>InterraIT - 06/Oct/2014 - Introduced the class </li>
 * </p>
 */
@Repository
public class RemoveDiagramDAOImpl extends DataAccessObject implements IRemoveDiagramDAO {

	@Autowired
	private SessionFactory sessionFactory;
	
	private static final Logger LOGGER = LoggerFactory.getLogger(RemoveDiagramDAOImpl.class);
	
	@Transactional(propagation = Propagation.REQUIRED)
	public List<RemoveDiagramDTO> getDiagrams(String emailId)
			throws DAOException {
		LOGGER.info(">>>>>>>> RemoveDiagramDAOImpl.getDiagrams");
		return sessionFactory.getCurrentSession().getNamedQuery("searchDiagramsToDiableByEmailId")
				.setParameter("emailId", emailId).list();
	}

	@Transactional(propagation = Propagation.REQUIRED)
	public int removeDiagram(RemoveDiagramVOList vo) throws DAOException {
		return sessionFactory.getCurrentSession().getNamedQuery("softDeleteBsAndAsDiagramByAdmin")
				.setParameter("emailId", vo.getEmailId())
				.setParameter("rowId", vo.getRowId())
				.setParameter("softDelStatus", 2)
				.setParameter("ts", new Date()).executeUpdate();
	}

	@Transactional(propagation = Propagation.REQUIRED)
	public List<RemoveDiagramDTO> populateAllDiagram(int recordIndex)
			throws DAOException {
		LOGGER.info(">>>>>>>> RemoveDiagramDAOImpl.populateAllDiagram");
		return sessionFactory.getCurrentSession().getNamedQuery("searchAllDiagrams")
				.setFirstResult(recordIndex).setMaxResults(4).list();
	}

	@Transactional(propagation = Propagation.REQUIRED)
	public String fetchUserName(String jobRefNo) throws DAOException {
		LOGGER.info(">>>>>>>> RemoveDiagramDAOImpl.fetchUserName");	
		return (String) sessionFactory.getCurrentSession().getNamedQuery("fetchCreatedByByJobRefNo")
				.setParameter("jobRefNo", jobRefNo).list().get(0);
	}

	@Transactional(propagation = Propagation.REQUIRED)
	public Long fetchAllDiagramCount() throws DAOException {
		LOGGER.info(">>>>>>>> RemoveDiagramDAOImpl.fetchAllDiagramCount");	
		return (Long) sessionFactory.getCurrentSession().getNamedQuery("fetchAllDiagramCount").uniqueResult();
	}

	@Transactional(propagation = Propagation.REQUIRED)
	public JctUser searchUserForRefundRequest(String emailId)
			throws DAOException {
		LOGGER.info(">>>>>>>> RemoveDiagramDAOImpl.searchUserForRefundRequest");
		
		JctUser obj = null;
		
		obj = (JctUser) sessionFactory.getCurrentSession().
				getNamedQuery("getExistingUserByEmailAndRoleId")
				.setParameter("email", emailId)
				.setParameter("roleId", 1).uniqueResult();
		/*return (JctUser) sessionFactory.getCurrentSession().
				getNamedQuery("getExistingUserByEmailAndRoleId")
				.setParameter("email", emailId)
				.setParameter("roleId", 1).list().get(0);*/
		return obj;
	}
	
	@Transactional(propagation = Propagation.REQUIRED)
	public boolean disableUserForRefundRequestByUserId(int userId)
			throws DAOException {
		LOGGER.info(">>>>>>>> RemoveDiagramDAOImpl.disableUserForRefundRequestByUserId");
		int effectedRows = 0; 
		/* update user table */
		effectedRows = sessionFactory.getCurrentSession().getNamedQuery("disableUserForRefundRequestByUserId")
				.setParameter("userId", userId).executeUpdate();
		/* update user dtls table */
		effectedRows += sessionFactory.getCurrentSession().getNamedQuery("disableUserDtlsForRefundRequestByUserId")
				.setParameter("userId", userId).executeUpdate();
		if( effectedRows < 2 ){			
			return false;
		}
		return true;		
	}
}
