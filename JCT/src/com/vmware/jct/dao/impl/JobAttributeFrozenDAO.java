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
import com.vmware.jct.dao.IJobAttributeFrozenDAO;
import com.vmware.jct.dao.dto.JobAttributeDTO;
import com.vmware.jct.exception.DAOException;
import com.vmware.jct.model.JctJobAttributeFrozen;

@Repository
public class JobAttributeFrozenDAO extends DataAccessObject implements IJobAttributeFrozenDAO {
	@Autowired
	private SessionFactory sessionFactory;
	
	private static final Logger LOGGER = LoggerFactory.getLogger(JobAttributeFrozenDAO.class);

	/**
	 * <p><b>Description:</b>  Method saves data in jct_job_attribute_frozen table </p>
	 * @param It will take jobReferenceNo as a input parameter
	 * @return  It returns string object
	 * @throws DAOException -it throws DAOException
	 * 
	 */
	@Transactional(propagation=Propagation.REQUIRED)
	public String saveFrozenAttributes(JctJobAttributeFrozen frozen)
			throws DAOException {
		return save(frozen);
	}
	/**
	 * <p><b>Description:</b>   Baserd on profile Id and user this method will 
	 * fetch all job attributes from frozen table</p>
	 * @param It will take profileId and username as a input parameter
	 * @return  It returns list of JobAttributeDTO data from database
	 * @throws DAOException -it throws DAOException
	 * 
	 */
	@Transactional(propagation=Propagation.REQUIRED,readOnly = true)
	public List<JobAttributeDTO> getFrozenJobAttribute(int profileId, String userName, int userId) throws DAOException {
		return sessionFactory.getCurrentSession().getNamedQuery("fetchFrozenAttrs")
				.setParameter("profileId", profileId)
				.setParameter("createdFor", userName)
				.setParameter("userId", userId).list();
	}
	
}
