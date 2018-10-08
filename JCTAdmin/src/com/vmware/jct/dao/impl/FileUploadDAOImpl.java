package com.vmware.jct.dao.impl;

import org.hibernate.SessionFactory;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Repository;
import org.springframework.transaction.annotation.Propagation;
import org.springframework.transaction.annotation.Transactional;

import com.vmware.jct.dao.DataAccessObject;
import com.vmware.jct.dao.IFileUploadDAO;
import com.vmware.jct.exception.DAOException;
import com.vmware.jct.model.JctOnetOccupationList;
/**
 * 
 * <p><b>Class name:</b> FileUploadDAOImpl.java</p>
 * <p><b>Author:</b> InterraIT</p>
 * <p><b>Purpose:</b> This is a FileUploadDAOImpl class. This artifact is Persistence Manager layer artifact.
 * FileUploadDAOImpl implement IFileUploadDAO interface and override the following  methods.
 * -saveONetData()
 * -getOccupationTitle()
 * -updateONetData()
 * -getOccupationObj()
 * </p>
 * <p><b>Description:</b> This class used to perform insert, modify ,fetch data from db and delete operation. Save, update, delete operations are performed in DataAccessObject class </p>
 * <p><b>Revision History:</b>
 * 	<li>InterraIT - 01/Oct/2014 - Implement Exception </li>
 * <li>InterraIT -  - 11/Oct/2014 - Implement comments, introduce Named query, Transactional attribute </li>
 * <li>InterraIT - 31/Oct/2014 - Updated comments </li>
 * </p>
 */
@Repository
public class FileUploadDAOImpl extends DataAccessObject implements IFileUploadDAO {
	
	@Autowired
	private SessionFactory sessionFactory;
	
	private static final Logger LOGGER = LoggerFactory.getLogger(ContentConfigDAOImpl.class);	

	/**
	 * Method saves new OnetOccupation data
	 * @param jctOnetOccupation
	 */
	@Transactional(propagation = Propagation.REQUIRED)
	public String saveONetData(JctOnetOccupationList jctOnetOccupation) throws DAOException {
		LOGGER.info(">>>>>> FileUploadDAOImpl.saveONetData");
		return save(jctOnetOccupation);
	}
	
	/**
	 * Method to fetch the occupation title against the occupation code
	 * @param occupationCode
	 * @return occupationTitle
	 */
	@Transactional(propagation = Propagation.REQUIRED)
	public String getOccupationTitle(String occupationCode) throws DAOException {
		LOGGER.info(">>>>>> FileUploadDAOImpl.getOccupationTitle");
		String occupationTitle  = "";
		try {
			occupationTitle = (String) sessionFactory.getCurrentSession().getNamedQuery("fetchOnetTitle").
					setParameter("occupationCode", occupationCode).uniqueResult();
		} catch (Exception e) {
			LOGGER.info(e.getMessage());
		}
		LOGGER.info("<<<<<< FileUploadDAOImpl.getOccupationTitle");
		return occupationTitle;
	}
	
	/**
	 * Method to update the occupation data
	 * @param jctOnetOccupation
	 */
	@Transactional(propagation = Propagation.REQUIRED)
	public String updateONetData(JctOnetOccupationList jctOnetOccupation) throws DAOException {
		LOGGER.info(">>>>>> FileUploadDAOImpl.updateONetData");
		return update(jctOnetOccupation);
	}
	
	/**
	 * Method to fetch the occupation object against the occupation code
	 * @param occupationCode
	 * @throws JctOnetOccupationList object
	 */
	@Transactional(propagation = Propagation.REQUIRED)
	public JctOnetOccupationList getOccupationObj(String occupationCode) throws DAOException {
		LOGGER.info(">>>>>> FileUploadDAOImpl.getOccupationObj");
		return (JctOnetOccupationList) sessionFactory.getCurrentSession().getNamedQuery("fetchOnetDataObj")
				.setParameter("occupationCode", occupationCode).list().get(0);
	}
}
