package com.vmware.jct.dao.impl;

import org.hibernate.SessionFactory;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Repository;
import org.springframework.transaction.annotation.Propagation;
import org.springframework.transaction.annotation.Transactional;
import com.vmware.jct.dao.DataAccessObject;
import com.vmware.jct.dao.IMyAccountServiceDAO;
import com.vmware.jct.exception.DAOException;
import com.vmware.jct.model.JctUser;
/**
 * 
 * <p><b>Class name:</b> MyAccountDAOImpl.java</p>
 * <p><b>Author:</b> InterraIT</p>
 * <p><b>Purpose:</b> This is a MyAccountDAOImpl class. This artifact is Persistence Manager layer artifact.
 * MyAccountDAOImpl implement IMyAccountServiceDAO interface and override the following  methods.
 * -populateExistingUserData()
 * -fetchUserData()
 * -updateFirstName()
 * -updateLastName()
 * -updateEmailId()
 * -updatePassword()
 * </p>
 * <p><b>Description:</b> This class used to perform insert, modify ,fetch data from db and delete operation. Save, update, delete operations are performed in DataAccessObject class </p>
 * <p><b>Revision History:</b>
 * 	<li>InterraIT - 19/Jan/2014 - Implement Exception </li>
 * <li>InterraIT -  - 31/Jan/2014 - Implement comments, introduce Named query, Transactional attribute </li>
 * <li>InterraIT - 16/May/2014 - Updated comments </li>
 * </p>
 */
@Repository
public class MyAccountDAOImpl extends DataAccessObject implements IMyAccountServiceDAO{

	@Autowired
	private SessionFactory sessionFactory;
	
	private static final Logger LOGGER = LoggerFactory.getLogger(MyAccountDAOImpl.class);
	
	/**
	 * @param emailId
	 * @throws DAOException
	 * 
	 */
	@Transactional(propagation = Propagation.REQUIRED)
	public JctUser populateExistingUserData()
			throws DAOException {
		LOGGER.info(">>>>>> MyAccountDAOImpl.populateExistingUserData");				
		return (JctUser) sessionFactory.getCurrentSession().
				getNamedQuery("fetchUserByEmail").list().get(0);		
	}

	
	@Transactional(propagation = Propagation.REQUIRED)
	public JctUser fetchUserData(Integer pk)
			throws DAOException {
		LOGGER.info(">>>>>> MyAccountDAOImpl.fetchUserData");				
		return (JctUser) sessionFactory.getCurrentSession().
				getNamedQuery("fetchUserByPk")
				.setParameter("pk", pk).list().get(0);		
	}
	
	@Transactional(propagation = Propagation.REQUIRED)
	public String updateFirstName(JctUser user) throws DAOException {
		LOGGER.info(">>>>>> MyAccountDAOImpl.updateFirstName");
		return update(user);
	}
	
	@Transactional(propagation = Propagation.REQUIRED)
	public String updateLastName(JctUser user) throws DAOException {
		LOGGER.info(">>>>>> MyAccountDAOImpl.updateLastName");
		return update(user);
	}
	
	@Transactional(propagation = Propagation.REQUIRED)
	public String updateEmailId(JctUser user) throws DAOException {
		LOGGER.info(">>>>>> MyAccountDAOImpl.updateEmailId");
		return update(user);
	}
	
	@Transactional(propagation = Propagation.REQUIRED)
	public String updatePassword(JctUser user) throws DAOException {
		LOGGER.info(">>>>>> MyAccountDAOImpl.updatePassword");
		return update(user);
	}
	
	@Transactional(propagation = Propagation.REQUIRED)
	public String saveUser(JctUser user) throws DAOException {
		LOGGER.info(">>>>>> MyAccountDAOImpl.saveUser");
		return save(user);
	}
}