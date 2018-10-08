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
import com.vmware.jct.dao.IFacilitatorAccountServiceDAO;
import com.vmware.jct.exception.DAOException;
import com.vmware.jct.model.JctCheckPaymentUserDetails;
import com.vmware.jct.model.JctFacilitatorDetails;
import com.vmware.jct.model.JctPaymentHeader;
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
public class FacilitatorAccountDAOImpl extends DataAccessObject implements IFacilitatorAccountServiceDAO{

	@Autowired
	private SessionFactory sessionFactory;
	
	private static final Logger LOGGER = LoggerFactory.getLogger(FacilitatorAccountDAOImpl.class);
	
	/**
	 * Method populates existing facilitator data based on email
	 * @param facilitatorEmail
	 * @return JctUser entity
	 */
	@Transactional(propagation = Propagation.REQUIRED)
	public JctUser populateExistingFacilitatorData(String facilitatorEmail)
			throws DAOException {
		LOGGER.info(">>>>>> FacilitatorAccountDAOImpl.populateExistingFacilitatorData");				
		return (JctUser) sessionFactory.getCurrentSession().
				getNamedQuery("fetchFacilitatorByEmail")
				.setParameter("facilitatorEmail", facilitatorEmail).list().get(0);		
	}

	/**
	 * Method populates existing user data based on primary key
	 * @param pk
	 * @return JctUser entity
	 */
	@Transactional(propagation = Propagation.REQUIRED)
	public JctUser fetchUserData(Integer pk)
			throws DAOException {
		LOGGER.info(">>>>>> MyAccountDAOImpl.fetchUserData");				
		return (JctUser) sessionFactory.getCurrentSession().
				getNamedQuery("fetchUserByPk")
				.setParameter("pk", pk).list().get(0);		
	}
	
	/**
	 * Method updates the user name
	 * @param user
	 * @return status
	 */
	@Transactional(propagation = Propagation.REQUIRED)
	public String updateUserNameFacilitator(JctUser user) throws DAOException {
		LOGGER.info(">>>>>> FacilitatorAccountDAOImpl.updateUserNameFacilitator");
		return update(user);
	}
	
	/**
	 * Method updates the email address
	 * @param user
	 * @return status
	 * @throws DAOException
	 */
	@Transactional(propagation = Propagation.REQUIRED)
	public String updateEmailIdFacilitator(JctUser user) throws DAOException {
		LOGGER.info(">>>>>> FacilitatorAccountDAOImpl.updateEmailIdFacilitator");
		return update(user);
	}
	
	/**
	 * Method updates the password
	 * @param valueDesc
	 * @param tablePkId
	 * @return MyAccountVO
	 */
	@Transactional(propagation = Propagation.REQUIRED)
	public String updatePasswordFacilitator(JctUser user) throws DAOException {
		LOGGER.info(">>>>>> FacilitatorAccountDAOImpl.updatePasswordFacilitator");
		return update(user);
	}
	
	/**
	 * Method save the user data
	 * @param user
	 * @return status
	 */
	@Transactional(propagation = Propagation.REQUIRED)
	public String saveFacilitator(JctUser user) throws DAOException {
		LOGGER.info(">>>>>> FacilitatorAccountDAOImpl.saveFacilitator");
		return save(user);
	}

	/**
	 * Method save the payment header
	 * @param paymentHeader
	 * @return status
	 */
	@Transactional(propagation = Propagation.REQUIRED)
	public String savePaymentHeader(JctPaymentHeader paymentHeader)
			throws DAOException {
		LOGGER.info(">>>>>> FacilitatorAccountDAOImpl.savePaymentHeader");
		return save(paymentHeader);
	}
	
	/**
	 * Method save the facilitator details
	 * @param facilitatorDetails
	 * @return status
	 */
	@Transactional(propagation = Propagation.REQUIRED)
	public String saveToFacilitatorDetails (JctFacilitatorDetails facilitatorDetails)
			throws DAOException {
		LOGGER.info(">>>>>> FacilitatorAccountDAOImpl.saveToFacilitatorDetails");
		return save(facilitatorDetails);
	}

	@Transactional(propagation = Propagation.REQUIRED)
	public String searchUserIdInDB(String facilitatorEmail) throws DAOException {
		LOGGER.info(">>>>>> FacilitatorAccountDAOImpl.searchUserIdInDB");
		String status = "failure";
		List list = null;
		try{
			list = sessionFactory.getCurrentSession().getNamedQuery("fetchJctUserByEmailId")
					.setParameter("emailId", facilitatorEmail).list();
			if(list.size() == 0){
				status = "notexist";
			} else {
				status = "exist";
			}			
		}catch(Exception ez){
			status = "notexist";
			LOGGER.error(ez.getLocalizedMessage());
		}
		LOGGER.info("<<<<<< FacilitatorAccountDAOImpl.searchUserIdInDB");
		return status;
	}
	
	/**
	 * Method save the JCT check more User subscription/renew table
	 * @param moreUserTemp
	 * @return status
	 */
	@Transactional(propagation = Propagation.REQUIRED)
	public String saveCheckPaymentUserDetails (JctCheckPaymentUserDetails obj)
			throws DAOException {
		LOGGER.info(">>>>>> FacilitatorAccountDAOImpl.saveCheckPaymentUserDetails");
		return save(obj);
	}
}