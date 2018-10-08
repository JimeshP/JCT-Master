package com.vmware.jct.dao.impl;

import java.util.ArrayList;
import java.util.List;

import org.hibernate.SessionFactory;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Repository;
import org.springframework.transaction.annotation.Propagation;
import org.springframework.transaction.annotation.Transactional;

import com.vmware.jct.dao.DataAccessObject;
import com.vmware.jct.dao.IPaymentDAO;
import com.vmware.jct.dao.dto.PaymentDetailsDTO;
import com.vmware.jct.exception.DAOException;
import com.vmware.jct.model.JctCheckPaymentUserDetails;
import com.vmware.jct.model.JctPaymentDetails;
/**
 * 
 * <p><b>Class name:</b> PaymentDAOImpl.java</p>
 * <p><b>Author:</b> InterraIT</p>
 * <p><b>Purpose:</b> This is a PaymentDAOImpl class. This artifact is Persistence Manager layer artifact.
 * PaymentDAOImpl implement IPaymentDAO interface and override the following  methods.
 * -savePaymentDetails
 * </p>
 * <p><b>Description:</b> This class used to perform insert, modify ,fetch data from db and delete operation. Save, update, delete operations are performed in DataAccessObject class </p>
 * <p><b>Revision History:</b>
 * 	<li></li>
 * </p>
 */
@Repository
public class PaymentDAOImpl extends DataAccessObject implements IPaymentDAO {

	@Autowired
	private SessionFactory sessionFactory;
	
	private static final Logger LOGGER = LoggerFactory.getLogger(PaymentDAOImpl.class);
	/**
	 * Method saves the payment details
	 */
	@Transactional(propagation = Propagation.REQUIRED)
	public String savePaymentDetails(JctPaymentDetails details)
			throws DAOException {
		LOGGER.info(">>>>>> PaymentDAOImpl.savePaymentDetails");
		return save(details);
	}
	@Transactional(propagation=Propagation.REQUIRED)
	public List<PaymentDetailsDTO> fetchExistingChequeUsersUniqueFields(String fetchType, String chequeNum, String userName) throws DAOException {
		LOGGER.info(">>>> PaymentDAOImpl.fetchExistingChequeUsers");
		List<PaymentDetailsDTO> userList = null;
		if(fetchType.equals("") ){
			userList = sessionFactory.getCurrentSession().getNamedQuery("fetchCheckPaymentUserUniqueDetails").list();		
		} else if(fetchType.equals("BY_CHECK") || fetchType.equals("BY_USERNAME")) {
			//	BY_CHECK and also for BY_USERNAME
			userList = sessionFactory.getCurrentSession().getNamedQuery("fetchCheckPaymentUserUniqueDetailsByChk")
									.setParameter("chequeNum", chequeNum).list();		
		} else if(fetchType.equals("BY_ALL")){
			userList = sessionFactory.getCurrentSession().getNamedQuery("CheckPymntUserUniqueDtlByChkAndUsrName")
					.setParameter("chequeNum", chequeNum)
					.setParameter("userName", userName).list();			
		}	
		return userList;
	}	
	@Transactional(propagation=Propagation.REQUIRED)
	public List<JctCheckPaymentUserDetails> searchByTranId(String tranId) 
			throws DAOException {
		LOGGER.info(">>>> PaymentDAOImpl.searchByTranId");
		return sessionFactory.getCurrentSession().getNamedQuery("fetchCheckUserDtlByTranId")
				.setParameter("tranId", tranId).list();
	}
	
	@Transactional(propagation=Propagation.REQUIRED)
	public List<PaymentDetailsDTO> fetchuserRemainingFields(String transId) throws DAOException {
		LOGGER.info(">>>> PaymentDAOImpl.fetchuserRemainingFields");
		List<PaymentDetailsDTO> userList = null;		
		userList = sessionFactory.getCurrentSession().getNamedQuery("fetchuserRemainingFields")
		.setParameter("transId", transId).list();
		return userList;
	}
	@Transactional(propagation=Propagation.REQUIRED)
	public String getCHkNoByUserName(String userName) 
			throws DAOException {
		LOGGER.info(">>>> PaymentDAOImpl.getCHkNoByUserName");
		return (String) sessionFactory.getCurrentSession().getNamedQuery("getChkNoByUserName")
				.setParameter("userName", userName).uniqueResult();
	}
}