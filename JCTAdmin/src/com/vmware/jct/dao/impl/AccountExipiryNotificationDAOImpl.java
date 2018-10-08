package com.vmware.jct.dao.impl;

import java.util.Date;
import java.util.List;

import org.hibernate.SessionFactory;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Repository;
import org.springframework.transaction.annotation.Propagation;
import org.springframework.transaction.annotation.Transactional;

import com.vmware.jct.dao.IAccountExpiryNotificationDAO;
import com.vmware.jct.dao.dto.AccountExpiryDTO;
import com.vmware.jct.exception.DAOException;
/**
 * 
 * <p><b>Class name:</b> AccountExipiryNotificationDAOImpl.java</p>
 * <p><b>Author:</b> InterraIT</p>
 * <p><b>Purpose:</b> This is a AccountExipiryNotificationDAOImpl class. This artifact is Persistence Manager layer artifact.
 * AccountExipiryNotificationDAOImpl implement IAccountExpiryNotificationDAO interface and override the following  methods.
 * -getAllActiveFacilitators()
 * -getExpirationDetails(int facilitatorId, Date startDate, Date endDate)
 * -getEmailIdByUserId(Integer facilitatorId, int role)
 * <p><b>Description:</b> This class used to perform insert, modify ,fetch data from db and delete operation. Save, update, delete operations are performed in DataAccessObject class </p>
 * <p><b>Revision History:</b>
 * </p>
 */
@Repository
public class AccountExipiryNotificationDAOImpl implements IAccountExpiryNotificationDAO {
	@Autowired
	private SessionFactory sessionFactory;

	@Transactional(propagation=Propagation.REQUIRED)
	public List<Integer> getAllActiveFacilitators() throws DAOException {
		return sessionFactory.getCurrentSession().getNamedQuery("getAllActiveFacilitatorWhoCreatedUsers")
				.list();
	}

	@Transactional(propagation=Propagation.REQUIRED)
	public List<AccountExpiryDTO> getExpirationDetails(int facilitatorId, Date startDate, Date endDate)
			throws DAOException {
		return sessionFactory.getCurrentSession().getNamedQuery("getAllToBeExpiredUsersRegByFacilitator")
				.setParameter("facilitatorId", facilitatorId)
				.setParameter("startDate", startDate)
				.setParameter("endDate", endDate)
				.list();
	}

	@Transactional(propagation=Propagation.REQUIRED)
	public String getEmailIdByUserId(Integer facilitatorId, int role)
			throws DAOException {
		return (String) sessionFactory.getCurrentSession().getNamedQuery("getEmailIdByUserId")
				.setParameter("jctUserId", facilitatorId)
				.setParameter("role", role)
				.uniqueResult();
	}
}