package com.vmware.jct.dao.impl;

import java.util.List;

import org.hibernate.SessionFactory;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Repository;
import org.springframework.transaction.annotation.Propagation;
import org.springframework.transaction.annotation.Transactional;

import com.vmware.jct.dao.DataAccessObject;
import com.vmware.jct.dao.IAppSettingsDAO;
import com.vmware.jct.exception.DAOException;
import com.vmware.jct.model.JctAppSettingsMaster;

/**
 * 
 * <p><b>Class name:</b> AppSettingsDAOImpl.java</p>
 * <p><b>Author:</b> InterraIT</p>
 * <p><b>Purpose:</b> This is a AppSettingsDAOImpl class. This artifact is Persistence Manager layer artifact.
 * AppSettingsDAOImpl implement IAppSettingsDAO interface and override the following  methods.
 * -getColors(int softDelete)
 * -updateSoftDelete(JctAppSettingsMaster master)
 * -saveWhiteLebel(JctAppSettingsMaster master)
  * <p><b>Description:</b> This class used to perform insert, modify ,fetch data from db and delete operation. Save, update, delete operations are performed in DataAccessObject class </p>
 * <p><b>Revision History:</b>
 * </p>
 */
@Repository
public class AppSettingsDAOImpl extends DataAccessObject implements IAppSettingsDAO{

	@Autowired
	private SessionFactory sessionFactory;
	
	@Transactional(propagation=Propagation.REQUIRED)
	public List<JctAppSettingsMaster> getColors(int softDelete)
			throws DAOException {
		return sessionFactory.getCurrentSession().getNamedQuery("fetchLastWhiteLebeling")
				.setInteger("softDelete", softDelete)
				.list();
	}

	@Transactional(propagation=Propagation.REQUIRED)
	public String updateSoftDelete(JctAppSettingsMaster master)
			throws DAOException {
		return update(master);
	}

	@Transactional(propagation=Propagation.REQUIRED)
	public String saveWhiteLebel(JctAppSettingsMaster master)
			throws DAOException {
		return save(master);
	}

}
