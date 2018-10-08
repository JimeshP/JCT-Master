package com.vmware.jct.dao.impl;

import java.util.List;

import org.hibernate.SessionFactory;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Repository;
import org.springframework.transaction.annotation.Propagation;
import org.springframework.transaction.annotation.Transactional;

import com.vmware.jct.dao.IAdminMenuDAO;
import com.vmware.jct.exception.DAOException;
import com.vmware.jct.model.JctAdminMenu;
/**
 * 
 * <p><b>Class name:</b> AdminMenuDAOImpl.java</p>
 * <p><b>Author:</b> InterraIT</p>
 * <p><b>Purpose:</b> This is a AdminMenuDAOImpl class. This artifact is Persistence Manager layer artifact.
 * AdminMenuDAOImpl implement IAdminMenuDAO interface and override the following  methods.
 * -getDistinctMenu( int roleId )
 * -getMenuDetails( int roleId, String mainMenu )
  * <p><b>Description:</b> This class used to perform insert, modify ,fetch data from db and delete operation. Save, update, delete operations are performed in DataAccessObject class </p>
 * <p><b>Revision History:</b>
 * </p>
 */
@Repository
public class AdminMenuDAOImpl implements IAdminMenuDAO {

	@Autowired
	private SessionFactory sessionFactory;
	
	/**
	 * The method fetches distinct main menu items depending upon the 
	 * role id of the user.
	 */
	@Transactional(propagation=Propagation.REQUIRED)
	public List<String> getDistinctMenu ( int roleId ) throws DAOException {
		return sessionFactory.getCurrentSession().getNamedQuery("getActiveDistinctMenuItems")
				.setInteger("roleId", roleId)
				.list();
	}
	/**
	 * The method fetches the menu list depending upon the 
	 * role id of the user and corresponding main menu.
	 */
	@Transactional(propagation=Propagation.REQUIRED)
	public List<JctAdminMenu> getMenuDetails( int roleId, String mainMenu )
			throws DAOException {
		return sessionFactory.getCurrentSession().getNamedQuery("getActiveMenuDetails")
				.setInteger("roleId", roleId)
				.setString("mainMenu", mainMenu)
				.list();
	}
}