package com.vmware.jct.dao;

import java.util.List;

import com.vmware.jct.exception.DAOException;
import com.vmware.jct.model.JctAdminMenu;

/**
 * 
 * <p><b>Class name:</b>IAdminMenuDAO.java</p>
 * <p><b>@author:</b> InterraIT </p>
 * <p><b>Purpose:</b> 		This DAO interface represents a IAdminMenuDAO interface
 * <p><b>Description:</b> 	This will have all methods mapped to the database and fetching data from database to be used across the application </p>
 * <p><b>Note:</b> 			This methods implemented by DAOImpl classes and execute the mapped queries and get the records.</p>  
 * <p><b>Copyrights:</b> 	All rights reserved by Interra IT and should only be used for its internal application development.</p>
 *  <pre>
 * <b>Revision History:</b>
 * </pre>
 */
public interface IAdminMenuDAO {
	/**
	 * Method fetches district menu
	 * @param userType
	 * @return
	 * @throws DAOException
	 */
	public List<String> getDistinctMenu ( int roleId ) throws DAOException;
	/**
	 * Method returns menu details
	 * @param roleId
	 * @param mainMenu
	 * @return
	 * @throws DAOException
	 */
	public List<JctAdminMenu> getMenuDetails ( int roleId, String mainMenu ) throws DAOException;
}
