package com.vmware.jct.dao;

import java.util.List;

import com.vmware.jct.exception.DAOException;
import com.vmware.jct.model.JctAppSettingsMaster;

/**
 * 
 * <p><b>Class name:</b>IAppSettingsDAO.java</p>
 * <p><b>@author:</b> InterraIT </p>
 * <p><b>Purpose:</b> 		This DAO interface represents a IAppSettingsDAO interface
 * <p><b>Description:</b> 	This will have all methods mapped to the database and fetching data from database to be used across the application </p>
 * <p><b>Note:</b> 			This methods implemented by DAOImpl classes and execute the mapped queries and get the records.</p>  
 * <p><b>Copyrights:</b> 	All rights reserved by Interra IT and should only be used for its internal application development.</p>
 *  <pre>
 * <b>Revision History:</b>
 * </pre>
 */
public interface IAppSettingsDAO {
	/**
	 * Method returns the colors
	 * @param softDelete
	 * @return List of JctAppSettingsMaster
	 * @throws DAOException
	 */
	public List<JctAppSettingsMaster> getColors(int softDelete) throws DAOException;
	/**
	 * Method updates the soft delete status
	 * @param master
	 * @return
	 * @throws DAOException
	 */
	public String updateSoftDelete(JctAppSettingsMaster master) throws DAOException;
	/**
	 * Method saves new while lebelling
	 * @param master
	 * @return
	 * @throws DAOException
	 */
	public String saveWhiteLebel(JctAppSettingsMaster master) throws DAOException;
}
