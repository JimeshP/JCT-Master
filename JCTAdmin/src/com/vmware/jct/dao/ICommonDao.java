package com.vmware.jct.dao;
import com.vmware.jct.exception.DAOException;

/**
 * 
 * <p><b>Class name:</b>ICommonDao.java</p>
 * <p><b>@author:</b> InterraIT </p>
 * <p><b>Purpose:</b> 		This DAO interface represents a ICommonDao interface
 * <p><b>Description:</b> 	This will have all methods mapped to the database and fetching data from database to be used across the application </p>
 * <p><b>Copyrights:</b> 	All rights reserved by Interra IT and should only be used for its internal application development.</p>
 *  <pre>
 * <b>Revision History:</b>
 * InterraIT - 31/Jan/2014 - Implement comments through out the application
 * </pre>
 */
public interface ICommonDao {
	
	/**
	 * 
	 * @param It will take string  object as a input parameter
	 * @return unique identification number
	 * @throws DAOException
	 */
	public Integer generateKey(String keyName) throws DAOException;
	/**
	 * Generates unique customer Id
	 * @param roleId
	 * @return
	 * @throws DAOException
	 */
	public String generateUniqueCustomerId ( int roleId ) throws DAOException;
}
