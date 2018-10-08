package com.vmware.jct.dao;

import com.vmware.jct.exception.DAOException;
import com.vmware.jct.model.JctUser;

/**
 * 
 * <p><b>Class name:</b>IContentServiceDAO.java</p>
 * <p><b>@author:</b> InterraIT </p>
 * <p><b>Purpose:</b> 		This DAO interface represents a IAuthenticatorDAO interface
 * <p><b>Description:</b> 	This will have all methods mapped to the database and fetching data from database to be used for Content service. </p>
 * <p><b>Note:</b> 			This methods implemented by DAOImpl classes and execute the mapped queries and get the records.</p>  
 * <p><b>Copyrights:</b> 	All rights reserved by InterraIT and should only be used for its internal application development.</p>
 *  <pre>
 * <b>Revision History:</b>
 * <li> </li>
 * </pre>
 */
public interface IMyAccountServiceDAO {
	/**
	 * Method populates existing user data based on email
	 * @return JctUser entity
	 * @throws DAOException
	 */
	public JctUser populateExistingUserData() throws DAOException;
	/**
	 * Method populates existing user data based on primary key
	 * @param primaryKey
	 * @return JctUser entity
	 * @throws DAOException
	 */
	public JctUser fetchUserData(Integer primaryKey) throws DAOException;
	/**
	 * Method updates the first name
	 * @param user
	 * @return status
	 * @throws DAOException
	 */
	public String updateFirstName(JctUser user) throws DAOException;
	/**
	 * Method updates last name
	 * @param user
	 * @return status
	 * @throws DAOException
	 */
	public String updateLastName(JctUser user) throws DAOException;
	/**
	 * Method updates email id
	 * @param user
	 * @return status
	 * @throws DAOException
	 */
	public String updateEmailId(JctUser user) throws DAOException;
	/**
	 * Method updates existing password
	 * @param user
	 * @return status
	 * @throws DAOException
	 */
	public String updatePassword(JctUser user) throws DAOException;
	/**
	 * Method save the user data
	 * @param user
	 * @return status
	 * @throws DAOException
	 */
	public String saveUser(JctUser user) throws DAOException;
}
