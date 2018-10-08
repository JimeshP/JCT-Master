package com.vmware.jct.dao;

import com.vmware.jct.exception.DAOException;
import com.vmware.jct.model.JctCheckPaymentUserDetails;
import com.vmware.jct.model.JctFacilitatorDetails;
import com.vmware.jct.model.JctPaymentHeader;
import com.vmware.jct.model.JctUser;

/**
 * 
 * <p><b>Class name:</b>IFacilitatorAccountServiceDAO.java</p>
 * <p><b>@author:</b> InterraIT </p>
 * <p><b>Purpose:</b> 		This DAO interface represents a IFacilitatorAccountServiceDAO interface
 * <p><b>Description:</b> 	This will have all methods mapped to the database and fetching data from database to be used. </p>
 * <p><b>Note:</b> 			This methods implemented by DAOImpl classes and execute the mapped queries and get the records.</p>  
 * <p><b>Copyrights:</b> 	All rights reserved by InterraIT and should only be used for its internal application development.</p>
 *  <pre>
 * <b>Revision History:</b>
 * <li> </li>
 * </pre>
 */
public interface IFacilitatorAccountServiceDAO {
	/**
	 * Method populates existing facilitator data based on email
	 * @param facilitatorEmail
	 * @return JctUser entity
	 * @throws DAOException
	 */
	public JctUser populateExistingFacilitatorData(String facilitatorEmail) throws DAOException;
	/**
	 * Method populates existing user data based on primary key
	 * @param primaryKey
	 * @return JctUser entity
	 * @throws DAOException
	 */
	public JctUser fetchUserData(Integer primaryKey) throws DAOException;
	/**
	 * Method updates the user name
	 * @param user
	 * @return status
	 * @throws DAOException
	 */
	public String updateUserNameFacilitator(JctUser user) throws DAOException;
	/**
	 * Method updates email address
	 * @param user
	 * @return status
	 * @throws DAOException
	 */
	public String updateEmailIdFacilitator(JctUser user) throws DAOException;
	/**
	 * Method updates existing password
	 * @param user
	 * @return status
	 * @throws DAOException
	 */
	public String updatePasswordFacilitator(JctUser user) throws DAOException;
	/**
	 * Method save the user data
	 * @param user
	 * @return status
	 * @throws DAOException
	 */
	public String saveFacilitator(JctUser user) throws DAOException;
	/**
	 * Method saves to payment Header Table
	 * @param details
	 * @return
	 * @throws DAOException
	 */
	public String savePaymentHeader (JctPaymentHeader paymentHeader) throws DAOException;
	/**
	 * Method saves to facilitator Details table
	 * @param details
	 * @return
	 * @throws DAOException
	 */
	public String saveToFacilitatorDetails (JctFacilitatorDetails facilitatorDetails) throws DAOException;
	/**
	 * Method search the user id is exist in database or not
	 * @param facilitatorEmail
	 * @return
	 * @throws DAOException
	 */
	public String searchUserIdInDB(String facilitatorEmail) throws DAOException;
	/**
	 * Method saves to JCT check more User subscription/renew table
	 * @param moreUserTemp
	 * @return
	 * @throws DAOException
	 */
	public String saveCheckPaymentUserDetails(JctCheckPaymentUserDetails obj) throws DAOException;	
}
