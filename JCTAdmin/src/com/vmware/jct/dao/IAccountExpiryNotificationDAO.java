package com.vmware.jct.dao;

import java.util.Date;
import java.util.List;

import com.vmware.jct.dao.dto.AccountExpiryDTO;
import com.vmware.jct.exception.DAOException;
/**
 * 
 * <p><b>Class name:</b>IAccountExpiryNotificationDAO.java</p>
 * <p><b>@author:</b> InterraIT </p>
 * <p><b>Purpose:</b> 		This DAO interface represents a IAccountExpiryNotificationDAO interface
 * <p><b>Description:</b> 	This will have all methods mapped to the database and fetching data from database to be used across the application </p>
 * <p><b>Note:</b> 			This methods implemented by DAOImpl classes and execute the mapped queries and get the records.</p>  
 * <p><b>Copyrights:</b> 	All rights reserved by Interra IT and should only be used for its internal application development.</p>
 *  <pre>
 * <b>Revision History:</b>
 * <li></li>
 * </pre>
 */
public interface IAccountExpiryNotificationDAO {
	/**
	 * Method fetches the user Id of all the active facilitators
	 * who have created any user. 
	 * @return List of facilitator User Id
	 * @throws DAOException
	 */
	public List<Integer> getAllActiveFacilitators() throws DAOException;
	/**
	 * Method fetches the user email ids and expire dates only if the user 
	 * account is getting expired in <b>3 days </b> 
	 * @param facilitatorId
	 * @param startDate
	 * @param endDate
	 * @return List AccountExpiryDTO
	 * @throws DAOException
	 */
	public List<AccountExpiryDTO> getExpirationDetails ( int facilitatorId, Date startDate, Date endDate ) throws DAOException;
	/**
	 * Fetches the email Id
	 * @param facilitatorId
	 * @param role
	 * @return
	 * @throws DAOException
	 */
	public String getEmailIdByUserId(Integer facilitatorId, int role) throws DAOException;
}