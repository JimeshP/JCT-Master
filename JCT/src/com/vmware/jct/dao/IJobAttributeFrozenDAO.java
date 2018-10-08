package com.vmware.jct.dao;

import java.util.List;

import com.vmware.jct.dao.dto.JobAttributeDTO;
import com.vmware.jct.exception.DAOException;
import com.vmware.jct.model.JctJobAttributeFrozen;

/**
 * 
 * <p><b>Class name:</b>IJobAttributeFrozenDAO.java</p>
 * <p><b>@author:</b> InterraIt </p>
 * <p><b>Purpose:</b> 		This DAO interface represents a IJobAttributeFrozenDAO interface
 * <p><b>Description:</b> 	This will have all methods mapped to the database and fetching data from database to be used across the application </p>
 * <p><b>Copyrights:</b> 	All rights reserved by Interra IT and should only be used for its internal application development.</p>
 *  <pre>
 * <b>Revision History:</b>
 * <li> 13/Oct/2014 - Introduced the class</li>
 * </pre>
 */
public interface IJobAttributeFrozenDAO {
	/**
	 * 
	 * @param JctJobAttributeFrozen
	 * @return
	 * @throws DAOException
	 */
	public String saveFrozenAttributes(JctJobAttributeFrozen frozen) throws DAOException;
	/**
	 * 
	 * @param profileId
	 * @param userName
	 * @return
	 * @throws DAOException
	 */
	public List<JobAttributeDTO> getFrozenJobAttribute(int profileId, String userName, int userId) throws DAOException;
}
