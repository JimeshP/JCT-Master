package com.vmware.jct.dao;

import com.vmware.jct.exception.DAOException;
import com.vmware.jct.model.JctPdfRecords;
/**
 * 
 * <p><b>Class name:</b>IPdfRecordsDAO.java</p>
 * <p><b>@author:</b> InterraIt </p>
 * <p><b>Purpose:</b> 		This DAO interface represents a IConfirmationDAO interface
 * <p><b>Description:</b> 	This will have all methods mapped to the database and fetching data from database to be used across the application </p>
 * <p><b>Copyrights:</b> 	All rights reserved by Interra IT and should only be used for its internal application development.</p>
 *  <pre>
 * <b>Revision History:</b>
 * </pre>
 */
public interface IPdfRecordsDAO {
	/**
	 * 
	 * @param records
	 * @return
	 * @throws DAOException
	 */
	public String savePDFReference(JctPdfRecords records) throws DAOException;
	/**
	 * 
	 * @param jobReferenceNo
	 * @return
	 * @throws DAOException
	 */
	public JctPdfRecords getMaxDocRow(String jobReferenceNo) throws DAOException;
}
