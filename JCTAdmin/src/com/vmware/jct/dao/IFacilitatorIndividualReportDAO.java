package com.vmware.jct.dao;

import java.util.List;

import com.vmware.jct.exception.DAOException;
import com.vmware.jct.model.JctUserLoginInfo;

/**
 * 
 * <p><b>Class name:</b>IFacilitatorIndividualReportDAO.java</p>
 * <p><b>@author:</b> InterraIT </p>
 * <p><b>Purpose:</b> 		This DAO interface represents a IFacilitatorIndividualReportDAO interface
 * <p><b>Description:</b> 	This will have all methods mapped to the database and fetching data from database to be used for report service. </p>
 * <p><b>Note:</b> 			This methods implemented by DAOImpl classes and execute the mapped queries and get the records.</p>  
 * <p><b>Copyrights:</b> 	All rights reserved by InterraIT and should only be used for its internal application development.</p>
 */

public interface IFacilitatorIndividualReportDAO {

	/**
	 *New Progress Report */
	
	public List<Object> populateAllProgressReports(String jctEmail) throws DAOException;	
	
		
	/**
	 * Method returns collection of roles in which the task falls
	 * @param usrType
	 * @return
	 * @throws DAOException
	 */
	public List<Object> getUserDetails(int usrType) throws DAOException;	
	/**
	 * Method fetches list of items required to plot misc report
	 * @return List of Object
	 * @throws DAOException
	 */
	public List<Object> createFaciIndiReport(String userType) throws DAOException;
	
	public List<Object> populateProgressReports(String jctEmail,String userGroupName) throws DAOException;	
	
	public List<JctUserLoginInfo> getTime (String userName) throws DAOException;
}
