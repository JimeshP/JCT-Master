package com.vmware.jct.service;

import com.vmware.jct.exception.JCTException;
/**
 * 
 * <p><b>Class name:</b>IFacilitatorIndividualReportService.java</p>
 * <p><b>@author:</b> InterraIT </p>
 * <p><b>Purpose:</b> 		This Service interface represents a IFacilitatorIndividualReportService interface
 * <p><b>Description:</b> 	This will have all methods where we implement all business logic and all service layer </p>
 * <p><b>Copyrights:</b> 	All rights reserved by InterraIT and should only be used for its internal application development.</p> 
 */

public interface IFacilitatorIndividualReportService {
	
	/**
	 * Creates New Progress Reports
	 * @return String
	 * @return JCTException
	 * 
	 */
		
	public String populateAllProgressReports(String jctEmail) throws JCTException;	
	
	/**
	 * Method gets List of questionnaire answers
	 * @param adminEmailId
	 * @param usrType
	 */	
	public String getUserDetails(int usrType) throws JCTException;
	
	/**
	 * Creates Facilitator/Individual overview report
	 * @return String
	 * @throws JCTException
	 */
	public String createFaciIndiReport(String userType) throws JCTException;
	
	
	public String populateProgressReports(String jctEmail,String userGroupName) throws JCTException;	

}
