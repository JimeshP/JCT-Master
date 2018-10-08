package com.vmware.jct.service.impl;



import java.math.BigDecimal;
import java.math.BigInteger;
import java.util.Calendar;
import java.util.List;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.context.MessageSource;
import org.springframework.stereotype.Service;
import org.springframework.transaction.annotation.Propagation;
import org.springframework.transaction.annotation.Transactional;

import com.vmware.jct.common.utility.CommonUtility;
import com.vmware.jct.dao.IFacilitatorIndividualReportDAO;
import com.vmware.jct.exception.DAOException;
import com.vmware.jct.exception.JCTException;
import com.vmware.jct.model.JctUserLoginInfo;
import com.vmware.jct.service.IFacilitatorIndividualReportService;

/**
 * <p><b>Class name:</b> FacilitatorIndividualReportServiceImpl.java</p>
 * <p><b>Author:</b> InterraIT</p>
 * <p><b>Purpose:</b> This is a FacilitatorIndividualReportServiceImpl class. This artifact is Business layer artifact.
 * FacilitatorIndividualReportServiceImpl implement IFacilitatorIndividualReportServiceImpl interface and override the following  methods.
 * -populateAllProgressReports()
 * -getUserDetails()
 * -createFaciIndiReport()
 * -populateProgressReports()
 */

@Service
public class FacilitatorIndividualReportServiceImpl implements IFacilitatorIndividualReportService{

	@Autowired
	private IFacilitatorIndividualReportDAO iFacilitatorIndividualReportDAO;
	
	@Autowired
	private MessageSource messageSource;
	
	private static final Logger LOGGER = LoggerFactory.getLogger(FacilitatorIndividualReportServiceImpl.class);
	/**
	* New Progress Report */
	@Transactional(propagation=Propagation.REQUIRED)
	public String populateAllProgressReports(String jctEmail) throws JCTException {
		LOGGER.info(">>>> FacilitatorIndividualReportServiceImpl.populateAllProgressReports");
		List<Object> viewList = null;
		List<Object> finalList1 = null;
		StringBuilder sb = new StringBuilder("");
		try {
			//viewList = reportDAO.getBeforeSketchListForExcel(profile);
			viewList = iFacilitatorIndividualReportDAO.populateAllProgressReports(jctEmail);
			for(int i=0;i < viewList.size();i++){
			finalList1 = (List) viewList.get(i);
			if(finalList1.size() > 0) {
 			for(int index = 0;index < finalList1.size();index++){
				Object[] innerObj = (Object[])finalList1.get(index);
				if(innerObj.length > 6) {
					// Fetch the list of time spent
					List<JctUserLoginInfo> list = (List<JctUserLoginInfo>) iFacilitatorIndividualReportDAO.getTime((String)innerObj[0]);
					String totalTime = calculateTotalTimeSpent(list);
					
					sb.append((String)innerObj[0]+"#");   	//User Name
					sb.append((String)innerObj[6]+"#");
					sb.append((java.sql.Timestamp)innerObj[1]+"#");   //Expiry
					sb.append((BigDecimal)innerObj[4]+"#");					
					sb.append((java.sql.Timestamp)innerObj[2]+"#");   // Start Time
					
					//sb.append((BigDecimal)innerObj[3]+"#"); // Total Time Spent
					sb.append(totalTime+"#");
					
					if(null != (String)innerObj[7] && (!innerObj[7].toString().equals(""))){ // last name
						sb.append((String)innerObj[7]+"#");
					} else {
						sb.append("N/A"+"#");
					}
					if(null != (String)innerObj[8] && (!innerObj[8].toString().equals(""))){ // first name
						sb.append((String)innerObj[8]+"#");
					} else {
						sb.append("N/A"+"#");
					}
					
					if(null != (String)innerObj[5] && (!innerObj[5].toString().equals(""))){ 
						sb.append((String)innerObj[5]+"#");
					} else {
						sb.append("N/A"+"#");
					}
				} else {					
					sb.append((String)innerObj[0]+"#");   	//User Name
					sb.append((String)innerObj[2]+"#");
					sb.append((java.sql.Timestamp)innerObj[1]+"#");   //Expiry
					sb.append("00"+"#");	
					sb.append("N/A"+"#");
					sb.append("00"+"#");	
					if(null != (String)innerObj[3] && (!innerObj[3].toString().equals(""))){ // last name
						sb.append((String)innerObj[3]+"#");
					} else {
						sb.append("N/A"+"#");
					}
					if(null != (String)innerObj[4] && (!innerObj[4].toString().equals(""))){ // first name
						sb.append((String)innerObj[4]+"#");
					} else {
						sb.append("N/A"+"#");
					}
					sb.append("N/A"+"#");
						
				}
			}
			sb.append("$$$");
			}
		}
	}
		 catch (DAOException e) {
			LOGGER.error(e.getLocalizedMessage());
		}
		LOGGER.info("<<<< FacilitatorIndividualReportServiceImpl.populateAllProgressReports");
		return sb.toString();
	}
	
	private String calculateTotalTimeSpent(List<JctUserLoginInfo> list) {
		long totalMillis = 0L;
		for (JctUserLoginInfo obj : list) {
			Calendar startCal = Calendar.getInstance();
			startCal.setTime(obj.getJctStartTime());
			
			Calendar endCal = Calendar.getInstance();
			endCal.setTime(obj.getJctEndTime());
			totalMillis += (endCal.getTimeInMillis() - startCal.getTimeInMillis());
		}
		return CommonUtility.convertMillis(totalMillis);
	}
	
	@Transactional(propagation=Propagation.REQUIRED)
	public String getUserDetails(int usrType) throws JCTException {
		LOGGER.info(">>>> ReportServiceImpl.getUserDetails");
		List<Object> userList = null;
		StringBuilder sBuilder = new StringBuilder("");
		try {
			userList = iFacilitatorIndividualReportDAO.getUserDetails(usrType);
			for(int index=0; index < userList.size(); index++){
					Object innerObj[] = (Object[])userList.get(index);
					sBuilder.append((String)innerObj[0]+"#");   			//USER NAME
					sBuilder.append((String)innerObj[1]+"#");   			//PROFILE NAME
					sBuilder.append((String)innerObj[2]+"#");   			//GROUP NAME
					//CREATED BY
					if((Integer)innerObj[6] == 3) {	//	if facilitator
						sBuilder.append(this.messageSource.getMessage("admin.admin.hash",null, null));					
					} else {
						if( ((String)innerObj[3]).substring(0, 2).equals("98") ){			//	99 = created by Admin
							sBuilder.append(this.messageSource.getMessage("admin.admin.hash",null, null));
						} else if ( ((String)innerObj[3]).substring(0, 2).equals("99") ){	//	98 = created by Facilitator
							sBuilder.append(this.messageSource.getMessage("admin.facilitator.hash",null, null));
						} else {
							sBuilder.append(this.messageSource.getMessage("admin.e-commerce.hash",null, null));
						}					
					}
					sBuilder.append((java.sql.Timestamp)innerObj[4]+"#");   //CREATED ON
				
					//	For Facilitator no Expiration date will be shown
					if( usrType == 3 ) {
						sBuilder.append("NA#");						
					} else if( usrType == 1 ) {
						sBuilder.append((java.sql.Timestamp)innerObj[5]+"#"); 						
					} 
					if( usrType == 0 ){		//	All user selected
						if( ((Integer)innerObj[6]) == 3 ){
							sBuilder.append("NA#"); 						
						} else {
							sBuilder.append((java.sql.Timestamp)innerObj[5]+"#");							
						}					
					}					
					sBuilder.append("$$$");
			}
		}
		catch (DAOException e) {
			LOGGER.error(e.getLocalizedMessage());
		}
		LOGGER.info("<<<< ReportServiceImpl.getUserDetails");
		return sBuilder.toString();	
	}
	
	@Transactional(propagation=Propagation.REQUIRED)
	public String createFaciIndiReport(String userType) throws JCTException {
		LOGGER.info(">>>> ReportServiceImpl.createFaciIndiReport");
		List<Object> viewList = null;
		StringBuilder sBuilder = new StringBuilder("");
		try{
			viewList = iFacilitatorIndividualReportDAO.createFaciIndiReport(userType);
			for(int index=0; index < viewList.size(); index++){
				Object[] innerObj = (Object[])viewList.get(index);
				sBuilder.append((String)innerObj[0]+"#");   			//USER NAME
				sBuilder.append((String)innerObj[1]+"#");   			//PROF NAME
				sBuilder.append((String)innerObj[2]+"#");   			//GROUP NAME				
																		//********* CREATED BY
				if( ((String)innerObj[3]).substring(0, 2).equals("98") ){			//	98 = Admin
					sBuilder.append(this.messageSource.getMessage("admin.admin.hash",null, null));  				
				} else if ( ((String)innerObj[3]).substring(0, 2).equals("99") ){	//	99 = Facilitator
					sBuilder.append(this.messageSource.getMessage("admin.facilitator.hash",null, null));					
				} else {
					sBuilder.append(this.messageSource.getMessage("admin.e-commerce.hash",null, null));
				}	
				sBuilder.append("$$$");
			}
		} catch (DAOException e){
			LOGGER.error(e.getLocalizedMessage());
		}
		LOGGER.info("<<<< ReportServiceImpl.createFaciIndiReport");
		return sBuilder.toString();
	}
	
	@Transactional(propagation=Propagation.REQUIRED)
	public String populateProgressReports(String jctEmail,String userGroupName) throws JCTException{
		List<Object> viewList = null;
		List<Object> finalList1 = null;
		StringBuilder sb = new StringBuilder("");
		try {
			//viewList = reportDAO.getBeforeSketchListForExcel(profile);
			viewList = iFacilitatorIndividualReportDAO.populateProgressReports(jctEmail,userGroupName);
			for(int i=0;i < viewList.size();i++){
			finalList1 = (List) viewList.get(i);
			if(finalList1.size() > 0) {
				for(int index = 0;index < finalList1.size();index++){
					Object[] innerObj = (Object[])finalList1.get(index);					
					if(innerObj.length > 6) {
						// Fetch the list of time spent
						List<JctUserLoginInfo> list = (List<JctUserLoginInfo>) iFacilitatorIndividualReportDAO.getTime((String)innerObj[0]);
						String totalTime = calculateTotalTimeSpent(list);
						
						sb.append((String)innerObj[0]+"#");   	//User Name
						sb.append((String)innerObj[6]+"#");
						sb.append((java.sql.Timestamp)innerObj[1]+"#");   //Expiry
						sb.append((BigDecimal)innerObj[4]+"#");					
						sb.append((java.sql.Timestamp)innerObj[2]+"#");   // Start Time
						
						//sb.append((BigDecimal)innerObj[3]+"#"); // Total Time Spent
						sb.append(totalTime+"#");
						
						if(null != (String)innerObj[7] && (!innerObj[7].toString().equals(""))){ // last name
							sb.append((String)innerObj[7]+"#");
						} else {
							sb.append("N/A"+"#");
						}
						if(null != (String)innerObj[8] && (!innerObj[8].toString().equals(""))){ // first name
							sb.append((String)innerObj[8]+"#");
						} else {
							sb.append("N/A"+"#");
						}
						
						if(null != (String)innerObj[5] && (!innerObj[5].toString().equals(""))){ 
							sb.append((String)innerObj[5]+"#");
						} else {
							sb.append("N/A"+"#");
						}
					} else {					
						sb.append((String)innerObj[0]+"#");   	//User Name
						sb.append((String)innerObj[2]+"#");
						sb.append((java.sql.Timestamp)innerObj[1]+"#");   //Expiry
						sb.append("00"+"#");	
						sb.append("N/A"+"#");
						sb.append("00"+"#");	
						if(null != (String)innerObj[3] && (!innerObj[3].toString().equals(""))){ // last name
							sb.append((String)innerObj[3]+"#");
						} else {
							sb.append("N/A"+"#");
						}
						if(null != (String)innerObj[4] && (!innerObj[4].toString().equals(""))){ // first name
							sb.append((String)innerObj[4]+"#");
						} else {
							sb.append("N/A"+"#");
						}
						sb.append("N/A"+"#");
							
					}
				}
				sb.append("$$$");
				}		
			}
		}
		 catch (DAOException e) {
			LOGGER.error(e.getLocalizedMessage());
		}
		LOGGER.info("<<<< FacilitatorIndividualReportServiceImpl.createNewProgressReport");
		return sb.toString();
	}
}
