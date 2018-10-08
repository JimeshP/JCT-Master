package com.vmware.jct.service.impl;

import java.math.BigInteger;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.HashSet;
import java.util.List;
import java.util.Set;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.context.MessageSource;
import org.springframework.stereotype.Service;
import org.springframework.transaction.annotation.Propagation;
import org.springframework.transaction.annotation.Transactional;

import com.vmware.jct.common.utility.CommonConstants;
import com.vmware.jct.dao.IReportDAO;
import com.vmware.jct.dao.dto.SurveyQuestionDTO;
import com.vmware.jct.exception.DAOException;
import com.vmware.jct.exception.JCTException;
import com.vmware.jct.service.IReportService;
import com.vmware.jct.service.vo.surveyQtnRedundentListVO;
import com.vmware.jct.service.vo.surveyQtnReportVO;
/**
 * 
 * <p><b>Class name:</b> ReportServiceImpl.java</p>
 * <p><b>Author:</b> InterraIT</p>
 * <p><b>Purpose:</b> This is a ReportServiceImpl class. This artifact is Business layer artifact.
 * ReportServiceImpl implement IReportService interface and override the following  methods.
 * -getBeforeSketchList()
 * -getBeforeSketchListForExcel()
 * -getAfterSketchListForExcel()
 * -getTotalCount()
 * -getActionPlanList()
 * -getActionPlanListForExcel()
 * -getTotalCountActionPlan()
 * -getASTotalCount()
 * -getAfterSketchList()
 * -getTotalCountBsToAs()
 * -getTotalTaskCount()
 * -getTotalMappingCount()
 * -getTotalRoleCount()
 * -getBsToAsList()
 * -getBsToAsListForExcel()
 * -createMiscHeaderReport()
 * -createMiscDetailesReport()
 * -getASDetailedViewOfMappings()
 * -getGroupProfileListForExcel()
 * -getProfileDescListForExcel()
 * -getLoginInfoDetails()
 * -getGroupCreationDate()
 * -getLoggedIndetails()
 * -getTotalTimeSpentOnBs()
 * -getTotalBsTasks()
 * -getBSDescriptionAndTimeEnergy()
 * -getTotalTimeSpentOnAs()
 * -getTotalAsTasks()
 * -getRoleFrame()
 * -getASDescriptionMainPersonAndTimeEnergy()
 * -getBsToAs()
 * -getStrValPassItems()
 * -getReflectionQuestions()
 * -populateActionPlan()
 * -populateSubQtnActionPlan()
 * -getActionPlanAnswers()
 * -populateASTaskLocation()
 * -createPaymentReport()
 * </p>
 * <p><b>Description:</b> This class used to perform insert, modify ,fetch data from db and delete operation. Save, update, delete operations are performed in DataAccessObject class </p>
 * <p><b>Revision History:</b>
 * 	<li>InterraIT - 19/Jan/2014 - Implement Exception </li>
 * <li>InterraIT - 31/Jan/2014 - Implement comments, introduce Named query, Transactional attribute </li>
 * </p>
 */
@Service
public class ReportServiceImpl implements IReportService{

	@Autowired
	private IReportDAO reportDAO;
	
	@Autowired  
	private MessageSource messageSource;
	
	private static final Logger LOGGER = LoggerFactory.getLogger(ReportServiceImpl.class);
	
@Transactional(propagation=Propagation.REQUIRED)
public List<Object> getBeforeSketchList(String occupationCode, int recordIndex) throws JCTException {
	LOGGER.info(">>>> ReportServiceImpl.getBeforeSketchList");
	List<Object> viewList = null;
	List<String> emailIdList = new ArrayList<String>();
	List<Object> totalViewList = new ArrayList<Object>();
	try {
		// search with 4, 0		
		viewList = reportDAO.getBeforeSketchList(occupationCode, recordIndex, 4, 0);
		for (int index = 0; index < viewList.size(); index++) {
			Object[] innerObj = (Object[])viewList.get(index);
			if (!emailIdList.contains((String)innerObj[0])) {
				emailIdList.add((String)innerObj[0]);
			}
		}
		totalViewList.addAll(viewList);
		viewList.clear();
		
		if (emailIdList.size() == 0) {
			emailIdList.clear();
			emailIdList.add("");
		}		
		viewList = reportDAO.getBeforeSketchList(occupationCode, recordIndex, 5, 0);
		for (int index = 0; index < viewList.size(); index++) {
			Object[] innerObj = (Object[])viewList.get(index);
			if (!emailIdList.contains((String)innerObj[0])) {
				emailIdList.add((String)innerObj[0]);
			}
		}
		totalViewList.addAll(viewList);
		viewList.clear();
	} catch (DAOException e) {
		LOGGER.error(e.getLocalizedMessage());
	}
	LOGGER.info("<<<<<< ReportServiceImpl.getBeforeSketchList");
	return totalViewList;
}


@Transactional(propagation=Propagation.REQUIRED)
public List<Object> getBeforeSketchListForExcel(String occupation) throws JCTException {
	LOGGER.info(">>>> ReportServiceImpl.getBeforeSketchListForExcel");
	List<Object> viewList = null;
	List<String> emailIdList = new ArrayList<String>();
	List<Object> totalViewList = new ArrayList<Object>();
	try {
		viewList = reportDAO.getBeforeSketchListForExcel(occupation, 4, 0);
		for (int index = 0; index < viewList.size(); index++) {
			Object[] innerObj = (Object[])viewList.get(index);
			if (!emailIdList.contains((String)innerObj[0])) {
				emailIdList.add((String)innerObj[0]);
			}
		}
		totalViewList.addAll(viewList);
		viewList.clear();				
		if (emailIdList.size() == 0) {
			emailIdList.clear();
			emailIdList.add("");
		}		
		viewList = reportDAO.getBeforeSketchListForExcel(occupation, 5, 0);
		for (int index = 0; index < viewList.size(); index++) {
			Object[] innerObj = (Object[])viewList.get(index);
			if (!emailIdList.contains((String)innerObj[0])) {
				emailIdList.add((String)innerObj[0]);
			}
		}
		totalViewList.addAll(viewList);
		viewList.clear();
		if (emailIdList.size() == 0) {
			emailIdList.clear();
			emailIdList.add("");
		}
		
		
		/*if (viewList.size() == 0) {
			viewList = reportDAO.getBeforeSketchListForExcel(function, jobLevel, 5, 0);
			if (viewList.size() == 0) {
				viewList = reportDAO.getBeforeSketchListForExcelMax(function, jobLevel);
			}
		}*/
		totalViewList.addAll(viewList);
	} catch (DAOException e) {
		LOGGER.error(e.getLocalizedMessage());
	}
	LOGGER.info("<<<< ReportServiceImpl.getBeforeSketchListForExcel");
	return totalViewList;
}
	
@Transactional(propagation=Propagation.REQUIRED)
public List<Object> getAfterSketchListForExcel(String occupation) throws JCTException {

	LOGGER.info(">>>> ReportServiceImpl.getAfterSketchListForExcel");
	List<Object> viewList = null;
	List<String> emailIdList = new ArrayList<String>();
	List<Object> totalViewList = new ArrayList<Object>();
	try {
		viewList = reportDAO.getAfterSketchListForExcel(null, occupation, 4, 0);
		for (int index = 0; index < viewList.size(); index++) {
			Object[] innerObj = (Object[])viewList.get(index);
			if (!emailIdList.contains((String)innerObj[0])) {
				emailIdList.add((String)innerObj[0]);
			}
		}
		totalViewList.addAll(viewList);
		viewList.clear();
		if (emailIdList.size() == 0) {
			emailIdList.clear();
			emailIdList.add("");
		}
		viewList = reportDAO.getAfterSketchListForExcel(emailIdList, occupation, 5, 0);
		for (int index = 0; index < viewList.size(); index++) {
			Object[] innerObj = (Object[])viewList.get(index);
			if (!emailIdList.contains((String)innerObj[0])) {
				emailIdList.add((String)innerObj[0]);
			}
		}
		totalViewList.addAll(viewList);
		viewList.clear();
		if (emailIdList.size() == 0) {
			emailIdList.clear();
			emailIdList.add("");
		}
		viewList = reportDAO.getAfterSketchListForExcel(emailIdList, occupation, 5, 1);
		totalViewList.addAll(viewList);
	} catch (DAOException e) {
		LOGGER.error(e.getLocalizedMessage());
	}
	LOGGER.info("<<<< ReportServiceImpl.getAfterSketchListForExcel");
	return totalViewList;

}

@Transactional(propagation=Propagation.REQUIRED)
public int getTotalCount(String function, String jobLevel) throws JCTException {
	LOGGER.info(">>>> ReportServiceImpl.getTotalCount");
	List<String> list = null;
	List<String> emailIdList = new ArrayList<String>();
	List<Object> totalViewList = new ArrayList<Object>();
	Set<String> emailSet = new HashSet<String>();
	try {
	list = reportDAO.getTotalCount(null, function, jobLevel, 4, 0);
	for (int index = 0; index < list.size(); index++) {
		String email = (String)list.get(index);
		if (!emailIdList.contains(email)) {
			emailIdList.add(email);
			emailSet.add(email);
		}
	}
	totalViewList.addAll(list);
	list.clear();
	if (emailIdList.size() == 0) {
		emailIdList.clear();
		emailIdList.add("");
	}
	
	list = reportDAO.getTotalCount(emailIdList, function, jobLevel, 5, 0);
	for (int index = 0; index < list.size(); index++) {
		String email = (String)list.get(index);
		if (!emailIdList.contains(email)) {
			emailIdList.add(email);
			emailSet.add(email);
		}
	}
	totalViewList.addAll(list);
	list.clear();
	if (emailIdList.size() == 0) {
		emailIdList.clear();
		emailIdList.add("");
	}
	
	list = reportDAO.getTotalCount(emailIdList, function, jobLevel, 5, 1);
	for (int index = 0; index < list.size(); index++) {
		String email = (String)list.get(index);
		if (!emailIdList.contains(email)) {
			emailIdList.add(email);
			emailSet.add(email);
		}
	}
	totalViewList.addAll(list);
	list.clear();
	if (emailIdList.size() == 0) {
		emailIdList.clear();
		emailIdList.add("");
	}
	
	
	
	
	/*count = list.size();
		if (count == 0) {
			list = reportDAO.getTotalCount(function, jobLevel, 5, 0);
			count = list.size();
			if (list.size() == 0) {
				list = reportDAO.getTotalCount(function, jobLevel, 5, 1);
				count = list.size();
			}
		}*/
	} catch(DAOException dao){
		LOGGER.error(dao.getLocalizedMessage());
		throw new JCTException();
	}
	LOGGER.info("<<<< ReportServiceImpl.getTotalCount");
	return emailSet.size();
}	

@Transactional(propagation=Propagation.REQUIRED)
public List<Object> getActionPlanList(String occupation, int recordIndex) throws JCTException {
	LOGGER.info(">>>> ReportServiceImpl.getActionPlanList");
	List<Object> viewList = null;
	List<String> emailIdList = new ArrayList<String>();
	List<Object> totalViewList = new ArrayList<Object>();
	try {		
		viewList = reportDAO.getActionPlanList(null,occupation, recordIndex, 4, 0);	
		for (int index = 0; index < viewList.size(); index++) {
			Object[] innerObj = (Object[])viewList.get(index);
			if (!emailIdList.contains((String)innerObj[0])) {
				emailIdList.add((String)innerObj[0]);
			}
		}
		totalViewList.addAll(viewList);
		viewList.clear();
		
		if (emailIdList.size() == 0) {
			emailIdList.clear();
			emailIdList.add("");
		}
		
		//viewList = reportDAO.getActionPlanList(emailIdList,occupation, recordIndex, 5, 0);	
		viewList = reportDAO.getActionPlanList(null,occupation, recordIndex, 5, 0);
		for (int index = 0; index < viewList.size(); index++) {
			Object[] innerObj = (Object[])viewList.get(index);
			if (!emailIdList.contains((String)innerObj[0])) {
				emailIdList.add((String)innerObj[0]);
			}
		}
		totalViewList.addAll(viewList);
		viewList.clear();
		
		if (emailIdList.size() == 0) {
			emailIdList.clear();
			emailIdList.add("");
		}
		
		//viewList = reportDAO.getActionPlanList(emailIdList, function, jobLevel, recordIndex, 5, 1);		
		/*if (viewList.size() == 0) {
			viewList = reportDAO.getActionPlanList(function, jobLevel, recordIndex, 5, 0);
			if (viewList.size() == 0) {
				viewList = reportDAO.getActionPlanListMax(function, jobLevel, recordIndex);
			}
		}*/
		totalViewList.addAll(viewList);	
	} catch (DAOException e) {
		LOGGER.error(e.getLocalizedMessage());
	}
	LOGGER.info("<<<< ReportServiceImpl.getActionPlanList");
	return totalViewList;
}

@Transactional(propagation=Propagation.REQUIRED)
public List<Object> getActionPlanListForExcel(String occupation) throws JCTException {
	LOGGER.info(">>>> ReportServiceImpl.getActionPlanListForExcel");
	List<Object> viewList = null;
	List<String> emailIdList = new ArrayList<String>();
	List<Object> totalViewList = new ArrayList<Object>();
	try {
		viewList = reportDAO.getActionPlanListForExcel(null,occupation, 4, 0);	
		for (int index = 0; index < viewList.size(); index++) {
			Object[] innerObj = (Object[])viewList.get(index);
			if (!emailIdList.contains((String)innerObj[0])) {
				emailIdList.add((String)innerObj[0]);
			}
		}
		totalViewList.addAll(viewList);
		viewList.clear();
		
		if (emailIdList.size() == 0) {
			emailIdList.clear();
			emailIdList.add("");
		}
		
		viewList = reportDAO.getActionPlanListForExcel(null,occupation, 5, 0);	
		for (int index = 0; index < viewList.size(); index++) {
			Object[] innerObj = (Object[])viewList.get(index);
			if (!emailIdList.contains((String)innerObj[0])) {
				emailIdList.add((String)innerObj[0]);
			}
		}		
		/*	viewList.clear();	
		if (emailIdList.size() == 0) {
			emailIdList.clear();
			emailIdList.add("");
		}
		viewList = reportDAO.getActionPlanListForExcel(null,function, jobLevel, 5, 1);	*/
		totalViewList.addAll(viewList);
		
	} catch (DAOException e) {
		LOGGER.error(e.getLocalizedMessage());
	}
	LOGGER.info("<<<< ReportServiceImpl.getActionPlanListForExcel");
	return totalViewList;
}

@Transactional(propagation=Propagation.REQUIRED)
public int getTotalCountActionPlan(String function) throws JCTException {
	LOGGER.info(">>>> ReportServiceImpl.getTotalCountActionPlan");
	List<String> list = null;
	List<String> emailIdList = new ArrayList<String>();
	List<String> totalViewList = new ArrayList<String>();
	Set<String> emailSet = new HashSet<String>();
	try{
		list = reportDAO.getTotalCountActionPlan(null, function, 4, 0);
		for (int index = 0; index < list.size(); index++) {
			String email = (String)list.get(index);
			if (!emailIdList.contains(email)) {
				emailIdList.add(email);
				emailSet.add(email);
			}
		}
		totalViewList.addAll(emailIdList);
		if (emailIdList.size() == 0) {
			emailIdList.clear();
			emailIdList.add("");
		}
		list = reportDAO.getTotalCountActionPlan(emailIdList, function, 5, 0);
		for (int index = 0; index < list.size(); index++) {
			String email = (String)list.get(index);
			if (!emailIdList.contains(email)) {
				emailIdList.add(email);
				emailSet.add(email);
			}
		}
		totalViewList.addAll(emailIdList);
		/*if (emailIdList.size() == 0) {
			emailIdList.clear();
			emailIdList.add("");
		}
		list = reportDAO.getTotalCountActionPlan(emailIdList, function, 5, 1);
		for (int index = 0; index < list.size(); index++) {
			String email = (String)list.get(index);
			if (!emailIdList.contains(email)) {
				emailIdList.add(email);
				emailSet.add(email);
			}
		}*/
		totalViewList.addAll(emailIdList);
		/*if (emailIdList.size() == 0) {
			emailIdList.clear();
			emailIdList.add("");
		}*/
		
	}catch(DAOException dao){
		LOGGER.error(dao.getLocalizedMessage());
		throw new JCTException();
	}
	LOGGER.info("<<<< ReportServiceImpl.getTotalCountActionPlan");
	return emailSet.size();
}

@Transactional(propagation=Propagation.REQUIRED)
public int getASTotalCount(String function, String jobLevel) throws JCTException {
	LOGGER.info(">>>> ReportServiceImpl.getASTotalCount");
	int count = 0;
	LOGGER.info("<<<< ReportServiceImpl.getASTotalCount");
	
	List<String> viewList = null;
	List<String> emailIdList = new ArrayList<String>();
	List<String> totalViewList = new ArrayList<String>();
	Set<String> emailSet = new HashSet<String>();
	try {
		viewList = reportDAO.getASTotalCount(null, function, jobLevel, 4, 0);
		for (int index = 0; index < viewList.size(); index++) {
			String email = (String)viewList.get(index);
			if (!emailIdList.contains(email)) {
				emailIdList.add(email);
				emailSet.add(email);
			}
		}
		totalViewList.addAll(emailIdList);
		if (emailIdList.size() == 0) {
			emailIdList.clear();
			emailIdList.add("");
		}
		viewList.clear();
		viewList = reportDAO.getASTotalCount(emailIdList, function, jobLevel, 5, 0);
		for (int index = 0; index < viewList.size(); index++) {
			String email = (String)viewList.get(index);
			if (!emailIdList.contains(email)) {
				emailIdList.add(email);
				emailSet.add(email);
			}
		}
		totalViewList.addAll(emailIdList);
		if (emailIdList.size() == 0) {
			emailIdList.clear();
			emailIdList.add("");
		}
		viewList.clear();
		viewList = reportDAO.getASTotalCount(null, function, jobLevel, 5, 1);
		for (int index = 0; index < viewList.size(); index++) {
			String email = (String)viewList.get(index);
			if (!emailIdList.contains(email)) {
				emailIdList.add(email);
				emailSet.add(email);
			}
		}
		totalViewList.addAll(emailIdList);
		//count = totalViewList.size();
		count = emailSet.size();
	} catch (DAOException ex) {
		
	}
	return count;
}

@Transactional(propagation=Propagation.REQUIRED)
public List<Object> getAfterSketchList(String function, String jobLevel, int recordIndex) throws JCTException {
	LOGGER.info(">>>> ReportServiceImpl.getAfterSketchList");
	List<Object> viewList = null;
	List<String> emailIdList = new ArrayList<String>();
	List<Object> totalViewList = new ArrayList<Object>();
	try {
		viewList = reportDAO.getAfterSketchList(null, function, jobLevel, recordIndex, 4, 0);
		for (int index = 0; index < viewList.size(); index++) {
			Object[] innerObj = (Object[])viewList.get(index);
			if (!emailIdList.contains((String)innerObj[0])) {
				emailIdList.add((String)innerObj[0]);
			}
		}
		totalViewList.addAll(viewList);
		viewList.clear();
		
		if (emailIdList.size() == 0) {
			emailIdList.clear();
			emailIdList.add("");
		}
		
		viewList = reportDAO.getAfterSketchList(emailIdList, function, jobLevel, recordIndex, 5, 0);
		for (int index = 0; index < viewList.size(); index++) {
			Object[] innerObj = (Object[])viewList.get(index);
			if (!emailIdList.contains((String)innerObj[0])) {
				emailIdList.add((String)innerObj[0]);
			}
		}
		totalViewList.addAll(viewList);
		viewList.clear();
		
		/*if (emailIdList.size() == 0) {
			emailIdList.clear();
			emailIdList.add("");
		}
		
		viewList = reportDAO.getAfterSketchList(emailIdList, function, jobLevel, recordIndex, 5, 1);
		totalViewList.addAll(viewList);*/
	 } catch (DAOException e) {
		LOGGER.error(e.getLocalizedMessage());
	}
	LOGGER.info("<<<< ReportServiceImpl.getAfterSketchList");
	return totalViewList;
}

@Transactional(propagation=Propagation.REQUIRED)
public int getTotalCountBsToAs(String function, String jobLevel) throws JCTException {
	LOGGER.info(">>>> ReportServiceImpl.getTotalCountBsToAs");
	int count = 0;
	try{
		List<String> list = reportDAO.getTotalCountBsToAs(function, jobLevel);
		count = list.size();
	}catch(DAOException dao){
		LOGGER.error(dao.getLocalizedMessage());
		throw new JCTException();
	}
	LOGGER.info("<<<< ReportServiceImpl.getTotalCountBsToAs");
	return count;
}

@Transactional(propagation=Propagation.REQUIRED)
public Integer getTotalTaskCount(String emailId) {
	LOGGER.info(">>>> ReportServiceImpl.getTotalTaskCount");
	Integer count = 0;
	try {
		BigInteger tempCount = (java.math.BigInteger) reportDAO.getTotalTaskCount(emailId, 4, 0);
		count = tempCount.intValue();
		if (count == 0) {
			BigInteger tempCount1 = (java.math.BigInteger) reportDAO.getTotalTaskCount(emailId, 5, 0);
			count = tempCount1.intValue();
			if (count == 0) {
				BigInteger tempCount2 = (java.math.BigInteger) reportDAO.getTotalTaskCount(emailId, 5, 1);
				count = tempCount2.intValue();
			}
		}
	} catch (DAOException e) {
		LOGGER.error(e.getLocalizedMessage());
	}
	LOGGER.info("<<<< ReportServiceImpl.getTotalTaskCount");
	return count;

}

@Transactional(propagation=Propagation.REQUIRED)
public Integer getTotalMappingCount(String emailId) {
	LOGGER.info(">>>> ReportServiceImpl.getTotalMappingCount");
	Integer count = 0;
	try {
		BigInteger tempCount = (java.math.BigInteger) reportDAO.getTotalMappingCount(emailId);;
		count = tempCount.intValue();
	} catch (DAOException e) {
		LOGGER.error(e.getLocalizedMessage());
	}
	LOGGER.info("<<<< ReportServiceImpl.getTotalMappingCount");
	return count;
}

@Transactional(propagation=Propagation.REQUIRED)
public Integer getTotalRoleCount(String emailId) {
	LOGGER.info(">>>> ReportServiceImpl.getTotalRoleCount");
	Integer count = 0;
	try {
		BigInteger tempCount = (java.math.BigInteger) reportDAO.getTotalRoleCount(emailId, 4, 0);
		count = tempCount.intValue();
		if (count == 0) {
			BigInteger tempCount1 = (java.math.BigInteger) reportDAO.getTotalRoleCount(emailId, 5, 0);
			count = tempCount1.intValue();
			if (count == 0) {
				BigInteger tempCount2 = (java.math.BigInteger) reportDAO.getTotalRoleCount(emailId, 5, 1);
				count = tempCount2.intValue();
			}
		}
	} catch (DAOException e) {
		LOGGER.error(e.getLocalizedMessage());
	}
	LOGGER.info("<<<< ReportServiceImpl.getTotalRoleCount");
	return count;
}

@Transactional(propagation=Propagation.REQUIRED)
public List<Object> getBsToAsList(String occupation, int recordIndex) throws JCTException {
	LOGGER.info(">>>> ReportServiceImpl.getBsToAsList");
	List<Object> viewList = null;
	try {
		viewList = reportDAO.getBsToAsList(occupation, recordIndex);
	} catch (DAOException e) {
		LOGGER.error(e.getLocalizedMessage());
	}
	LOGGER.info("<<<< ReportServiceImpl.getBsToAsList");
	return viewList;
}

@Transactional(propagation=Propagation.REQUIRED)
public List<Object> getBsToAsListForExcel(String occupation) throws JCTException {
	LOGGER.info(">>>> ReportServiceImpl.getBsToAsListForExcel");
	List<Object> viewList = null;
	try {
		viewList = reportDAO.getBsToAsListForExcel(occupation);
	} catch (DAOException e) {
		LOGGER.error(e.getLocalizedMessage());
	}
	LOGGER.info("<<<< ReportServiceImpl.getBsToAsListForExcel");
	return viewList;
}

@Transactional(propagation=Propagation.REQUIRED)
public String createMiscHeaderReport() throws JCTException {
	LOGGER.info(">>>> ReportServiceImpl.createMiscHeaderReport");
	List<Object> viewList = null;
	StringBuilder sBuilder = new StringBuilder("");
	try{
		viewList = reportDAO.createMiscHeaderReport();
		for(int index=0; index < viewList.size(); index++){
			Object[] innerObj = (Object[])viewList.get(index);
			sBuilder.append((String)innerObj[0]+"#");   			//EMAIL
			sBuilder.append((String)innerObj[1]+"#");   			//JRNOS
			sBuilder.append((java.sql.Timestamp)innerObj[2]+"#");   //MIN
			sBuilder.append((java.sql.Timestamp)innerObj[3]+"#");   //MAX
			sBuilder.append((java.math.BigInteger)innerObj[4]+"#"); 	//COUNT
			sBuilder.append((String)innerObj[6]+", "+(String)innerObj[5]+"#"); //First and Last Name
			java.math.BigInteger statusSearchCount = reportDAO.getSearchableCount((String)innerObj[1]);
			if (statusSearchCount.intValue() > 0) {
				sBuilder.append("Yes"); 	//COUNT
			} else {
				sBuilder.append("No");
			}
			sBuilder.append("$$$");
		}
	} catch (DAOException e){
		LOGGER.error(e.getLocalizedMessage());
	}
	LOGGER.info("<<<< ReportServiceImpl.createMiscHeaderReport");
	return sBuilder.toString();
}
	
@Transactional(propagation=Propagation.REQUIRED)
public String createMiscDetailesReport(String jobReferenceNo) throws JCTException{
	LOGGER.info(">>>> ReportServiceImpl.createMiscDetailesReport");
	List<Object> viewList = null;
	StringBuilder sBuilder = new StringBuilder("");
	try{
		viewList = reportDAO.createMiscDetailedReport(jobReferenceNo);
		for(int index=0; index < viewList.size(); index++){
			Object[] innerObj = (Object[])viewList.get(index);
			sBuilder.append((java.sql.Timestamp)innerObj[0]+"#");   //START TIME
			sBuilder.append((java.sql.Timestamp)innerObj[1]+"#");   //END TIME
			sBuilder.append("$$$");
		}
	} catch (DAOException e){
		LOGGER.error(e.getLocalizedMessage());
	}
	LOGGER.info("<<<< ReportServiceImpl.createMiscDetailesReport");
	return sBuilder.toString();
}

@Transactional(propagation=Propagation.REQUIRED)
public String getASDetailedViewOfMappings (String emailId) {
	LOGGER.info(">>>> ReportServiceImpl.getASDetailedViewOfMappings");
	StringBuilder mappingBuilder = new StringBuilder("");
	try {
		List<Object> strObject = null;
		strObject = reportDAO.getASDetailedViewForMappings(emailId,CommonConstants.JOB_ATTRIBUTE_STRENGTH, 4, 0);
		if (strObject.size() == 0) {
			strObject = reportDAO.getASDetailedViewForMappings(emailId,CommonConstants.JOB_ATTRIBUTE_STRENGTH, 5, 0);
			if (strObject.size() == 0) {
				strObject = reportDAO.getASDetailedViewForMappings(emailId,CommonConstants.JOB_ATTRIBUTE_STRENGTH, 5, 1);
			} 
		}
		mappingBuilder.append(this.getMappingDetailText(strObject));
		mappingBuilder.append(" &&&");
		
		List<Object> valObject = null;
		valObject = reportDAO.getASDetailedViewForMappings(emailId,CommonConstants.JOB_ATTRIBUTE_VALUE, 4, 0);
		if (valObject.size() == 0) {
			valObject = reportDAO.getASDetailedViewForMappings(emailId,CommonConstants.JOB_ATTRIBUTE_VALUE, 5, 0);
			if (valObject.size() == 0) {
				valObject = reportDAO.getASDetailedViewForMappings(emailId,CommonConstants.JOB_ATTRIBUTE_VALUE, 5, 1);
			}
		}
		mappingBuilder.append(this.getMappingDetailText(valObject));
		mappingBuilder.append(" &&&");
		List<Object> psnObject = null;
		psnObject = reportDAO.getASDetailedViewForMappings(emailId,CommonConstants.JOB_ATTRIBUTE_PASSION, 4, 0);
		if (psnObject.size() == 0) {
			psnObject = reportDAO.getASDetailedViewForMappings(emailId,CommonConstants.JOB_ATTRIBUTE_PASSION, 5, 0);
			if (psnObject.size() == 0) {
				psnObject = reportDAO.getASDetailedViewForMappings(emailId,CommonConstants.JOB_ATTRIBUTE_PASSION, 5, 1);
			} 
		}
		mappingBuilder.append(this.getMappingDetailText(psnObject));
		mappingBuilder.append(" &&&");
	} catch (DAOException daoEx) {
		LOGGER.error(daoEx.getLocalizedMessage());
	}
	LOGGER.info("<<<< ReportServiceImpl.getASDetailedViewOfMappings");
	return mappingBuilder.toString();
}

private String getMappingDetailText (List<Object> object) {
	LOGGER.info(">>>> ReportServiceImpl.getMappingDetailText");
	StringBuilder sb = new StringBuilder("");
	for (int index = 0; index < object.size(); index++) {
		sb.append((String) object.get(index));
		sb.append("#");
	}
	LOGGER.info("<<<< ReportServiceImpl.getMappingDetailText");
	return sb.toString();
}

@Transactional(propagation=Propagation.REQUIRED)
public List<Object> getGroupProfileListForExcel(int profile) throws JCTException {
	LOGGER.info(">>>> ReportServiceImpl.getGroupProfileListForExcel");
	List<Object> viewList = null;
	try {
		//viewList = reportDAO.getBeforeSketchListForExcel(profile);
		viewList = reportDAO.getGroupProfileListForExcel(profile);
		
	} catch (DAOException e) {
		LOGGER.error(e.getLocalizedMessage());
	}
	LOGGER.info("<<<< ReportServiceImpl.getGroupProfileListForExcel");
	return viewList;
}

@Transactional(propagation=Propagation.REQUIRED)
public List<Object> getProfileDescListForExcel(int profile) throws JCTException {
	LOGGER.info(">>>> ReportServiceImpl.getProfileDescListForExcel");
	List<Object> viewList = null;
	try {
		//viewList = reportDAO.getBeforeSketchListForExcel(profile);
		viewList = reportDAO.getProfileDescListForExcel(profile);
		
	} catch (DAOException e) {
		LOGGER.error(e.getLocalizedMessage());
	}
	LOGGER.info("<<<< ReportServiceImpl.getProfileDescListForExcel");
	return viewList;
}

@Transactional(propagation=Propagation.REQUIRED)
public List<Object> getLoginInfoDetails(String emailId) throws JCTException {
	LOGGER.info(">>>> ReportServiceImpl.getLoginInfoDetails");
	List<Object> viewList = null;
	try {
		viewList = reportDAO.getLoginInfoDetails(emailId);
	} catch (DAOException e) {
		LOGGER.error(e.getLocalizedMessage());
	}
	LOGGER.info("<<<< ReportServiceImpl.getLoginInfoDetails");
	return viewList;
}

@Transactional(propagation=Propagation.REQUIRED)
public java.sql.Timestamp getGroupCreationDate(Integer groupName) throws JCTException {
	LOGGER.info(">>>> ReportServiceImpl.getGroupCreationDate");
	java.sql.Timestamp crationDate = null;
	try {
		crationDate = reportDAO.getGroupCreationDate(groupName);
	} catch (DAOException e) {
		LOGGER.error(e.getLocalizedMessage());
	}
	LOGGER.info("<<<< ReportServiceImpl.getGroupCreationDate");
	return crationDate;
}

@Transactional(propagation=Propagation.REQUIRED)
public List<Object> getLoggedIndetails(String email) throws JCTException {
	LOGGER.info(">>>> ReportServiceImpl.getLoggedIndetails");
	List<Object> viewList = null;
	try {
		viewList = reportDAO.getLoggedIndetails(email);
	} catch (DAOException e) {
		LOGGER.error(e.getLocalizedMessage());
	}
	LOGGER.info("<<<< ReportServiceImpl.getLoggedIndetails");
	return viewList;
}

@Transactional(propagation=Propagation.REQUIRED)
public Integer getTotalTimeSpentOnBs(String emailId) throws JCTException {
	LOGGER.info(">>>> ReportServiceImpl.getTotalTimeSpentOnBs");
	Integer count = 0;
	try {
		count = reportDAO.getTotalTimeSpentOnBs(emailId, 4, 0);
		if (count == null) {
			count = reportDAO.getTotalTimeSpentOnBs(emailId, 5, 0);
			if (count == null) {
				count = reportDAO.getTotalTimeSpentOnBs(emailId, 5, 1);
			}
		}
	} catch (DAOException e) {
		LOGGER.error(e.getLocalizedMessage());
	}
	LOGGER.info("<<<< ReportServiceImpl.getTotalTimeSpentOnBs");
	return count;
}

@Transactional(propagation=Propagation.REQUIRED)
public BigInteger getTotalBsTasks(String emailId) throws JCTException {
	LOGGER.info(">>>> ReportServiceImpl.getTotalBsTasks");
	BigInteger nosOfBsTasks = new BigInteger("0");
	try {
		nosOfBsTasks = reportDAO.getTotalBsTasks(emailId);
	} catch (DAOException e) {
		LOGGER.error(e.getLocalizedMessage());
	}
	LOGGER.info("<<<< ReportServiceImpl.getTotalBsTasks");
	return nosOfBsTasks;
}

@Transactional(propagation=Propagation.REQUIRED)
public List<Object> getBSDescriptionAndTimeEnergy(String emailId)
		throws JCTException {
	LOGGER.info(">>>> ReportServiceImpl.getBSDescriptionAndTimeEnergy");
	List<Object> viewList = null;
	try {
		viewList = reportDAO.getBSDescriptionAndTimeEnergy(emailId);
	} catch (DAOException e) {
		LOGGER.error(e.getLocalizedMessage());
	}
	LOGGER.info("<<<< ReportServiceImpl.getBSDescriptionAndTimeEnergy");
	return viewList;
}

@Transactional(propagation=Propagation.REQUIRED)
public Integer getTotalTimeSpentOnAs(String emailId) throws JCTException {
	LOGGER.info(">>>> ReportServiceImpl.getTotalTimeSpentOnAs");
	Integer count = 0;
	try {
		count = reportDAO.getTotalTimeSpentOnAs(emailId, 4, 0);
		if (count == null) {
			count = reportDAO.getTotalTimeSpentOnAs(emailId, 5, 0);
			if (count == null) {
				count = reportDAO.getTotalTimeSpentOnAs(emailId, 5, 1);
			}
		}
	} catch (DAOException e) {
		LOGGER.error(e.getLocalizedMessage());
	}
	LOGGER.info("<<<< ReportServiceImpl.getTotalTimeSpentOnAs");
	return count;
}

@Transactional(propagation=Propagation.REQUIRED)
public List<Object> getTotalAsTasks(String emailId) throws JCTException {
	LOGGER.info(">>>> ReportServiceImpl.getTotalAsTasks");
	List<Object> nosOfAsTasksAndRoleFrame = null;
	try {
		nosOfAsTasksAndRoleFrame = reportDAO.getTotalAsTasks(emailId);
	} catch (DAOException e) {
		LOGGER.error(e.getLocalizedMessage());
	}
	LOGGER.info("<<<< ReportServiceImpl.getTotalAsTasks");
	return nosOfAsTasksAndRoleFrame;
}

@Transactional(propagation=Propagation.REQUIRED)
public List<Object> getRoleFrame (String emailId) throws JCTException {
	LOGGER.info(">>>> ReportServiceImpl.getRoleFrame");
	List<Object> roleFrameList = null;
	try {
		roleFrameList = reportDAO.getRoleFrame(emailId);
	} catch (DAOException e) {
		LOGGER.error(e.getLocalizedMessage());
	}
	LOGGER.info("<<<< ReportServiceImpl.getRoleFrame");
	return roleFrameList;
}

@Transactional(propagation=Propagation.REQUIRED)
public List<Object> getASDescriptionMainPersonAndTimeEnergy(String emailId)
		throws JCTException {
	LOGGER.info(">>>> ReportServiceImpl.getASDescriptionMainPersonAndTimeEnergy");
	List<Object> viewList = null;
	try {
		viewList = reportDAO.getASDescriptionMainPersonAndTimeEnergy(emailId);
	} catch (DAOException e) {
		LOGGER.error(e.getLocalizedMessage());
	}
	LOGGER.info("<<<< ReportServiceImpl.getASDescriptionMainPersonAndTimeEnergy");
	return viewList;
}

@Transactional(propagation=Propagation.REQUIRED)
public List<Object> getBsToAs(String emailId) throws JCTException {
	LOGGER.info(">>>> ReportServiceImpl.getBsToAs");
	List<Object> bsToAsList = null;
	try {
		bsToAsList = reportDAO.getBsToAs(emailId);
	} catch (DAOException e) {
		LOGGER.error(e.getLocalizedMessage());
	}
	LOGGER.info("<<<< ReportServiceImpl.getBsToAs");
	return bsToAsList;
}

@Transactional(propagation=Propagation.REQUIRED)
public List<Object> getStrValPassItems(String emailId) throws JCTException {
	LOGGER.info(">>>> ReportServiceImpl.getStrValPassItems");
	List<Object> mappingList = null;
	try {
		mappingList = reportDAO.getStrValPassItems(emailId);
	} catch (DAOException e) {
		LOGGER.error(e.getLocalizedMessage());
	}
	LOGGER.info("<<<< ReportServiceImpl.getStrValPassItems");
	return mappingList;
}

@Transactional(propagation=Propagation.REQUIRED)
public List<Object> getReflectionQuestions(String emailId)
		throws JCTException {
	LOGGER.info(">>>> ReportServiceImpl.getReflectionQuestions");
	List<Object> reflectionQtnList = null;
	try {
		reflectionQtnList = reportDAO.getReflectionQuestions(emailId, 4, 0, "N");
		if (reflectionQtnList.size() == 0) {
			reflectionQtnList = reportDAO.getReflectionQuestions(emailId, 5, 0, "Y");
			if (reflectionQtnList.size() == 0) {
				reflectionQtnList = reportDAO.getReflectionQuestions(emailId, 5, 1, "Y");
			}
		}
	} catch (DAOException e) {
		LOGGER.error(e.getLocalizedMessage());
	}
	LOGGER.info("<<<< ReportServiceImpl.getReflectionQuestions");
	return reflectionQtnList;
}

@Transactional(propagation=Propagation.REQUIRED)
public List<Object> populateActionPlan(String emailId) throws JCTException {
	LOGGER.info(">>>> ReportServiceImpl.populateActionPlan");
	List<Object> actionPlanList = null;
	try {
		actionPlanList = reportDAO.populateActionPlan(emailId, 4, 0);
		if (actionPlanList.size() == 0) {
			actionPlanList = reportDAO.populateActionPlan(emailId, 5, 0);
			if (actionPlanList.size() == 0) {
				actionPlanList = reportDAO.populateActionPlan(emailId, 5, 1);
			}
		}
	} catch (DAOException e) {
		LOGGER.error(e.getLocalizedMessage());
	}
	LOGGER.info("<<<< ReportServiceImpl.populateActionPlan");
	return actionPlanList;

}

@Transactional(propagation=Propagation.REQUIRED)
public List<Object> populateSubQtnActionPlan(String mainQuestion,
		String emailId) throws JCTException {
	LOGGER.info(">>>> ReportServiceImpl.populateSubQtnActionPlan");
	List<Object> actionPlanSubQtnList = null;
	try {
		actionPlanSubQtnList = reportDAO.populateSubQtnActionPlan(mainQuestion, emailId, "N", 4, 0);
		if (actionPlanSubQtnList.size() == 0) {
			actionPlanSubQtnList = reportDAO.populateSubQtnActionPlan(mainQuestion, emailId, "N", 5, 0);
			if (actionPlanSubQtnList.size() == 0) {
				actionPlanSubQtnList = reportDAO.populateSubQtnActionPlan(mainQuestion, emailId, "Y", 5, 1);
			}
		}
		
	} catch (DAOException e) {
		LOGGER.error(e.getLocalizedMessage());
	}
	LOGGER.info("<<<< ReportServiceImpl.populateSubQtnActionPlan");
	return actionPlanSubQtnList;
}

@Transactional(propagation=Propagation.REQUIRED)
public List<Object> getActionPlanAnswers(String mainQuestion,
		String subQuestion, String emailId) throws JCTException {
	LOGGER.info(">>>> ReportServiceImpl.getActionPlanAnswers");
	List<Object> actionPlanAnsList = null;
	try {
		actionPlanAnsList = reportDAO.getActionPlanAnswers(mainQuestion, subQuestion, emailId, "N", 4, 0);
		if (actionPlanAnsList.size() == 0) {
			actionPlanAnsList = reportDAO.getActionPlanAnswers(mainQuestion, subQuestion, emailId, "N", 5, 0);
			if (actionPlanAnsList.size() == 0) {
				actionPlanAnsList = reportDAO.getActionPlanAnswers(mainQuestion, subQuestion, emailId, "y", 5, 1);
			}
		}
		
	} catch (DAOException e) {
		LOGGER.error(e.getLocalizedMessage());
	}
	LOGGER.info("<<<< ReportServiceImpl.getActionPlanAnswers");
	return actionPlanAnsList;
}

@Transactional(propagation=Propagation.REQUIRED)
public List<Object> populateASTaskLocation(String emailId)
		throws JCTException {
	LOGGER.info(">>>> ReportServiceImpl.populateASTaskLocation");
	List<Object> asTaskLocationList = null;
	try {
		asTaskLocationList = reportDAO.populateASTaskLocation(emailId);
	} catch (DAOException e) {
		LOGGER.error(e.getLocalizedMessage());
	}
	LOGGER.info("<<<< ReportServiceImpl.populateASTaskLocation");
	return asTaskLocationList;
}

@Transactional(propagation=Propagation.REQUIRED)
public List<Object> populateSubQtnQuestionnaire(String mainQuestion,
		String emailId) throws JCTException {
	LOGGER.info(">>>> ReportServiceImpl.populateSubQtnQuestionnaire");
	List<Object> questionnaireSubQtnList = null;
	try {
		questionnaireSubQtnList = reportDAO.populateSubQtnQuestionnaire(mainQuestion, emailId, "N", 4, 0);
		if (questionnaireSubQtnList.size() == 0) {
			questionnaireSubQtnList = reportDAO.populateSubQtnQuestionnaire(mainQuestion, emailId, "N", 5, 0);
			if (questionnaireSubQtnList.size() == 0) {
				questionnaireSubQtnList = reportDAO.populateSubQtnQuestionnaire(mainQuestion, emailId, "Y", 5, 1);
			}
		}
	} catch (DAOException e) {
		LOGGER.error(e.getLocalizedMessage());
	}
	LOGGER.info("<<<< ReportServiceImpl.populateSubQtnQuestionnaire");
	return questionnaireSubQtnList;
}

@Transactional(propagation=Propagation.REQUIRED)
public List<Object> getQuestionnaireAnswers(String mainQuestion,
		String subQuestion, String emailId) throws JCTException {
	LOGGER.info(">>>> ReportServiceImpl.getQuestionnaireAnswers");
	List<Object> questionnaireAnsList = null;
	try {
		questionnaireAnsList = reportDAO.getQuestionnaireAnswers(mainQuestion, subQuestion, emailId, "N", 4, 0);
		if (questionnaireAnsList.size() == 0) {
			questionnaireAnsList = reportDAO.getQuestionnaireAnswers(mainQuestion, subQuestion, emailId, "N", 5, 0);
			if (questionnaireAnsList.size() == 0) {
				questionnaireAnsList = reportDAO.getQuestionnaireAnswers(mainQuestion, subQuestion, emailId, "Y", 5, 1);
			}
		}
	} catch (DAOException e) {
		LOGGER.error(e.getLocalizedMessage());
	}
	LOGGER.info("<<<< ReportServiceImpl.getQuestionnaireAnswers");
	return questionnaireAnsList;
}

@Transactional(propagation=Propagation.REQUIRED)
public List<String> getRolesForTasks(String emailId, String taskDesc)
		throws JCTException {
	LOGGER.info(">>>> ReportServiceImpl.getRolesForTasks");
	List<String> list = null;
	try {
		list = reportDAO.getRolesForTasks(emailId, taskDesc, 4, 0);
		if (list.size() == 0) {
			list = reportDAO.getRolesForTasks(emailId, taskDesc, 5, 0);
			if (list.size() == 0) {
				list = reportDAO.getRolesForTasks(emailId, taskDesc, 5, 1);
			}
		}
	} catch (DAOException e) {
		LOGGER.error(e.getLocalizedMessage());
	}
	LOGGER.info("<<<< ReportServiceImpl.getRolesForTasks");
	return list;
}

@Transactional(propagation=Propagation.REQUIRED)
public String createSurveyReport(int usrType) throws JCTException {
	LOGGER.info(">>>> ReportServiceImpl.createSurveyReport");
	List<Object> viewList = null;
	StringBuilder sBuilder = new StringBuilder("");
	String userName = "";
	try{
		viewList = reportDAO.createSurveyReport(usrType);
		for(int index=0; index < viewList.size(); index++){
			Object[] innerObj = (Object[])viewList.get(index);
			if(!userName.equalsIgnoreCase((String)innerObj[0])) {					
				sBuilder.append("@@@");
				sBuilder.append((String)innerObj[0]+"#");   			//NAME
				sBuilder.append((Integer)innerObj[1]+"#");				//CREATED BY
				sBuilder.append((java.sql.Timestamp)innerObj[2]+"#");   			//CREATED ON
				sBuilder.append((java.sql.Timestamp)innerObj[3]+"#");   //A/c EXP DATE
				sBuilder.append((String)innerObj[4]+"#");   //SURVEY QTN
				sBuilder.append((String)innerObj[5]+"#"); 	//ANSWER
				sBuilder.append("$$$");					
			} else {
				sBuilder.append((String)innerObj[0]+"#");   			//NAME
				sBuilder.append((Integer)innerObj[1]+"#");				//CREATED BY
				sBuilder.append((java.sql.Timestamp)innerObj[2]+"#");   			//CREATED ON
				sBuilder.append((java.sql.Timestamp)innerObj[3]+"#");   //A/c EXP DATE
				sBuilder.append((String)innerObj[4]+"#");   //SURVEY QTN
				sBuilder.append((String)innerObj[5]+"#"); 	//ANSWER
				sBuilder.append("$$$");
			}
			userName = (String)innerObj[0];
			}
	} catch (DAOException e){
		LOGGER.error(e.getLocalizedMessage());
	}
	LOGGER.info("<<<< ReportServiceImpl.createSurveyReport");
	return sBuilder.toString();
}

@Transactional(propagation=Propagation.REQUIRED)
public List<SurveyQuestionDTO> fetchUserIdFromSurvey(int userType) throws JCTException {
	LOGGER.info(">>>> ReportServiceImpl.fetchUserIdFromSurvey");
	List<SurveyQuestionDTO> uIds = null;
	try{
		uIds = reportDAO.fetchUserIdFromSurvey(userType);			
	} catch (DAOException e){
		LOGGER.error(e.getLocalizedMessage());	
	}
	LOGGER.info("<<<< ReportServiceImpl.fetchUserIdFromSurvey");
	return uIds;
}
/**
 * Method to built survey qtn List-String
 * @param answeredUserIdList
 * @throws JCTException
 * @return List
 * */	
@Transactional(propagation=Propagation.REQUIRED)
public List<surveyQtnReportVO> getSurveyListForExcel(List<SurveyQuestionDTO> answeredUserIdList) throws JCTException {
	LOGGER.info(">>>> ReportServiceImpl.getSurveyListForExcel");
	List<surveyQtnReportVO> surveyQtnReport = new ArrayList<surveyQtnReportVO>();
	try {
		for(int j = 0 ; j < answeredUserIdList.size() ; j++){			
			int uId = (Integer) answeredUserIdList.get(j).getSurveyQtnUserId();	
			surveyQtnReportVO vo = new surveyQtnReportVO();
			Object[] viewList = (Object[]) reportDAO.getSurveyListCommonFieldsForExcel(uId);
			
			if((viewList.length > 0) || (viewList != null)){
				SimpleDateFormat dateFormat = new SimpleDateFormat("dd MMMM, yyyy");
				int tempId = Integer.parseInt(viewList[3].toString().substring(0,2));						
				String createdBy = tempId == 98	? "Admin" : (tempId == 99 ? "Facilitator" : "E-commerce");
				String userType = (Integer)viewList[4] == 1 ? "Individual User" : "Facilitator User" ;				
				/********************** Common fields for each user **************************/	
				vo.setName(viewList[7].toString()+","+viewList[6].toString());
				vo.setUserName(viewList[0].toString());
				vo.setProfileName(viewList[1].toString());
				vo.setGroupName(viewList[2].toString());
				vo.setCustId(viewList[3].toString());
				vo.setUserType(userType);
				vo.setCreatedBy(createdBy);
				vo.setExpirationDate(dateFormat.format(viewList[5]));				
				/************* append Survey_Ques|type|ans (redundent fields) *****************/
				List<Object> redundentList = (List<Object>) reportDAO.getSurveyListRedundentFields(uId);
				List<surveyQtnRedundentListVO> innerList = new ArrayList<surveyQtnRedundentListVO>();
				for(int i = 0 ; i < redundentList.size() ; i++){
					surveyQtnRedundentListVO redundantVO = new surveyQtnRedundentListVO();
					Object[] obj = (Object[]) redundentList.get(i);
					String quesType = (Integer)obj[1] == 1 ? "Text" : ((Integer)obj[1] == 2 ? "Radio Button" : ((Integer)obj[1] == 3) ? "Dropdown" : "Checkbox");					
					redundantVO.setQtn(obj[0].toString());
					redundantVO.setQtnType(quesType);
					redundantVO.setAnswer(obj[2].toString());					
					innerList.add(redundantVO);				
				}
				vo.setQtnList(innerList);
			}
			surveyQtnReport.add(vo);
		}	//	for
		
	} catch (DAOException e){
		LOGGER.error(e.getLocalizedMessage());
	}
	LOGGER.info("<<<< ReportServiceImpl.getSurveyListForExcel");
	return surveyQtnReport;
}
/**
 * method to generate excel body which will contain all payment details
 * @param userType
 * @return String
 **/	
/*@Transactional(propagation=Propagation.REQUIRED)
public String createPaymentReport( int userType ) throws JCTException {
	LOGGER.info(">>>> ReportServiceImpl.createPaymentReport");
	List<Object> viewList = null;
	StringBuilder sBuilder = new StringBuilder("");
	String typeOfUser = "";
	try{
		viewList = reportDAO.createUserPaymentReport(userType);
		if(viewList.size() != 0){
			for(int index=0; index < viewList.size(); index++) {
				Object[] innerObj = (Object[])viewList.get(index);
				
				sBuilder.append((String)innerObj[0]+"~");		//	user name
				sBuilder.append((String)innerObj[1]+"~");		//	user profile name
				sBuilder.append((String)innerObj[2]+"~");		//	user group name
				sBuilder.append((String)innerObj[3]+"~");		//	cust id
				
				if((Integer)innerObj[4] == 1){	//	individual
					typeOfUser = this.messageSource.getMessage("label.user.type.individual",null, null);
				} else if ((Integer)innerObj[4] == 3) { 	//	facilitator
					typeOfUser = this.messageSource.getMessage("label.user.type.facilitator",null, null);	
				}
				sBuilder.append(typeOfUser+"~");				//	user type				
				sBuilder.append((String)innerObj[5]+"~");		//	created by
				sBuilder.append((Double)innerObj[6]+"~");		//	Amount
				sBuilder.append((String)innerObj[7]+"~");		//	payment type/ mode
				sBuilder.append((java.sql.Timestamp)innerObj[8]+"~");//	payment date
				
				if(((String)innerObj[7]).equals("CHECK")) {					
					sBuilder.append((String)innerObj[9]+"~");		//	chech no
				} else {					
					sBuilder.append("N/A"+"~");		//	if not check payment				
				}
				sBuilder.append((String)innerObj[10]+"~");		//	trans Id			
				//	no of users
				if( (Integer)innerObj[4] == 3 ) {	//	if facilitator then fetch no of total user
					int faciTotalLimit = reportDAO.fetchFacilitatorTotalLimit((String)innerObj[3],(Integer)innerObj[12]);// total user no by customer id, payment_hdr_id
					sBuilder.append(faciTotalLimit+"~");		//	no of user
					
					String userSubsType = ((String)innerObj[11]).equals("AD") ? "Subscribed" : "Renewed";	//	subscription type
					sBuilder.append(userSubsType);	
				} else {
					sBuilder.append("1"+"~");		//	no of user
				}
				if((Integer)innerObj[4] == 1){ //individual
					sBuilder.append((String)innerObj[12]+","+(String)innerObj[11]+"~");
				}
				else if ((Integer)innerObj[4] == 3) //faci
				{
					sBuilder.append("~"+(String)innerObj[14]+","+(String)innerObj[13]+"~");
				}
				sBuilder.append("#");
			}
		}
	} catch (DAOException e) {
		LOGGER.error(e.getLocalizedMessage());
	}
	LOGGER.info("<<<< ReportServiceImpl.createPaymentReport");
	return sBuilder.toString();
}*/
@Transactional(propagation=Propagation.REQUIRED)
public String createPaymentReport( int userType ) throws JCTException {
 LOGGER.info(">>>> ReportServiceImpl.createPaymentReport");
 List<Object> viewList = null;
 StringBuilder sBuilder = new StringBuilder("");
 String typeOfUser = "";
 try{
  viewList = reportDAO.createUserPaymentReport(userType);
  if(viewList.size() != 0){
   
   LOGGER.info(">>>> ReportServiceImpl.createPaymentReport viewList.size()" + viewList.size());
   for(int index=0; index < viewList.size(); index++) {
    Object[] innerObj = (Object[])viewList.get(index);
    
    sBuilder.append((String)innerObj[0]+"~");  // user name
    sBuilder.append((String)innerObj[1]+"~");  // user profile name
    sBuilder.append((String)innerObj[2]+"~");  // user group name
    sBuilder.append((String)innerObj[3]+"~");  // cust id
    
    if((Integer)innerObj[4] == 1){ // individual
     typeOfUser = this.messageSource.getMessage("label.user.type.individual",null, null);
    } else if ((Integer)innerObj[4] == 3) {  // facilitator
     typeOfUser = this.messageSource.getMessage("label.user.type.facilitator",null, null); 
    }
    sBuilder.append(typeOfUser+"~");    // user type    
    sBuilder.append((String)innerObj[5]+"~");  // created by
    sBuilder.append((Double)innerObj[6]+"~");  // Amount
    sBuilder.append((String)innerObj[7]+"~");  // payment type/ mode
    sBuilder.append((java.sql.Timestamp)innerObj[8]+"~");// payment date
    
    if(((String)innerObj[7]).equals("CHECK")) {     
     sBuilder.append((String)innerObj[9]+"~");  // chech no
    } else {     
     sBuilder.append("N/A"+"~");  // if not check payment    
    }
    sBuilder.append((String)innerObj[10]+"~");  // trans Id
    //sBuilder.append((String)innerObj[12]+","+(String)innerObj[11]+"~");//last name + first name
    // no of users
    if( (Integer)innerObj[4] == 3 ) { // if facilitator then fetch no of total user
     int faciTotalLimit = reportDAO.fetchFacilitatorTotalLimit((String)innerObj[3],(Integer)innerObj[14]);// total user no by customer id, payment_hdr_id
     sBuilder.append(faciTotalLimit+"~");  // no of user
     
     String userSubsType = ((String)innerObj[13]).equals("AD") ? "Subscribed" : "Renewed"; // subscription type
     sBuilder.append(userSubsType); 
    } else {
     sBuilder.append("1"+"~");  // no of user
    }
    
    sBuilder.append((String)innerObj[12]+","+(String)innerObj[11]+"~");//last name + first name
    sBuilder.append("#");
    LOGGER.info(">>>> ReportServiceImpl.createPaymentReport sBuilder" + sBuilder);
   }
  }
 } catch (DAOException e) {
  LOGGER.error(e.getLocalizedMessage());
 }
 LOGGER.info("<<<< ReportServiceImpl.createPaymentReport");
 return sBuilder.toString();
}

	@Transactional(propagation=Propagation.REQUIRED)
	public Integer getSurveyQuestionCount() throws JCTException {
		LOGGER.info(">>>> ReportServiceImpl.getSurveyQuestionCount");
		Integer count = 0;
		try {
			BigInteger tempCount = (java.math.BigInteger) reportDAO.getSurveyQuestionCount();
			if (tempCount != null) {
				count = tempCount.intValue();
			}
		} catch (DAOException e) {
			LOGGER.error(e.getLocalizedMessage());
		}
		LOGGER.info("<<<< ReportServiceImpl.getSurveyQuestionCount");
		return count;
	}


	@Transactional(propagation=Propagation.REQUIRED)
	public List<Object> getSurveyMainQtns(String emailId) throws JCTException {
		LOGGER.info(">>>> ReportServiceImpl.getSurveyMainQtns");
		List<Object> viewList = null;
		try {
			viewList = reportDAO.getSurveyMainQtns(emailId);
		} catch (DAOException e) {
			LOGGER.error(e.getLocalizedMessage());
			viewList = new ArrayList<Object>();
		}
		LOGGER.info("<<<< ReportServiceImpl.getSurveyMainQtns");
		return viewList;
	}
	
	@Transactional(propagation=Propagation.REQUIRED)
	public List<String> getSurveyAnswers(String emailId, String mainQtn, Integer ansType) throws JCTException {
		LOGGER.info(">>>> ReportServiceImpl.getSurveyAnswers");
		List<String> ansList = null;
		try {
			ansList = reportDAO.getSurveyAnswers(emailId, mainQtn, ansType);
		} catch (DAOException e) {
			LOGGER.error(e.getLocalizedMessage());
			ansList = new ArrayList<String>();
		}
		LOGGER.info("<<<< ReportServiceImpl.getSurveyAnswers");
		return ansList;
	}
	
	
	
	@Transactional(propagation=Propagation.REQUIRED)
	public int getTotalCount(String occupationCode) throws JCTException {
		LOGGER.info(">>>> ReportServiceImpl.getTotalCount");
		List<String> list = null;
		List<String> emailIdList = new ArrayList<String>();
		List<Object> totalViewList = new ArrayList<Object>();
		Set<String> emailSet = new HashSet<String>();
		try {
		list = reportDAO.getTotalCount(null, occupationCode, 4, 0);
		for (int index = 0; index < list.size(); index++) {
			String email = (String)list.get(index);
			if (!emailIdList.contains(email)) {
				emailIdList.add(email);
				emailSet.add(email);
			}
		}
		totalViewList.addAll(list);
		list.clear();
		if (emailIdList.size() == 0) {
			emailIdList.clear();
			emailIdList.add("");
		}
		
		list = reportDAO.getTotalCount(emailIdList, occupationCode, 5, 0);
		for (int index = 0; index < list.size(); index++) {
			String email = (String)list.get(index);
			if (!emailIdList.contains(email)) {
				emailIdList.add(email);
				emailSet.add(email);
			}
		}
		totalViewList.addAll(list);
		list.clear();
		if (emailIdList.size() == 0) {
			emailIdList.clear();
			emailIdList.add("");
		}
		
		list = reportDAO.getTotalCount(emailIdList, occupationCode, 5, 1);
		for (int index = 0; index < list.size(); index++) {
			String email = (String)list.get(index);
			if (!emailIdList.contains(email)) {
				emailIdList.add(email);
				emailSet.add(email);
			}
		}
		totalViewList.addAll(list);
		list.clear();
		if (emailIdList.size() == 0) {
			emailIdList.clear();
			emailIdList.add("");
		}
		} catch(DAOException dao){
			LOGGER.error(dao.getLocalizedMessage());
			throw new JCTException();
		}
		LOGGER.info("<<<< ReportServiceImpl.getTotalCount");
		return emailSet.size();
	}
	
	
	
	@Transactional(propagation=Propagation.REQUIRED)
	public List<Object> getAfterSketchList(String occupationCode, int recordIndex) throws JCTException {
		LOGGER.info(">>>> ReportServiceImpl.getAfterSketchList");
		List<Object> viewList = null;
		List<String> emailIdList = new ArrayList<String>();
		List<Object> totalViewList = new ArrayList<Object>();
		try {
			viewList = reportDAO.getAfterSketchList(null, occupationCode, recordIndex, 4, 0);
			for (int index = 0; index < viewList.size(); index++) {
				Object[] innerObj = (Object[])viewList.get(index);
				if (!emailIdList.contains((String)innerObj[0])) {
					emailIdList.add((String)innerObj[0]);
				}
			}
			totalViewList.addAll(viewList);
			viewList.clear();
			
			if (emailIdList.size() == 0) {
				emailIdList.clear();
				emailIdList.add("");
			}
			
			viewList = reportDAO.getAfterSketchList(emailIdList, occupationCode, recordIndex, 5, 0);
			for (int index = 0; index < viewList.size(); index++) {
				Object[] innerObj = (Object[])viewList.get(index);
				if (!emailIdList.contains((String)innerObj[0])) {
					emailIdList.add((String)innerObj[0]);
				}
			}
			totalViewList.addAll(viewList);
			viewList.clear();
		 } catch (DAOException e) {
			LOGGER.error(e.getLocalizedMessage());
		}
		LOGGER.info("<<<< ReportServiceImpl.getAfterSketchList");
		return totalViewList;
	}
	
	
	
	@Transactional(propagation=Propagation.REQUIRED)
	public int getASTotalCount(String occupationCode) throws JCTException {
		LOGGER.info(">>>> ReportServiceImpl.getASTotalCount");
		int count = 0;
		LOGGER.info("<<<< ReportServiceImpl.getASTotalCount");
		
		List<String> viewList = null;
		List<String> emailIdList = new ArrayList<String>();
		List<String> totalViewList = new ArrayList<String>();
		Set<String> emailSet = new HashSet<String>();
		try {
			viewList = reportDAO.getASTotalCount(null, occupationCode, 4, 0);
			for (int index = 0; index < viewList.size(); index++) {
				String email = (String)viewList.get(index);
				if (!emailIdList.contains(email)) {
					emailIdList.add(email);
					emailSet.add(email);
				}
			}
			totalViewList.addAll(emailIdList);
			if (emailIdList.size() == 0) {
				emailIdList.clear();
				emailIdList.add("");
			}
			viewList.clear();
			viewList = reportDAO.getASTotalCount(emailIdList, occupationCode, 5, 0);
			for (int index = 0; index < viewList.size(); index++) {
				String email = (String)viewList.get(index);
				if (!emailIdList.contains(email)) {
					emailIdList.add(email);
					emailSet.add(email);
				}
			}
			totalViewList.addAll(emailIdList);
			if (emailIdList.size() == 0) {
				emailIdList.clear();
				emailIdList.add("");
			}
			viewList.clear();
			viewList = reportDAO.getASTotalCount(null, occupationCode, 5, 1);
			for (int index = 0; index < viewList.size(); index++) {
				String email = (String)viewList.get(index);
				if (!emailIdList.contains(email)) {
					emailIdList.add(email);
					emailSet.add(email);
				}
			}
			totalViewList.addAll(emailIdList);
			//count = totalViewList.size();
			count = emailSet.size();
		} catch (DAOException ex) {
			
		}
		return count;
	}


	@Transactional(propagation=Propagation.REQUIRED)
	public String getUserGroupNameById(int groupId) throws JCTException {
		LOGGER.info(">>>> ReportServiceImpl.getUserGroupNameById");
		String userGroupName = "";
		try {
			userGroupName = reportDAO.getJctUserGroupNameById(groupId);
		} catch (DAOException e) {
			LOGGER.error(e.getLocalizedMessage());
		}
		LOGGER.info("<<<< ReportServiceImpl.getUserGroupNameById");
		return userGroupName;
	}
}