package com.vmware.jct.dao;

import java.math.BigInteger;
import java.util.List;

import com.vmware.jct.dao.dto.SurveyQuestionDTO;
import com.vmware.jct.exception.DAOException;
/**
 * 
 * <p><b>Class name:</b>IReportDAO.java</p>
 * <p><b>@author:</b> InterraIT </p>
 * <p><b>Purpose:</b> 		This DAO interface represents a IReportDAO interface
 * <p><b>Description:</b> 	This will have all methods mapped to the database and fetching data from database to be used for report service. </p>
 * <p><b>Note:</b> 			This methods implemented by DAOImpl classes and execute the mapped queries and get the records.</p>  
 * <p><b>Copyrights:</b> 	All rights reserved by InterraIT and should only be used for its internal application development.</p>
 *  <pre>
 * <b>Revision History:</b>
 * <li> </li>
 * </pre>
 */
public interface IReportDAO {
	/**
	 * Method returns list of before sketch object
	 * @param function
	 * @param jobLevel
	 * @param recordIndex
	 * @return List of Object
	 * @throws DAOException
	 */
	public List<Object> getBeforeSketchList(String function, String jobLevel, int recordIndex, int status, int softDelete) throws DAOException;
	/**
	 * Method returns list of before sketch object
	 * @param function
	 * @param jobLevel
	 * @param recordIndex
	 * @return List of Object
	 * @throws DAOException
	 */
	public List<Object> getBeforeSketchListMax(String function, String jobLevel, int recordIndex) throws DAOException;
	/**
	 * Method returns total count
	 * @param function
	 * @param jobLevel
	 * @param statusId
	 * @param softDel
	 * @return List of String
	 * @throws DAOException
	 */
	public List<String> getTotalCount(List<String> emailIdList, String function, String jobLevel, int status, int softDel) throws DAOException;
	/**
	 * Method returns total count
	 * @param function
	 * @param jobLevel
	 * @return List of String
	 * @throws DAOException
	 */
	public List<String> getTotalCountMax(String function, String jobLevel) throws DAOException;
	/**
	 * Method returns before sketch list specifically for excel
	 * @param function
	 * @param jobLevel
	 * @return List of Object
	 * @throws DAOException
	 */
	public List<Object> getBeforeSketchListForExcel(String occupation, int status, int softDel) throws DAOException;
	/**
	 * Method returns before sketch list specifically for excel
	 * @param function
	 * @param jobLevel
	 * @return List of Object
	 * @throws DAOException
	 */
	public List<Object> getBeforeSketchListForExcelMax(String function, String jobLevel) throws DAOException;
	/**
	 * Method returns group profile list specifically for Excel
	 * @param profile
	 * @return List of Object
	 * @throws DAOException
	 */
	public List<Object> getGroupProfileListForExcel(int profile) throws DAOException;
	/**
	 * Method returns profile description list specifically for excel
	 * @param profile
	 * @return List of Object
	 * @throws DAOException
	 */
	public List<Object> getProfileDescListForExcel(int profile) throws DAOException;
	/**
	 * Method fetches action plan list
	 * @param emailIdList
	 * @param function
	 * @param jobLevel
	 * @param recordIndex
	 * @return List of Object
	 * @throws DAOException
	 */
	public List<Object> getActionPlanList(List<String> emailIdList, String occupation, int recordIndex, int statusId, int softDel) throws DAOException;
	/**
	 * Method fetches action plan list
	 * @param function
	 * @param jobLevel
	 * @param recordIndex
	 * @return List of Object
	 * @throws DAOException
	 */
	public List<Object> getActionPlanListMax(String function, String jobLevel, int recordIndex) throws DAOException;
	/**
	 * Method fetches total count of action plan
	 * @param function
	 * @param jobLevel
	 * @return List of String
	 * @throws DAOException
	 */
	public List<String> getTotalCountActionPlan(List<String> emailIdList, String occupation, int statusId, int softDel) throws DAOException;
	/**
	 * Method fetches total count of action plan
	 * @param function
	 * @param jobLevel
	 * @return List of String
	 * @throws DAOException
	 */
	public List<String> getTotalCountActionPlanMax(String function, String jobLevel) throws DAOException;
	/**
	 * Method fetches action plan list specifically for excel
	 * @param function
	 * @param emailIdList
	 * @param jobLevel
	 * @return
	 * @throws DAOException
	 */
	public List<Object> getActionPlanListForExcel(List<String> emailIdList,String occupation, int statusId, int softDel) throws DAOException;
	/**
	 * Method fetches action plan list specifically for excel
	 * @param function
	 * @param jobLevel
	 * @return
	 * @throws DAOException
	 */
	public List<Object> getActionPlanListForExcelMax(String function, String jobLevel) throws DAOException;
	/**
	 * Method fetches after sketch list
	 * @param function
	 * @param jobLevel
	 * @param recordIndex
	 * @return List of Object
	 * @throws DAOException
	 */
	public List<Object> getAfterSketchList(List<String> emailIdList, String function, String jobLevel, int recordIndex, int statusId, int softDel) throws DAOException;
	/**
	 * Method fetches after sketch list
	 * @param function
	 * @param jobLevel
	 * @param recordIndex
	 * @return List of Object
	 * @throws DAOException
	 */
	public List<Object> getAfterSketchListMax(String function, String jobLevel, int recordIndex) throws DAOException;
	/**
	 * Method fetches after sketch total count list
	 * @param function
	 * @param jobLevel
	 * @return List of String
	 * @throws DAOException
	 */
	public List<String> getASTotalCount(List<String> emailId, String function, String jobLevel, int statusId, int softDel) throws DAOException;
	/**
	 * Method fetches after sketch total count list
	 * @param function
	 * @param jobLevel
	 * @return List of String
	 * @throws DAOException
	 */
	public List<String> getASTotalCountMax(String function, String jobLevel) throws DAOException;
	/**
	 * Method fetches after sketch list specifically for excel
	 * @param function
	 * @param jobLevel
	 * @return List of Object
	 * @throws DAOException
	 */
	public List<Object> getAfterSketchListForExcel(List<String> emailIdList, String ocuupation, int statusId, int softDel) throws DAOException;
	/**
	 * Method fetches after sketch list specifically for excel
	 * @param function
	 * @param jobLevel
	 * @return List of Object
	 * @throws DAOException
	 */
	public List<Object> getAfterSketchListForExcelMax(String function, String jobLevel) throws DAOException;
	/**
	 * Method fetches Before sketch to After sketch list
	 * @param function
	 * @param jobLevel
	 * @param recordIndex
	 * @return List of Object
	 * @throws DAOException
	 */
	public List<Object> getBsToAsList(String occupation, int recordIndex) throws DAOException;
	/**
	 * Method feteches list of Before sketch to After sketch count
	 * @param function
	 * @param jobLevel
	 * @return List of String
	 * @throws DAOException
	 */
	public List<String> getTotalCountBsToAs(String function, String jobLevel) throws DAOException;
	/**
	 * Method fetches Before sketch to After sketch list specifically for excel
	 * @param function
	 * @param jobLevel
	 * @return List of Object
	 * @throws DAOException
	 */
	public List<Object> getBsToAsListForExcel(String occupation) throws DAOException;
	/**
	 * Method fetches list of items required to plot misc report
	 * @return List of Object
	 * @throws DAOException
	 */
	public List<Object> createMiscHeaderReport() throws DAOException;
	/**
	 * Method returns list of items required to plot detailed misc report 
	 * based on job reference number
	 * @param jobRefNo
	 * @return List of Object
	 * @throws DAOException
	 */
	public List<Object> createMiscDetailedReport(String jobRefNo) throws DAOException;
	/**
	 * Method returns count of number diagrams made searchable for the job 
	 * reference number
	 * @param jrno
	 * @return count
	 * @throws DAOException
	 */
	public java.math.BigInteger getSearchableCount(String jrno) throws DAOException;
	/**
	 * Method returns total number of tasks for supplied email id
	 * @param emailId
	 * @return count
	 * @throws DAOException
	 */
	public java.math.BigInteger getTotalTaskCount(String emailId, int statusId, int softDel) throws DAOException;
	/**
	 * Method returns total number of mapping for supplied email id
	 * @param emailId
	 * @return count
	 * @throws DAOException
	 */
	public java.math.BigInteger getTotalMappingCount(String emailId) throws DAOException;
	/**
	 * Method returns count of role for supplied email id
	 * @param emailId
	 * @return count
	 * @throws DAOException
	 */
	public java.math.BigInteger getTotalRoleCount(String emailId, int statusId, int softDel) throws DAOException;
	/**
	 * Method returns list of detailed data of mapping attributes for the supplied 
	 * email id and type
	 * @param emailId
	 * @param elementCode (STR / VAL / PAS)
	 * @return List of Object
	 * @throws DAOException
	 */
	public List<Object> getASDetailedViewForMappings(String emailId, String elementCode, int statusId, int softDelete) throws DAOException;
	/**
	 * Method returns list of Login info details for email id provided
	 * @param emailId
	 * @return List of Object
	 * @throws DAOException
	 */
	public List<Object> getLoginInfoDetails(String emailId) throws DAOException;
	/**
	 * Method return group creation date time
	 * @param groupName
	 * @return sql.timestamp
	 * @throws DAOException
	 */
	public java.sql.Timestamp getGroupCreationDate(Integer groupName) throws DAOException;
	/**
	 * Method returns list of Logged in details for the supplied 
	 * email id
	 * @param email
	 * @return List of Object
	 * @throws DAOException
	 */
	public List<Object> getLoggedIndetails(String email) throws DAOException;
	/**
	 * Method returns total time spent on before sketch page
	 * @param emailId
	 * @return time spent - Integer
	 * @throws DAOException
	 */
	public Integer getTotalTimeSpentOnBs(String emailId, int statusId, int softDel) throws DAOException;
	/**
	 * Method returns number of total tasks in Before sketch
	 * @param emailId
	 * @return total bs tasks - BigInteger
	 * @throws DAOException
	 */
	public BigInteger getTotalBsTasks(String emailId) throws DAOException;
	/**
	 * Method fetches list of before sketch description and time energy allocation 
	 * for the provided email id
	 * @param emailId
	 * @return list of Object
	 * @throws DAOException
	 */
	public List<Object> getBSDescriptionAndTimeEnergy(String emailId) throws DAOException;
	/**
	 * Method fetches total time spent on before diagram for the provided 
	 * email id
	 * @param emailId
	 * @return total time - Integer
	 * @throws DAOException
	 */
	public Integer getTotalTimeSpentOnAs(String emailId, int statusId, int softDel) throws DAOException;
	/**
	 * Method returns list of after sketch tasks for the provided email id
	 * @param emailId
	 * @return List of Object
	 * @throws DAOException
	 */
	public List<Object> getTotalAsTasks(String emailId) throws DAOException;
	/**
	 * Method returns list of role frame for the provided email id
	 * @param emailId
	 * @return List of Object
	 * @throws DAOException
	 */
	public List<Object> getRoleFrame (String emailId) throws DAOException;
	/**
	 * Method fetches total time spent on after diagram for the provided 
	 * email id
	 * @param emailId
	 * @return List of Object
	 * @throws DAOException
	 */
	public List<Object> getASDescriptionMainPersonAndTimeEnergy(String emailId) throws DAOException;
	/**
	 * Method returns list of Before sketch to After sketch based on email id
	 * @param emailId
	 * @return List of Object
	 * @throws DAOException
	 */
	public List<Object> getBsToAs(String emailId) throws DAOException;
	/**
	 * Method returns list of mapping attributes based on email id
	 * @param emailId
	 * @return List of Object
	 * @throws DAOException
	 */
	public List<Object> getStrValPassItems(String emailId) throws DAOException;
	/**
	 * Method returns reflection question list based on email id
	 * @param emailId
	 * @return List of Object
	 * @throws DAOException
	 */
	public List<Object> getReflectionQuestions(String emailId, int statusId, int softDel, String diff) throws DAOException;
	/**
	 * Method returns list of action plan based on email id
	 * @param emailId
	 * @return List of Object
	 * @throws DAOException
	 */
	public List<Object> populateActionPlan(String emailId, int statusId, int softDel) throws DAOException;
	/**
	 * Method lists sub questions based on main question and email id
	 * @param mainQuestion
	 * @param emailId
	 * @return List of Object
	 * @throws DAOException
	 */
	public List<Object> populateSubQtnActionPlan(String mainQuestion, String emailId, String diff, int statusId, int softDel) throws DAOException;
	/**
	 * Method returns lists of answers based on main question, sub question and 
	 * email id
	 * @param mainQuestion
	 * @param subQuestion
	 * @param emailId
	 * @return List of Object
	 * @throws DAOException
	 */
	public List<Object> getActionPlanAnswers(String mainQuestion, String subQuestion, String emailId, String diff, int statusId, int softDel) throws DAOException;
	/**
	 * Method fetches after sketch task location screen coordinates
	 * @param emailId
	 * @return List of Object
	 * @throws DAOException
	 */
	public List<Object> populateASTaskLocation(String emailId) throws DAOException;
	/**
	 * Method lists sub questions based on main question and email id
	 * @param mainQuestion
	 * @param emailId
	 * @return List of Object
	 * @throws DAOException
	 */
	public List<Object> populateSubQtnQuestionnaire(String mainQuestion, String emailId, String diff, int statusId, int softDel) throws DAOException;
	/**
	 * Method returns lists of answers based on main question, sub question and 
	 * email id
	 * @param mainQuestion
	 * @param subQuestion
	 * @param emailId
	 * @return List of Object
	 * @throws DAOException
	 */
	public List<Object> getQuestionnaireAnswers(String mainQuestion, String subQuestion, String emailId, String diff, int statusId, int softDel) throws DAOException;
	/**
	 * Method returns collection of roles in which the task falls
	 * @param emailId
	 * @param taskDesc
	 * @return
	 * @throws DAOException
	 */
	public List<String> getRolesForTasks(String emailId, String taskDesc, int statusId, int softDel) throws DAOException;

	/**
	 * Method returns list of items required to plot detailed survey report 
	 * based on job reference number
	 * @param jobRefNo
	 * @return List of Object
	 * @throws DAOException
	 */
	public List<Object> createSurveyReport(int usrType) throws DAOException;
	
	/**
	 * Method fetches after sketch list specifically for excel
	 * @param answeredUserId
	 * @return List of Object
	 * @throws DAOException
	 */
	public Object getSurveyListCommonFieldsForExcel(int answeredUserId) throws DAOException;
	/**
	 * Method fetches after sketch list specifically for excel
	 * @param userType
	 * @return List of Object
	 * @throws DAOException
	 */
	public List<Object> createUserPaymentReport(int userType) throws DAOException;
	/**
	 * Method fetches fecilitator total limit / user strength
	 * @param custId
	 * @param paymentHdrNo
	 * @return count
	 * @throws DAOException
	 */	
	public int fetchFacilitatorTotalLimit(String custId, int paymentHdrId)  throws DAOException;
	
	/**
	 * Method to fetch fetch list of userId from survey table
	 * @param userType
	 * @throws DAOException 
	 * */
	public List<SurveyQuestionDTO> fetchUserIdFromSurvey(int userType) throws DAOException;	

	/**
	 * Method to fetch fetch list of userId from survey table
	 * @param userType
	 * @throws DAOException 
	 * */
	public Object getSurveyListRedundentFields(int uId) throws DAOException;	
	/**
	 * 
	 * @return
	 * @throws DAOException
	 */
	public java.math.BigInteger getSurveyQuestionCount() throws DAOException;
	/**
	 * 
	 * @param emailId
	 * @return
	 * @throws DAOException
	 */
	public List<Object> getSurveyMainQtns(String emailId) throws DAOException;
	/**
	 * 
	 * @param emailId
	 * @return
	 * @throws DAOException
	 */
	public List<String> getSurveyAnswers(String emailId, String mainQtn, Integer ansType) throws DAOException;
	
	public List<Object> getBeforeSketchList(String occupationCode, int recordIndex, int status, int softDelete) throws DAOException;
	
	public List<String> getTotalCount(List<String> emailIdList, String occupationCode, int status, int softDel) throws DAOException;
	
	public List<Object> getAfterSketchList(List<String> emailIdList, String occupationCode, int recordIndex, int statusId, int softDel) throws DAOException;
	
	public List<String> getASTotalCount(List<String> emailId, String occupationCode, int statusId, int softDel) throws DAOException;
	
	public String getJctUserGroupNameById(int groupId) throws DAOException;
}
