package com.vmware.jct.service;

import java.util.List;

import com.vmware.jct.dao.dto.SurveyQuestionDTO;
import com.vmware.jct.exception.JCTException;
import com.vmware.jct.service.vo.surveyQtnReportVO;
/**
 * 
 * <p><b>Class name:</b>IReportService.java</p>
 * <p><b>@author:</b> InterraIT </p>
 * <p><b>Purpose:</b> 		This Service interface represents a IReportService interface
 * <p><b>Description:</b> 	This will have all methods where we implement all business logic and all service layer </p>
 * <p><b>Copyrights:</b> 	All rights reserved by InterraIT and should only be used for its internal application development.</p>
 *  <pre>
 * <b>Revision History:</b>
 * InterraIT - 31/Jan/2014 - Implement exception through out the application
 * </pre>
 */
public interface IReportService {
	/**
	 * Method fetches the before sketch list
	 * @param fnGrp
	 * @param jLevel
	 * @param recordIndex
	 * @return List of objects
	 * @throws JCTException
	 */
	public List<Object> getBeforeSketchList(String occupationCode, int recordIndex) throws JCTException;
	/**
	 * Method returns the total count
	 * @param function
	 * @param jobLevel
	 * @return count
	 * @throws JCTException
	 */
	public int getTotalCount(String function, String jobLevel) throws JCTException;
	/**
	 * Method returns the before sketch list for specifically generating excel
	 * @param function
	 * @param jobLevel
	 * @return List of objects
	 * @throws JCTException
	 */
	public List<Object> getBeforeSketchListForExcel(String occupation) throws JCTException;
	/**
	 * Method returns the Group and profile list for specifically generating excel
	 * @param profile
	 * @return List of objects
	 * @throws JCTException
	 */
	public List<Object> getGroupProfileListForExcel(int profile) throws JCTException;
	/**
	 * Method returns the user profile list for specifically generating excel
	 * @param profile
	 * @return List of objects
	 * @throws JCTException
	 */
	public List<Object> getProfileDescListForExcel(int profile) throws JCTException;
	/**
	 * Method returns the action plan list
	 * @param fnGrp
	 * @param jLevel
	 * @param recordIndex
	 * @return List of objects
	 * @throws JCTException
	 */
	public List<Object> getActionPlanList(String occupation, int recordIndex) throws JCTException;
	/**
	 * Method returns the total count of action plan
	 * @param function
	 * @param jobLevel
	 * @return count
	 * @throws JCTException
	 */
	public int getTotalCountActionPlan(String occupation) throws JCTException;
	/**
	 * Method returns the action plan list for specifically generating excel
	 * @param function
	 * @param jobLevel
	 * @return List of objects
	 * @throws JCTException
	 */
	public List<Object> getActionPlanListForExcel(String occupation) throws JCTException;
	/**
	 * Method returns the after sketch list
	 * @param fnGrp
	 * @param jLevel
	 * @param recordIndex
	 * @return List of objects
	 * @throws JCTException
	 */
	public List<Object> getAfterSketchList(String fnGrp, String jLevel, int recordIndex) throws JCTException;
	/**
	 * Method returns after sketch total count
	 * @param function
	 * @param jobLevel
	 * @return count
	 * @throws JCTException
	 */
	public int getASTotalCount(String function, String jobLevel) throws JCTException;
	/**
	 * Method returns the after sketch list for specifically generating excel
	 * @param function
	 * @param jobLevel
	 * @return List of objects
	 * @throws JCTException
	 */
	public List<Object> getAfterSketchListForExcel(String occupation) throws JCTException;
	/**
	 * Method returns before sketch to after sketch list
	 * @param fnGrp
	 * @param jLevel
	 * @param recordIndex
	 * @return List of objects
	 * @throws JCTException
	 */
	public List<Object> getBsToAsList(String occupationVal, int recordIndex) throws JCTException;
	/**
	 * Method returns before sketch to after sketch total count
	 * @param function
	 * @param jobLevel
	 * @return count
	 * @throws JCTException
	 */
	public int getTotalCountBsToAs(String function, String jobLevel) throws JCTException;
	/**
	 * Method returns before sketch to after sketch list specifically for excel
	 * @param function
	 * @param jobLevel
	 * @return List of objects
	 * @throws JCTException
	 */
	public List<Object> getBsToAsListForExcel(String occupation) throws JCTException;
	/**
	 * Creates Misc overview report
	 * @return String
	 * @throws JCTException
	 */
	public String createMiscHeaderReport() throws JCTException;
	/**
	 * Method creates Misc detailed report based on the selected 
	 * job reference number
	 * @param jobReferenceNo
	 * @return String
	 * @throws JCTException
	 */
	public String createMiscDetailesReport(String jobReferenceNo) throws JCTException;
	/**
	 * Method gets the total task count
	 * @param emailId
	 * @return count
	 */
	public Integer getTotalTaskCount(String emailId);
	/**
	 * Method gets total attribute mapping count
	 * @param emailId
	 * @return count
	 */
	public Integer getTotalMappingCount(String emailId);
	/**
	 * Method gets total number of role
	 * @param emailId
	 * @return count
	 */
	public Integer getTotalRoleCount(String emailId);
	/**
	 * Method creates string containing the all the job attribute 
	 * which the email id has selected in after sketch diagram
	 * @param emailId
	 * @return string
	 */
	public String getASDetailedViewOfMappings(String emailId);
	/**
	 * Method returns list of users from user login table
	 * @param emailId
	 * @return List of objects
	 * @throws JCTException
	 */
	public List<Object> getLoginInfoDetails(String emailId) throws JCTException;
	/**
	 * Method returns the group creation date
	 * @param groupName
	 * @return java.sql.Timestamp
	 * @throws JCTException
	 */
	public java.sql.Timestamp getGroupCreationDate(Integer groupName) throws JCTException;
	/**
	 * Method returns list of users who have logged in atleast once
	 * @param email
	 * @return List of objects
	 * @throws JCTException
	 */
	public List<Object> getLoggedIndetails(String email) throws JCTException;
	/**
	 * Method returns total time spent of before sketch
	 * @param emailId
	 * @return int
	 * @throws JCTException
	 */
	public Integer getTotalTimeSpentOnBs (String emailId) throws JCTException;
	/**
	 * Method returns total number of before sketch tasks
	 * @param emailId
	 * @return
	 * @throws JCTException
	 */
	public java.math.BigInteger getTotalBsTasks (String emailId) throws JCTException;	
	/**
	 * Method returns Before sketch description, additional user input and corresponding 
	 * time and energy 
	 * @param emailId
	 * @return List of objects
	 * @throws JCTException
	 */
	public List<Object> getBSDescriptionAndTimeEnergy (String emailId) throws JCTException;
	/**
	 * Method returns total time spent of After sketch
	 * @param emailId
	 * @return int
	 * @throws JCTException
	 */
	public Integer getTotalTimeSpentOnAs (String emailId) throws JCTException;
	/**
	 * Method returns Total After sketch tasks
	 * @param emailId
	 * @return List of objects
	 * @throws JCTException
	 */
	public List<Object> getTotalAsTasks (String emailId) throws JCTException;
	/**
	 * Method returns list of role frames
	 * @param emailId
	 * @return List of objects
	 * @throws JCTException
	 */
	public List<Object> getRoleFrame (String emailId) throws JCTException;
	/**
	 * Method returns After sketch description, additional user input and corresponding 
	 * time and energy
	 * @param emailId
	 * @return List of objects
	 * @throws JCTException
	 */
	public List<Object> getASDescriptionMainPersonAndTimeEnergy (String emailId) throws JCTException;
	/**
	 * Method returns list of roles for tasks
	 * @param emailId
	 * @param taskDesc
	 * @return
	 * @throws JCTException
	 */
	public List<String> getRolesForTasks (String emailId, String taskDesc) throws JCTException;
	/**
	 * Method returns list of Before sketch to after sketch
	 * @param emailId
	 * @return List of objects
	 * @throws JCTException
	 */
	public List<Object> getBsToAs (String emailId) throws JCTException;
	/**
	 * Method returns the mapping items
	 * @param emailId
	 * @return List of objects
	 * @throws JCTException
	 */
	public List<Object> getStrValPassItems (String emailId) throws JCTException;
	/**
	 * Method returns list of reflection questions and answers
	 * @param emailId
	 * @return List of objects
	 * @throws JCTException
	 */
	public List<Object> getReflectionQuestions (String emailId) throws JCTException;
	/**
	 * Method populates action plan
	 * @param emailId
	 * @return List of objects
	 * @throws JCTException
	 */
	public List<Object> populateActionPlan (String emailId) throws JCTException;
	/**
	 * Method populates list of sub questions
	 * @param mainQuestion
	 * @param emailId
	 * @return List of objects
	 * @throws JCTException
	 */
	public List<Object> populateSubQtnActionPlan (String mainQuestion, String emailId) throws JCTException;
	/**
	 * Method gets List of action plan answers
	 * @param mainQuestion
	 * @param subQuestion
	 * @param emailId
	 * @return List of objects
	 * @throws JCTException
	 */
	public List<Object> getActionPlanAnswers (String mainQuestion, String subQuestion, String emailId) throws JCTException;
	/**
	 * Method returns the list of after sketch task locations
	 * @param emailId
	 * @return List of objects
	 * @throws JCTException
	 */
	public List<Object> populateASTaskLocation (String emailId) throws JCTException;
	/**
	 * Method populates list of sub questions
	 * @param mainQuestion
	 * @param emailId
	 * @return List of objects
	 * @throws JCTException
	 */
	public List<Object> populateSubQtnQuestionnaire (String mainQuestion, String emailId) throws JCTException;
	/**
	 * Method gets List of questionnaire answers
	 * @param mainQuestion
	 * @param subQuestion
	 * @param emailId
	 * @return List of objects
	 * @throws JCTException
	 */
	public List<Object> getQuestionnaireAnswers (String mainQuestion, String subQuestion, String emailId) throws JCTException;
	
	/**
	 * Method creates Survey report based on the selected 
	 * job reference number
	 * @param jobReferenceNo
	 * @return String
	 * @throws JCTException
	 */
	public String createSurveyReport(int usrType) throws JCTException;
	
	/**
	 * Method returns Survey Report list
	 * @param userType
	 * @param jLevel
	 * @param recordIndex
	 * @return List of objects
	 * @throws JCTException
	 */
	public List<surveyQtnReportVO> getSurveyListForExcel(List<SurveyQuestionDTO> answeredUserId) throws JCTException;
	

	/**
	 * Method returns Payment Report list
	 * @param userType
	 * @param recordIndex
	 * @return List of objects
	 * @throws JCTException
	 */
	public String createPaymentReport(int userType) throws JCTException;
	
	/**
	 * Method to fetch userId(s) from survey table
	 * @param userType
	 * @return List<Object>  
	 */
	
	public List<SurveyQuestionDTO> fetchUserIdFromSurvey(int userType) throws JCTException;
	/**
	 * 
	 * @return
	 * @throws JCTException
	 */
	public Integer getSurveyQuestionCount() throws JCTException;
	/**
	 * 
	 * @param emailId
	 * @return
	 * @throws JCTException
	 */
	public List<Object> getSurveyMainQtns (String emailId) throws JCTException;
	/**
	 * 
	 * @param emailId
	 * @return
	 * @throws JCTException
	 */
	public List<String> getSurveyAnswers (String emailId, String mainQtn, Integer ansType) throws JCTException;
	
	public int getTotalCount(String occupationCode) throws JCTException;
	
	public List<Object> getAfterSketchList(String occupationCode, int recordIndex) throws JCTException;
	
	public int getASTotalCount(String occupationCode) throws JCTException;
	
	public String getUserGroupNameById(int groupId) throws JCTException;
}
