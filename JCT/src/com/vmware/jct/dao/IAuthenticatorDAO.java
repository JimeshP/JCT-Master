package com.vmware.jct.dao;

import java.util.List;

import com.vmware.jct.dao.dto.DemographicDTO;
import com.vmware.jct.dao.dto.FunctionDTO;
import com.vmware.jct.dao.dto.LevelDTO;
import com.vmware.jct.dao.dto.OccupationListDTO;
import com.vmware.jct.dao.dto.RegionDTO;
import com.vmware.jct.exception.DAOException;
import com.vmware.jct.exception.DAONullException;
import com.vmware.jct.exception.JCTException;
import com.vmware.jct.model.JctActionPlan;
import com.vmware.jct.model.JctAfterSketchHeader;
import com.vmware.jct.model.JctAfterSketchHeaderTemp;
import com.vmware.jct.model.JctBeforeSketchHeader;
import com.vmware.jct.model.JctBeforeSketchHeaderTemp;
import com.vmware.jct.model.JctBeforeSketchQuestion;
import com.vmware.jct.model.JctInstructionBar;
import com.vmware.jct.model.JctPopupInstruction;
import com.vmware.jct.model.JctStatusSearch;
import com.vmware.jct.model.JctSurveyQuestions;
import com.vmware.jct.model.JctUser;
import com.vmware.jct.model.JctUserInfo;
import com.vmware.jct.model.JctUserLoginInfo;
import com.vmware.jct.model.JctUserProfile;
import com.vmware.jct.service.vo.MyAccountVO;
import com.vmware.jct.service.vo.UserVO;
//import com.vmware.jct.exception.JCTException;

/**
 * 
 * <p><b>Class name:</b>IAuthenticatorDAO.java</p>
 * <p><b>@author:</b> InterraIt </p>
 * <p><b>Purpose:</b> 		This DAO interface represents a IAuthenticatorDAO interface
 * <p><b>Description:</b> 	This will have all methods mapped to the database and fetching data from database to be used across the application </p>
 * <p><b>Note:</b> 			This methods implemented by DAOImpl classes and execute the mapped queries and get the records.</p>  
 * <p><b>Copyrights:</b> 	All rights reserved by Interra IT and should only be used for its internal application development.</p>
 *  <pre>
 * <b>Revision History:</b>
 * InterraIt - 19/Jan/2014 - Implement exception through out the appliaction
 * </pre>
 */

public interface IAuthenticatorDAO {
	/**
	 * @param It will take User VO object as a input parameter
	 * @return It returns list of all user data from database
	 * @throws DAOException -it throws DAOException
	 */
	public List<JctUser> authenticateUser(UserVO userVO) throws DAOException;
	/**
	 * @param It will take User VO object as a input parameter
	 * @return It returns list of all user data from database
	 * @throws DAOException -it throws DAOException
	 */
	public List<JctUser> checkUser(UserVO userVO) throws DAOException;
	/**
	 * @param It will take User VO object as a input parameter
	 * @return It returns list of all user data from database
	 * @throws DAOException -it throws DAOException
	 */
	public List<JctUser> checkFacilitatorUser(UserVO userVO) throws DAOException;
	/**
	 * @param It will take User VO object as a input parameter
	 * @return It returns list of  user data from database
	 * @throws DAOException -it throws DAOException
	 */
	public List<JctUser> checkUsers(UserVO userVO) throws DAOException;
	/**
	 * @param It will take User  object as a input parameter
	 * @return It will returns list of all user data from database
	 * @throws DAOException -it throws DAOException
	 */
	public String register(JctUser user) throws DAONullException, DAOException;
	/**
	 * @param It will take User  object as a input parameter
	 * @return It will returns string object
	 * @throws DAOException -it throws DAOException
	 */
	public String updateUser(JctUser user) throws DAOException;
	/**
	 * @param It will take User  object as a input parameter
	 * @return It will returns string object
	 * @throws DAOException -it throws DAOException
	 */
	public String merge(JctUser user) throws DAONullException, DAOException;
	/**
	 * @param It will take User  object as a input parameter
	 * @return It will returns string object
	 * @throws DAOException -it throws DAOException
	 */
	public List<JctUserLoginInfo> getTaskPendingDetails(int userId) throws DAOException;
	/**
	 * @param It will take JctUserLoginInfo object as a input parameter
	 * @return It will returns string object
	 * @throws DAOException -it throws DAOException
	 */
	public String saveLoginInfo(JctUserLoginInfo info) throws DAOException;
	/**
	 * @param It will take Job reference number and row id as a input parameters
	 * @return It will returns string object
	 * @throws DAOException -it throws DAOException
	 */
	public String updateLoginInfo(String jobRefNo, int rowId, String lastActivePage) throws DAOException;
	/**
	 * @param Null
	 * @return It will returns List<RegionDTO> object
	 * @throws DAOException -it throws DAOException
	 */
	public List<RegionDTO> getJctRegionList() throws DAOException;
	/**
	 * @param Null
	 * @return It will returns List<RegionDTO> object
	 * @throws DAOException -it throws DAOException
	 */
	public List<FunctionDTO> getJctFunctionList() throws DAOException;
	
	/**
	 * @param Null
	 * @return It will returns List<LevelDTO> object
	 * @throws DAOException -it throws DAOException
	 */
	public List<LevelDTO> getJctLevelList() throws DAOException;
	/**
	 * 
	 * @param userName
	 * @return
	 * @throws DAOException
	 */
	public List<JctUser> getUserList(String userName, int roleId) throws DAOException;
	/**
	 * 
	 * @param jobReferenceNo
	 * @return
	 * @throws DAOException
	 */
	public int getMaxLoginInfoId(String jobReferenceNo) throws DAOException;
	/**
	 * 
	 * @param id
	 * @param page
	 * @return
	 * @throws DAOException
	 */
	public String updateLoginInfo(int id, String page) throws DAOException;
	/**
	 * 
	 * @param refNo
	 * @return
	 * @throws DAOException
	 */
	public Long getIsCompleted(String refNo) throws DAOException;
	/**
	 * 
	 * @param jobReferenceNo
	 * @param nextPage
	 * @return
	 * @throws DAOException
	 */
	public int updateNextPage(String jobReferenceNo, String nextPage) throws DAOException;
	
	/**
	 * @param It will take user id and user name as a input parameter
	 * @return It will returns JctUserInfo list
	 	 * @throws DAOException -it throws DAOException
	 */
	
	public List<JctUserInfo> fetchUserInfo (int userId, String userName) throws DAOException;
	
	/**
	 * @param It will takeuser name as a input parameter
	 * @return It will returns JctUserInfo list
	 	 * @throws DAOException -it throws DAOException
	 */
	public List<JctUserInfo> fetchUserInfoByEmail (String userName) throws DAOException;
	
	/**
	 * @param It will take user info as a input parameter
	 * @return It will returns string
	 	 * @throws DAOException -it throws DAOException
	 */
	public String saveUserInfo (JctUserInfo info) throws DAOException;
	/**
	 * @param It will take profile id and relatedPage as a input parameter
	 * @return It will returns JctInstructionBar
	 * @throws DAOException -it throws DAOException
	 */
	public JctInstructionBar populateExistingInstruction(int profileId, String relatedPage) throws DAOException;
	/**
	 * @param It will take email as a input parameter
	 * @return It will returns Integer
	 * @throws DAOException -it throws DAOException
	 */
	public Integer fetchProfileId(String emailId) throws DAOException;
	/**
	 * 
	 * @param emailId
	 * @return List<DemographicDTO>
	 */
	public List<DemographicDTO> getDemographicInformation ( String emailId ) throws DAOException;
	/**
	 * 
	 * @param rowNo
	 * @return
	 * @throws DAOException
	 */
	public String getEmailIdByRowNos ( int rowNo ) throws DAOException;
	/**
	 * 
	 * @param userVO
	 * @return
	 * @throws DAOException
	 */
	public List<JctUser> checkInactiveUser(UserVO userVO) throws DAOException;
	/**
	 * 
	 * @param jobReferenceNos
	 * @return
	 * @throws DAOException
	 */
	public List<JctBeforeSketchHeader> getBeforeSketchDummyEntry (String jobReferenceNos) throws DAOException;
	/**
	 * 
	 * @param header
	 * @throws DAOException
	 */
	public int deleteBSData (String jobRefNo) throws DAOException;
	
	/**
	 * 
	 * @param header
	 * @throws DAOException
	 */
	public int deleteBSChildData (String jobRefNo) throws DAOException;
	/**
	 * 
	 * @param jobReferenceNos
	 * @return
	 * @throws DAOException
	 */
	public List<JctBeforeSketchHeaderTemp> getBeforeSketchTempDummyEntry (String jobReferenceNos) throws DAOException;
	/**
	 * 
	 * @param header
	 * @throws DAOException
	 */
	public int deleteBSTempData (String jobRefNos) throws DAOException;
	/**
	 * 
	 * @param header
	 * @throws DAOException
	 */
	public int deleteBSTempChildData (String jobRefNos) throws DAOException;
	/**
	 * 
	 * @param jobReferenceNos
	 * @return
	 * @throws DAOException
	 */
	public List<JctBeforeSketchQuestion> getQuestionnaireDummyEntry (String jobReferenceNos) throws DAOException;
	/**
	 * 
	 * @param header
	 * @throws DAOException
	 */
	public int deleteQuestionnaireTempData (String jobRefNo) throws DAOException;
	/**
	 * 
	 * @param jobReferenceNos
	 * @return
	 * @throws DAOException
	 */
	public List<JctAfterSketchHeader> getAfterSketchDummyEntry (String jobReferenceNos) throws DAOException;
	/**
	 * 
	 * @param header
	 * @throws DAOException
	 */
	public int deleteASData (String jobRefNo) throws DAOException;
	/**
	 * 
	 * @param jobRefNos
	 * @return
	 * @throws DAOException
	 */
	public int deleteASPageOneData (String jobRefNos) throws DAOException;
	/**
	 * 
	 * @param jobRefNos
	 * @return
	 * @throws DAOException
	 */
	public int deleteASPageTwoData (String jobRefNos) throws DAOException;
	/**
	 * 
	 * @param jobRefNos
	 * @return
	 * @throws DAOException
	 */
	public int deleteASPageOneTempData (String jobRefNos) throws DAOException;
	/**
	 * 
	 * @param jobRefNos
	 * @return
	 * @throws DAOException
	 */
	public int deleteASPageTwoTempData (String jobRefNos) throws DAOException;
	/**
	 * 
	 * @param jobReferenceNos
	 * @return
	 * @throws DAOException
	 */
	public List<JctAfterSketchHeaderTemp> getAfterSketchDummyTempEntry (String jobReferenceNos) throws DAOException;
	/**
	 * 
	 * @param header
	 * @throws DAOException
	 */
	public int deleteASTempData (String jobRefNo) throws DAOException;
	/**
	 * 
	 * @param jobReferenceNos
	 * @return
	 * @throws DAOException
	 */
	public List<JctActionPlan> getActionDummyEntry (String jobReferenceNos) throws DAOException;
	/**
	 * 
	 * @param plan
	 * @throws DAOException
	 */
	public int deleteActionPlanTempData (String jobRefNo) throws DAOException;
	/**
	 * 
	 * @param jobReferenceNos
	 * @return
	 * @throws DAOException
	 */
	public List<JctStatusSearch> getStatusSearchDummyEntry(String jobReferenceNos) throws DAOException;
	/**
	 * 
	 * @param search
	 * @throws DAOException
	 */
	public int deleteStatusSearchDummyData(String jobRefNo) throws DAOException;
	/**
	 * 
	 * @param search
	 * @throws DAOException
	 */
	public int deleteCompletionDummyData(String jobRefNo) throws DAOException;
	/**
	 * 
	 * @param jobRefNo
	 * @return
	 * @throws DAOException
	 */
	public int deleteDummyLoginInfoData(String jobRefNo) throws DAOException;
	/**
	 * 
	 * @param userVO
	 * @return
	 * @throws DAOException
	 */
	public List<JctUser> checkActiveUser(UserVO userVO) throws DAOException;
	/**
	 * 
	 * @param userId
	 * @return
	 * @throws DAOException
	 */
	public String getEmailIdByUserId(int userId) throws DAOException;
	/**
	 * 
	 * @return
	 * @throws DAOException
	 */
	//public List<String> getTextSurveyQuestion() throws DAOException;
	public List<Object[]> getTextSurveyQuestion() throws DAOException;
	/**
	 * 
	 * @return
	 * @throws DAOException
	 */
	//public List<String> getSurveyMainQuestion(int ansType) throws DAOException;
	public List<Object[]> getSurveyMainQuestion(int ansType) throws DAOException;
	/**
	 * 
	 * @param ansType
	 * @param mainQtn
	 * @return
	 * @throws DAOException
	 */
	public List<String> getAllSubQtns(int ansType, String mainQtn) throws DAOException;
	/**
	 * 
	 * @param userId
	 * @return
	 * @throws DAOException
	 */
	public int getDaysToExpire(int userId) throws DAOException;
	/**
	 * 
	 * @param searchString
	 * @return
	 * @throws DAOException
	 */
	public List<OccupationListDTO> searchOccupationList (String searchString) throws DAOException;
	/**
	 * 
	 * @param searchString
	 * @return
	 * @throws DAOException
	 */
	public List<String> getOnetDesc (String searchString) throws DAOException;
	/**
	 * 
	 * @param row
	 * @return
	 * @throws DAOException
	 */
	public int getUserId (int row) throws DAOException;
	
	/*
	 * Populate pop up
	 */

	public JctPopupInstruction populatePopUp(int profileId,String pageSequence) throws DAOException;
	
	/**
	 * 
	 * @param searchString
	 * @return
	 * @throws DAOException
	 */
	public String getOnetCodeByDesc (String searchString) throws DAOException;
	
	
	/**
	 * 03.03.15
	 */
	public String checkInstrList(int profileId,String relatedPage) throws DAOException;
	
	/*Jira Id : JCTP-35*/
	public String getMailedPassword(UserVO userVO) throws DAOException;
	
	/* Bugzilla JCTP-9045 */
	public String getInitialPassword(UserVO userVO) throws DAOException;
	
	/**
	 * Method to fetch existing term & condition from DB 
	 * @param userProfile
	 * @param userType
	 * @throws DAOException 
	 * **/
	public String getTermsAndConditions(JctUserProfile userProfile, int userType) throws JCTException, DAOException;
	/**
	 * Method to get user profiel by PK
	 * @param userProfileId
	 * **/
	public JctUserProfile getUserProfileByPk(int userProfileId) throws JCTException;
	/**
	 * 
	 * @param custId
	 * @return
	 * @throws DAOException
	 */
	public List<MyAccountVO> getFacilitatorDetails (String custId) throws DAOException;
	/**
	 * 
	 * @param searchString
	 * @return
	 * @throws DAOException
	 */
	public List searchOccupationListNew(String searchString) throws DAOException;
	/**
	 * 
	 * @param val
	 * @return
	 */
	public List<OccupationListDTO> getOccupationListDto(String val) throws DAOException;
	
	public List<Object[]> getAllSurveyQuestion() throws DAOException;
	
	public List<JctSurveyQuestions> alreadyRegisteredOnce(String emailId) throws DAOException;
	
	public String silentUpdateTime(int rowIdentity) throws DAOException;
	
}
