package com.vmware.jct.service;

import java.util.List;

import com.vmware.jct.dao.dto.FunctionDTO;
import com.vmware.jct.dao.dto.LevelDTO;
import com.vmware.jct.dao.dto.OccupationListDTO;
import com.vmware.jct.dao.dto.RegionDTO;
import com.vmware.jct.exception.JCTException;
import com.vmware.jct.model.JctPopupInstruction;
import com.vmware.jct.model.JctUser;
import com.vmware.jct.model.JctUserLoginInfo;
import com.vmware.jct.service.vo.NavigationVO;
import com.vmware.jct.service.vo.PopulateInstructionVO;
import com.vmware.jct.service.vo.ReturnVO;
import com.vmware.jct.service.vo.UserLoginInfoVO;
import com.vmware.jct.service.vo.UserVO;

/**
 * 
 * <p><b>Class name:</b>IAuthenticatorService.java</p>
 * <p><b>@author:</b> InterreaIt </p>
 * <p><b>Purpose:</b> 		This Service interface represents a IAuthenticatorService interface
 * <p><b>Description:</b> 	This will have all methods where we implement all business logic and all dao layer </p>
 * <p><b>Copyrights:</b> 	All rights reserved by Interra IT and should only be used for its internal application development.</p>
 *  <pre>
 * <b>Revision History:</b>
 * InterreaIt - 31/Jan/2014 - Implement exception through out the appliaction
 * </pre>
 */
public interface IAuthenticatorService {
	/**
	 * @param It will take User VO object as a input parameter
	 * @return It returns list of all user data from database
	 * @throws it throws JCTException
	 */
	public List<JctUser> authenticateUser(UserVO userVO) throws JCTException;
	/**
	 * @param It will take User VO object as a input parameter
	 * @return It returns list of all user data from database
	 * @throws it throws JCTException
	 */
	public List<JctUser> checkUser(UserVO userVO) throws JCTException;
	/**
	 * @param It will take User VO object as a input parameter
	 * @return It returns list of all user data from database
	 * @throws it throws JCTException
	 */
	public List<JctUser> checkFacilitatorUser(UserVO userVO) throws JCTException;
	/**
	 * @param It will take User VO object as a input parameter
	 * @return It returns string object
	 * @throws it throws JCTException
	 */
	public String register(UserVO userVO)  throws JCTException;
	/**
	 * @param It will take User VO object as a input parameter
	 * @param dist
	 * @return It returns string object
	 * @throws it throws JCTException
	 */
	public String resetPassword(UserVO userVO, int dist) throws JCTException;
	/**
	 * @param It will take User VO object as a input parameter
	 * @return It returns Return VO object
	 * @throws it throws JCTException
	 */
	public ReturnVO forgotPassword(UserVO userVO) throws JCTException;	
	/**
	 * @param It will take user id as a input parameter
	 * @return It returns JctUserLoginInfo object
	 * @throws it throws JCTException
	 */
	public List<JctUserLoginInfo> getTaskPendingDetails(int userId) throws JCTException;
	/**
	 * @param NULL
	 * @return It returns Int object
	 * @throws it throws JCTException
	 */
	public int getGeneratedJRId() throws JCTException;
	/**
	 * @param It will take UserLoginInfoVO as a input parameter
	 * @return It returns string object
	 * @throws it throws JCTException
	 */
	public String saveLoginInfo(UserLoginInfoVO loginVO) throws JCTException;
	/**
	 * @param It will take jobRefNo,rowId,lastActivePage as a input parameter
	 * @return It returns String object
	 * @throws it throws JCTException
	 */
	public String updateLoginInfo(String jobRefNo, int rowId, String lastActivePage) throws JCTException;
	
	/**
	 * @param Null
	 * @return It returns List<RegionDTO> object
	 * @throws it throws JCTException
	 */
	public List<RegionDTO> jctRegionList() throws JCTException;
	
	
	/**
	 * @param Null
	 * @return It returns List<FunctionDTO> object
	 * @throws it throws JCTException
	 */
	public List<FunctionDTO> jctFunctionList() throws JCTException;
	/**
	 * @param Null
	 * @return It returns List<LevelDTO> object
	 * @throws it throws JCTException
	 */
	public List<LevelDTO> jctLevelList() throws JCTException;
	/**
	 * @param It will take refNo as a input parameter
	 * @return It returns int object
	 * @throws it throws JCTException
	 */
	
	public int getIsCompleted(String refNo) throws JCTException;
	/**
	 * @param It will take jobRefNo,nextPage as a input parameter
	 * @return It returns String object
	 * @throws it throws JCTException
	 */
	public String updateNextPage(String jobRefNo, String nextPage) throws JCTException;
	/**
	 * @param It will take userId,userEmail as a input parameter
	 * @return It returns int object
	 * @throws it throws JCTException
	 */
	public int saveUserValidationFails(int userId, String userEmail) throws JCTException;
	/**
	 * @param It will take userId,userName as a input parameter
	 * @return It returns int object
	 * @throws it throws JCTException
	 */
	public int getTotalLoginFailedCounts (int userId, String userName) throws JCTException;
	/**
	 * @param It will take userId,userName as a input parameter
	 * @return It returns string object
	 * @throws it throws JCTException
	 */
	public String changeInfoStatus (int userId, String userName) throws JCTException;
	/**
	 * @param It will take userName as a input parameter
	 * @return It returns boolean object
	 * @throws it throws JCTException
	 */
	public boolean isForgotPwdAllowed (String userName) throws JCTException;
	/**
	 * @param It will take profileId,relatedPage as a input parameter
	 * @return It returns boolean object
	 * @throws it throws JCTException
	 */
	public NavigationVO populateInstruction(int profileId, String relatedPage) throws JCTException;
	/**
	 * @param It will take emailId as a input parameter
	 * @return It returns Integer object
	 * @throws it throws JCTException
	 */
	public Integer fetchProfileId(String emailId) throws JCTException;
	/**
	 * @param It will take User VO object as a input parameter
	 * @return int
	 * @throws it throws JCTException
	 */
	public int updateDemographicInfo(UserVO userVO) throws JCTException;
	/**
	 * 
	 * @param emailId
	 * @return
	 * @throws JCTException
	 */
	public String getDemographicInformation ( String emailId ) throws JCTException;
	/**
	 * 
	 * @param rowNo
	 * @return
	 * @throws JCTException
	 */
	public String getEmailIdByRowNos ( int rowNo ) throws JCTException;
	/**
	 * 
	 * @param userVO
	 * @return
	 * @throws JCTException
	 */
	public List<JctUser> checkInactiveUser(UserVO userVO) throws JCTException;
	/**
	 * Method deletes the dummy data from different tables
	 * @param jobReferenceNos
	 * @return
	 * @throws JCTException
	 */
	public String deleteDummyData (String jobReferenceNos) throws JCTException;
	/**
	 * Method fetches the survey question
	 * @param UserVO
	 * @return
	 * @throws JCTException
	 */
	public UserVO getSurveyQuestions (UserVO userVO) throws JCTException;
	/**
	 * 
	 * @param userId
	 * @return
	 * @throws JCTException
	 */
	public int getDaysToExpire(int userId) throws JCTException;
	/**'
	 * 
	 * @param searchString
	 * @return
	 * @throws JCTException
	 */
	public List<OccupationListDTO> searchOccupationList( String searchString ) throws JCTException;
	/**
	 * 
	 * @param row
	 * @return
	 * @throws JCTException
	 */
	public int getUserId (int row) throws JCTException;
	
	/*
	 * Populate pop up*/
	
	public PopulateInstructionVO populatePopUp(int profileId,String relatedPage) throws JCTException;
	
	/*Jira JCTP-35*/
	public String getMailedPassword(UserVO userVo) throws JCTException;
	
	/* Bugzilla JCTP-9045 */
	public String getInitialPassword(UserVO userVo) throws JCTException;
	/**'
	 * Method fetch occu[pation data
	 * @param emailId
	 * @return
	 * @throws JCTException
	 */
	public String getOccupationData(String emailId) throws JCTException;
	/**
	 * Method to fetch existing term & condition from DB 
	 * @param userProfile
	 * @param userType
	 * **/
	public String getTermsAndConditions(int userProfile, int userType) throws JCTException;
	
	public boolean alreadyRegisteredOnce(String emailId) throws JCTException;
	
	public String silentUpdateTime (int rowIdentity) throws JCTException;
	
}
