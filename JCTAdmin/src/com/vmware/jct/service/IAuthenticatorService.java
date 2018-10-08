package com.vmware.jct.service;

import java.util.List;

import com.vmware.jct.dao.dto.FunctionDTO;
import com.vmware.jct.dao.dto.LevelDTO;
import com.vmware.jct.dao.dto.OccupationListDTO;
import com.vmware.jct.dao.dto.UserProfileDTO;
import com.vmware.jct.exception.JCTException;
import com.vmware.jct.model.JctUser;
import com.vmware.jct.service.vo.UserVO;
/**
 * 
 * <p><b>Class name:</b>IAuthenticatorService.java</p>
 * <p><b>@author:</b> InterraIT </p>
 * <p><b>Purpose:</b> 		This Service interface represents a IAuthenticatorService interface
 * <p><b>Description:</b> 	This will have all methods where we implement all business logic and all service layer </p>
 * <p><b>Copyrights:</b> 	All rights reserved by InterraIT and should only be used for its internal application development.</p>
 *  <pre>
 * <b>Revision History:</b>
 * InterraIT - 31/Jan/2014 - Implement exception through out the application
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
	 * @param It will take jobRefNo,rowId,lastActivePage as a input parameter
	 * @return It returns String object
	 * @throws it throws JCTException
	 */
	public String updateLoginInfo(String jobRefNo, int rowId, String lastActivePage) throws JCTException;
	/**
	 * Method will return list of Function DTO
	 * @return list of FunctionDTO
	 * @throws JCTException
	 */
	public List<FunctionDTO> jctFunctionList() throws JCTException;
	/**
	 * Method will return list of Level DTO
	 * @return list of LevelDTO
	 * @throws JCTException
	 */
	public List<LevelDTO> jctLevelList() throws JCTException;
	/**
	 * Method will return list of User profile
	 * @return list of UserProfileDTO
	 * @throws JCTException
	 */
	public List<UserProfileDTO> jctUserProfileList() throws JCTException;
	
	/**
	 * Method fetches the survey question
	 * @param UserVO
	 * @return
	 * @throws JCTException
	 */
	public UserVO getSurveyQuestions (UserVO userVO) throws JCTException;
	
	public List<OccupationListDTO> searchOccupationList( String searchString ) throws JCTException;
	
	public String updateUser(String password, int activeFlag, String email) throws JCTException;
	
	public List<JctUser> getUser(String email) throws JCTException;
	/**
	 * Method to fetch Terms & Conditions by profile Id for facilitator
	 * @param userProfile
	 * @param userType
	 * @throws JCTException 
	 */
	public String getTermsAndConditions(int userProfile, int userType) throws JCTException;
}
