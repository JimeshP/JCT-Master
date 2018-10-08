package com.vmware.jct.service.impl;

import java.util.ArrayList;
import java.util.Arrays;
import java.util.Calendar;
import java.util.Collections;
import java.util.Date;
import java.util.Iterator;
import java.util.List;

import org.apache.commons.lang3.StringUtils;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.context.MessageSource;
import org.springframework.stereotype.Service;
import org.springframework.transaction.annotation.Propagation;
import org.springframework.transaction.annotation.Transactional;

import com.vmware.jct.common.utility.CommonConstants;
import com.vmware.jct.common.utility.PetComparator;
import com.vmware.jct.dao.IAfterSketchDAO;
import com.vmware.jct.dao.IAuthenticatorDAO;
import com.vmware.jct.dao.IBeforeSketchDAO;
import com.vmware.jct.dao.ICommonDao;
import com.vmware.jct.dao.ISketchSearchDAO;
import com.vmware.jct.dao.dto.DemographicDTO;
import com.vmware.jct.dao.dto.FunctionDTO;
import com.vmware.jct.dao.dto.LevelDTO;
import com.vmware.jct.dao.dto.OccupationListDTO;
import com.vmware.jct.dao.dto.RegionDTO;
import com.vmware.jct.exception.DAOException;
import com.vmware.jct.exception.JCTException;
import com.vmware.jct.model.JctInstructionBar;
import com.vmware.jct.model.JctPopupInstruction;
import com.vmware.jct.model.JctStatus;
import com.vmware.jct.model.JctSurveyQuestions;
import com.vmware.jct.model.JctUser;
import com.vmware.jct.model.JctUserDetails;
import com.vmware.jct.model.JctUserInfo;
import com.vmware.jct.model.JctUserLoginInfo;
import com.vmware.jct.model.JctUserProfile;
import com.vmware.jct.service.IAuthenticatorService;
import com.vmware.jct.service.vo.NavigationVO;
import com.vmware.jct.service.vo.PopulateInstructionVO;
import com.vmware.jct.service.vo.ReturnVO;
import com.vmware.jct.service.vo.UserLoginInfoVO;
import com.vmware.jct.service.vo.UserVO;


/**
 * 
 * <p><b>Class name:</b> AuthenticatorServiceImpl.java</p>
 * <p><b>Author:</b> InterreIt</p>
 * <p><b>Purpose:</b> This is a AuthenticatorServiceImpl class. This artifact is Business layer artifact.
 * AuthenticatorServiceImpl implement IAuthenticatorService interface and override the following  methods.
 * -authenticateUser(UserVO userVO)
 * -checkUser(UserVO userVO)
 * -register(UserVO userVO)
 * -resetPassword(UserVO userVO)
 * -forgotPassword(UserVO userVO)
 * -getTaskPendingDetails(int userId)
 * -getGeneratedJRId()
 * -saveLoginInfo(UserLoginInfoVO loginVO)
 * -updateLoginInfo(String jobRefNo, int rowId, String lastActivePage)
 * </p>
 * <p><b>Description:</b> This class used to perform insert, modify ,fetch data from db and delete operation. Save, update, delete operations are performed in DataAccessObject class </p>
 * <p><b>Revision History:</b>
 * 	<li>InterreIt - 19/Jan/2014 - Implement Exception </li>
 * <li>InterreIt - InterraIt - 31/Jan/2014 - Implement comments, introduce Named query, Transactional attribute </li>
 * </p>
 */
@Service
public class AuthenticatorServiceImpl implements IAuthenticatorService {

	@Autowired
	private IAuthenticatorDAO authenticatorDAO;
	
	@Autowired
	private IBeforeSketchDAO beforeSketchDAO;
	
	@Autowired
	private IAfterSketchDAO afterSketchDAO;
	
	@Autowired
	private ICommonDao commonDaoImpl;
	
	@Autowired  
	private MessageSource messageSource;
	
	@Autowired
	private ISketchSearchDAO sketchSearchDao;
	
	private static final Logger logger = LoggerFactory.getLogger(AuthenticatorServiceImpl.class);
	/**
	 *<p><b>Description:</b> Method is used to authenticate the user</p>
	 * @param It will take UserVO as a input parameter
	 * @return It returns List<JctUser>
	 * @throws JCTException
	 * 
	 */
	@Transactional(propagation=Propagation.REQUIRED)
	public List<JctUser> authenticateUser(UserVO userVO) throws JCTException {	
		logger.info(">>>> AuthenticatorServiceImpl.authenticateUser");
		List<JctUser> jctUser=new ArrayList<JctUser>();
		try{
			jctUser=authenticatorDAO.authenticateUser(userVO);
		}catch(DAOException daoException){
			//daoException.printStackTrace();
			throw new JCTException(daoException.getMessage());
		}
		logger.info("<<<< AuthenticatorServiceImpl.authenticateUser");
		return jctUser;
	}
	/**
	 *<p><b>Description:</b> Method is used to check user existence. UNUSED IN JCT TOOL</p>
	 * @param It will take UserVO as a input parameter
	 * @return It returns List<JctUser>
	 * @throws JCTException
	 * 
	 */
	@Transactional(propagation=Propagation.REQUIRED)
	public List<JctUser> checkUser(UserVO userVO) throws JCTException {
		List<JctUser> jctUser=new ArrayList<JctUser>();
		try{
			jctUser= authenticatorDAO.checkUser(userVO);
		}catch(DAOException daoException){
			logger.error(daoException.getLocalizedMessage());
			throw new JCTException(daoException.getMessage());
		}
		return jctUser;
	}
	/**
	 *<p><b>Description:</b> Method is used to check user existence. UNUSED IN JCT TOOL</p>
	 * @param It will take UserVO as a input parameter
	 * @return It returns List<JctUser>
	 * @throws JCTException
	 * 
	 */
	@Transactional(propagation=Propagation.REQUIRED)
	public List<JctUser> checkFacilitatorUser(UserVO userVO) throws JCTException {
		List<JctUser> jctUser=new ArrayList<JctUser>();
		try{
			jctUser= authenticatorDAO.checkFacilitatorUser(userVO);
		}catch(DAOException daoException){
			logger.error(daoException.getLocalizedMessage());
			throw new JCTException(daoException.getMessage());
		}
		return jctUser;
	}
	/**
	 *<p><b>Description:</b> Method is used to check inactive user existence.</p>
	 * @param It will take UserVO as a input parameter
	 * @return It returns List<JctUser>
	 * @throws JCTException
	 * 
	 */
	@Transactional(propagation=Propagation.REQUIRED)
	public List<JctUser> checkInactiveUser(UserVO userVO) throws JCTException {
		List<JctUser> jctUser=new ArrayList<JctUser>();
		try{
			jctUser= authenticatorDAO.checkInactiveUser(userVO);
		}catch(DAOException daoException){
			logger.error(daoException.getLocalizedMessage());
			throw new JCTException(daoException.getMessage());
		}
		return jctUser;
	}
	//
	/**
	 *<p><b>Description:</b> Method is used to register and self activate the user</p>
	 * @param It will take UserVO as a input parameter
	 * @return String
	 * @throws JCTException
	 * 
	 */
	@Transactional(propagation=Propagation.REQUIRED)
	public String register(UserVO userVO) throws JCTException {
		logger.info(">>>> AuthenticatorServiceImpl.register");
		//Get the user Obj
		String result = "";
		try{
			// Check if onet occupation exists in our database
			List<String> onetDescList = authenticatorDAO.getOnetDesc(userVO.getOccupationVal().trim());
			if (onetDescList.size() > 0) {
				String onetCode = authenticatorDAO.getOnetCodeByDesc(userVO.getOccupationVal().trim());
				List<JctUser> userList = authenticatorDAO.checkInactiveUser(userVO);
				if(userList.size()>0){
					JctUser userObj = (JctUser)userList.get(0);
					//Get the user details
					JctUserDetails details = userObj.getJctUserDetails();
					details.setJctUserDetailsFirstName(userVO.getFirstName());
					details.setJctUserDetailsLastName(userVO.getLastName());
					details.setJctUserDetailsGender(userVO.getGender());
					details.setJctUserDetailsSupervisePeople(userVO.getSupervise());
					details.setJctUserDetailsFunctionGroup(userVO.getFunctionGrpSelected());
					details.setJctUserDetailsTenure(userVO.getTenure());
					details.setJctUserDetailsRegion(userVO.getLocation());
					details.setJctUserDetailsLevels(userVO.getJobLevelSelected());
					details.setJctUserDetailsLastmodifiedTs(new Date());
					details.setJctUserOnetOccupation(onetCode);
					details.setNosOfYrs(userVO.getNosOfYrs());
					userObj.setLastmodifiedTs(new Date());
					userObj.setJctActiveYn(CommonConstants.ACTIVATION_COMPLETE);
					details.setJctUser(userObj);
					userObj.setJctUserDetails(details);
					result = authenticatorDAO.updateUser(userObj);
				}else{
					result = "not-found";
				}
			} else {
				result = "onet-not-found";
			}
			} catch (DAOException daoException) {
				logger.error(daoException.getLocalizedMessage());
				throw new JCTException(daoException.getMessage());
			}
		logger.info("<<<< AuthenticatorServiceImpl.register");
		return result;
	}
	/**
	 *<p><b>Description:</b> Method is used to reset the user password</p>
	 * @param It will take UserVO as a input parameter
	 * @return String
	 * @throws JCTException
	 * 
	 */
	@Transactional(propagation=Propagation.REQUIRED)
	public String resetPassword(UserVO userVO, int dist) throws JCTException {
		logger.info(">>>> AuthenticatorServiceImpl.resetPassword");
		//Get the user Obj
		String result = "";
		try{
			List<JctUser> userList = null;
			//List<JctUser> userList = authenticatorDAO.checkUsers(userVO);
			if (dist == 1) {
				userList = authenticatorDAO.checkInactiveUser(userVO);
			} else {
				userList = authenticatorDAO.checkActiveUser(userVO);
				if (userList.size() == 0) { // If the user has not registered but opted to reset the password
					userList = authenticatorDAO.checkInactiveUser(userVO);
				}
			}
			
			if(userList.size()>0){
				JctUser userObj = (JctUser)userList.get(0);
				
				// Check if the user is normal user only
				if (userObj.getJctUserRole().getJctRoleId() == 1) {
					if (userVO.getClickStatus().equals("normal")) {
						int activationStatus = userObj.getJctActiveYn();
						/*if ( (activationStatus == CommonConstants.ACTIVATED_ONLY_EMAIL) || 
								(activationStatus == CommonConstants.ACTIVATION_COMPLETE) ) {*/
							userObj.setJctPassword(userVO.getPassword());
							userObj.setJctActiveYn(userVO.getActiveYn()); //2
							result = authenticatorDAO.updateUser(userObj);
						/*} else {
							result = "not-complete";
						}*/
					} else {
						userObj.setJctPassword(userVO.getPassword());
						userObj.setJctActiveYn(userVO.getActiveYn()); //2
						result = authenticatorDAO.updateUser(userObj);
					}
				}
			} else {
				result = "not-found";
			}
			}catch(DAOException daoException) {
				logger.error(daoException.getLocalizedMessage());
				throw new JCTException(daoException.getMessage());
			}
		logger.info("<<<< AuthenticatorServiceImpl.resetPassword");
		return result;
	}
	/**
	 *<p><b>Description:</b> Method is used to re generate the password when the user forgot it</p>
	 * @param It will take UserVO as a input parameter
	 * @return String
	 * @throws JCTException
	 * 
	 */
	@Transactional(propagation=Propagation.REQUIRED)
	public ReturnVO forgotPassword(UserVO userVO) throws JCTException {
		logger.info(">>>> AuthenticatorServiceImpl.forgotPassword");
		ReturnVO retrunVO = new ReturnVO();
		String status = "";
		retrunVO.setUserName(userVO.getUserName());
		try{
			List<JctUser> userList = authenticatorDAO.checkUser(userVO);
			if (userList.size() > 0) {
				JctUser user = (JctUser) userList.get(0);
				if (user.getJctUserRole().getJctRoleId() == 1) {
					/*int activationStatus = user.getJctActiveYn();
					 if ( (activationStatus == CommonConstants.ACTIVATED_ONLY_EMAIL) || 
							(activationStatus == CommonConstants.ACTIVATION_COMPLETE) ) {*/
						user.setJctPassword(userVO.getPassword());
						user.setJctActiveYn(CommonConstants.CREATED); //1
						user.setLastmodifiedTs(new Date());
						if(null != user.getJctUserDetails().getJctUserDetailsFirstName()) {
							retrunVO.setfName(user.getJctUserDetails().getJctUserDetailsFirstName());
						} else {
							retrunVO.setfName("User");
						}
						if(null != user.getJctUserDetails().getJctUserDetailsLastName()) {
							retrunVO.setlName(user.getJctUserDetails().getJctUserDetailsLastName());
						} else {
							retrunVO.setlName("");
						}
						status = authenticatorDAO.merge(user);
					/*} else {
						status = "not-complete";
					}*/
				} else {
					status = "not-found";
				}
				
				if (status.equals("success")) {
					retrunVO.setStatusCode(CommonConstants.SUCCESSFUL);
					retrunVO.setStatusMsg(status);
					retrunVO.setStatusDesc(this.messageSource.getMessage("error.passwordChangedSuccessfully",null, null));
				} else if (status.equals("not-complete")) {
					retrunVO.setStatusCode(CommonConstants.PASSWORD_RESET_NOT_COMPLETE);
					retrunVO.setStatusMsg(status);
					retrunVO.setStatusDesc(this.messageSource.getMessage("error.passwordResetNotComplete",null, null));
				} else if (status.equals("not-found")) {
					retrunVO.setStatusCode(CommonConstants.USER_NOT_FOUND);
					retrunVO.setStatusMsg("failure");
					retrunVO.setStatusDesc(this.messageSource.getMessage("error.userPasswordNotMatch",null, null));
				} else {
					retrunVO.setStatusCode(CommonConstants.PASSWORD_RESET_FAILED);
					retrunVO.setStatusMsg("failure");
					retrunVO.setStatusDesc(this.messageSource.getMessage("error.changePasswordOnceAgain",null, null));
				}
				
				/*if (status.equalsIgnoreCase("success")) {
					retrunVO.setStatusCode(CommonConstants.SUCCESSFUL);
					retrunVO.setStatusDesc(this.messageSource.getMessage("error.confirmation",null, null));
					retrunVO.setPathFinder(user.getJctUserEmail());
					retrunVO.setUserName(user.getJctUserDetails().getJctUserDetailsFirstName()+" "+user.getJctUserDetails().getJctUserDetailsLastName());
				} else if (status.equalsIgnoreCase("not-complete")) {
					retrunVO.setStatusCode(CommonConstants.PASSWORD_RESET_NOT_COMPLETE);
					retrunVO.setStatusDesc(this.messageSource.getMessage("error.passwordResetNotComplete",null, null));					
				} else {
					retrunVO.setStatusCode(CommonConstants.FAILURE); //Server update problemthis.messageSource.getMessage("error.confirmation",null, null)
					retrunVO.setStatusDesc(this.messageSource.getMessage("error.updationError",null, null));
				}*/
			} else {
				retrunVO.setStatusCode(CommonConstants.USER_NOT_FOUND); //user not present
				retrunVO.setStatusDesc(this.messageSource.getMessage("error.userNotAvaliable",null, null));
			}
		}catch(DAOException daoException){
			//daoException.printStackTrace();
			logger.error(daoException.getLocalizedMessage());
			throw new JCTException(daoException.getMessage());
			
		}
		logger.info("<<<< AuthenticatorServiceImpl.forgotPassword");
		return retrunVO;
	}
	/**
	 *<p><b>Description:</b> This method is used for fetching  user info data based on user id and status</p>
	 * @param userId
	 * @return List<JctUserLoginInfo>
	 * @throws JCTException
	 * 
	 */
	@Transactional(propagation=Propagation.REQUIRED)
	public List<JctUserLoginInfo> getTaskPendingDetails(int userId)
			throws JCTException {
		logger.info(">>>> AuthenticatorServiceImpl.getTaskPendingDetails");
		List<JctUserLoginInfo> jctUserLoginInfo= new ArrayList<JctUserLoginInfo>();
		try{
			jctUserLoginInfo=authenticatorDAO.getTaskPendingDetails(userId);
		}catch(DAOException daoException){
			//daoException.printStackTrace();
			logger.error(daoException.getLocalizedMessage());
			throw new JCTException(daoException.getMessage());
			
		}
		logger.info("<<<< AuthenticatorServiceImpl.getTaskPendingDetails");
		return jctUserLoginInfo;
	}
	/**
	 *<p><b>Description:</b> This method is used to generate new job reference number</p>
	 * @return Integer
	 * @throws JCTException
	 * 
	 */
	@Transactional(propagation=Propagation.REQUIRED)
	public int getGeneratedJRId() throws JCTException {
		logger.info(">>>> AuthenticatorServiceImpl.getGeneratedJRId");
		int generatedId= 0;
		try{
			generatedId=commonDaoImpl.generateKey("jct_job_reference_no");
		}catch(DAOException daoException){
			//daoException.printStackTrace();
			logger.error(daoException.getLocalizedMessage());
			throw new JCTException(daoException.getMessage());
			
		}
		logger.info("<<<< AuthenticatorServiceImpl.getGeneratedJRId");
		return generatedId;
	}
	/**
	 *<p><b>Description:</b> This method is used to save user login information to the login info table</p>
	 * @param UserLoginInfoVO
	 * @return String
	 * @throws JCTException
	 * 
	 */
	@Transactional(propagation=Propagation.REQUIRED)
	public String saveLoginInfo(UserLoginInfoVO loginVO) throws JCTException {
		logger.info(">>>> AuthenticatorServiceImpl.saveLoginInfo");
		JctUserLoginInfo info = new JctUserLoginInfo();
		try{
			// Fetch the user email
			String emailId = authenticatorDAO.getEmailIdByUserId(loginVO.getUserId());
			
			info.setJctUserId(loginVO.getUserId());
			info.setJctJobrefNo(loginVO.getJobReferenceNo());
			info.setJctPageInfo(loginVO.getPageInfo());
			info.setJctStartTime(new Date());
			JctStatus status = new JctStatus();
			status.setJctStatusId(4);
			info.setJctStatus(status); // Pending
			info.setJctPageInfo(loginVO.getPageInfo());
			info.setJctAsCreatedTs(new Date());
			info.setJctAsLastmodifiedTs(new Date());
			info.setJctAsCreatedBy(emailId);
			info.setJctAsLastmodifiedBy(emailId);
			info.setJctCompleted(loginVO.getIsCompleted());
			
			//We need the index also to put the exact end date hence concatenating
			String resultString = authenticatorDAO.saveLoginInfo(info);
			logger.info("<<<< AuthenticatorServiceImpl.saveLoginInfo");
			return resultString+"#"+info.getJctUserLoginInfoId();
		} catch(DAOException daoException ){
			logger.error(daoException.getLocalizedMessage());
			throw new JCTException(daoException.getMessage());
		}
	}
	/**
	 *<p><b>Description:</b> This method is used to update user login information to the 
	 *login info table after the user logs out</p>
	 * @param job reference number
	 * @param rowId
	 * @param last active page
	 * @return String
	 * @throws JCTException
	 * 
	 */
	@Transactional(propagation=Propagation.REQUIRED)
	public String updateLoginInfo(String jobRefNo, int rowId, String lastActivePage) throws JCTException {
		logger.info(">>>> AuthenticatorServiceImpl.updateLoginInfo");
		try{
			logger.info("<<<< AuthenticatorServiceImpl.updateLoginInfo");
			return authenticatorDAO.updateLoginInfo(jobRefNo, rowId, lastActivePage);
		}catch(DAOException daoException ){
			logger.error(daoException.getLocalizedMessage());
			//daoException.printStackTrace();
			throw new JCTException(daoException.getMessage());
		}
	}
	/**
	 *<p><b>Description:</b> This method fetches list of active regions</p>
	 * @return List<RegionDTO>
	 * @throws JCTException
	 * 
	 */
	@Transactional(propagation=Propagation.REQUIRED)
	public List<RegionDTO> jctRegionList() throws JCTException {
		logger.info(">>>> AuthenticatorServiceImpl.jctRegionList");
		List<RegionDTO> jctRegionList=null;
		try{
			jctRegionList= authenticatorDAO.getJctRegionList();
		}catch(DAOException daoException ){
			logger.error(daoException.getLocalizedMessage());
		//	daoException.printStackTrace();
			throw new JCTException(daoException.getMessage());
		}
		logger.info("<<<< AuthenticatorServiceImpl.jctRegionList");
		return jctRegionList;
	}
	/**
	 *<p><b>Description:</b> This method fetches list of active function group</p>
	 * @return List<FunctionDTO>
	 * @throws JCTException
	 * 
	 */
	@Transactional(propagation=Propagation.REQUIRED)
	public List<FunctionDTO> jctFunctionList() throws JCTException {
		logger.info(">>>> AuthenticatorServiceImpl.jctFunctionList");
		List<FunctionDTO> jctFunctionList=null;
		try{
			jctFunctionList= authenticatorDAO.getJctFunctionList();
		}catch(DAOException daoException ){
			logger.error(daoException.getLocalizedMessage());
			//daoException.printStackTrace();
			throw new JCTException(daoException.getMessage());
		}
		logger.info("<<<< AuthenticatorServiceImpl.jctFunctionList");
		return jctFunctionList;
	}
	/**
	 *<p><b>Description:</b> This method fetches list of active job level</p>
	 * @return List<LevelDTO>
	 * @throws JCTException
	 * 
	 */
	@Transactional(propagation=Propagation.REQUIRED)
	public List<LevelDTO> jctLevelList() throws JCTException {
		logger.info(">>>> AuthenticatorServiceImpl.jctLevelList");
		List<LevelDTO> jctLevelList=null;
		try{
			jctLevelList= authenticatorDAO.getJctLevelList();
		}catch(DAOException daoException ){
			//daoException.printStackTrace();
			logger.error(daoException.getLocalizedMessage());
			throw new JCTException(daoException.getMessage());
		}
		logger.info("<<<< AuthenticatorServiceImpl.jctLevelList");
		return jctLevelList;
	}
	/**
	 *<p><b>Description:</b> This method checks if user has completed his after sketch or not 
	 *</p>
	 *@param job reference number
	 * @return List<LevelDTO>
	 * @throws JCTException
	 * 
	 */
	@Transactional(propagation=Propagation.REQUIRED)
	public int getIsCompleted(String refNo) throws JCTException {
		logger.info(">>>> AuthenticatorServiceImpl.getIsCompleted");
		int count = 0;
		try{
			count = authenticatorDAO.getIsCompleted(refNo).intValue();
		} catch(DAOException daoException){
			//daoException.printStackTrace();
			logger.error(daoException.getLocalizedMessage());
			throw new JCTException();
		}
		logger.info("<<<< AuthenticatorServiceImpl.getIsCompleted");
		return count;
	}
	/**
	 *<p><b>Description:</b> This method updates the next page information</p>
	 * @param job reference number
	 * @param next page
	 * @return String
	 * @throws JCTException
	 * 
	 */
	@Transactional(propagation=Propagation.REQUIRED)
	public String updateNextPage(String jobRefNo, String nextPage) throws JCTException {
		logger.info(">>>> AuthenticatorServiceImpl.updateNextPage");
		String result = "failed";
		try{
			authenticatorDAO.updateNextPage(jobRefNo, nextPage);
			result = "success";
		}catch(DAOException daoException ){
			logger.error(daoException.getLocalizedMessage());
			//daoException.printStackTrace();
			throw new JCTException(daoException.getMessage());
		}
		logger.info("<<<< AuthenticatorServiceImpl.updateNextPage");
		return result;
	}
	/**
	 *<p><b>Description:</b> Method is used to store the user information if the user has 
	 *provided wrong password. Eventually after 5 attempt, the account will get locked.</p>
	 * @param user id
	 * @param user email id
	 * @return Integer
	 * @throws JCTException
	 * 
	 */
	@Transactional(propagation=Propagation.REQUIRED)
	public int saveUserValidationFails(int userId, String userEmail)
			throws JCTException {
		logger.info(">>>> AuthenticatorServiceImpl.saveUserValidationFails");
		int totalFailedCount = 1;
		//Fetch the object.
		try {
			List<JctUserInfo> infoList = authenticatorDAO.fetchUserInfo(userId, userEmail);
			if (infoList.size() > 0) {
				JctUserInfo info = (JctUserInfo) infoList.get(0);
				Calendar systemTime = Calendar.getInstance();
				Calendar infoObject = Calendar.getInstance();
				infoObject.setTime(info.getJctCreatedTs());
				totalFailedCount = info.getJctNoOfCount() + 1;
				//Check the difference
				long totalMillis = systemTime.getTimeInMillis() - infoObject.getTimeInMillis();
				//86400000 = 24 hours
				if (totalMillis < 86400000l) {
					int chance = info.getJctNoOfCount() + 1;
					info.setJctNoOfCount(chance);
					//Merge the result
					String result = authenticatorDAO.saveUserInfo(info);
				} else {
					//soft delete the existing and insert new
					info.setJctSoftDelete(CommonConstants.ENTRY_SOFT_DELETED);
					authenticatorDAO.saveUserInfo(info);
					createAndSaveFailedLogin(userId, userEmail);
					totalFailedCount = 99;
				}
			} else { //Insert a new record
				createAndSaveFailedLogin(userId, userEmail);
			}
		} catch (DAOException daoException) {
			logger.error(daoException.getLocalizedMessage());
			//daoException.printStackTrace();
			throw new JCTException();
		}
		logger.info("<<<< AuthenticatorServiceImpl.saveUserValidationFails");
		return totalFailedCount;
	}
	/**
	 *<p><b>Description:</b> Method is used to check total number of login failed counts.</p>
	 * @param user id
	 * @param user email id
	 * @return Integer
	 * @throws JCTException
	 * 
	 */
	@Transactional(propagation=Propagation.REQUIRED)
	public int getTotalLoginFailedCounts (int userId, String userName) throws JCTException {
		logger.info(">>>> AuthenticatorServiceImpl.getTotalLoginFailedCounts");
		int totalFailedCount = 0;
		try {
			List<JctUserInfo> infoList = authenticatorDAO.fetchUserInfo(userId, userName);
			if (infoList.size() > 0) {
				JctUserInfo info = (JctUserInfo) infoList.get(0);
				Calendar systemTime = Calendar.getInstance();
				Calendar infoObject = Calendar.getInstance();
				infoObject.setTime(info.getJctCreatedTs());
				long totalMillis = systemTime.getTimeInMillis() - infoObject.getTimeInMillis();
				//86400000 = 24 hours
				if (totalMillis > 86400000l) { //old data somehow has not been soft deleted
					info.setJctSoftDelete(CommonConstants.ENTRY_SOFT_DELETED);
					authenticatorDAO.saveUserInfo(info);
				} else {
					totalFailedCount = info.getJctNoOfCount();
				}
			}
		} catch (DAOException daoException) {
			logger.error(daoException.getLocalizedMessage());
			//daoException.printStackTrace();
			throw new JCTException();
		}
		logger.info("<<<< AuthenticatorServiceImpl.getTotalLoginFailedCounts");
		return totalFailedCount;
	}
	/**
	 *<p><b>Description:</b> Method is used to check if the user is allowed to click the forgot password link or not.
	 *Only the registered users will be able to click it.</p>
	 * @param user email
	 * @return Boolean
	 * @throws JCTException
	 * 
	 */
	@Transactional(propagation=Propagation.REQUIRED)
	public boolean isForgotPwdAllowed(String userName) throws JCTException {
		logger.info(">>>> AuthenticatorServiceImpl.isForgotPwdAllowed");
		int totalFailedCount = 0;
		boolean isAllowed = false;
		try {
			List<JctUserInfo> infoList = authenticatorDAO.fetchUserInfoByEmail(userName);
			if (infoList.size() > 0) {
				JctUserInfo info = (JctUserInfo) infoList.get(0);
				Calendar systemTime = Calendar.getInstance();
				Calendar infoObject = Calendar.getInstance();
				infoObject.setTime(info.getJctCreatedTs());
				long totalMillis = systemTime.getTimeInMillis() - infoObject.getTimeInMillis();
				//86400000 = 24 hours
				if (totalMillis > 86400000l) { //old data somehow has not been soft deleted
					info.setJctSoftDelete(CommonConstants.ENTRY_SOFT_DELETED);
					authenticatorDAO.saveUserInfo(info);
				} else {
					totalFailedCount = info.getJctNoOfCount();
				}
			}
		} catch (DAOException daoException) {
			logger.error(daoException.getLocalizedMessage());
			//daoException.printStackTrace();
			throw new JCTException();
		}
		if (totalFailedCount < 5) {
			isAllowed = true;
		}
		logger.info("<<<< AuthenticatorServiceImpl.isForgotPwdAllowed");
		return isAllowed;
	}
	
	private String createAndSaveFailedLogin(int userId, String userEmail) throws DAOException {
		logger.info(">>>> AuthenticatorServiceImpl.createAndSaveFailedLogin");
		JctUserInfo info = new JctUserInfo();
		//info.setJctUserInfoId(commonDaoImpl.generateKey("jct_user_info_id"));
		info.setJctUserId(userId);
		info.setJctUserEmail(userEmail);
		info.setJctCreatedTs(new Date());
		info.setJctNoOfCount(1); //initial count
		info.setJctSoftDelete(CommonConstants.ENTRY_NOT_DELETED);
		logger.info("<<<< AuthenticatorServiceImpl.createAndSaveFailedLogin");
		return authenticatorDAO.saveUserInfo(info);
	}
	/**
	 *<p><b>Description:</b> Method is used to change the user information status.</p>
	 * @param user id
	 * @param user email id
	 * @return String
	 * @throws JCTException
	 * 
	 */
	@Transactional(propagation=Propagation.REQUIRED)
	public String changeInfoStatus(int userId, String userName)
			throws JCTException {
		logger.info(">>>> AuthenticatorServiceImpl.changeInfoStatus");
		String status = "success";
		try {
			List<JctUserInfo> infoList = authenticatorDAO.fetchUserInfo(userId, userName);
			if (infoList.size() > 0) {
				JctUserInfo info = (JctUserInfo) infoList.get(0);
				info.setJctSoftDelete(CommonConstants.ENTRY_SOFT_DELETED);
				authenticatorDAO.saveUserInfo(info);
			}
		} catch (DAOException daoException) {
			status = "failure";
			logger.error(daoException.getLocalizedMessage());
			//daoException.printStackTrace();
			throw new JCTException();
		}
		logger.info("<<<< AuthenticatorServiceImpl.changeInfoStatus");
		return status;
	}
	/**
	 *<p><b>Description:</b> Method is used to populate instruction bar data.</p>
	 * @param profile id
	 * @param page for which instruction data is to be fetched
	 * @return NavigationVO
	 * @throws JCTException
	 * 
	 */
	@Transactional(propagation = Propagation.REQUIRED)
	public NavigationVO populateInstruction(int profileId, String relatedPage) {	
		logger.info(">>>> AuthenticatorServiceImpl.populateInstruction");
		NavigationVO vo = new NavigationVO();
		try {
			JctInstructionBar instruction = authenticatorDAO.populateExistingInstruction(profileId, relatedPage);
			if (null != instruction) {						
				vo.setInstructionDesc(instruction.getJctInstructionBarDesc());
				vo.setVideoUrl(instruction.getJctVideoPath());
				vo.setStatusCode(Integer.parseInt(CommonConstants.SUCCESSFUL));
			} else {
				vo.setInstructionDesc("");
				vo.setVideoUrl(null);
				vo.setStatusCode(Integer.parseInt(CommonConstants.USER_REGISTRATION_FAILED));				
			}
		} catch (DAOException e) {
			//e.printStackTrace();
			logger.error(e.getLocalizedMessage());
		}
		logger.info("<<<< AuthenticatorServiceImpl.populateInstruction");
		return vo;
	}
	/**
	 *<p><b>Description:</b> Method fetches only profile id.</p>
	 * @param email id
	 * @return profile id - integer
	 * @throws JCTException
	 * 
	 */
	@Transactional(propagation = Propagation.REQUIRED)
	public Integer fetchProfileId(String emailId) throws JCTException {
		int profileId = 0;
		try {
			profileId = authenticatorDAO.fetchProfileId(emailId);
		} catch (DAOException exp) {
			logger.error(exp.getLocalizedMessage());
		}
		return profileId;
	}
	/**
	 *<p><b>Description:</b> Method updates the demographic information.</p>
	 * @param UserVO
	 * @return int
	 * @throws JCTException
	 * 
	 */
	@Transactional(propagation = Propagation.REQUIRED)
	public int updateDemographicInfo(UserVO userVO) throws JCTException {
		int status = 500; //failed
		logger.info(">>>> AuthenticatorServiceImpl.updateDemographicInfo");
		try {
			List<JctUser> userList = authenticatorDAO.checkUser(userVO);
			if (userList.size() > 0) {
				JctUser user = (JctUser) userList.get(0);
				//Get the user details object
				JctUserDetails details = user.getJctUserDetails();
				
				/*details.setJctUserDetailsRegion(userVO.getLocation());
				details.setJctUserDetailsFunctionGroup(userVO.getFunctionGrpSelected());
				details.setJctUserDetailsLevels(userVO.getJobLevelSelected());*/
				
				details.setJctUserOnetOccupation(authenticatorDAO.getOnetCodeByDesc(userVO.getOccupationVal().trim()));			
				
				details.setJctUser(user);
				user.setJctUserDetails(details);
			
				authenticatorDAO.updateUser(user);
				//Fetch Before Sketch
				/*List<JctBeforeSketchHeader> headers = beforeSketchDAO.getBeforeSketchActiveObject(userVO);
				JctBeforeSketchHeader header = null;
				if (headers.size() > 1) {
					logger.error("DATA INCONSISTENCY FOR user: "+userVO.getEmail()+" IN JCT BEFORE SKETCH HEADER... FOUND MULTIPLE ROWS WITH 4, 0 STATUS. LAST ROW WILL BE CONSIDERED");
				}
				header = (JctBeforeSketchHeader) headers.get(headers.size()-1); //Always last one is considered
				header.setJctBsUserJobTitle(userVO.getJobLevelSelected());
				beforeSketchDAO.saveOrUpdateBeforeSketch(header);
				//Fetch After Sketch
				List<JctAfterSketchHeader> asHeaders = afterSketchDAO.getAfterSketchActiveObject(userVO);
				JctAfterSketchHeader asHeader = null;
				if (asHeaders.size() > 1) {
					logger.error("DATA INCONSISTENCY FOR user: "+userVO.getEmail()+" IN JCT BEFORE AFTER HEADER... FOUND MULTIPLE ROWS WITH 4, 0 STATUS. LAST ROW WILL BE CONSIDERED");
				}
				asHeader = (JctAfterSketchHeader) asHeaders.get(asHeaders.size()-1); //Always last one is considered
				asHeader.setJctAsUserJobTitle(userVO.getJobLevelSelected());
				afterSketchDAO.updateHeaderPage(asHeader);*/
				status = 200;
			} else {
				status = 600; // some how data got deleted
			}
		} catch (DAOException jctEx) {
			logger.error(jctEx.getLocalizedMessage());
		}
		logger.info("<<<< AuthenticatorServiceImpl.updateDemographicInfo");
		return status;
	}
	/**
	 * Method fetches dto containing demographic information based on email id
	 * @param emailId
	 * @return String
	 * @throws JCTException
	 */
	@Transactional(propagation = Propagation.REQUIRED)
	public String getDemographicInformation ( String emailId ) throws JCTException {
		StringBuffer sb = new StringBuffer("<p><table align='center'><width='96%'>");
		try {
			List<DemographicDTO> demographicDTO = (List<DemographicDTO>) authenticatorDAO.getDemographicInformation(emailId);
			Iterator<DemographicDTO> itr = demographicDTO.iterator();
			while ( itr.hasNext() ) {
				DemographicDTO dto = (DemographicDTO) itr.next();
				sb.append("<tr><td align='left'><b>");
				sb.append(this.messageSource.getMessage("label.demo.details.onet.data",null, null));
				sb.append("</b></td><td align='left'>");
				if (null != dto.getOnetOccupationData()) {
					sb.append(sketchSearchDao.getOccupationDesc(dto.getOnetOccupationData()));
				} else {
					sb.append("N/A");
				}
				sb.append("</td></tr>");
								
				/*sb.append("<tr><td align='left'><b>");
				sb.append(this.messageSource.getMessage("label.demo.details.function.group",null, null));
				sb.append("</b></td><td align='left'>");
				if (null != dto.getFunctionGroup()) {
					sb.append(dto.getFunctionGroup());
				} else {
					sb.append("N/A");
				}
				sb.append("</td></tr>");
				
				sb.append("<tr><td align='left'><b>");
				sb.append(this.messageSource.getMessage("label.demo.job.level",null, null));
				sb.append("</b></td><td align='left'>");
				if (null != dto.getJobLevel()) {
					sb.append(dto.getJobLevel());
				} else {
					sb.append("N/A");
				}
				sb.append("</td></tr>");*/
			}
		} catch (DAOException ex) {
			logger.error(ex.getLocalizedMessage());
			throw new JCTException();
		}
		sb.append("</table></p>");
		return sb.toString();
	}
	/**
	 * Method fetches email id
	 * @param rowId
	 * @return String
	 * @throws JCTException
	 */
	@Transactional(propagation = Propagation.REQUIRED)
	public String getEmailIdByRowNos ( int rowNo ) throws JCTException {
		StringBuilder sb = new StringBuilder("");
		try {
			sb.append(authenticatorDAO.getEmailIdByRowNos(rowNo));
		} catch (DAOException ex) {
			logger.error(ex.getLocalizedMessage());
			throw new JCTException();
		}
		return sb.toString();
	}
	/**
	 * Method deletes all the dummy data
	 */
	@Transactional(propagation = Propagation.REQUIRED)
	public String deleteDummyData(String jobReferenceNos) throws JCTException {
		// Remove Questionnaire
		deleteDummyQuestionnaire(jobReferenceNos);
		// Remove before sketch entries
		deleteBeforeSketchDummyEntry(jobReferenceNos);
		deleteBeforeSketchTempDummyEntry(jobReferenceNos);
		// Remove Action Plan
		deleteDummyActionPlan(jobReferenceNos);
		// Remove After Sketch
		deleteDummyAfterSketchData(jobReferenceNos);
		deleteDummyAfterSketchTempData(jobReferenceNos);
		// Remove From Status Search Table
		deleteDummyStatusSearch(jobReferenceNos);
		// Remove From Completion Status
		deleteFromCompletionStatus(jobReferenceNos);
		// Remove From Login Info
		deleteFromLoginInfo(jobReferenceNos);
		return null;
	}
	
	private void deleteDummyQuestionnaire (String jobRefNo) {
		try {
			authenticatorDAO.deleteQuestionnaireTempData(jobRefNo);
			logger.info("---- ALL QUESTIONNAIRE DUMMY TEMP DATA DELETED ----");
		} catch (DAOException ex) {
			logger.error("UNABLE TO DELETE DUMMY QUESTIONNAIRE DATA");
		}
	}
	private void deleteBeforeSketchDummyEntry (String jobRefNo) {
		try {
			authenticatorDAO.deleteBSChildData(jobRefNo);
			authenticatorDAO.deleteBSData(jobRefNo);
			logger.info("---- ALL BEFORE SKETCH DUMMY DATA DELETED ----");
		} catch (DAOException ex) {
			logger.error("UNABLE TO DELETE DUMMY BEFORE SKETCH DATA");
		}
	}
	private void deleteBeforeSketchTempDummyEntry(String jobRefNo) {
		try {
			authenticatorDAO.deleteBSTempChildData(jobRefNo);
			authenticatorDAO.deleteBSTempData(jobRefNo);
			logger.info("---- ALL BEFORE SKETCH DUMMY TEMP DATA DELETED ----");
		} catch (DAOException ex) {
			logger.error("UNABLE TO DELETE DUMMY BEFORE SKETCH TEMP DATA");
		}
	}
	private void deleteDummyActionPlan(String jobReferenceNos) {
		try {
			authenticatorDAO.deleteActionPlanTempData(jobReferenceNos);
			logger.info("---- ALL ACTION PLAN DUMMY DATA DELETED ----");
		} catch (DAOException ex) {
			logger.error("UNABLE TO DELETE DUMMY ACTION PLAN DATA");
		}
	}
	
	private void deleteDummyAfterSketchData(String jobRefNo) {
		try {
			authenticatorDAO.deleteASPageOneData(jobRefNo);
			authenticatorDAO.deleteASPageTwoData(jobRefNo);
			authenticatorDAO.deleteASData(jobRefNo);
			logger.info("---- ALL AFTER SKETCH DUMMY DATA DELETED ----");
		} catch (DAOException ex) {
			logger.error("UNABLE TO DELETE DUMMY AFTER SKETCH DATA");
		}
	}
	private void deleteDummyAfterSketchTempData(String jobReferenceNos) {
		try {
			authenticatorDAO.deleteASPageOneTempData(jobReferenceNos);
			authenticatorDAO.deleteASPageTwoTempData(jobReferenceNos);
			authenticatorDAO.deleteASTempData(jobReferenceNos);
			logger.info("---- ALL AFTER SKETCH DUMMY TEMP DATA DELETED ----");
		} catch (DAOException ex) {
			logger.error("UNABLE TO DELETE DUMMY AFTER SKETCH TEMP DATA");
		}
	}
	
	private void deleteDummyStatusSearch(String jobReferenceNos) {
		try {
			authenticatorDAO.deleteStatusSearchDummyData(jobReferenceNos);
			logger.info("---- ALL STATUS SEARCH DUMMY DATA DELETED ----");
		} catch (DAOException ex) {
			logger.error("UNABLE TO DELETE DUMMY STATUS SEARCH DATA");
		}
	}
	
	private void deleteFromCompletionStatus(String jobReferenceNos) {
		try {
			authenticatorDAO.deleteCompletionDummyData(jobReferenceNos);
			logger.info("---- ALL COMPLETION DUMMY DATA DELETED ----");
		} catch (DAOException ex) {
			logger.error("UNABLE TO DELETE DUMMY COMPLETION STATUS DATA");
		}
	}
	
	private void deleteFromLoginInfo(String jobReferenceNos) {
		try {
			authenticatorDAO.deleteDummyLoginInfoData(jobReferenceNos);
			logger.info("---- ALL LOGIN INFO DUMMY DATA DELETED ----");
		} catch (DAOException ex) {
			logger.error("UNABLE TO DELETE DUMMY LOGIN INFO DATA");
		}
	}
	/**
	 * Method fetches survey question for general user
	 */
	@Transactional(propagation = Propagation.REQUIRED)
	public UserVO getSurveyQuestions(UserVO userVO) throws JCTException {
		logger.info(">>>>>>>> AuthenticatorServiceImpl.getSurveyQuestions");
		try {
			// Get all the question based on the order as set by admin
			List<Object[]> allSurveyQtns = authenticatorDAO.getAllSurveyQuestion();
			StringBuffer qtnBuffer = new StringBuffer();
			int loopCounter = 0;
			for (int index = 0; index <allSurveyQtns.size(); index++) {
				loopCounter = loopCounter + 1;
				// Get the object
				Object[] obj = (Object[]) allSurveyQtns.get(index);
				// Check the type of question
				int type = (Integer) obj[2];
				if (type == 1) {			// text
					qtnBuffer.append("1`");
					qtnBuffer.append((String) obj[0]);	//	ques
					qtnBuffer.append("#");
					
					qtnBuffer.append((String) obj[1]);	//	mandatory field
					qtnBuffer.append("<>");
				} else if (type == 2) {		// Radio
					qtnBuffer.append("2`");
					String mainQtn = (String) obj[0];
					qtnBuffer.append(mainQtn);			//	main ques
					qtnBuffer.append("#");				
					qtnBuffer.append((String) obj[1]);	//	mandatory field					
					qtnBuffer.append("~");		
					
					List<String> subQtnDtoList = authenticatorDAO.getAllSubQtns(2, mainQtn);
					int loop = 0;
					for (int i=0; i<subQtnDtoList.size(); i++) {
						String subQtn = (String) subQtnDtoList.get(i);
						loop = loop + 1;
						qtnBuffer.append(subQtn);
						if (loop < subQtnDtoList.size()) {
							qtnBuffer.append("^");
						}
					}
					qtnBuffer.append("<>");
				} else if (type == 3) {		// Dropdown
					qtnBuffer.append("3`");
					String mainQtn = (String) obj[0];
					qtnBuffer.append(mainQtn);			//	main ques
					qtnBuffer.append("#");				
					qtnBuffer.append((String) obj[1]);	//	mandatory field					
					qtnBuffer.append("~");		
					
					List<String> subQtnDtoList = authenticatorDAO.getAllSubQtns(3, mainQtn);
					int loop = 0;
					for (int i=0; i<subQtnDtoList.size(); i++) {
						String subQtn = (String) subQtnDtoList.get(i);
						loop = loop + 1;
						qtnBuffer.append(subQtn);
						if (loop < subQtnDtoList.size()) {
							qtnBuffer.append("^");
						}
					}
					qtnBuffer.append("<>");
				} else {					// Checkbox
					qtnBuffer.append("4`");
					String mainQtn = (String) obj[0];
					qtnBuffer.append(mainQtn);			//	main ques
					qtnBuffer.append("#");				
					qtnBuffer.append((String) obj[1]);	//	mandatory field					
					qtnBuffer.append("~");		
					
					List<String> subQtnDtoList = authenticatorDAO.getAllSubQtns(4, mainQtn);
					int loop = 0;
					for (int i=0; i<subQtnDtoList.size(); i++) {
						String subQtn = (String) subQtnDtoList.get(i);
						loop = loop + 1;
						qtnBuffer.append(subQtn);
						if (loop < subQtnDtoList.size()) {
							qtnBuffer.append("^");
						}
					}
					qtnBuffer.append("<>");
				}
			}
			userVO.setSurveyTextList(qtnBuffer.toString());
			//userVO.setSurveyRadioList(getSurveyQtnsWithSubQtns(2));
			//userVO.setSurveyDropdownList(getSurveyQtnsWithSubQtns(3));
			//userVO.setSurveyCheckboxList(getSurveyQtnsWithSubQtns(4));
		} catch (DAOException ex) {
			logger.error("UNABLE TO FETCH THE SURVEY QUESTIONS: "+ex.getLocalizedMessage());
		}
		logger.info("<<<<<<<< AuthenticatorServiceImpl.getSurveyQuestions");
		return userVO;
	}
	
	
	private String populateTextSurveyQtn () {
		return null;
	}
	
	
	private String getSurveyQtnsWithSubQtns (int answerType) {
		logger.info(">>>>>>>> AuthenticatorServiceImpl.getSurveyQtnsWithSubQtns");
		
		StringBuilder txtBuilder = new StringBuilder();
		try {
			List<Object[]> dropdownSurveyQtns = authenticatorDAO.getSurveyMainQuestion(answerType);		
				
			for (int index = 0; index <dropdownSurveyQtns.size(); index++) {	
				Object[] obj = (Object[]) dropdownSurveyQtns.get(index);				
				ArrayList<Object> newObj = new ArrayList<Object>(Arrays.asList(obj));
				
				String mainQtn = (String) newObj.get(0);
				txtBuilder.append(mainQtn);					//	main ques
				txtBuilder.append("#");				
				txtBuilder.append((String) newObj.get(1));	//	mandatory field					
				txtBuilder.append("~");				
				
				List<String> subQtnDtoList = authenticatorDAO.getAllSubQtns(answerType, mainQtn);
				int loopCounter = 0;
				for (int i=0; i<subQtnDtoList.size(); i++) {
					String subQtn = (String) subQtnDtoList.get(i);
					loopCounter = loopCounter + 1;
					txtBuilder.append(subQtn);
					if (loopCounter < subQtnDtoList.size()) {
						txtBuilder.append("^");
					}
				}
				txtBuilder.append("`");
			}
		} catch (DAOException ex) {
			logger.error("UNABLE TO FETCH THE SURVEY QUESTIONS: "+ex.getLocalizedMessage());
		}
		logger.info("<<<<<<<< AuthenticatorServiceImpl.getSurveyQtnsWithSubQtns");
		return txtBuilder.toString();
	}

	@Transactional(propagation = Propagation.REQUIRED)
	public int getDaysToExpire(int userId) throws JCTException {
		int daysToExpire = 0;
		try {
			daysToExpire = authenticatorDAO.getDaysToExpire(userId);
		} catch (DAOException ex) {
			logger.error("UNABLE TO FETCH DAYS TO EXPIRE: "+ex.getLocalizedMessage());
		}
		return daysToExpire;
	}
	
	@Transactional(propagation = Propagation.REQUIRED)
	public List<OccupationListDTO> searchOccupationList(String searchString)
			throws JCTException {
		StringBuffer sb = new StringBuffer("");
		List<OccupationListDTO> occListByCode = new ArrayList<OccupationListDTO>();		

		String[] items = searchString.split(",");
		int itemCount = items.length;
		   
		try {
			//occListByCode = authenticatorDAO.searchOccupationList(searchString);
			List l = authenticatorDAO.searchOccupationListNew(searchString);
			
			List occupationListTemp = new ArrayList();			
			for (int i=0; i<l.size(); i++) {
				int finalCountTemp = 0;
				int finalCount = 0;
				Object[] obj = (Object[]) l.get(i);				
				if (searchString.contains(",")) {
					String[] tok = searchString.split(",");
					for (int index=0; index < tok.length; index++) {								
						int count = StringUtils.countMatches(((String) obj[1]).toLowerCase()+ " "+((String) obj[2]).toLowerCase(), ((String) tok[index].trim()).toLowerCase());
						if(count > 0) {
							finalCountTemp = finalCountTemp + 1;
						}
					}
				}
				finalCount = itemCount - finalCountTemp;
				occupationListTemp.add(new PetComparator(finalCount, (String) obj[1]));
			}
			
			Collections.sort(occupationListTemp, new PetComparator());
			for(Iterator itr = occupationListTemp.iterator() ; itr.hasNext();){
				PetComparator elemenet = (PetComparator)itr.next();
				List<OccupationListDTO> resultDto = authenticatorDAO.getOccupationListDto(elemenet.toString());
				OccupationListDTO dto = new OccupationListDTO(resultDto.get(0).getCode(), resultDto.get(0).getTitle(), resultDto.get(0).getDesc());			
				occListByCode.add(dto);
			}
		} catch (DAOException ex) {
			logger.error("UNABLE TO FETCH ONET OCCUPATION LIST: "+ex.getLocalizedMessage());
		}
		return occListByCode;
	}
	@Transactional(propagation = Propagation.REQUIRED)
	public int getUserId(int row) throws JCTException {
		int userId = 0;
		try {
			userId = authenticatorDAO.getUserId(row);
		} catch (DAOException ex) {
			logger.error("UNABLE TO FETCH USER ID: "+ex.getLocalizedMessage());
		}
		return userId;
	}
	
	@Transactional(propagation = Propagation.REQUIRED)
	public PopulateInstructionVO populatePopUp(int profileId, String relatedPage) 
			throws JCTException {	
		logger.info(">>>> AuthenticatorServiceImpl.populateInstruction");
		PopulateInstructionVO vo = new PopulateInstructionVO();
		String status = "";
		try {
			status = authenticatorDAO.checkInstrList(profileId, relatedPage);
			if(status.equalsIgnoreCase("success")){
			JctPopupInstruction instruction = authenticatorDAO.populatePopUp (profileId,relatedPage);
			if (null != instruction) {
				vo.setPopupType(instruction.getJctPopupInstructionType());
				vo.setStatusCode(Integer.parseInt(CommonConstants.SUCCESSFUL));
				/*if(instruction.getJctPopupInstructionType().equalsIgnoreCase("TEXT")){
					vo.setTextOrVideoDesc(instruction.getJctPopupInstructionTextIns());
				}
				else if(instruction.getJctPopupInstructionType().equalsIgnoreCase("VIDEO")){
					vo.setTextOrVideoDesc(instruction.getJctPopupInstructionVideoName());
				} else if(instruction.getJctPopupInstructionType().equalsIgnoreCase("TEXTANDVIDEO")){
					vo.setOnlyTextDesc(instruction.getJctPopupInstructionTextIns());
					vo.setOnlyVideoDesc(instruction.getJctPopupInstructionVideoName());
				}*/
				vo.setInstructionTextBeforeVideo(instruction.getJctPopupInstructionTextBeforeVideo());
				vo.setInstructionTextAfterVideo(instruction.getJctPopupInstructionTextAfterVideo());
				vo.setInstructionVideo(instruction.getJctPopupInstructionVideoName());			
				}				
			}
			else if(status.equalsIgnoreCase("failure")){
				vo.setStatusCode(318);
			}
		} catch (DAOException e) {
			//e.printStackTrace();
			logger.error(e.getLocalizedMessage());
		}
		logger.info("<<<< AuthenticatorServiceImpl.populatePopUp");
		return vo;
	}
	
	@Transactional(propagation = Propagation.REQUIRED)
	public String getMailedPassword(UserVO userVo) 
			throws JCTException {
	logger.info(">>>> AuthenticatorServiceImpl.populateInstruction");
	String status = "failure";
	String mailedPwd = "";
	try{
		mailedPwd = authenticatorDAO.getMailedPassword(userVo);
		if(userVo.getMailedPassword().equals(decriptString(mailedPwd))){
			status = "success";
		}else{
			status = "failure";
		}
	}catch(DAOException e){
		logger.error(e.getLocalizedMessage());
	}
	return status;
	}
	
	@Transactional(propagation = Propagation.REQUIRED)
	public String getInitialPassword(UserVO userVo) 
			throws JCTException {
	logger.info(">>>> AuthenticatorServiceImpl.getInitialPassword");
	String status = "failure";
	String initalPwd = "";
	try{
		initalPwd = authenticatorDAO.getInitialPassword(userVo);
		if(userVo.getInitialPassword().equals(initalPwd)){
			status = "success";
		}else{
			status = "failure";
		}
	}catch(DAOException e){
		logger.error(e.getLocalizedMessage());
	}
	return status;
	}
	
	private String decriptString(String string){	  
		String ref = "P!QR#S$T%U&VW(X)Y*Z+[,-].^/_0`1a2b3c4d5e6f7g8h9i:j;k<l=m>n?o@pAqBrCsDtEuFvGwHxIyJzK{L|M}N~O";
        StringBuilder result = new StringBuilder("");
        for(int i=0;i<string.length()/2;i++){
                        String charStr     = string.substring (i, i+1);
                        int num = ref.indexOf(charStr);                                              
                        String original = ref.substring(num-1, num);
                        result.append(original);                                
        }
        return result.toString();
	}

	@Transactional(propagation = Propagation.REQUIRED)
	public String getOccupationData(String emailId) throws JCTException {
		StringBuffer sb = new StringBuffer();
		try {
			List<DemographicDTO> demographicDTO = (List<DemographicDTO>) authenticatorDAO.getDemographicInformation(emailId);
			Iterator<DemographicDTO> itr = demographicDTO.iterator();
			while ( itr.hasNext() ) {
				DemographicDTO dto = (DemographicDTO) itr.next();
				sb.append(sketchSearchDao.getOccupationDesc(dto.getOnetOccupationData()));				
			}
		} catch (DAOException ex) {
			logger.error(ex.getLocalizedMessage());
			throw new JCTException();
		}
		return sb.toString();
	}	
	
	@Transactional(propagation = Propagation.REQUIRED)
	public String getTermsAndConditions(int userProfileId, int userType) 
			throws JCTException{
		logger.info(">>>> AuthenticatorServiceImpl.getTermsAndConditions");
		String tnC = null;
		JctUserProfile userProfile = new JctUserProfile();	
		try {
			userProfile = authenticatorDAO.getUserProfileByPk(userProfileId);
			tnC = authenticatorDAO.getTermsAndConditions(userProfile, userType);
		} catch (DAOException e) {
			logger.error(e.getLocalizedMessage());
		}		
		logger.info("<<<< AuthenticatorServiceImpl.getTermsAndConditions");
        return tnC;
	}
	@Transactional(propagation = Propagation.REQUIRED)
	public boolean alreadyRegisteredOnce(String emailId) throws JCTException {
		boolean result = false;
		List<JctSurveyQuestions> qtns = null;
		try {
			qtns = authenticatorDAO.alreadyRegisteredOnce(emailId);
			if (qtns.size() > 0) {
				result = true;
			}
		} catch (DAOException e) {
			logger.error(e.getLocalizedMessage());
		}
		return result;
	}
	
	@Transactional(propagation = Propagation.REQUIRED)
	public String silentUpdateTime(int rowIdentity) throws JCTException {
		logger.info(">>>> AuthenticatorServiceImpl.silentUpdateTime");
		String status = "failure";
		try {
			status = authenticatorDAO.silentUpdateTime(rowIdentity);
		} catch (DAOException e) {
			logger.error(e.getLocalizedMessage());
		}
		logger.info("<<<< AuthenticatorServiceImpl.silentUpdateTime");
		return status;
	}
}
