package com.vmware.jct.service.impl;

import java.util.Arrays;
import java.util.List;
import java.util.ArrayList;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.context.MessageSource;
import org.springframework.stereotype.Service;
import org.springframework.transaction.annotation.Propagation;
import org.springframework.transaction.annotation.Transactional;
import com.vmware.jct.dao.IAuthenticatorDAO;
import com.vmware.jct.dao.ICommonDao;
import com.vmware.jct.dao.dto.FunctionDTO;
import com.vmware.jct.dao.dto.LevelDTO;
import com.vmware.jct.dao.dto.OccupationListDTO;
import com.vmware.jct.dao.dto.UserProfileDTO;
import com.vmware.jct.exception.DAOException;
import com.vmware.jct.exception.JCTException;
import com.vmware.jct.model.JctUser;
import com.vmware.jct.model.JctUserProfile;
import com.vmware.jct.service.IAuthenticatorService;
import com.vmware.jct.service.vo.UserVO;

/**
 * 
 * <p><b>Class name:</b> AuthenticatorServiceImpl.java</p>
 * <p><b>Author:</b> InterraIT</p>
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
 * 	<li>InterraIT - 19/Jan/2014 - Implement Exception </li>
 * <li>InterraIT- 31/Jan/2014 - Implement comments, introduce Named query, Transactional attribute </li>
 * </p>
 */
@Service
public class AuthenticatorServiceImpl implements IAuthenticatorService {
	
	private static final Logger LOGGER = LoggerFactory.getLogger(AuthenticatorServiceImpl.class);

	@Autowired
	private IAuthenticatorDAO authenticatorDAO;
	
	@Autowired
	private ICommonDao commonDaoImpl;
	
	@Autowired  
	private MessageSource messageSource;
	
	@Transactional(propagation=Propagation.REQUIRED)
	public List<JctUser> authenticateUser(UserVO userVO) throws JCTException {	
		LOGGER.info(">>>> AuthenticatorServiceImpl.authenticateUser");
		List<JctUser> jctUser=new ArrayList<JctUser>();
		try{
			jctUser=authenticatorDAO.authenticateUser(userVO);
		}catch(DAOException daoException){
			LOGGER.error(daoException.getLocalizedMessage());
			throw new JCTException(daoException.getMessage());
		}
		LOGGER.info("<<<< AuthenticatorServiceImpl.authenticateUser");
		return jctUser;
	}
	
	@Transactional(propagation=Propagation.REQUIRED)
	public List<JctUser> checkUser(UserVO userVO) throws JCTException {
		LOGGER.info(">>>> AuthenticatorServiceImpl.checkUser");
		List<JctUser> jctUser=new ArrayList<JctUser>();
		try{
			jctUser= authenticatorDAO.checkUser(userVO);
		}catch(DAOException daoException){
			LOGGER.error(daoException.getLocalizedMessage());
			throw new JCTException(daoException.getMessage());
		}
		LOGGER.info("<<<< AuthenticatorServiceImpl.checkUser");
		return jctUser;
	}

	
	
	@Transactional(propagation=Propagation.REQUIRED)
	public String updateLoginInfo(String jobRefNo, int rowId, String lastActivePage) throws JCTException {
		LOGGER.info(">>>> AuthenticatorServiceImpl.updateLoginInfo");
		try{
			LOGGER.info("<<<< AuthenticatorServiceImpl.updateLoginInfo");
			return authenticatorDAO.updateLoginInfo(jobRefNo, rowId, lastActivePage);
		}catch(DAOException daoException ){
			LOGGER.error(daoException.getLocalizedMessage());
			throw new JCTException(daoException.getMessage());
		}
	}

	@Transactional(propagation=Propagation.REQUIRED)
	public List<FunctionDTO> jctFunctionList() throws JCTException {
		LOGGER.info(">>>> AuthenticatorServiceImpl.jctFunctionList");
		List<FunctionDTO> jctFunctionList=null;
		try{
			jctFunctionList= authenticatorDAO.getJctFunctionList();
		}catch(DAOException daoException ){
			LOGGER.error(daoException.getLocalizedMessage());
			throw new JCTException(daoException.getMessage());
		}
		LOGGER.info("<<<< AuthenticatorServiceImpl.jctFunctionList");
		return jctFunctionList;
	}
	
	@Transactional(propagation=Propagation.REQUIRED)
	public List<LevelDTO> jctLevelList() throws JCTException {
		LOGGER.info(">>>> AuthenticatorServiceImpl.jctLevelList");
		List<LevelDTO> jctLevelList=null;
		try{
			jctLevelList= authenticatorDAO.getJctLevelList();
		}catch(DAOException daoException ){
			LOGGER.error(daoException.getLocalizedMessage());
			throw new JCTException(daoException.getMessage());
		}
		LOGGER.info("<<<< AuthenticatorServiceImpl.jctLevelList");
		return jctLevelList;
	}
	
	@Transactional(propagation=Propagation.REQUIRED)
	public List<UserProfileDTO> jctUserProfileList() throws JCTException {
		LOGGER.info(">>>> AuthenticatorServiceImpl.jctUserProfileList");
		List<UserProfileDTO> jctUserPrifileList=null;
		try{
			jctUserPrifileList= authenticatorDAO.getJctUserProfileList();
		}catch(DAOException daoException ){
			LOGGER.error(daoException.getLocalizedMessage());
			throw new JCTException(daoException.getMessage());
		}
		LOGGER.info("<<<< AuthenticatorServiceImpl.jctUserProfileList");
		return jctUserPrifileList;
	}

	/**
	 * Method fetches survey question for general user
	 */
	@Transactional(propagation = Propagation.REQUIRED)
	public UserVO getSurveyQuestions(UserVO userVO) throws JCTException {/*
		LOGGER.info(">>>>>>>> AuthenticatorServiceImpl.getSurveyQuestions");
		try {
			List<Object[]> textSurveyQtns = authenticatorDAO.getTextSurveyQuestion();
			StringBuilder txtBuilder = new StringBuilder();
			int loopCounter = 0;
			for (int index = 0; index <textSurveyQtns.size(); index++) {
				loopCounter = loopCounter + 1;
				// Get the object
				Object[] obj = (Object[]) textSurveyQtns.get(index);
				ArrayList<Object> newObj = new ArrayList<Object>(Arrays.asList(obj));
				
				txtBuilder.append((String) newObj.get(0));	//	ques
				txtBuilder.append("#");
				
				txtBuilder.append((String) newObj.get(1));	//	mandatory field
				if (loopCounter < textSurveyQtns.size()) {
					txtBuilder.append("`");
				}
			}
			userVO.setSurveyTextList(txtBuilder.toString());
			userVO.setSurveyRadioList(getSurveyQtnsWithSubQtns(2));
			userVO.setSurveyDropdownList(getSurveyQtnsWithSubQtns(3));
			userVO.setSurveyCheckboxList(getSurveyQtnsWithSubQtns(4));
		} catch (DAOException ex) {
			LOGGER.error("UNABLE TO FETCH THE SURVEY QUESTIONS: "+ex.getLocalizedMessage());
		}
		LOGGER.info("<<<<<<<< AuthenticatorServiceImpl.getSurveyQuestions");
		return userVO;
	*/

		LOGGER.info(">>>>>>>> AuthenticatorServiceImpl.getSurveyQuestions");
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
			LOGGER.error("UNABLE TO FETCH THE SURVEY QUESTIONS: "+ex.getLocalizedMessage());
		}
		LOGGER.info("<<<<<<<< AuthenticatorServiceImpl.getSurveyQuestions");
		return userVO;
		
	}
	
	private String getSurveyQtnsWithSubQtns (int answerType) {
		LOGGER.info(">>>>>>>> AuthenticatorServiceImpl.getSurveyQtnsWithSubQtns");	
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
			LOGGER.error("UNABLE TO FETCH THE SURVEY QUESTIONS: "+ex.getLocalizedMessage());
		}
		LOGGER.info("<<<<<<<< AuthenticatorServiceImpl.getSurveyQtnsWithSubQtns");
		return txtBuilder.toString();
	}
	
	@Transactional(propagation = Propagation.REQUIRED)
	public List<OccupationListDTO> searchOccupationList(String searchString)
			throws JCTException {
		List<OccupationListDTO> occListByCode = new ArrayList<OccupationListDTO>();
		try {
			occListByCode = authenticatorDAO.searchOccupationList(searchString);
		} catch (DAOException ex) {
			LOGGER.error("UNABLE TO FETCH ONET OCCUPATION LIST: "+ex.getLocalizedMessage());
		}
		return occListByCode;
	}

	@Transactional(propagation=Propagation.REQUIRED)
	public String updateUser(String password, int activeFlag, String email)
			throws JCTException {
		String status = "failure";
		try {
			status = authenticatorDAO.updateUser(password, activeFlag, email);
		} catch (DAOException ex) {
			LOGGER.error("UNABLE TO FETCH ONET OCCUPATION LIST: "+ex.getLocalizedMessage());
		}
		return status;
	}

	@Transactional(propagation=Propagation.REQUIRED)
	public List<JctUser> getUser(String email) throws JCTException {
		List<JctUser> usrList = null;
		try {
			usrList = authenticatorDAO.getUser(email);
		} catch (DAOException ex) {
			LOGGER.error("UNABLE TO FETCH ONET OCCUPATION LIST: "+ex.getLocalizedMessage());
			usrList = new ArrayList<JctUser>();
		}
		return usrList;
	}
	
	@Transactional(propagation = Propagation.REQUIRED)
	public String getTermsAndConditions(int userProfileId, int userType) 
			throws JCTException {
		LOGGER.info(">>>> AuthenticatorServiceImpl.getTermsAndConditions");
		String tnC = null;
		JctUserProfile userProfile = new JctUserProfile();	
		try {
			userProfile = authenticatorDAO.getUserProfileByPk(userProfileId);
			tnC = authenticatorDAO.getTermsAndConditions(userProfile,userType);
		} catch (DAOException e) {
			LOGGER.error(e.getLocalizedMessage());
		}		
		LOGGER.info("<<<< AuthenticatorServiceImpl.getTermsAndConditions");
        return tnC;
	} 
	
}