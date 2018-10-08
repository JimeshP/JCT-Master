package com.vmware.jct.service.impl;

import java.util.Date;
import java.util.Iterator;
import java.util.List;
import java.util.regex.Pattern;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.context.MessageSource;
import org.springframework.stereotype.Service;
import org.springframework.transaction.annotation.Propagation;
import org.springframework.transaction.annotation.Transactional;

import com.vmware.jct.common.utility.CommonConstants;
import com.vmware.jct.common.utility.StatusConstants;
import com.vmware.jct.dao.ISurveyQuestionDAO;
import com.vmware.jct.dao.dto.SurveyQuestionDTO;
import com.vmware.jct.exception.DAOException;
import com.vmware.jct.exception.JCTException;
import com.vmware.jct.model.JctSurveyQuestionMaster;
import com.vmware.jct.model.JctSurveyQuestions;
import com.vmware.jct.model.JctUser;
import com.vmware.jct.service.ISurveyQuestionService;
import com.vmware.jct.service.vo.SurveyQuestionsVO;

/**
 * 
 * <p><b>Class name:</b> SurveyQuestionServiceImpl.java</p>
 * <p><b>Author:</b> InterraIT</p>
 * <p><b>Purpose:</b> This is a SurveyQuestionServiceImpl class. This artifact is Business layer artifact.
 * SurveyQuestionServiceImpl implement ISurveyQuestionService interface and override the following  methods.
 * - saveFreeText(SurveyQuestionsVO surveyQtnsVO)
 * </p>
 * <p><b>Description:</b> This class used to perform insert, modify ,fetch data from db and delete operation. Save, update, delete operations are performed in DataAccessObject class </p>
 * <p><b>Revision History:</b>
 * 	<li>InterraIT - 21/Oct/2014 - Implemented the class</li>
 * </p>
 */
@Service
public class SurveyQuestionServiceImpl implements ISurveyQuestionService {
	
	@Autowired
	private ISurveyQuestionDAO surveyQuestionDAO;
	
	@Autowired  
	private MessageSource messageSource;
	
	private static final Logger LOGGER = LoggerFactory.getLogger(SurveyQuestionServiceImpl.class);
	/**
	 * Method saves free text survey question
	 * @param SurveyQuestionVO
	 * @return SurveyQuestionVO
	 * @throws JCTException
	 */
	@Transactional(propagation = Propagation.REQUIRED)
	public SurveyQuestionsVO saveFreeText(SurveyQuestionsVO surveyQtnsVO)
			throws JCTException {
		LOGGER.info(">>>>>>>> SurveyQuestionServiceImpl.saveFreeText");
		// Break the free text and create the entity
		int failureCount = 0;
		int totalCount = 0;
		int successCount = 0;		               
        String[] parts = surveyQtnsVO.getSurveyTextQuestion().split("##");
        for(String row : parts) {
        	
        	String[] innerParts = row.split("~");
        	String qtn = innerParts[0];
        	String mandatoryChoice = innerParts[1]; 
        	String exists = checkIfSurveyQtnExists(qtn, CommonConstants.FREE_TEXT, surveyQtnsVO.getUserType());
        	totalCount = totalCount + 1;
        	String[] split = surveyQtnsVO.getUserProfile().split("!");
        	if (exists.equals("N")) {
        		JctSurveyQuestionMaster master = new JctSurveyQuestionMaster();
                master.setJctSurveyQtnMasterAnsType(CommonConstants.FREE_TEXT);
                master.setJctSurveyQtnMasterCreatedBy(surveyQtnsVO.getCreatedBy());
                master.setJctSurveyQtnMasterCreatedTs(new Date());
                master.setJctSurveyQtnMasterMainQtn(qtn);
                master.setJctSurveyQtnMasterModifiedBy(surveyQtnsVO.getCreatedBy());
                master.setJctSurveyQtnMasterModifiedTs(new Date());
                master.setJctSurveyQtnMasterSubQtn("NA");
                master.setJctSurveyQtnMasterUserType(surveyQtnsVO.getUserType());
                master.setJctSurveyQtnMasterMandatory(mandatoryChoice);
                
                master.setJctSurveyQtnMasterProfileId(Integer.parseInt(split[0]));
                master.setJctSurveyQtnMasterProfileName(split[1]);              
                //Store in db
                try {
                	master.setJctSurveyQtnMasterOrder(surveyQuestionDAO.generateOrderSurvQtn(Integer.parseInt(split[0]), surveyQtnsVO.getUserType()));
    				surveyQuestionDAO.saveSurveyQuestion(master);
    				successCount = successCount + 1;
    			} catch (DAOException e) {
    				failureCount = failureCount + 1;
    				LOGGER.error("UNABLE TO STORE FREE TEXT SURVEY QUESTION FOR USER: "
    						+surveyQtnsVO.getCreatedBy()+". REASON: "+e.getLocalizedMessage());
    			}
        	} else {
        		failureCount = failureCount + 1;
        	}
        }
        //Create the message
		if (failureCount == 0) { // TOTAL SUCCESS
			surveyQtnsVO.setMessage(this.messageSource.getMessage("msg.survey.qtns.text.header.total.success", null, null));
		} else if (successCount == 0) { // TOTAL FAILURE
			surveyQtnsVO.setMessage(this.messageSource.getMessage("msg.survey.qtns.text.header.total.success", null, null));
		} else {
			StringBuilder sb = new StringBuilder(this.messageSource.getMessage("msg.survey.qtns.text.header.partial.one", null, null));
			sb.append(totalCount);
			sb.append(this.messageSource.getMessage("msg.survey.qtns.text.header.partial.two", null, null));
			sb.append(successCount);
			sb.append(this.messageSource.getMessage("msg.survey.qtns.text.header.partial.three", null, null));
			surveyQtnsVO.setMessage(sb.toString());
		}
		LOGGER.info("<<<<<<<< SurveyQuestionServiceImpl.saveFreeText");
		return surveyQtnsVO;
	}
	/**
	 * Method saves radio survey question
	 * @param SurveyQuestionVO
	 * @return SurveyQuestionVO
	 * @throws JCTException
	 */
	@Transactional(propagation = Propagation.REQUIRED)
	public SurveyQuestionsVO saveRadio(SurveyQuestionsVO surveyQtnsVO)
			throws JCTException {
		LOGGER.info(">>>>>>>> SurveyQuestionServiceImpl.saveRadio");
		int failureCount = 0;
		int totalCount = 0;
		int successCount = 0;
		int qtnOrder = 0;
		// Break the sub question text and create the entities
		String exists = checkIfSurveyQtnExists(surveyQtnsVO.getMainQuestion(), CommonConstants.RADIO_BTN, surveyQtnsVO.getUserType());
		if (exists.equals("N")) {
			final String regularExp = "`~!~`";
			Pattern p = Pattern.compile(regularExp);
			String[] items = p.split(surveyQtnsVO.getSubQuestionsString());
			String[] split = surveyQtnsVO.getUserProfile().split("!");
			for(String subQtn : items) {
				totalCount = totalCount + 1;
				JctSurveyQuestionMaster master = new JctSurveyQuestionMaster();
		        master.setJctSurveyQtnMasterAnsType(CommonConstants.RADIO_BTN);
		        master.setJctSurveyQtnMasterCreatedBy(surveyQtnsVO.getCreatedBy());
		        master.setJctSurveyQtnMasterCreatedTs(new Date());
		        master.setJctSurveyQtnMasterMainQtn(surveyQtnsVO.getMainQuestion());
		        master.setJctSurveyQtnMasterModifiedBy(surveyQtnsVO.getCreatedBy());
		        master.setJctSurveyQtnMasterModifiedTs(new Date());
		        master.setJctSurveyQtnMasterSubQtn(subQtn);
		        master.setJctSurveyQtnMasterUserType(surveyQtnsVO.getUserType());
		        master.setJctSurveyQtnMasterMandatory(surveyQtnsVO.getIsMandatory());
		        
                master.setJctSurveyQtnMasterProfileId(Integer.parseInt(split[0]));
                master.setJctSurveyQtnMasterProfileName(split[1]);
		        //Store in db
		        try {
		        	if (totalCount == 1) {
		        		qtnOrder = surveyQuestionDAO.generateOrderSurvQtn(Integer.parseInt(split[0]), surveyQtnsVO.getUserType());
		        	} 		        	
		        	master.setJctSurveyQtnMasterOrder(qtnOrder);
		        	surveyQuestionDAO.saveSurveyQuestion(master);
					successCount = successCount + 1;
		        } catch (DAOException e) {
					failureCount = failureCount + 1;
					LOGGER.error("UNABLE TO STORE RADIO SURVEY QUESTION FOR USER: "
							+surveyQtnsVO.getCreatedBy()+". REASON: "+e.getLocalizedMessage());
				}
			}
			//Create the message
			if (failureCount == 0) { // TOTAL SUCCESS
				surveyQtnsVO.setMessage(this.messageSource.getMessage("msg.survey.qtns.text.header.total.success", null, null));
			} else if (successCount == 0) { // TOTAL FAILURE
				surveyQtnsVO.setMessage(this.messageSource.getMessage("msg.survey.qtns.text.header.total.success", null, null));
			} else {
				StringBuilder sb = new StringBuilder(this.messageSource.getMessage("msg.survey.qtns.text.header.partial.one", null, null));
				sb.append(totalCount);
				sb.append(this.messageSource.getMessage("msg.survey.qtns.text.header.partial.two", null, null));
				sb.append(successCount);
				sb.append(this.messageSource.getMessage("msg.survey.qtns.text.header.partial.three", null, null));
				surveyQtnsVO.setMessage(sb.toString());
			}
		} else {
			StringBuilder sb = new StringBuilder(this.messageSource.getMessage("msg.survey.qtns.already.exist.one", null, null));
			sb.append(surveyQtnsVO.getMainQuestion());
			sb.append(this.messageSource.getMessage("msg.survey.qtns.already.exist.two", null, null));
			surveyQtnsVO.setMessage(sb.toString());
		}
		LOGGER.info("<<<<<<<< SurveyQuestionServiceImpl.saveRadio");
		return surveyQtnsVO;
	}
	/**
	 * Method saves drop-down survey question
	 * @param SurveyQuestionVO
	 * @return SurveyQuestionVO
	 * @throws JCTException
	 */
	@Transactional(propagation = Propagation.REQUIRED)
	public SurveyQuestionsVO saveDropDown(SurveyQuestionsVO surveyQtnsVO)
			throws JCTException {
		LOGGER.info(">>>>>>>> SurveyQuestionServiceImpl.saveDropDown");
		int failureCount = 0;
		int totalCount = 0;
		int successCount = 0;
		int qtnOrder = 0;
		String exists = checkIfSurveyQtnExists(surveyQtnsVO.getMainQuestion(), CommonConstants.DROP_DOWN, surveyQtnsVO.getUserType());
		if (exists.equals("N")) {
			// Break the sub question text and create the entities
			final String regularExp = "`~!~`";
			Pattern p = Pattern.compile(regularExp);
			String[] items = p.split(surveyQtnsVO.getSubQuestionsString());
			String[] split = surveyQtnsVO.getUserProfile().split("!");
			for(String subQtn : items) {
				totalCount = totalCount + 1;
				JctSurveyQuestionMaster master = new JctSurveyQuestionMaster();
		        master.setJctSurveyQtnMasterAnsType(CommonConstants.DROP_DOWN);
		        master.setJctSurveyQtnMasterCreatedBy(surveyQtnsVO.getCreatedBy());
		        master.setJctSurveyQtnMasterCreatedTs(new Date());
		        master.setJctSurveyQtnMasterMainQtn(surveyQtnsVO.getMainQuestion());
		        master.setJctSurveyQtnMasterModifiedBy(surveyQtnsVO.getCreatedBy());
		        master.setJctSurveyQtnMasterModifiedTs(new Date());
		        master.setJctSurveyQtnMasterSubQtn(subQtn);
		        master.setJctSurveyQtnMasterUserType(surveyQtnsVO.getUserType());
		        master.setJctSurveyQtnMasterMandatory(surveyQtnsVO.getIsMandatory());
		        
                master.setJctSurveyQtnMasterProfileId(Integer.parseInt(split[0]));
                master.setJctSurveyQtnMasterProfileName(split[1]);
		        //Store in db
		        try {
		        	//master.setJctSurveyQtnMasterOrder(surveyQuestionDAO.generateOrderSurvQtn(Integer.parseInt(split[0]), surveyQtnsVO.getUserType()));
		        	if (totalCount == 1) {
		        		qtnOrder = surveyQuestionDAO.generateOrderSurvQtn(Integer.parseInt(split[0]), surveyQtnsVO.getUserType());
		        	} 
		        	master.setJctSurveyQtnMasterOrder(qtnOrder);
		        	surveyQuestionDAO.saveSurveyQuestion(master);
					successCount = successCount + 1;
		        } catch (DAOException e) {
					failureCount = failureCount + 1;
					LOGGER.error("UNABLE TO STORE DROP DOWN SURVEY QUESTION FOR USER: "
							+surveyQtnsVO.getCreatedBy()+". REASON: "+e.getLocalizedMessage());
				}
			}	
			//Create the message
			if (failureCount == 0) { // TOTAL SUCCESS
				surveyQtnsVO.setMessage(this.messageSource.getMessage("msg.survey.qtns.text.header.total.success", null, null));
			} else if (successCount == 0) { // TOTAL FAILURE
				surveyQtnsVO.setMessage(this.messageSource.getMessage("msg.survey.qtns.text.header.total.success", null, null));
			} else {
				StringBuilder sb = new StringBuilder(this.messageSource.getMessage("msg.survey.qtns.text.header.partial.one", null, null));
				sb.append(totalCount);
				sb.append(this.messageSource.getMessage("msg.survey.qtns.text.header.partial.two", null, null));
				sb.append(successCount);
				sb.append(this.messageSource.getMessage("msg.survey.qtns.text.header.partial.three", null, null));
				surveyQtnsVO.setMessage(sb.toString());
			}
		} else {
			StringBuilder sb = new StringBuilder(this.messageSource.getMessage("msg.survey.qtns.already.exist.one", null, null));
			sb.append(surveyQtnsVO.getMainQuestion());
			sb.append(this.messageSource.getMessage("msg.survey.qtns.already.exist.two", null, null));
			surveyQtnsVO.setMessage(sb.toString());
		}
		LOGGER.info("<<<<<<<< SurveyQuestionServiceImpl.saveDropDown");
		return surveyQtnsVO;
	}
	/**
	 * Method saves check box survey question
	 * @param SurveyQuestionVO
	 * @return SurveyQuestionVO
	 * @throws JCTException
	 */
	@Transactional(propagation = Propagation.REQUIRED)
	public SurveyQuestionsVO saveCheckBox(SurveyQuestionsVO surveyQtnsVO)
			throws JCTException {
		LOGGER.info(">>>>>>>> SurveyQuestionServiceImpl.saveCheckBox");
		int failureCount = 0;
		int totalCount = 0;
		int successCount = 0;
		int qtnOrder = 0;
		String exists = checkIfSurveyQtnExists(surveyQtnsVO.getMainQuestion(), CommonConstants.CHECK_BOX, surveyQtnsVO.getUserType());
		if (exists.equals("N")) {
			// Break the sub question text and create the entities
			final String regularExp = "`~!~`";
			Pattern p = Pattern.compile(regularExp);
			String[] items = p.split(surveyQtnsVO.getSubQuestionsString());
			String[] split = surveyQtnsVO.getUserProfile().split("!");
			for(String subQtn : items) {
				totalCount = totalCount + 1;
				JctSurveyQuestionMaster master = new JctSurveyQuestionMaster();
		        master.setJctSurveyQtnMasterAnsType(CommonConstants.CHECK_BOX);
		        master.setJctSurveyQtnMasterCreatedBy(surveyQtnsVO.getCreatedBy());
		        master.setJctSurveyQtnMasterCreatedTs(new Date());
		        master.setJctSurveyQtnMasterMainQtn(surveyQtnsVO.getMainQuestion());
		        master.setJctSurveyQtnMasterModifiedBy(surveyQtnsVO.getCreatedBy());
		        master.setJctSurveyQtnMasterModifiedTs(new Date());
		        master.setJctSurveyQtnMasterSubQtn(subQtn);
		        master.setJctSurveyQtnMasterUserType(surveyQtnsVO.getUserType());
		        master.setJctSurveyQtnMasterMandatory(surveyQtnsVO.getIsMandatory());
		        
		        master.setJctSurveyQtnMasterProfileId(Integer.parseInt(split[0]));
                master.setJctSurveyQtnMasterProfileName(split[1]);
		        //Store in db
		        try {
		        	if (totalCount == 1) {
		        		qtnOrder = surveyQuestionDAO.generateOrderSurvQtn(Integer.parseInt(split[0]), surveyQtnsVO.getUserType());
		        	} 
		        	master.setJctSurveyQtnMasterOrder(qtnOrder);
		        	//master.setJctSurveyQtnMasterOrder(surveyQuestionDAO.generateOrderSurvQtn(Integer.parseInt(split[0]), surveyQtnsVO.getUserType()));
		        	surveyQuestionDAO.saveSurveyQuestion(master);
					successCount = successCount + 1;
		        } catch (DAOException e) {
					failureCount = failureCount + 1;
					LOGGER.error("UNABLE TO STORE CHECK BOX SURVEY QUESTION FOR USER: "
							+surveyQtnsVO.getCreatedBy()+". REASON: "+e.getLocalizedMessage());
				}
			}	
			//Create the message
			if (failureCount == 0) { // TOTAL SUCCESS
				surveyQtnsVO.setMessage(this.messageSource.getMessage("msg.survey.qtns.text.header.total.success", null, null));
			} else if (successCount == 0) { // TOTAL FAILURE
				surveyQtnsVO.setMessage(this.messageSource.getMessage("msg.survey.qtns.text.header.total.success", null, null));
			} else {
				StringBuilder sb = new StringBuilder(this.messageSource.getMessage("msg.survey.qtns.text.header.partial.one", null, null));
				sb.append(totalCount);
				sb.append(this.messageSource.getMessage("msg.survey.qtns.text.header.partial.two", null, null));
				sb.append(successCount);
				sb.append(this.messageSource.getMessage("msg.survey.qtns.text.header.partial.three", null, null));
				surveyQtnsVO.setMessage(sb.toString());
			}
		} else {
			StringBuilder sb = new StringBuilder(this.messageSource.getMessage("msg.survey.qtns.already.exist.one", null, null));
			sb.append(surveyQtnsVO.getMainQuestion());
			sb.append(this.messageSource.getMessage("msg.survey.qtns.already.exist.two", null, null));
			surveyQtnsVO.setMessage(sb.toString());
		}
		LOGGER.info("<<<<<<<< SurveyQuestionServiceImpl.saveCheckBox");
		return surveyQtnsVO;
	}
	/**
	 * Method fetches all existing survey questions
	 * @return String
	 * @throws JCTException
	 */
	@Transactional(propagation = Propagation.REQUIRED)
	public String getExistingAllSurveyQuestion() throws JCTException {
		StringBuilder questionStringBuilder = new StringBuilder("");
		try {
			List<SurveyQuestionDTO> mainQtnDtoList = surveyQuestionDAO.getExistingAllSurveyQuestion();
			Iterator<SurveyQuestionDTO> mainQtnItr = mainQtnDtoList.iterator();
			while (mainQtnItr.hasNext()) {
				SurveyQuestionDTO dto = (SurveyQuestionDTO) mainQtnItr.next();
				questionStringBuilder.append(dto.getAnswerType());
				questionStringBuilder.append("`");
				questionStringBuilder.append(dto.getJctUserType());
				questionStringBuilder.append("`");
				questionStringBuilder.append(dto.getMainQuestion());
				questionStringBuilder.append("`");
				List<SurveyQuestionDTO> subQtnDtoList = surveyQuestionDAO.getAllSubQtns(dto.getAnswerType(), dto.getJctUserType(), dto.getMainQuestion());
				int loopCounter = 0;
				for (int index=0; index<subQtnDtoList.size(); index++) {
					SurveyQuestionDTO subQtnDto = (SurveyQuestionDTO) subQtnDtoList.get(index);
					loopCounter = loopCounter + 1;
					questionStringBuilder.append(subQtnDto.getSubQtns());
					if (loopCounter < subQtnDtoList.size()) {
						questionStringBuilder.append("^");
					}
				}
				questionStringBuilder.append("`");
				// Get the number of sub question
				questionStringBuilder.append(surveyQuestionDAO.getNumberOfSubQuestions(dto.getAnswerType(), dto.getJctUserType(), dto.getMainQuestion()));
				// String user type
				questionStringBuilder.append("`");
				questionStringBuilder.append(this.messageSource.getMessage("msg.user.type."+dto.getJctUserType(), null, null));
				// String user type
				questionStringBuilder.append("`");
				questionStringBuilder.append(this.messageSource.getMessage("msg.answer.type."+dto.getAnswerType(), null, null));
				// String mandatory
				questionStringBuilder.append("`");
				questionStringBuilder.append(dto.getIsMandatory());
				// User Profile
				questionStringBuilder.append("`");
				questionStringBuilder.append(dto.getProfileName());
				
				questionStringBuilder.append("~");
			}
		} catch (DAOException dao) {
			LOGGER.error("UNABLE TO FETCH EXISTING SURVEY QUESTIONS: "+dao.getLocalizedMessage());
		}
		return questionStringBuilder.toString();
	}
	/**
	 * Method fetches all existing survey questions by user type
	 * @param userType
	 * @return String
	 * @throws JCTException
	 */
	@Transactional(propagation = Propagation.REQUIRED)
	public String fetchAllExistingSurveyQtnsByUserType(int userType)
			throws JCTException {
		StringBuilder questionStringBuilder = new StringBuilder("");
		try {
			List<SurveyQuestionDTO> mainQtnDtoList = surveyQuestionDAO.fetchAllExistingSurveyQtnsByUserType(userType);
			Iterator<SurveyQuestionDTO> mainQtnItr = mainQtnDtoList.iterator();
			while (mainQtnItr.hasNext()) {
				SurveyQuestionDTO dto = (SurveyQuestionDTO) mainQtnItr.next();
				questionStringBuilder.append(dto.getAnswerType());
				questionStringBuilder.append("`");
				questionStringBuilder.append(dto.getJctUserType());
				questionStringBuilder.append("`");
				questionStringBuilder.append(dto.getMainQuestion());
				questionStringBuilder.append("`");
				List<SurveyQuestionDTO> subQtnDtoList = surveyQuestionDAO.getAllSubQtns(dto.getAnswerType(), userType, dto.getMainQuestion());
				int loopCounter = 0;
				for (int index=0; index<subQtnDtoList.size(); index++) {
					SurveyQuestionDTO subQtnDto = (SurveyQuestionDTO) subQtnDtoList.get(index);
					loopCounter = loopCounter + 1;
					questionStringBuilder.append(subQtnDto.getSubQtns());
					if (loopCounter < subQtnDtoList.size()) {
						questionStringBuilder.append("^");
					}
				}
				questionStringBuilder.append("`");
				// Get the number of sub question
				questionStringBuilder.append(surveyQuestionDAO.getNumberOfSubQuestions(dto.getAnswerType(), userType, dto.getMainQuestion()));
				// String user type
				questionStringBuilder.append("`");
				questionStringBuilder.append(this.messageSource.getMessage("msg.user.type."+userType, null, null));
				// String user type
				questionStringBuilder.append("`");
				questionStringBuilder.append(this.messageSource.getMessage("msg.answer.type."+dto.getAnswerType(), null, null));
				//  String mandatory
				questionStringBuilder.append("`");
				questionStringBuilder.append(dto.getIsMandatory());
				// User profile
				questionStringBuilder.append("`");
				questionStringBuilder.append(dto.getProfileName());
				questionStringBuilder.append("~");
			}
		} catch (DAOException dao) {
			LOGGER.error("UNABLE TO FETCH EXISTING SURVEY QUESTIONS BY USER TYPE: "+dao.getLocalizedMessage());
		}
		return questionStringBuilder.toString();
	}
	/**
	 * Method fetches all existing survey questions by user type and question type
	 * @param userType
	 * @param qtnType
	 * @return String
	 * @throws JCTException
	 */
	@Transactional(propagation = Propagation.REQUIRED)
	public String fetchAllExistingSurveyQtnsByUserTypeAndQtnType(int userType,
			int qtnType) throws JCTException {
		StringBuilder questionStringBuilder = new StringBuilder("");
		try {
			List<SurveyQuestionDTO> mainQtnDtoList = surveyQuestionDAO.fetchAllExistingSurveyQtnsByUserTypeAndQtnType(userType, qtnType);
			Iterator<SurveyQuestionDTO> mainQtnItr = mainQtnDtoList.iterator();
			while (mainQtnItr.hasNext()) {
				SurveyQuestionDTO dto = (SurveyQuestionDTO) mainQtnItr.next();
				questionStringBuilder.append(dto.getAnswerType());
				questionStringBuilder.append("`");
				questionStringBuilder.append(dto.getJctUserType());
				questionStringBuilder.append("`");
				questionStringBuilder.append(dto.getMainQuestion());
				questionStringBuilder.append("`");
				List<SurveyQuestionDTO> subQtnDtoList = surveyQuestionDAO.getAllSubQtns(qtnType, userType, dto.getMainQuestion());
				int loopCounter = 0;
				for (int index=0; index<subQtnDtoList.size(); index++) {
					SurveyQuestionDTO subQtnDto = (SurveyQuestionDTO) subQtnDtoList.get(index);
					loopCounter = loopCounter + 1;
					questionStringBuilder.append(subQtnDto.getSubQtns());
					if (loopCounter < subQtnDtoList.size()) {
						questionStringBuilder.append("^");
					}
				}
				questionStringBuilder.append("`");
				// Get the number of sub question
				questionStringBuilder.append(surveyQuestionDAO.getNumberOfSubQuestions(qtnType, userType, dto.getMainQuestion()));
				// String user type
				questionStringBuilder.append("`");
				questionStringBuilder.append(this.messageSource.getMessage("msg.user.type."+userType, null, null));
				// String user type
				questionStringBuilder.append("`");
				questionStringBuilder.append(this.messageSource.getMessage("msg.answer.type."+qtnType, null, null));
				// String mandatory
				questionStringBuilder.append("`");
				questionStringBuilder.append(dto.getIsMandatory());
				// user profile
				questionStringBuilder.append("`");
				questionStringBuilder.append(dto.getProfileName());
				
				questionStringBuilder.append("~");
			}
		} catch (DAOException dao) {
			LOGGER.error("UNABLE TO FETCH EXISTING SURVEY QUESTIONS BY USER TYPE AND QUESTION TYPE: "+dao.getLocalizedMessage());
		}
		return questionStringBuilder.toString();
	}
	/**
	 * Method fetches all existing survey questions by user type and question type
	 * @param userType
	 * @param qtnType
	 * @return String
	 * @throws JCTException
	 */
	@Transactional(propagation = Propagation.REQUIRED)
	public String fetchAllExistingSurveyQtnsByQtnType(int qtnType) throws JCTException {
		StringBuilder questionStringBuilder = new StringBuilder("");
		try {
			List<SurveyQuestionDTO> mainQtnDtoList = surveyQuestionDAO.fetchAllExistingSurveyQtnsByQtnType(qtnType);
			Iterator<SurveyQuestionDTO> mainQtnItr = mainQtnDtoList.iterator();
			while (mainQtnItr.hasNext()) {
				SurveyQuestionDTO dto = (SurveyQuestionDTO) mainQtnItr.next();
				questionStringBuilder.append(dto.getAnswerType());
				questionStringBuilder.append("`");
				questionStringBuilder.append(dto.getJctUserType());
				questionStringBuilder.append("`");
				questionStringBuilder.append(dto.getMainQuestion());
				questionStringBuilder.append("`");
				List<SurveyQuestionDTO> subQtnDtoList = surveyQuestionDAO.getAllSubQtnsByAnsType(qtnType, dto.getMainQuestion());
				int loopCounter = 0;
				for (int index=0; index<subQtnDtoList.size(); index++) {
					SurveyQuestionDTO subQtnDto = (SurveyQuestionDTO) subQtnDtoList.get(index);
					loopCounter = loopCounter + 1;
					questionStringBuilder.append(subQtnDto.getSubQtns());
					if (loopCounter < subQtnDtoList.size()) {
						questionStringBuilder.append("^");
					}
				}
				questionStringBuilder.append("`");
				// Get the number of sub question
				questionStringBuilder.append(surveyQuestionDAO.getNumberOfSubQuestionsByAnsType(qtnType, dto.getMainQuestion()));
				// String user type
				questionStringBuilder.append("`");
				questionStringBuilder.append(this.messageSource.getMessage("msg.user.type."+dto.getJctUserType(), null, null));
				// String user type
				questionStringBuilder.append("`");
				questionStringBuilder.append(this.messageSource.getMessage("msg.answer.type."+qtnType, null, null));
				questionStringBuilder.append("`");
				questionStringBuilder.append(dto.getIsMandatory());
				// user profile
				questionStringBuilder.append("`");
				questionStringBuilder.append(dto.getProfileName());
				questionStringBuilder.append("~");
			}
		} catch (DAOException dao) {
			LOGGER.error("UNABLE TO FETCH EXISTING SURVEY QUESTIONS BY USER TYPE AND QUESTION TYPE: "+dao.getLocalizedMessage());
		}
		return questionStringBuilder.toString();
	}
	/**
	 * Method soft deletes survey questions
	 * @param userType
	 * @param answerType
	 * @param qtnDescription
	 * @return String
	 * @throws JCTException
	 */
	@Transactional(propagation = Propagation.REQUIRED)
	public String deleteSurveyQtn(int userType, int answerType,
			String qtnDescription) throws JCTException {
		String result = "success";
		try {
			int status = surveyQuestionDAO.deleteSurveyQtn(userType, answerType, qtnDescription);
			LOGGER.info("DELETION RESULT: "+status);
		} catch (DAOException ex) {
			result = "failure";
			LOGGER.error("UNABLE TO SOFT DELETE THE SURVEY QUESTION. "+ex.getLocalizedMessage());
		}
		return result;
	}
	/**
	 * Method updates existing text survey questions
	 * @param userType
	 * @param answerType
	 * @param originalQuestion
	 * @param updatedQuestion
	 * @return String
	 * @throws JCTException
	 */
	@Transactional(propagation = Propagation.REQUIRED)
	public String updateTextSurveyQtn(int userType, int answerType,
			String originalQuestion, String updatedQuestion,
			String isMandatory)throws JCTException {
		String result = "success";
		try {
			List<JctSurveyQuestionMaster> masterList = surveyQuestionDAO.getSurveyQuestionList(userType, answerType, originalQuestion);
			Iterator<JctSurveyQuestionMaster> masterItr = masterList.iterator();
			while (masterItr.hasNext()) {
				JctSurveyQuestionMaster master = (JctSurveyQuestionMaster) masterItr.next();
				JctSurveyQuestionMaster copy = new JctSurveyQuestionMaster();
				copy.setJctSurveyQtnMasterAnsType(CommonConstants.FREE_TEXT);
				copy.setJctSurveyQtnMasterCreatedBy(master.getJctSurveyQtnMasterCreatedBy());
				copy.setJctSurveyQtnMasterCreatedTs(master.getJctSurveyQtnMasterCreatedTs());
				copy.setJctSurveyQtnMasterMainQtn(updatedQuestion);
				copy.setJctSurveyQtnMasterModifiedBy(master.getJctSurveyQtnMasterModifiedBy());
				copy.setJctSurveyQtnMasterModifiedTs(new Date());
				copy.setJctSurveyQtnMasterSubQtn(master.getJctSurveyQtnMasterSubQtn());
				copy.setJctSurveyQtnMasterUserType(userType);
				copy.setJctSurveyQtnMasterMandatory(isMandatory);
				copy.setJctSurveyQtnMasterProfileId(master.getJctSurveyQtnMasterProfileId());
				copy.setJctSurveyQtnMasterProfileName(master.getJctSurveyQtnMasterProfileName());
				copy.setJctSurveyQtnMasterOrder(master.getJctSurveyQtnMasterOrder());
				// Merge the old
				master.setJctSurveyQtnMasterSoftDel(CommonConstants.ENTRY_SOFT_DELETED);
				surveyQuestionDAO.updateSurveyQuestion(master);
				// Save the new
				surveyQuestionDAO.saveSurveyQuestion(copy);
			}
		} catch (DAOException exception) {
			result = "failure";
			LOGGER.error("UNABLE TO UPDATE FREE TEXT. "+exception.getLocalizedMessage());
		}
		return result;
	}
	/**
	 * Method updates existing radio survey questions
	 * @param userType
	 * @param answerType
	 * @param originalQuestion
	 * @param updatedMainQuestion
	 * @param subQuestion
	 * @param createdBy
	 * @return String
	 * @throws JCTException
	 */
	@Transactional(propagation = Propagation.REQUIRED)
	public String updateTextRadioQtn(int userType, int answerType,
			String originalQuestion, String updatedMainQuestion,
			String subQuestion, String createdBy, String isMandatory) throws JCTException {
		LOGGER.info(">>>>>>>> SurveyQuestionServiceImpl.updateTextRadioQtn");
		String result = "success";
		// Update all the entries with main question and answer type and user type
		try {
			List<JctSurveyQuestionMaster> masterList = surveyQuestionDAO.getSurveyQuestionList(userType, answerType, originalQuestion);
			surveyQuestionDAO.deleteSurveyQtn(userType, answerType, originalQuestion);
			final String regularExp = "`~!~`";
			Pattern p = Pattern.compile(regularExp);
			String[] items = p.split(subQuestion);
			for(String subQtn : items) {
				JctSurveyQuestionMaster master = new JctSurveyQuestionMaster();
		        master.setJctSurveyQtnMasterAnsType(answerType);
		        master.setJctSurveyQtnMasterCreatedBy(createdBy);
		        master.setJctSurveyQtnMasterCreatedTs(new Date());
		        master.setJctSurveyQtnMasterMainQtn(updatedMainQuestion);
		        master.setJctSurveyQtnMasterModifiedBy(createdBy);
		        master.setJctSurveyQtnMasterModifiedTs(new Date());
		        master.setJctSurveyQtnMasterSubQtn(subQtn);
		        master.setJctSurveyQtnMasterUserType(userType);
		        master.setJctSurveyQtnMasterMandatory(isMandatory);
		        master.setJctSurveyQtnMasterProfileId(masterList.get(0).getJctSurveyQtnMasterProfileId());
		        master.setJctSurveyQtnMasterProfileName(masterList.get(0).getJctSurveyQtnMasterProfileName());
		        master.setJctSurveyQtnMasterOrder(masterList.get(0).getJctSurveyQtnMasterOrder());
		        surveyQuestionDAO.saveSurveyQuestion(master);
			}
		} catch (DAOException exception) {
			result = "failure";
			LOGGER.error("UNABLE TO UPDATE RADIO. "+exception.getLocalizedMessage());
		}		
		LOGGER.info("<<<<<<<< SurveyQuestionServiceImpl.updateTextRadioQtn");
		return result;
	}
	/**
	 * Method updates existing drop-down survey questions
	 * @param userType
	 * @param answerType
	 * @param originalQuestion
	 * @param updatedMainQuestion
	 * @param subQuestion
	 * @param createdBy
	 * @return String
	 * @throws JCTException
	 */
	@Transactional(propagation = Propagation.REQUIRED)
	public String updateDropDownSurveyQtn(int userType, int answerType,
			String originalQuestion, String updatedMainQuestion,
			String subQuestion, String createdBy, String isMandatory) throws JCTException {
		LOGGER.info(">>>>>>>> SurveyQuestionServiceImpl.updateDropDownSurveyQtn");
		String result = "success";
		// Update all the entries with main question and answer type and user type
		try {
			List<JctSurveyQuestionMaster> masterList = surveyQuestionDAO.getSurveyQuestionList(userType, answerType, originalQuestion);
			surveyQuestionDAO.deleteSurveyQtn(userType, answerType, originalQuestion);
			final String regularExp = "`~!~`";
			Pattern p = Pattern.compile(regularExp);
			String[] items = p.split(subQuestion);
			for(String subQtn : items) {
				JctSurveyQuestionMaster master = new JctSurveyQuestionMaster();
		        master.setJctSurveyQtnMasterAnsType(answerType);
		        master.setJctSurveyQtnMasterCreatedBy(createdBy);
		        master.setJctSurveyQtnMasterCreatedTs(new Date());
		        master.setJctSurveyQtnMasterMainQtn(updatedMainQuestion);
		        master.setJctSurveyQtnMasterModifiedBy(createdBy);
		        master.setJctSurveyQtnMasterModifiedTs(new Date());
		        master.setJctSurveyQtnMasterSubQtn(subQtn);
		        master.setJctSurveyQtnMasterUserType(userType);
		        master.setJctSurveyQtnMasterMandatory(isMandatory);
		        master.setJctSurveyQtnMasterProfileId(masterList.get(0).getJctSurveyQtnMasterProfileId());
		        master.setJctSurveyQtnMasterProfileName(masterList.get(0).getJctSurveyQtnMasterProfileName());
		        master.setJctSurveyQtnMasterOrder(masterList.get(0).getJctSurveyQtnMasterOrder());
		        surveyQuestionDAO.saveSurveyQuestion(master);
			}
		} catch (DAOException exception) {
			result = "failure";
			LOGGER.error("UNABLE TO UPDATE RADIO. "+exception.getLocalizedMessage());
		}		
		LOGGER.info("<<<<<<<< SurveyQuestionServiceImpl.updateDropDownSurveyQtn");
		return result;
	}
	/**
	 * Method updates existing text survey questions
	 * @param userType
	 * @param answerType
	 * @param originalQuestion
	 * @param updatedMainQuestion
	 * @param subQuestion
	 * @param createdBy
	 * @return String
	 * @throws JCTException
	 */
	@Transactional(propagation = Propagation.REQUIRED)
	public String updateCheckboxSurveyQtn(int userType, int answerType,
			String originalQuestion, String updatedMainQuestion,
			String subQuestion, String createdBy, String isMandatory) throws JCTException {
		LOGGER.info(">>>>>>>> SurveyQuestionServiceImpl.updateCheckboxSurveyQtn");
		String result = "success";
		// Update all the entries with main question and answer type and user type
		try {
			List<JctSurveyQuestionMaster> masterList = surveyQuestionDAO.getSurveyQuestionList(userType, answerType, originalQuestion);
			surveyQuestionDAO.deleteSurveyQtn(userType, answerType, originalQuestion);
			final String regularExp = "`~!~`";
			Pattern p = Pattern.compile(regularExp);
			String[] items = p.split(subQuestion);
			for(String subQtn : items) {
				JctSurveyQuestionMaster master = new JctSurveyQuestionMaster();
		        master.setJctSurveyQtnMasterAnsType(answerType);
		        master.setJctSurveyQtnMasterCreatedBy(createdBy);
		        master.setJctSurveyQtnMasterCreatedTs(new Date());
		        master.setJctSurveyQtnMasterMainQtn(updatedMainQuestion);
		        master.setJctSurveyQtnMasterModifiedBy(createdBy);
		        master.setJctSurveyQtnMasterModifiedTs(new Date());
		        master.setJctSurveyQtnMasterSubQtn(subQtn);
		        master.setJctSurveyQtnMasterUserType(userType);
		        master.setJctSurveyQtnMasterMandatory(isMandatory);
		        master.setJctSurveyQtnMasterProfileId(masterList.get(0).getJctSurveyQtnMasterProfileId());
		        master.setJctSurveyQtnMasterProfileName(masterList.get(0).getJctSurveyQtnMasterProfileName());
		        master.setJctSurveyQtnMasterOrder(masterList.get(0).getJctSurveyQtnMasterOrder());
		        surveyQuestionDAO.saveSurveyQuestion(master);
			}
		} catch (DAOException exception) {
			result = "failure";
			LOGGER.error("UNABLE TO UPDATE RADIO. "+exception.getLocalizedMessage());
		}		
		LOGGER.info("<<<<<<<< SurveyQuestionServiceImpl.updateCheckboxSurveyQtn");
		return result;
	}
	/**
	 * Method checks if the same active survey question exists with the same combination. 
	 * @param question
	 * @param ansType
	 * @param userType
	 * @return
	 */
	private String checkIfSurveyQtnExists (String question, int ansType, int userType) {
		String ans = "Y";
		try {
			List<JctSurveyQuestionMaster> list = surveyQuestionDAO.getSurveyQuestionList(userType, ansType, question);
			if (list.size() == 0) {
				ans = "N";
			}
		} catch (DAOException e) {
			LOGGER.error("UNABLE TO FETCH LIST. "+e.getLocalizedMessage());
		}
		return ans;
	}
	
	/**
	 * Method saves all the survey questions.
	 * @param text
	 * @param radio
	 * @param drop
	 * @param check
	 * @return
	 * @throws JCTException
	 */
	@Transactional(propagation=Propagation.REQUIRED)
	public String saveSurveyQuestions(String text, String radio, String drop,
			String check, String emailId, int userId) throws JCTException {
		LOGGER.info(">>>>>>>> SurveyQuestionServiceImpl.saveSurveyQuestions");
		String result = "failure";
		try {
			if (text.trim().length() > 0) {
				saveTextSurveyAns(text, emailId, userId);
			}
			if (radio.trim().length() > 0) {
				saveRadioSurveyAns(radio, emailId, userId);
			}
			if (drop.trim().length() > 0) {
				saveDropSurveyAns(drop, emailId, userId);
			}
			if (check.trim().length() > 0) {
				saveCheckSurveyAns(check, emailId, userId);
			}
			result = "success";
			if(result.equalsIgnoreCase("success")) {
				JctUser user = surveyQuestionDAO.fetchUserData(userId);
				user.setJctActiveYn(2);
				surveyQuestionDAO.updateUser(user);
			}
		} catch (DAOException ex) {
			LOGGER.error("UNABLE TO SAVE TO SAVE SURVEY QUESTIONS: "+ex.getLocalizedMessage());
		}
		LOGGER.info("<<<<<<<< SurveyQuestionServiceImpl.saveSurveyQuestions");
		return result;
	}
	
	/**
	 * Method saves only the free text surveys.
	 * @param text
	 * @return
	 * @throws DAOException
	 */
	@Transactional(propagation=Propagation.REQUIRED)
	private String saveTextSurveyAns (String text, String emailId, int userId) throws DAOException {
		LOGGER.info(">>>>>>>> SurveyQuestionServiceImpl.saveTextSurveyAns");
		//What do you think of the tool?`ONE~~How asked you to join?`NOTANSWERED~~HEHE`TWO~~HOHO`NOTANSWERED~~HIHI`THREE~~HUHU`NOTANSWERED~~Test`NOTANSWERED~~
		String status = "failure";
		String[] qtnSplit = text.split("~~");
		for (int index = 0; index < qtnSplit.length; index++) {
			String[] innerSplit = qtnSplit[index].split("`"); // Left = qtn | Right = ans
			String mainQtn = innerSplit[0];
			String answer = innerSplit[1];
			if (!answer.equals("NOTANSWERED")) {
				JctSurveyQuestions surveyQuestions = new JctSurveyQuestions();
				surveyQuestions.setJctSurveyAnswer(answer);
				surveyQuestions.setJctSurveyAnswerType(CommonConstants.FREE_TEXT);
				surveyQuestions.setJctSurveyCreatedTs(new Date());
				surveyQuestions.setJctSurveyMainQtn(mainQtn);
				surveyQuestions.setJctSurveySubQtn("NA");
				surveyQuestions.setJctSurveyTakenBy(emailId);
				surveyQuestions.setJctSurveyTakenByUserId(userId);
				surveyQuestions.setJctSurveyUserType(CommonConstants.FACILITATOR_USER);
				surveyQuestionDAO.saveSurveyQuestion(surveyQuestions);
			}
			status = "success";
		}
		LOGGER.info("<<<<<<<< SurveyQuestionServiceImpl.saveTextSurveyAns");
		return status;
	}
	
	/**
	 * Method saves only the radio surveys.
	 * @param text
	 * @return
	 * @throws DAOException
	 */
	@Transactional(propagation=Propagation.REQUIRED)
	private String saveRadioSurveyAns (String radio, String emailId, int userId) throws DAOException {
		LOGGER.info(">>>>>>>> SurveyQuestionServiceImpl.saveRadioSurveyAns");
		// RADIO WITH 5 OPTIONS`OPTION 1~~Are you happy?`Yes~~
		String status = "failure";
		String[] qtnSplit = radio.split("~~");
		for (int index = 0; index < qtnSplit.length; index++) {
			try {
			String[] innerSplit = qtnSplit[index].split("`"); // Left = qtn | Right = ans
			String ans = (String) innerSplit[1];
			if(!ans.trim().equals("NOTANSWERED")){
				if (ans.trim().length() > 0) {
					JctSurveyQuestions surveyQuestions = new JctSurveyQuestions();
					surveyQuestions.setJctSurveyAnswer((String) innerSplit[1]);
					surveyQuestions.setJctSurveyAnswerType(CommonConstants.RADIO_BTN);
					surveyQuestions.setJctSurveyCreatedTs(new Date());
					surveyQuestions.setJctSurveyMainQtn((String) innerSplit[0]);
					surveyQuestions.setJctSurveySubQtn((String) innerSplit[0]);
					surveyQuestions.setJctSurveyTakenBy(emailId);
					surveyQuestions.setJctSurveyTakenByUserId(userId);
					surveyQuestions.setJctSurveyUserType(CommonConstants.FACILITATOR_USER);
					surveyQuestionDAO.saveSurveyQuestion(surveyQuestions);
				}				
			}
				
			} catch (Exception e) {
				LOGGER.error("NO OPTION SELECTED FOR RADIO");
			}
			status = "success";
		}
		LOGGER.info("<<<<<<<< SurveyQuestionServiceImpl.saveRadioSurveyAns");
		return status;
	}
	/**
	 * Method saves only the drop down surveys.
	 * @param text
	 * @return
	 * @throws DAOException
	 */
	@Transactional(propagation=Propagation.REQUIRED)
	private String saveDropSurveyAns (String drop, String emailId, int userId) throws DAOException {
		LOGGER.info(">>>>>>>> SurveyQuestionServiceImpl.saveDropSurveyAns");
		// What is your job profile?`HR~~AGAIN A DROP DOWN`VERY FUNNY~~ 
		String status = "failure";
		String[] qtnSplit = drop.split("~~");
		for (int index = 0; index < qtnSplit.length; index++) {
			String[] innerSplit = qtnSplit[index].split("`"); // Left = qtn | Right = ans
			String ans = (String) innerSplit[1];
			if(!ans.trim().equals("NOTANSWERED")){
				if (ans.trim().length() > 0) {
					JctSurveyQuestions surveyQuestions = new JctSurveyQuestions();
					surveyQuestions.setJctSurveyAnswer((String) innerSplit[1]);
					surveyQuestions.setJctSurveyAnswerType(CommonConstants.DROP_DOWN);
					surveyQuestions.setJctSurveyCreatedTs(new Date());
					surveyQuestions.setJctSurveyMainQtn((String) innerSplit[0]);
					surveyQuestions.setJctSurveySubQtn((String) innerSplit[0]);
					surveyQuestions.setJctSurveyTakenBy(emailId);
					surveyQuestions.setJctSurveyTakenByUserId(userId);
					surveyQuestions.setJctSurveyUserType(CommonConstants.FACILITATOR_USER);
					surveyQuestionDAO.saveSurveyQuestion(surveyQuestions);
				}
				status = "success";				
			}
			
		}
		LOGGER.info("<<<<<<<< SurveyQuestionServiceImpl.saveDropSurveyAns");
		return status;
	}
	/**
	 * Method saves only the check box surveys.
	 * @param text
	 * @return
	 * @throws DAOException
	 */
	@Transactional(propagation=Propagation.REQUIRED)
	private String saveCheckSurveyAns (String check, String emailId, int userId) throws DAOException {
		LOGGER.info(">>>>>>>> SurveyQuestionServiceImpl.saveCheckSurveyAns");
		String status = "failure";
		String[] qtnSplit = check.split("~~");
		for (int index = 0; index < qtnSplit.length; index++) {
			String[] innerSplit = qtnSplit[index].split("`"); // Left = qtn | Right = answers (multiple)
			try {
				String[] answerSplit = innerSplit[1].split("#");				
				for (int ansIndex = 0; ansIndex < answerSplit.length; ansIndex++) {
					JctSurveyQuestions surveyQuestions = new JctSurveyQuestions();
					if(((String) answerSplit[ansIndex]).equals("NOTANSWERED")){
						surveyQuestions.setJctSurveyAnswer((String) answerSplit[ansIndex]);
						surveyQuestions.setJctSurveyAnswerType(CommonConstants.CHECK_BOX);
						surveyQuestions.setJctSurveyCreatedTs(new Date());
						surveyQuestions.setJctSurveyMainQtn((String) innerSplit[0]);
						surveyQuestions.setJctSurveySubQtn((String) answerSplit[ansIndex]);
						surveyQuestions.setJctSurveyTakenBy(emailId);
						surveyQuestions.setJctSurveyTakenByUserId(userId);
						surveyQuestions.setJctSurveyUserType(CommonConstants.FACILITATOR_USER);
						surveyQuestionDAO.saveSurveyQuestion(surveyQuestions);
						status = "success";						
					}					
				}
			} catch (Exception ex) {
				LOGGER.error("NO SUB ELEMENT FOUND");
			}
		}
		LOGGER.info("<<<<<<<< SurveyQuestionServiceImpl.saveCheckSurveyAns");
		return status;
	}

	@Transactional(propagation=Propagation.REQUIRED)
	public String updateActiveStatusFacilitator(String emailId, int userId)
			throws JCTException {
		LOGGER.info(">>>>>>>> SurveyQuestionServiceImpl.updateActiveStatusFacilitator");
		String result = "failure";
		try {
			JctUser user = surveyQuestionDAO.fetchUserData(userId);
			user.setJctActiveYn(2);
			user.setLastmodifiedTs(new Date());
			user.setJctLastmodifiedBy(emailId);
			//user.getJctUserDetails().setJctUserOnetOccupation("ACTIVATED");
			surveyQuestionDAO.updateUser(user);
			result = "success";
		} catch (DAOException ex) {
			result = "failure";
			LOGGER.error("UNABLE TO SAVE TO SAVE SURVEY QUESTIONS: "+ex.getLocalizedMessage());
		}
		LOGGER.info("<<<<<<<< SurveyQuestionServiceImpl.updateActiveStatusFacilitator");
		return result;
	}
	
	/**
	 * Method fetches all existing survey questions by user type and question type
	 * @param userType
	 * @param qtnType
	 * @return String
	 * @throws JCTException
	 */
	@Transactional(propagation = Propagation.REQUIRED)
	public String fetchgSurveyQtnsByAllParams (int userType, int qtnType, int profileId) throws JCTException {
		StringBuilder questionStringBuilder = new StringBuilder("");
		try {
			List<SurveyQuestionDTO> mainQtnDtoList = surveyQuestionDAO.fetchgSurveyQtnsByAllParams(userType, qtnType, profileId);
			Iterator<SurveyQuestionDTO> mainQtnItr = mainQtnDtoList.iterator();
			while (mainQtnItr.hasNext()) {
				SurveyQuestionDTO dto = (SurveyQuestionDTO) mainQtnItr.next();
				questionStringBuilder.append(dto.getAnswerType());
				questionStringBuilder.append("`");
				questionStringBuilder.append(dto.getJctUserType());
				questionStringBuilder.append("`");
				questionStringBuilder.append(dto.getMainQuestion());
				questionStringBuilder.append("`");
				List<SurveyQuestionDTO> subQtnDtoList = surveyQuestionDAO.getAllSubQtns(qtnType, userType, dto.getMainQuestion());
				int loopCounter = 0;
				for (int index=0; index<subQtnDtoList.size(); index++) {
					SurveyQuestionDTO subQtnDto = (SurveyQuestionDTO) subQtnDtoList.get(index);
					loopCounter = loopCounter + 1;
					questionStringBuilder.append(subQtnDto.getSubQtns());
					if (loopCounter < subQtnDtoList.size()) {
						questionStringBuilder.append("^");
					}
				}
				questionStringBuilder.append("`");
				// Get the number of sub question
				questionStringBuilder.append(surveyQuestionDAO.getNumberOfSubQuestions(qtnType, userType, dto.getMainQuestion()));
				// String user type
				questionStringBuilder.append("`");
				questionStringBuilder.append(this.messageSource.getMessage("msg.user.type."+userType, null, null));
				// String user type
				questionStringBuilder.append("`");
				questionStringBuilder.append(this.messageSource.getMessage("msg.answer.type."+qtnType, null, null));
				// String mandatory
				questionStringBuilder.append("`");
				questionStringBuilder.append(dto.getIsMandatory());
				// user profile
				questionStringBuilder.append("`");
				questionStringBuilder.append(dto.getProfileName());
				
				questionStringBuilder.append("~");
			}
		} catch (DAOException dao) {
			LOGGER.error("UNABLE TO FETCH EXISTING SURVEY QUESTIONS BY USER TYPE AND QUESTION TYPE: "+dao.getLocalizedMessage());
		}
		return questionStringBuilder.toString();
	}
	
	@Transactional(propagation = Propagation.REQUIRED)
	public String fetchAllExistingSurveyQtnsOnlyUserProfile(int qtnType,
			int profileId) throws JCTException {
		StringBuilder questionStringBuilder = new StringBuilder("");
		try {
			List<SurveyQuestionDTO> mainQtnDtoList = surveyQuestionDAO.fetchAllExistingSurveyQtnsOnlyUserProfile(qtnType, profileId);
			Iterator<SurveyQuestionDTO> mainQtnItr = mainQtnDtoList.iterator();
			while (mainQtnItr.hasNext()) {
				SurveyQuestionDTO dto = (SurveyQuestionDTO) mainQtnItr.next();
				questionStringBuilder.append(dto.getAnswerType());
				questionStringBuilder.append("`");
				questionStringBuilder.append(dto.getJctUserType());
				questionStringBuilder.append("`");
				questionStringBuilder.append(dto.getMainQuestion());
				questionStringBuilder.append("`");
				List<SurveyQuestionDTO> subQtnDtoList = surveyQuestionDAO.getAllSubQtnsByAnsType(qtnType, dto.getMainQuestion());
				int loopCounter = 0;
				for (int index=0; index<subQtnDtoList.size(); index++) {
					SurveyQuestionDTO subQtnDto = (SurveyQuestionDTO) subQtnDtoList.get(index);
					loopCounter = loopCounter + 1;
					questionStringBuilder.append(subQtnDto.getSubQtns());
					if (loopCounter < subQtnDtoList.size()) {
						questionStringBuilder.append("^");
					}
				}
				questionStringBuilder.append("`");
				// Get the number of sub question
				questionStringBuilder.append(surveyQuestionDAO.getNumberOfSubQuestionsByAnsType(qtnType, dto.getMainQuestion()));
				// String user type
				questionStringBuilder.append("`");
				questionStringBuilder.append(this.messageSource.getMessage("msg.user.type."+dto.getJctUserType(), null, null));
				// String user type
				questionStringBuilder.append("`");
				questionStringBuilder.append(this.messageSource.getMessage("msg.answer.type."+qtnType, null, null));
				questionStringBuilder.append("`");
				questionStringBuilder.append(dto.getIsMandatory());
				// user profile
				questionStringBuilder.append("`");
				questionStringBuilder.append(dto.getProfileName());
				questionStringBuilder.append("~");
			}
		} catch (DAOException dao) {
			LOGGER.error("UNABLE TO FETCH EXISTING SURVEY QUESTIONS BY USER TYPE AND QUESTION TYPE: "+dao.getLocalizedMessage());
		}
		return questionStringBuilder.toString();
	}
	
	@Transactional(propagation = Propagation.REQUIRED)
	public String fetchSurveyQtnsByProfileAndUserType(int userType,
			int profileId) throws JCTException {/*
		StringBuilder questionStringBuilder = new StringBuilder("");
		try {
			List<SurveyQuestionDTO> mainQtnDtoList = questionDAO.fetchSurveyQtnsByProfileAndUserType(userType, profileId);
			Iterator<SurveyQuestionDTO> mainQtnItr = mainQtnDtoList.iterator();
			while (mainQtnItr.hasNext()) {
				SurveyQuestionDTO dto = (SurveyQuestionDTO) mainQtnItr.next();
				questionStringBuilder.append(dto.getAnswerType());
				questionStringBuilder.append("`");
				questionStringBuilder.append(dto.getJctUserType());
				questionStringBuilder.append("`");
				questionStringBuilder.append(dto.getMainQuestion());
				questionStringBuilder.append("`");
				List<SurveyQuestionDTO> subQtnDtoList = questionDAO.getAllSubQtnsByUserTypeAndProfile(userType, profileId, dto.getMainQuestion());
				int loopCounter = 0;
				for (int index=0; index<subQtnDtoList.size(); index++) {
					SurveyQuestionDTO subQtnDto = (SurveyQuestionDTO) subQtnDtoList.get(index);
					loopCounter = loopCounter + 1;
					questionStringBuilder.append(subQtnDto.getSubQtns());
					if (loopCounter < subQtnDtoList.size()) {
						questionStringBuilder.append("^");
					}
				}
				questionStringBuilder.append("`");
				// Get the number of sub question
				questionStringBuilder.append(questionDAO.getNumberOfSubQuestionsByAnsType(qtnType, dto.getMainQuestion()));
				// String user type
				questionStringBuilder.append("`");
				questionStringBuilder.append(this.messageSource.getMessage("msg.user.type."+dto.getJctUserType(), null, null));
				// String user type
				questionStringBuilder.append("`");
				questionStringBuilder.append(this.messageSource.getMessage("msg.answer.type."+qtnType, null, null));
				questionStringBuilder.append("`");
				questionStringBuilder.append(dto.getIsMandatory());
				// user profile
				questionStringBuilder.append("`");
				questionStringBuilder.append(dto.getProfileName());
				questionStringBuilder.append("~");
			}
		} catch (DAOException dao) {
			LOGGER.error("UNABLE TO FETCH EXISTING SURVEY QUESTIONS BY USER TYPE AND QUESTION TYPE: "+dao.getLocalizedMessage());
		}
		return questionStringBuilder.toString();
	*/

		StringBuilder questionStringBuilder = new StringBuilder("");
		try {
			List<SurveyQuestionDTO> mainQtnDtoList = surveyQuestionDAO.fetchSurveyQtnsByProfileAndUserType(userType, profileId);
			Iterator<SurveyQuestionDTO> mainQtnItr = mainQtnDtoList.iterator();
			while (mainQtnItr.hasNext()) {
				SurveyQuestionDTO dto = (SurveyQuestionDTO) mainQtnItr.next();
				questionStringBuilder.append(dto.getAnswerType());
				questionStringBuilder.append("`");
				questionStringBuilder.append(dto.getJctUserType());
				questionStringBuilder.append("`");
				questionStringBuilder.append(dto.getMainQuestion());
				questionStringBuilder.append("`");
				List<SurveyQuestionDTO> subQtnDtoList = surveyQuestionDAO.getAllSubQtnsByUserTypeAndProfile(profileId, dto.getAnswerType(), userType, dto.getMainQuestion());
				int loopCounter = 0;
				for (int index=0; index<subQtnDtoList.size(); index++) {
					SurveyQuestionDTO subQtnDto = (SurveyQuestionDTO) subQtnDtoList.get(index);
					loopCounter = loopCounter + 1;
					questionStringBuilder.append(subQtnDto.getSubQtns());
					if (loopCounter < subQtnDtoList.size()) {
						questionStringBuilder.append("^");
					}
				}
				questionStringBuilder.append("`");
				// Get the number of sub question
				questionStringBuilder.append(surveyQuestionDAO.getNumberOfSubQuestions(dto.getAnswerType(), userType, dto.getMainQuestion()));
				// String user type
				questionStringBuilder.append("`");
				questionStringBuilder.append(this.messageSource.getMessage("msg.user.type."+userType, null, null));
				// String user type
				questionStringBuilder.append("`");
				questionStringBuilder.append(this.messageSource.getMessage("msg.answer.type."+dto.getAnswerType(), null, null));
				//  String mandatory
				questionStringBuilder.append("`");
				questionStringBuilder.append(dto.getIsMandatory());
				// User profile
				questionStringBuilder.append("`");
				questionStringBuilder.append(dto.getProfileName());
				questionStringBuilder.append("~");
			}
		} catch (DAOException dao) {
			LOGGER.error("UNABLE TO FETCH EXISTING SURVEY QUESTIONS BY USER TYPE: "+dao.getLocalizedMessage());
		}
		return questionStringBuilder.toString();			
	}
	
	@Transactional(propagation = Propagation.REQUIRED)
	public SurveyQuestionsVO saveSurvryQtnOrder(String builder,
			Integer userProfile, Integer userType,String builderAnsType) throws JCTException {
		LOGGER.info(">>>>>> SurveyQuestionServiceImpl.saveSurvryQtnOrder");
		SurveyQuestionsVO vo = new SurveyQuestionsVO();
		String retString = null;
		try {
			retString = surveyQuestionDAO.saveSurvryQtnOrder(builder, userProfile, userType, builderAnsType);
		} catch (DAOException e) {
			LOGGER.error("----"+e.getLocalizedMessage()+" ----");
			throw new JCTException(e.getLocalizedMessage());
		}
		
		if (!retString.equals("success")) {
			vo.setStatusCode(StatusConstants.STATUS_FAILURE);
			vo.setMessage(this.messageSource.
					getMessage("exception.dao.user.group.update", null, null));
		} else {
			vo.setStatusCode(StatusConstants.STATUS_SUCCESS_WITH_RESULT);
			vo.setMessage(retString);
		}
		LOGGER.info("<<<<<< SurveyQuestionServiceImpl.saveSurvryQtnOrder");
		return vo;
	}
	
	
}