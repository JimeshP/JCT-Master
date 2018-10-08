package com.vmware.jct.controller;

import javax.servlet.http.HttpServletRequest;

import org.codehaus.jackson.JsonNode;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.context.MessageSource;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.RequestBody;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestMethod;
import org.springframework.web.bind.annotation.ResponseBody;

import com.vmware.jct.common.utility.CommonConstants;
import com.vmware.jct.common.utility.CommonUtility;
import com.vmware.jct.common.utility.StatusConstants;
import com.vmware.jct.exception.JCTException;
import com.vmware.jct.service.ISurveyQuestionService;
import com.vmware.jct.service.vo.ContentConfigVO;
import com.vmware.jct.service.vo.SurveyQuestionsVO;
/**
 * 
 * <p><b>Class name:</b> ConfigureSurveyQuestionController.java</p>
 * <p><b>Author:</b> InterraIT</p>
 * <p><b>Purpose:</b> This class acts as a controller for configuring survey questions..
 * ConfigureSurveyQuestionController extends BasicController  and has following  methods.
 * - saveFreeText()
 * <p><b>Description:</b> This class acts as a controller for configuring survey questions. </p>
 * <p><b>Copyrights:</b>  All rights reserved by Interra IT and should only be used for its internal application development.</p>
 * <p><b>Created Date:</b> 21/Oct/2014</p>
 * <p><b>Revision History:</b>
 * 	<li>InterraIT - 21/Oct/2014 - Implemented class and comments</li>
 * </p>
 */
@Controller
@RequestMapping(value="/surveyQtns")
public class ConfigureSurveyQuestionController extends BasicController {
	
	@Autowired  
	private MessageSource messageSource;
	
	@Autowired
	private ISurveyQuestionService surveyService;
	
	private static final Logger LOGGER = LoggerFactory.getLogger(ConfigureSurveyQuestionController.class);
	/**
	 * Method stores free text
	 * @param body
	 * @param request
	 * @return
	 */
	@RequestMapping(value = ACTION_SAVE_FREE_QTN_SURVEY_QTN, method = RequestMethod.POST)
	@ResponseBody()
	public SurveyQuestionsVO saveFreeText(@RequestBody String body, HttpServletRequest request) {
		LOGGER.info(">>>>>>>> ConfigureSurveyQuestionController.saveFreeText");
		JsonNode node = CommonUtility.getNode(body);
		int userType = node.get("userType").asInt();
		String textQuestions = node.get("textQtns").asText();
		String createdBy = node.get("createdBy").asText();
		String userProfile = node.get("userProfile").asText();
		String[] userProfileSplit = userProfile.split("!");
		int profileId = Integer.parseInt(userProfileSplit[0]);
		
		SurveyQuestionsVO vo = new SurveyQuestionsVO();
		vo.setCreatedBy(createdBy);
		vo.setSurveyTextQuestion(textQuestions);
		vo.setUserType(userType);
		vo.setUserProfile(userProfile);
		try {
			vo = surveyService.saveFreeText(vo);
			vo.setResultList(surveyService.fetchgSurveyQtnsByAllParams(userType, CommonConstants.FREE_TEXT, profileId));
		} catch (JCTException e) {
			LOGGER.error("UNABLE TO SAVE FREE TEXT SURVEY QUESTION: "+e.getLocalizedMessage());
		}
		LOGGER.info("<<<<<<<< ConfigureSurveyQuestionController.saveFreeText");
		return vo;
	}
	/**
	 * Method stores radio
	 * @param body
	 * @param request
	 * @return
	 */
	@RequestMapping(value = ACTION_SAVE_RADIO_ANSWERS, method = RequestMethod.POST)
	@ResponseBody()
	public SurveyQuestionsVO saveRadio(@RequestBody String body, HttpServletRequest request) {
		LOGGER.info(">>>>>>>> ConfigureSurveyQuestionController.saveRadio");
		JsonNode node = CommonUtility.getNode(body);
		int userType = node.get("userType").asInt();
		String mainQuestion = node.get("mainQuestion").asText();
		String subQtns = node.get("subQtns").asText();
		String createdBy = node.get("createdBy").asText();
		String mandatoryChkVal = node.get("mandatoryChkVal").asText();	
		String userProfile = node.get("userProfile").asText();
		String[] userProfileSplit = userProfile.split("!");
		int profileId = Integer.parseInt(userProfileSplit[0]);
		SurveyQuestionsVO vo = new SurveyQuestionsVO();
		vo.setCreatedBy(createdBy);
		vo.setMainQuestion(mainQuestion);
		vo.setSubQuestionsString(subQtns);
		vo.setUserType(userType);
		vo.setIsMandatory(mandatoryChkVal);
		vo.setUserProfile(userProfile);
		try {
			vo = surveyService.saveRadio(vo);
			vo.setResultList(surveyService.fetchgSurveyQtnsByAllParams(userType, CommonConstants.RADIO_BTN, profileId));
		} catch (JCTException e) {
			LOGGER.error("UNABLE TO SAVE RADIO SURVEY QUESTION: "+e.getLocalizedMessage());
		}
		LOGGER.info("<<<<<<<< ConfigureSurveyQuestionController.saveRadio");
		return vo;
	}
	/**
	 * Method stores drop down
	 * @param body
	 * @param request
	 * @return
	 */
	@RequestMapping(value = ACTION_SAVE_DROP_DOWN, method = RequestMethod.POST)
	@ResponseBody()
	public SurveyQuestionsVO saveDropDown(@RequestBody String body, HttpServletRequest request) {
		LOGGER.info(">>>>>>>> ConfigureSurveyQuestionController.saveDropDown");
		JsonNode node = CommonUtility.getNode(body);
		int userType = node.get("userType").asInt();
		String mainQuestion = node.get("mainQuestion").asText();
		String subQtns = node.get("subQtns").asText();
		String createdBy = node.get("createdBy").asText();
		String mandatoryChkVal = node.get("mandatoryChkVal").asText();
		String userProfile = node.get("userProfile").asText();
		String[] userProfileSplit = userProfile.split("!");
		int profileId = Integer.parseInt(userProfileSplit[0]);
		SurveyQuestionsVO vo = new SurveyQuestionsVO();
		vo.setCreatedBy(createdBy);
		vo.setMainQuestion(mainQuestion);
		vo.setSubQuestionsString(subQtns);
		vo.setUserType(userType);
		vo.setIsMandatory(mandatoryChkVal);
		vo.setUserProfile(userProfile);
		try {
			vo = surveyService.saveDropDown(vo);
			vo.setResultList(surveyService.fetchgSurveyQtnsByAllParams(userType, CommonConstants.DROP_DOWN, profileId));
		} catch (JCTException e) {
			LOGGER.error("UNABLE TO SAVE DROP DOWN SURVEY QUESTION: "+e.getLocalizedMessage());
		}
		LOGGER.info("<<<<<<<< ConfigureSurveyQuestionController.saveDropDown");
		return vo;
	}
	/**
	 * Method stores check box
	 * @param body
	 * @param request
	 * @return
	 */
	@RequestMapping(value = ACTION_SAVE_CHECK_BOX, method = RequestMethod.POST)
	@ResponseBody()
	public SurveyQuestionsVO saveCheckBox(@RequestBody String body, HttpServletRequest request) {
		LOGGER.info(">>>>>>>> ConfigureSurveyQuestionController.saveCheckBox");
		JsonNode node = CommonUtility.getNode(body);
		int userType = node.get("userType").asInt();
		String mainQuestion = node.get("mainQuestion").asText();
		String subQtns = node.get("subQtns").asText();
		String createdBy = node.get("createdBy").asText();
		String mandatoryChkVal = node.get("mandatoryChkVal").asText();
		String userProfile = node.get("userProfile").asText();
		String[] userProfileSplit = userProfile.split("!");
		int profileId = Integer.parseInt(userProfileSplit[0]);
		SurveyQuestionsVO vo = new SurveyQuestionsVO();
		vo.setCreatedBy(createdBy);
		vo.setMainQuestion(mainQuestion);
		vo.setSubQuestionsString(subQtns);
		vo.setUserType(userType);
		vo.setIsMandatory(mandatoryChkVal);
		vo.setUserProfile(userProfile);
		try {
			vo = surveyService.saveCheckBox(vo);
			vo.setResultList(surveyService.fetchgSurveyQtnsByAllParams(userType, CommonConstants.CHECK_BOX, profileId));
		} catch (JCTException e) {
			LOGGER.error("UNABLE TO SAVE CHECK BOX SURVEY QUESTION: "+e.getLocalizedMessage());
		}
		LOGGER.info("<<<<<<<< ConfigureSurveyQuestionController.saveCheckBox");
		return vo;
	}
	/**
	 * Method fetches all existing survey question
	 * @param body
	 * @param request
	 * @return
	 */
	@RequestMapping(value = ACTION_FETCH_ALL_EXISTING_SURVEY_QTNS, method = RequestMethod.POST)
	@ResponseBody()
	public SurveyQuestionsVO fetchAllExistingSurveyQtns(@RequestBody String body, HttpServletRequest request) {
		LOGGER.info(">>>>>>>> ConfigureSurveyQuestionController.fetchAllExistingSurveyQtns");
		SurveyQuestionsVO vo = new SurveyQuestionsVO();
		try {
			vo.setResultList(surveyService.getExistingAllSurveyQuestion());
		} catch (JCTException e) {
			LOGGER.error("UNABLE TO FETCH SURVEY QUESTION: "+e.getLocalizedMessage());
		}
		LOGGER.info("<<<<<<<< ConfigureSurveyQuestionController.fetchAllExistingSurveyQtns");
		return vo;
	}
	/**
	 * Method fetches all existing survey question by user type
	 * @param body
	 * @param request
	 * @return
	 */
	@RequestMapping(value = ACTION_FETCH_ALL_EXISTING_SURVEY_QTNS_BY_USR_TYP, method = RequestMethod.POST)
	@ResponseBody()
	public SurveyQuestionsVO fetchAllExistingSurveyQtnsByUserType(@RequestBody String body, HttpServletRequest request) {
		LOGGER.info(">>>>>>>> ConfigureSurveyQuestionController.fetchAllExistingSurveyQtnsByUserType");
		JsonNode node = CommonUtility.getNode(body);
		int userType = node.get("userType").asInt();
		SurveyQuestionsVO vo = new SurveyQuestionsVO();
		try {
			vo.setResultList(surveyService.fetchAllExistingSurveyQtnsByUserType(userType));
		} catch (JCTException e) {
			LOGGER.error("UNABLE TO FETCH SURVEY QUESTION: "+e.getLocalizedMessage());
		}
		LOGGER.info("<<<<<<<< ConfigureSurveyQuestionController.fetchAllExistingSurveyQtnsByUserType");
		return vo;
	}
	/**
	 * Method fetches all existing survey question by user type and question type
	 * @param body
	 * @param request
	 * @return
	 */
	@RequestMapping(value = ACTION_FETCH_ALL_SURV_QTNS_BY_USR_TYPE_QTN_TYPE, method = RequestMethod.POST)
	@ResponseBody()
	public SurveyQuestionsVO fetchAllExistingSurveyQtnsByUserTypeAndQtnType(@RequestBody String body, HttpServletRequest request) {
		LOGGER.info(">>>>>>>> ConfigureSurveyQuestionController.fetchAllExistingSurveyQtnsByUserTypeAndQtnType");
		JsonNode node = CommonUtility.getNode(body);
		int userType = node.get("userType").asInt();
		int qtnType = node.get("qtnType").asInt();
		
		SurveyQuestionsVO vo = new SurveyQuestionsVO();
		try {
			vo.setResultList(surveyService.fetchAllExistingSurveyQtnsByUserTypeAndQtnType(userType, qtnType));
		} catch (JCTException e) {
			LOGGER.error("UNABLE TO FETCH SURVEY QUESTION: "+e.getLocalizedMessage());
		}
		LOGGER.info("<<<<<<<< ConfigureSurveyQuestionController.fetchAllExistingSurveyQtnsByUserTypeAndQtnType");
		return vo;
	}
	/**
	 * Method fetches all existing survey question by question type
	 * @param body
	 * @param request
	 * @return
	 */
	@RequestMapping(value = ACTION_FETCH_ALL_SURV_QTNS_BY_QTN_TYPE, method = RequestMethod.POST)
	@ResponseBody()
	public SurveyQuestionsVO fetchAllExistingSurveyQtnsByQtnType(@RequestBody String body, HttpServletRequest request) {
		LOGGER.info(">>>>>>>> ConfigureSurveyQuestionController.fetchAllExistingSurveyQtnsByQtnType");
		JsonNode node = CommonUtility.getNode(body);
		int qtnType = node.get("qtnType").asInt();
		
		SurveyQuestionsVO vo = new SurveyQuestionsVO();
		try {
			vo.setResultList(surveyService.fetchAllExistingSurveyQtnsByQtnType(qtnType));
		} catch (JCTException e) {
			LOGGER.error("UNABLE TO FETCH SURVEY QUESTION: "+e.getLocalizedMessage());
		}
		LOGGER.info("<<<<<<<< ConfigureSurveyQuestionController.fetchAllExistingSurveyQtnsByQtnType");
		return vo;
	}
	/**
	 * Method deletes the survey question
	 * @param body
	 * @param request
	 * @return
	 */
	@RequestMapping(value = ACTION_DELETE_SURVEY_QUESTION, method = RequestMethod.POST)
	@ResponseBody()
	public SurveyQuestionsVO deleteSurveyQtn(@RequestBody String body, HttpServletRequest request) {
		LOGGER.info(">>>>>>>> ConfigureSurveyQuestionController.deleteSurveyQtn");
		JsonNode node = CommonUtility.getNode(body);
		int userType = node.get("userType").asInt();
		int answerType = node.get("answerType").asInt();
		String qtnDescription = node.get("qtnDescription").asText();
		String deciderParam = node.get("deciderParam").asText();
		String deciderParamArr[] = deciderParam.split("`");
		int dispUserType = Integer.parseInt(deciderParamArr[0]);
		int radioSelected = Integer.parseInt(deciderParamArr[1]);
		SurveyQuestionsVO vo = new SurveyQuestionsVO();
		try {
			String result = surveyService.deleteSurveyQtn(userType, answerType, qtnDescription);
			if (dispUserType == 0 && radioSelected == 5) { 								 // NONE SELECTED
				vo.setResultList(surveyService.getExistingAllSurveyQuestion());
			} else if ((dispUserType == 1 || dispUserType == 3) && radioSelected == 5) { // ONLY USER TYPE SELECTED
				vo.setResultList(surveyService.fetchAllExistingSurveyQtnsByUserType (dispUserType));
			}  else if ((dispUserType == 1 || dispUserType == 3) && radioSelected != 5) { // BOTH SELECTED
				vo.setResultList(surveyService.fetchAllExistingSurveyQtnsByUserTypeAndQtnType (dispUserType, radioSelected));
			}
			vo.setMessage(this.messageSource.getMessage("msg.success.survey.deletion",null, null));
		} catch (JCTException e) {
			vo.setMessage(this.messageSource.getMessage("msg.failure.survey.deletion",null, null));
			LOGGER.error("UNABLE TO DELETE SURVEY QUESTION: "+e.getLocalizedMessage());
		}
		LOGGER.info("<<<<<<<< ConfigureSurveyQuestionController.deleteSurveyQtn");
		return vo;
	}
	/**
	 * Method updates the text content of the survey question
	 * @param body
	 * @param request
	 * @return
	 */
	@RequestMapping(value = ACTION_UPDATE_TEXT_SURVEY_QUESTION, method = RequestMethod.POST)
	@ResponseBody()
	public SurveyQuestionsVO updateTextSurveyQtn(@RequestBody String body, HttpServletRequest request) {
		LOGGER.info(">>>>>>>> ConfigureSurveyQuestionController.updateTextSurveyQtn");
		JsonNode node = CommonUtility.getNode(body);
		int userType = node.get("userType").asInt();
		int answerType = node.get("answerType").asInt();
		String originalQuestion = node.get("originalQuestion").asText();		
		String updatedQuestion = node.get("updatedQuestion").asText();
		String isMandatory = node.get("isMandatory").asText();
		String isMandatoryOld = node.get("isMandatoryOld").asText();
		SurveyQuestionsVO vo = new SurveyQuestionsVO();
		try {
			if (!( originalQuestion.equals(updatedQuestion) &&  isMandatoryOld.equals(isMandatory) )) {
				surveyService.updateTextSurveyQtn(userType, answerType, originalQuestion, updatedQuestion,isMandatory);
				vo.setMessage(this.messageSource.getMessage("msg.success.survey.updation",null, null));
			} else {
				vo.setMessage(this.messageSource.getMessage("msg.failure.survey.updation.same",null, null));
			}
			vo.setResultList(surveyService.fetchAllExistingSurveyQtnsByUserTypeAndQtnType (userType, answerType));
			
			
		} catch (JCTException e) {
			vo.setMessage(this.messageSource.getMessage("msg.failure.survey.updation",null, null));
			LOGGER.error("UNABLE TO UPDATE SURVEY QUESTION: "+e.getLocalizedMessage());
		}
		LOGGER.info("<<<<<<<< ConfigureSurveyQuestionController.updateTextSurveyQtn");
		return vo;
	}
	/**
	 * Method updates the radio content of the survey question
	 * @param body
	 * @param request
	 * @return
	 */
	@RequestMapping(value = ACTION_UPDATE_RADIO_SURVEY_QUESTION, method = RequestMethod.POST)
	@ResponseBody()
	public SurveyQuestionsVO updateRadioSurveyQtn(@RequestBody String body, HttpServletRequest request) {
		LOGGER.info(">>>>>>>> ConfigureSurveyQuestionController.updateRadioSurveyQtn");
		JsonNode node = CommonUtility.getNode(body);
		int userType = node.get("userType").asInt();
		int answerType = node.get("answerType").asInt();
		String originalQuestion = node.get("originalQuestion").asText();
		String updatedMainQuestion = node.get("updatedMainQuestion").asText();
		String subQuestion = node.get("subQuestion").asText();
		String createdBy = node.get("createdBy").asText();
		String isMandatory = node.get("isMandatory").asText();
		
		SurveyQuestionsVO vo = new SurveyQuestionsVO();
		try {
			surveyService.updateTextRadioQtn(userType, answerType, originalQuestion, updatedMainQuestion, subQuestion, createdBy, isMandatory);
			vo.setResultList(surveyService.fetchAllExistingSurveyQtnsByUserTypeAndQtnType (userType, answerType));
			vo.setMessage(this.messageSource.getMessage("msg.success.survey.updation",null, null));
		} catch (JCTException e) {
			vo.setMessage(this.messageSource.getMessage("msg.failure.survey.updation",null, null));
			LOGGER.error("UNABLE TO UPDATE SURVEY QUESTION: "+e.getLocalizedMessage());
		}
		LOGGER.info("<<<<<<<< ConfigureSurveyQuestionController.updateRadioSurveyQtn");
		return vo;
	}
	/**
	 * Method updates the drop-down content of the survey question
	 * @param body
	 * @param request
	 * @return
	 */
	@RequestMapping(value = ACTION_UPDATE_DROP_DOWN_ENTRY, method = RequestMethod.POST)
	@ResponseBody()
	public SurveyQuestionsVO updateDropdownSurveyQtn(@RequestBody String body, HttpServletRequest request) {
		LOGGER.info(">>>>>>>> ConfigureSurveyQuestionController.updateDropdownSurveyQtn");
		JsonNode node = CommonUtility.getNode(body);
		int userType = node.get("userType").asInt();
		int answerType = node.get("answerType").asInt();
		String originalQuestion = node.get("originalQuestion").asText();
		String updatedMainQuestion = node.get("updatedMainQuestion").asText();
		String subQuestion = node.get("subQuestion").asText();
		String createdBy = node.get("createdBy").asText();
		String isMandatory = node.get("isMandatory").asText();
			
		SurveyQuestionsVO vo = new SurveyQuestionsVO();
		try {
			surveyService.updateDropDownSurveyQtn(userType, answerType, originalQuestion, updatedMainQuestion, subQuestion, createdBy, isMandatory);
			vo.setResultList(surveyService.fetchAllExistingSurveyQtnsByUserTypeAndQtnType (userType, answerType));
			vo.setMessage(this.messageSource.getMessage("msg.success.survey.updation",null, null));
		} catch (JCTException e) {
			vo.setMessage(this.messageSource.getMessage("msg.failure.survey.updation",null, null));
			LOGGER.error("UNABLE TO UPDATE SURVEY QUESTION: "+e.getLocalizedMessage());
		}
		LOGGER.info("<<<<<<<< ConfigureSurveyQuestionController.updateDropdownSurveyQtn");
		return vo;
	}
	/**
	 * Method updates the checkbox content of the survey question
	 * @param body
	 * @param request
	 * @return
	 */
	@RequestMapping(value = ACTION_UPDATE_CHECK_BOX_ENTRY, method = RequestMethod.POST)
	@ResponseBody()
	public SurveyQuestionsVO updateCheckboxSurveyQtn(@RequestBody String body, HttpServletRequest request) {
		LOGGER.info(">>>>>>>> ConfigureSurveyQuestionController.updateCheckboxSurveyQtn");
		JsonNode node = CommonUtility.getNode(body);
		int userType = node.get("userType").asInt();
		int answerType = node.get("answerType").asInt();
		String originalQuestion = node.get("originalQuestion").asText();
		String updatedMainQuestion = node.get("updatedMainQuestion").asText();
		String subQuestion = node.get("subQuestion").asText();
		String createdBy = node.get("createdBy").asText();
		String isMandatory = node.get("isMandatory").asText();
				
		SurveyQuestionsVO vo = new SurveyQuestionsVO();
		try {
			surveyService.updateCheckboxSurveyQtn(userType, answerType, originalQuestion, updatedMainQuestion, subQuestion, createdBy, isMandatory);
			vo.setResultList(surveyService.fetchAllExistingSurveyQtnsByUserTypeAndQtnType (userType, answerType));
			vo.setMessage(this.messageSource.getMessage("msg.success.survey.updation",null, null));
		} catch (JCTException e) {
			vo.setMessage(this.messageSource.getMessage("msg.failure.survey.updation",null, null));
			LOGGER.error("UNABLE TO UPDATE SURVEY QUESTION: "+e.getLocalizedMessage());
		}
		LOGGER.info("<<<<<<<< ConfigureSurveyQuestionController.updateCheckboxSurveyQtn");
		return vo;
	}
	/**
	 * Method fetches all existing survey question by user type and question type
	 * @param body
	 * @param request
	 * @return
	 */
	@RequestMapping(value = ACTION_FETCH_ALL_SURV_QTNS_ALL_SELECTED, method = RequestMethod.POST)
	@ResponseBody()
	public SurveyQuestionsVO fetchAllExistingSurveyQtnsAllSelected(@RequestBody String body, HttpServletRequest request) {
		LOGGER.info(">>>>>>>> ConfigureSurveyQuestionController.fetchAllExistingSurveyQtnsAllSelected");
		JsonNode node = CommonUtility.getNode(body);
		int userType = node.get("userType").asInt();
		int qtnType = node.get("qtnType").asInt();
		int profileId = node.get("profile").asInt();
		SurveyQuestionsVO vo = new SurveyQuestionsVO();
		try {
			vo.setResultList(surveyService.fetchgSurveyQtnsByAllParams(userType, qtnType, profileId));
		} catch (JCTException e) {
			LOGGER.error("UNABLE TO FETCH SURVEY QUESTION: "+e.getLocalizedMessage());
		}
		LOGGER.info("<<<<<<<< ConfigureSurveyQuestionController.fetchAllExistingSurveyQtnsAllSelected");
		return vo;
	}
	/**
	 * Method fetches all existing survey question by user type and question type
	 * @param body
	 * @param request
	 * @return
	 */
	@RequestMapping(value = ACTION_FETCH_ALL_SURV_QTNS_USER_PROFILE, method = RequestMethod.POST)
	@ResponseBody()
	public SurveyQuestionsVO fetchAllExistingSurveyQtnsOnlyUserProfile(@RequestBody String body, HttpServletRequest request) {
		LOGGER.info(">>>>>>>> ConfigureSurveyQuestionController.fetchAllExistingSurveyQtnsAllSelected");
		JsonNode node = CommonUtility.getNode(body);
		int qtnType = node.get("qtnType").asInt();
		int profileId = node.get("profile").asInt();
		SurveyQuestionsVO vo = new SurveyQuestionsVO();
		try {
			vo.setResultList(surveyService.fetchAllExistingSurveyQtnsOnlyUserProfile(qtnType, profileId));
		} catch (JCTException e) {
			LOGGER.error("UNABLE TO FETCH SURVEY QUESTION: "+e.getLocalizedMessage());
		}
		LOGGER.info("<<<<<<<< ConfigureSurveyQuestionController.fetchAllExistingSurveyQtnsAllSelected");
		return vo;
	}
	/**
	 * Method fetches all existing survey question by user type and question type
	 * @param body
	 * @param request
	 * @return
	 */
	@RequestMapping(value = ACTION_FETCH_ALL_SURV_QTNS_USER_PROFILE_TYPE, method = RequestMethod.POST)
	@ResponseBody()
	public SurveyQuestionsVO fetchSurveyQtnsByProfileAndUserType(@RequestBody String body, HttpServletRequest request) {
		LOGGER.info(">>>>>>>> ConfigureSurveyQuestionController.fetchSurveyQtnsByProfileAndUserType");
		JsonNode node = CommonUtility.getNode(body);
		int userType = node.get("userType").asInt();
		int profileId = node.get("profile").asInt();
		SurveyQuestionsVO vo = new SurveyQuestionsVO();
		try {
			vo.setResultList(surveyService.fetchSurveyQtnsByProfileAndUserType(userType, profileId));
			vo.setStatusCode(StatusConstants.STATUS_SUCCESS_WITH_RESULT);
		} catch (JCTException e) {
			vo = new SurveyQuestionsVO();
			vo.setStatusCode(StatusConstants.STATUS_FAILURE);
			LOGGER.error("UNABLE TO FETCH SURVEY QUESTION: "+e.getLocalizedMessage());
		}
		LOGGER.info("<<<<<<<< ConfigureSurveyQuestionController.fetchSurveyQtnsByProfileAndUserType");
		return vo;
	}
	
	
	/**
	 * Method update attribute list order
	 * @param body
	 * @param request
	 * @return return status code only.
	 */
	@RequestMapping(value = ACTION_SAVE_SURVEY_QTN_ORDER, method = RequestMethod.POST)
	@ResponseBody()
	public SurveyQuestionsVO saveSurvryQtnOrder(@RequestBody String body, HttpServletRequest request) throws JCTException{
		LOGGER.info(">>>>>> ConfigureSurveyQuestionController.saveSurvryQtnOrder");
		SurveyQuestionsVO vo = new SurveyQuestionsVO();		
		JsonNode node = CommonUtility.getNode(body);
		String builder = node.get("builder").toString().replaceAll("\"" , "").trim();			
		Integer userProfile = Integer.parseInt(node.get("userProfile").toString().replaceAll("\"" , ""));	
		Integer userType = Integer.parseInt(node.get("userType").toString().replaceAll("\"" , ""));	
		String builderAnsType = node.get("builderAnsType").toString().replaceAll("\"" , "").trim();
			try {
				vo = surveyService.saveSurvryQtnOrder(builder, userProfile, userType, builderAnsType);
				//SurveyQuestionsVO unused = this.surveyService.getExistingAllSurveyQuestion();
				//contentConfigVO.setExistingMappingList(unused.getExistingMappingList());
				//unused = null;
				vo.setResultList(surveyService.getExistingAllSurveyQuestion());
			} catch (JCTException e) {
				vo = new SurveyQuestionsVO();
				vo.setStatusCode(StatusConstants.STATUS_FAILURE);				
				LOGGER.error(e.getLocalizedMessage());
			}
		
		LOGGER.info("<<<<<< ContentConfigController.saveAttributeOrder");
		return vo;
	}
}
