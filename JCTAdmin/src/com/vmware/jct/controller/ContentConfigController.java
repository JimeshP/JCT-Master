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

import com.vmware.jct.common.utility.CommonUtility;
import com.vmware.jct.common.utility.StatusConstants;
import com.vmware.jct.exception.JCTException;
import com.vmware.jct.model.JctTermsAndConditions;
import com.vmware.jct.service.IContentConfigService;
import com.vmware.jct.service.vo.ContentConfigVO;
import com.vmware.jct.service.vo.InstructionVO;
import com.vmware.jct.service.vo.JobAttributeVO;
import com.vmware.jct.service.vo.QuestionearVO;
import com.vmware.jct.service.vo.TermsAndConditionsVO;

/**
 * 
 * <p><b>Class name:</b> CententConfigController.java</p>
 * <p><b>Author:</b> InterraIT</p>
 * <p><b>Purpose:</b> This class acts as a controller for all User Content Configuration on Admin Side..
 * CententConfigController extends BasicController  and has following  methods.
 * -populateExistingUserProfile()
 * <p><b>Description:</b> This class acts as a controller for all Content Configuration on Admin Side. </p>
 * <p><b>Copyrights:</b>All rights reserved by Interra IT and should only be used for its internal 
 * application development.</p>
 * <p><b>Created Date:</b> 28/Mar/2014</p>
 * <p><b>Revision History:</b>
 * 	<li></li>
 * </p>
 */
@Controller
@RequestMapping(value = "/contentconfig")
public class ContentConfigController extends BasicController {
	
	@Autowired
	private IContentConfigService service;
	
	@Autowired  
	private MessageSource messageSource;
	
	private static final Logger LOGGER = LoggerFactory.getLogger(ContentConfigController.class);
	
	/**
	 * Method populates the Reflection Question List.
	 * @param body
	 * @param request
	 * @return ContentConfigVO containing Questions.
	 */
	@RequestMapping(value = ACTION_POPULATE_EXTNG_REF_QTN, method = RequestMethod.POST)
	@ResponseBody()
	public ContentConfigVO populateExistingRefQtn(@RequestBody String body, HttpServletRequest request) {
		LOGGER.info(">>>>>> CententConfigController.populateExistingReflectionQtn");
		JsonNode node = CommonUtility.getNode(body);
		int profileId = Integer.parseInt(node.get("profileId").toString().replaceAll("\"" , ""));
		String relatedPage = node.get("relatedPage").toString().replaceAll("\"" , "");	
		ContentConfigVO contentCongifVo = new ContentConfigVO();
		try {						
			contentCongifVo.setUserProfileMap(this.service.populateUserProfile());			
			ContentConfigVO unused = this.service.populateExistingRefQtn(profileId, relatedPage);
			contentCongifVo.setExistingRefQtnList(unused.getExistingRefQtnList());
			contentCongifVo.setNoOfSubQtn(this.messageSource.
					getMessage("label.subQuestionNumber", null , null));
			unused = null;
			contentCongifVo.setStatusCode(StatusConstants.STATUS_SUCCESS_WITH_RESULT);
		} catch (JCTException e) {
			contentCongifVo = new ContentConfigVO();
			contentCongifVo.setStatusCode(StatusConstants.STATUS_FAILURE);
			contentCongifVo.setStatusDesc(this.messageSource.getMessage("exception.dao.error", null, null));
			LOGGER.error(e.getLocalizedMessage());
		}
		LOGGER.info("<<<<<< CententConfigController.populateExistingReflectionQtn");
		return contentCongifVo;
	}
	
	
	/**
	 * Method saves new reflection question
	 * @param body
	 * @param request
	 * @return return status code only.
	 */
	@RequestMapping(value = ACTION_SAVE_REF_QTN, method = RequestMethod.POST)
	@ResponseBody()
	public ContentConfigVO saveRefQtn(@RequestBody String body, HttpServletRequest request) 
					throws JCTException {
		LOGGER.info(">>>>>> CententConfigController.saveRefQtn");
		ContentConfigVO contentConfigVO = new ContentConfigVO();
		QuestionearVO questionearVO = null;
		JsonNode node = CommonUtility.getNode(body);
		String refQtnDesc = node.get("refQtnDesc").toString().replaceAll("\"" , "").toUpperCase().trim();
		String subQtnDesc1 = node.get("subQtn").toString().replaceAll("\"" , "").toUpperCase().trim();
		String subQtnDesc = node.get("subQtn").toString().replaceAll("\"" , "").trim();
		String[] splitSubQtnDesc = subQtnDesc.split("~~");
		int sunQtnLength = splitSubQtnDesc.length;		
		String userProfileValue = node.get("userProfileVal").toString().replaceAll("\"" , "");
		Integer userProfileId = Integer.parseInt(node.get("userProfileId").toString().replaceAll("\"" , ""));
		//Integer questionOrder = Integer.parseInt(node.get("questionOrder").toString().replaceAll("\"" , ""));
		String relatedPage = node.get("relatedPage").toString().replaceAll("\"" , "");
		String createdBy = node.get("createdBy").toString().replaceAll("\"" , "");
		Integer subQtnOrder = null;
		questionearVO = new QuestionearVO();
		String updationSuccess = service.validateExistenceActionPlan(refQtnDesc, subQtnDesc1, userProfileId, "BS", "ADD");
		if (updationSuccess.equals("success")) {
			contentConfigVO.setStatusCode(StatusConstants.ALREADY_EXISTS);
		} else if (updationSuccess.equals("maximum")) {
			contentConfigVO.setStatusCode(StatusConstants.MAX_LIMIT);
		} else {
			contentConfigVO.setStatusCode(StatusConstants.DOES_NOT_EXIST);
			for (int i = 0; i < sunQtnLength; i++) {
				questionearVO.setQuestionsDesc(node.get("refQtnDesc").toString().replaceAll("\"" , "").trim());								
				questionearVO.setUserProfileName(userProfileValue);
				questionearVO.setUserProfileId(userProfileId);
				questionearVO.setCreatedBy(createdBy);
				questionearVO.setNoOfSubQtn(sunQtnLength);
				//questionearVO.setQuestionOrder(questionOrder);
				if (!subQtnDesc.trim().equalsIgnoreCase("")){
					questionearVO.setQuestionSubDesc(splitSubQtnDesc[i]);
					
				} else {
					questionearVO.setQuestionSubDesc("NA");
				}
				if (i == 0) {
					//subQtnOrder = questionOrder + ".a";
					subQtnOrder = 1;
				} else if (i == 1) {
					//subQtnOrder = questionOrder + ".b";
					subQtnOrder = 2;
				} else if (i == 2) {
					//subQtnOrder = questionOrder + ".c";
					subQtnOrder = 3;
				} else if (i == 3) {
					//subQtnOrder = questionOrder + ".d";
					subQtnOrder = 4;
				}					
				questionearVO.setSubQuestionOrder(subQtnOrder);
				try {
					contentConfigVO = service.saveRefQtn(questionearVO);
					ContentConfigVO unused = service.populateExistingRefQtn(userProfileId, relatedPage);
					contentConfigVO.setExistingRefQtnList(unused.getExistingRefQtnList());
					unused = null;
				} catch (JCTException e) {
					contentConfigVO = new ContentConfigVO();
					contentConfigVO.setStatusCode(StatusConstants.STATUS_FAILURE);
					contentConfigVO.setStatusDesc(this.messageSource.
							getMessage("exception.dao.user.group.save", null, null));
					LOGGER.error(e.getLocalizedMessage());
				}			
				
			}
		}
		LOGGER.info("<<<<<< CententConfigController.saveRefQtn");
		return contentConfigVO;
	}
	/**
	 * Method update existing reflection question
	 * @param body
	 * @param request
	 * @return return status code only.
	 */
	@RequestMapping(value = ACTION_UPDATE_REF_QTN, method = RequestMethod.POST)
	@ResponseBody()
	public ContentConfigVO updateRefQtn(@RequestBody String body, HttpServletRequest request) 
			throws JCTException {
		LOGGER.info(">>>>>> ContentConfigController.updateRefQtn");				
		ContentConfigVO contentConfigVO = new ContentConfigVO();
		QuestionearVO questionearVO = null;
		JsonNode node = CommonUtility.getNode(body);
		String reflectnQtn = node.get("reflectnQtnVal").toString().replaceAll("\"" , "").toUpperCase().trim();
		String reflectnQtn1 = node.get("reflectnQtnVal").toString().replaceAll("\"" , "").trim();
		String subQtn = node.get("subQtnVal").toString().replaceAll("\"" , "").toUpperCase().trim();
		String subQtn1 = node.get("subQtnVal").toString().replaceAll("\"" , "").trim();;
		String userProfileValue = node.get("userProfileVal").toString().replaceAll("\"" , "");
		Integer userProfileId = Integer.parseInt(node.get("userProfileId").toString().replaceAll("\"" , ""));
		Integer tablePkId = Integer.parseInt(node.get("tablePkId").toString().replaceAll("\"" , ""));
		/***** ADDED FOR JCT PUBLIC VERSION *****/
		String chechBoxVal = node.get("chechBoxVal").toString().replaceAll("\"" , "").trim();
		/****** ENDED *****/
		questionearVO = new QuestionearVO();
		questionearVO.setQuestionsDesc(reflectnQtn1);
		questionearVO.setQuestionSubDesc(subQtn1);
		questionearVO.setUserProfileName(userProfileValue);
		questionearVO.setUserProfileId(userProfileId);
		questionearVO.setPrimaryKeyVal(tablePkId);
		String updationSuccess = service.validateExistenceActionPlan(reflectnQtn, subQtn, userProfileId, "BS" , "UPDATE");
		if (updationSuccess.equals("success")) {
			contentConfigVO.setStatusCode(StatusConstants.ALREADY_EXISTS);
		} else {
			contentConfigVO.setStatusCode(StatusConstants.DOES_NOT_EXIST);
			try {
				contentConfigVO = service.updateRefQtn(questionearVO, null, chechBoxVal);
				ContentConfigVO unused = service.populateExistingRefQtn(userProfileId, "BS");
				contentConfigVO.setExistingRefQtnList(unused.getExistingRefQtnList());
				unused = null;
			} catch (JCTException e) {
				contentConfigVO = new ContentConfigVO();
				contentConfigVO.setStatusCode(StatusConstants.STATUS_FAILURE);
				contentConfigVO.setStatusDesc(this.messageSource.
						getMessage("exception.dao.user.group.update", null, null));
				LOGGER.error(e.getLocalizedMessage());
			}
		}
		LOGGER.info("<<<<<< ContentConfigController.updateRefQtn");
		return contentConfigVO;
	}
	/**
	 * Method populates the reflection question List.
	 * @param body
	 * @param request
	 * @return ContentConfigVO containing reflection question.
	 */
	@RequestMapping(value = ACTION_POPULATE_USER_PROFILE_REF_QTN, method = RequestMethod.POST)
	@ResponseBody()
	public ContentConfigVO populateUserProfileRefQtn(@RequestBody String body, HttpServletRequest request) {
		LOGGER.info(">>>>>> ContentConfigController.populateUserProfileRefQtn");
		ContentConfigVO contentConfigVO = new ContentConfigVO();
		try {
			contentConfigVO.setUserProfileMap(service.populateUserProfile());
			contentConfigVO.setStatusCode(StatusConstants.STATUS_SUCCESS_WITH_RESULT);
		} catch (JCTException e) {
			contentConfigVO = new ContentConfigVO();
			contentConfigVO.setStatusCode(StatusConstants.STATUS_FAILURE);
			contentConfigVO.setStatusDesc(this.messageSource.getMessage("exception.dao.error", null, null));
			LOGGER.error(e.getLocalizedMessage());
		}
		LOGGER.info("<<<<<< ContentConfigController.populateUserProfileRefQtn");
		return contentConfigVO;
	}
	/**
	 * Method deletes reflection question
	 * @param body
	 * @param request
	 * @return return status code only.
	 */
	@RequestMapping(value = ACTION_DELETE_REF_QTN, method = RequestMethod.POST)
	@ResponseBody()
	public ContentConfigVO deleteRefQtn(@RequestBody String body, HttpServletRequest request) {
		LOGGER.info(">>>>>> ManageUserController.saveUserGroup");
		ContentConfigVO contentConfigVO = null;
		QuestionearVO questionearVO = null;
		JsonNode node = CommonUtility.getNode(body);
		Integer tablePkId = Integer.parseInt(node.get("tablePkId").toString().replaceAll("\"" , ""));
		String reflectnQtn = node.get("questnDesc").toString().replaceAll("\"" , "");
		String relatedPage = node.get("relatedPage").toString().replaceAll("\"" , "");
		Integer profileId = Integer.parseInt(node.get("profileId").toString().replaceAll("\"" , ""));
		questionearVO = new QuestionearVO();
		questionearVO.setQuestionsDesc(reflectnQtn);		
		questionearVO.setPrimaryKeyVal(tablePkId);
		try {
			contentConfigVO = service.deleteRefQtn(questionearVO, null);
			ContentConfigVO unused = service.populateExistingRefQtn(profileId, relatedPage);
			contentConfigVO.setExistingRefQtnList(unused.getExistingRefQtnList());
			unused = null;
		} catch (JCTException e) {
			contentConfigVO = new ContentConfigVO();
			contentConfigVO.setStatusCode(StatusConstants.STATUS_FAILURE);
			contentConfigVO.setStatusDesc(this.messageSource.
					getMessage("exception.dao.user.group.update", null, null));
			LOGGER.error(e.getLocalizedMessage());
		}
		LOGGER.info("<<<<<< ContentConfigController.updateRefQtn");
		return contentConfigVO;		
	}
	/**
	 * Method saves new action plan
	 * @param body
	 * @param request
	 * @return return status code only.
	 */
	@RequestMapping(value = ACTION_SAVE_ACTION_PLAN, method = RequestMethod.POST)
	@ResponseBody()
	public ContentConfigVO saveActionPlan(@RequestBody String body, HttpServletRequest request) throws JCTException {
		LOGGER.info(">>>>>> CententConfigController.saveActionPlan");
		ContentConfigVO contentConfigVO = new ContentConfigVO();
		QuestionearVO questionearVO = null;
		JsonNode node = CommonUtility.getNode(body);
		String refQtnDesc = node.get("refQtnDesc").toString().replaceAll("\"" , "").toUpperCase().trim();
		String subQtnDesc1 = node.get("subQtn").toString().replaceAll("\"" , "").toUpperCase().trim();
		String subQtnDesc = node.get("subQtn").toString().replaceAll("\"" , "").trim();
		String[] splitSubQtnDesc = subQtnDesc.split("~~");
		int sunQtnLength = splitSubQtnDesc.length;		
		String userProfileValue = node.get("userProfileVal").toString().replaceAll("\"" , "");
		Integer userProfileId = Integer.parseInt(node.get("userProfileId").toString().replaceAll("\"" , ""));
		//Integer questionOrder = Integer.parseInt(node.get("questionOrder").toString().replaceAll("\"" , ""));
		String relatedPage = node.get("relatedPage").toString().replaceAll("\"" , "");
		String createdBy = node.get("createdBy").toString().replaceAll("\"" , "");
		questionearVO = new QuestionearVO();
		Integer subQtnOrder = null;
		String updationSuccess = service.validateExistenceActionPlan(refQtnDesc, subQtnDesc1, userProfileId, "AS", "ADD");
		if (updationSuccess.equals("success")) {
			contentConfigVO.setStatusCode(StatusConstants.ALREADY_EXISTS);
		} else if (updationSuccess.equals("maximum")) {
			contentConfigVO.setStatusCode(StatusConstants.MAX_LIMIT);
		} else {
			contentConfigVO.setStatusCode(StatusConstants.DOES_NOT_EXIST);
			for (int i = 0; i < sunQtnLength; i++) {
				questionearVO.setQuestionsDesc(node.get("refQtnDesc").
						toString().replaceAll("\"" , "").trim());		
				questionearVO.setUserProfileName(userProfileValue);
				questionearVO.setUserProfileId(userProfileId);
				questionearVO.setCreatedBy(createdBy);
				questionearVO.setNoOfSubQtn(sunQtnLength);
				//questionearVO.setQuestionOrder(questionOrder);
				if (!subQtnDesc.trim().equalsIgnoreCase("")){
					questionearVO.setQuestionSubDesc(splitSubQtnDesc[i]);					
				} else {
					questionearVO.setQuestionSubDesc("NA");
				}
				if (i == 0) {
					//subQtnOrder = questionOrder + ".a";
					subQtnOrder = 1;
				} else if (i == 1) {
					//subQtnOrder = questionOrder + ".b";
					subQtnOrder = 2;
				} else if (i == 2) {
					//subQtnOrder = questionOrder + ".c";
					subQtnOrder = 3;
				} else if (i == 3) {
					//subQtnOrder = questionOrder + ".d";
					subQtnOrder = 4;
				}
				questionearVO.setSubQuestionOrder(subQtnOrder);
				try {
					contentConfigVO = service.saveActionPlan(questionearVO);
					ContentConfigVO unused = service.populateExistingRefQtn(userProfileId, relatedPage);
					contentConfigVO.setExistingRefQtnList(unused.getExistingRefQtnList());
					unused = null;
				} catch (JCTException e) {
					contentConfigVO = new ContentConfigVO();
					contentConfigVO.setStatusCode(StatusConstants.STATUS_FAILURE);
					contentConfigVO.setStatusDesc(this.messageSource.
							getMessage("exception.dao.user.group.save", null, null));
					LOGGER.error(e.getLocalizedMessage());
				}			
				
			}
		}
		LOGGER.info("<<<<<< CententConfigController.saveActionPlan");
		return contentConfigVO;
	}
	/**
	 * Method update existing Action Plan
	 * @param body
	 * @param request
	 * @return return status code only.
	 */
	@RequestMapping(value = ACTION_UPDATE_ACTION_PLAN, method = RequestMethod.POST)
	@ResponseBody()
	public ContentConfigVO updateActionPlan(@RequestBody String body, HttpServletRequest request) throws JCTException{
		LOGGER.info(">>>>>> ContentConfigController.updateActionPlan");
		ContentConfigVO contentConfigVO = new ContentConfigVO();
		QuestionearVO questionearVO = null;
		JsonNode node = CommonUtility.getNode(body);
		String reflectnQtn = node.get("reflectnQtnVal").toString().replaceAll("\"" , "").toUpperCase().trim();
		String reflectnQtn1 = node.get("reflectnQtnVal").toString().replaceAll("\"" , "").trim();
		String subQtn = node.get("subQtnVal").toString().replaceAll("\"" , "").toUpperCase().trim();
		String subQtn1 = node.get("subQtnVal").toString().replaceAll("\"" , "").trim();;
		String userProfileValue = node.get("userProfileVal").toString().replaceAll("\"" , "");
		Integer userProfileId = Integer.parseInt(node.get("userProfileId").toString().replaceAll("\"" , ""));
		Integer tablePkId = Integer.parseInt(node.get("tablePkId").toString().replaceAll("\"" , ""));
		/***** ADDED FOR JCT PUBLIC VERSION *****/
		String chechBoxVal = node.get("chechBoxVal").toString().replaceAll("\"" , "").trim();
		/****** ENDED *****/
		questionearVO = new QuestionearVO();
		questionearVO.setQuestionsDesc(reflectnQtn1);
		questionearVO.setQuestionSubDesc(subQtn1);
		questionearVO.setUserProfileName(userProfileValue);
		questionearVO.setUserProfileId(userProfileId);
		questionearVO.setPrimaryKeyVal(tablePkId);
		String updationSuccess = service.validateExistenceActionPlan(reflectnQtn, subQtn, userProfileId, "AS", "UPDATE");
		if (updationSuccess.equals("success")) {
			contentConfigVO.setStatusCode(StatusConstants.ALREADY_EXISTS);
		} else {
			contentConfigVO.setStatusCode(StatusConstants.DOES_NOT_EXIST);
			try {
				contentConfigVO = service.updateActionPlan(questionearVO, null, chechBoxVal);
				ContentConfigVO unused = service.populateExistingRefQtn(userProfileId, "AS");
				contentConfigVO.setExistingRefQtnList(unused.getExistingRefQtnList());
				unused = null;
			} catch (JCTException e) {
				contentConfigVO = new ContentConfigVO();
				contentConfigVO.setStatusCode(StatusConstants.STATUS_FAILURE);
				contentConfigVO.setStatusDesc(this.messageSource.
						getMessage("exception.dao.user.group.update", null, null));
				LOGGER.error(e.getLocalizedMessage());
			}
		}
		LOGGER.info("<<<<<< ContentConfigController.updateActionPlan");
		return contentConfigVO;
	}
	/**
	 * Method deletes delete Action Plan
	 * @param body
	 * @param request
	 * @return return status code only.
	 */
	@RequestMapping(value = ACTION_DELETE_ACTION_PLAN, method = RequestMethod.POST)
	@ResponseBody()
	public ContentConfigVO deleteActionPlan(@RequestBody String body, HttpServletRequest request) {
		LOGGER.info(">>>>>> ManageUserController.deleteActionPlan");
		ContentConfigVO contentConfigVO = null;
		QuestionearVO questionearVO = null;
		JsonNode node = CommonUtility.getNode(body);
		Integer tablePkId = Integer.parseInt(node.get("tablePkId").toString().replaceAll("\"" , ""));
		String reflectnQtn = node.get("questnDesc").toString().replaceAll("\"" , "");
		Integer profileId = Integer.parseInt(node.get("profileId").toString().replaceAll("\"" , ""));
		String relatedPage = node.get("relatedPage").toString().replaceAll("\"" , "");
		questionearVO = new QuestionearVO();
		questionearVO.setQuestionsDesc(reflectnQtn);		
		questionearVO.setPrimaryKeyVal(tablePkId);
		try {
			contentConfigVO = service.deleteActionPlan(questionearVO, null);
			ContentConfigVO unused = service.populateExistingRefQtn(profileId, relatedPage);
			contentConfigVO.setExistingRefQtnList(unused.getExistingRefQtnList());
			unused = null;
		} catch (JCTException e) {
			contentConfigVO = new ContentConfigVO();
			contentConfigVO.setStatusCode(StatusConstants.STATUS_FAILURE);
			contentConfigVO.setStatusDesc(this.messageSource.
					getMessage("exception.dao.user.group.update", null, null));
			LOGGER.error(e.getLocalizedMessage());
		}
		LOGGER.info("<<<<<< ContentConfigController.deleteActionPlan");
		return contentConfigVO;		
	}
	/**
	 * Method populates the Function Group List.
	 * @param body
	 * @param request
	 * @return ContentConfigVO containing Function Group.
	 */
	@RequestMapping(value = ACTION_POPULATE_EXTNG_FUNC_GRP, method = RequestMethod.POST)
	@ResponseBody()
	public ContentConfigVO populateExistingFunctionGroup(@RequestBody String body, HttpServletRequest request) {
		LOGGER.info(">>>>>> CententConfigController.populateExistingFunctionGroup");
		ContentConfigVO contentCongifVo = new ContentConfigVO();
		try {											
			contentCongifVo = service.populateExistingFunctionGroup();
		} catch (JCTException e) {
			contentCongifVo = new ContentConfigVO();
			contentCongifVo.setStatusCode(StatusConstants.STATUS_FAILURE);
			contentCongifVo.setStatusDesc(this.messageSource.getMessage("exception.dao.error", null, null));
			LOGGER.error(e.getLocalizedMessage());
		}
		LOGGER.info("<<<<<< CententConfigController.populateExistingFunctionGroup");
		return contentCongifVo;
	}
	/**
	 * Method populates the Job Level List.
	 * @param body
	 * @param request
	 * @return ContentConfigVO containing Job Level.
	 */
	@RequestMapping(value = ACTION_POPULATE_EXTNG_JOB_LEVEL, method = RequestMethod.POST)
	@ResponseBody()
	public ContentConfigVO populateExistingJobLevel(@RequestBody String body, HttpServletRequest request) {
		LOGGER.info(">>>>>> CententConfigController.populateExistingJobLevel");
		ContentConfigVO contentCongifVo = new ContentConfigVO();
		try {						
			contentCongifVo = service.populateExistingJobLevel();
		} catch (JCTException e) {
			contentCongifVo = new ContentConfigVO();
			contentCongifVo.setStatusCode(StatusConstants.STATUS_FAILURE);
			contentCongifVo.setStatusDesc(this.messageSource.getMessage("exception.dao.error", null, null));
			LOGGER.error(e.getLocalizedMessage());
		}
		LOGGER.info("<<<<<< CententConfigController.populateExistingJobLevel");
		return contentCongifVo;
	}
	/**
	 * Method saves new function group
	 * @param body
	 * @param request
	 * @return return status code only.
	 */
	@RequestMapping(value = ACTION_SAVE_FUNC_GRP, method = RequestMethod.POST)
	@ResponseBody()
	public ContentConfigVO saveFunctionGroup(@RequestBody String body, HttpServletRequest request) 
			throws JCTException {
		LOGGER.info(">>>>>> CententConfigController.saveFunctionGroup");
		ContentConfigVO contentConfigVO = new ContentConfigVO();
		JsonNode node = CommonUtility.getNode(body);
		String functionGrp = node.get("functionGrpVal").toString().replaceAll("\"" , "").toUpperCase().trim();
		String createdBy = node.get("createdBy").toString().replaceAll("\"" , "");
		String updationSuccess = service.validateExistenceFuncGrp(functionGrp);
		if (updationSuccess.equals("success")) {
			contentConfigVO.setStatusCode(StatusConstants.ALREADY_EXISTS);
		} else {
			contentConfigVO.setStatusCode(StatusConstants.DOES_NOT_EXIST);
			try {
				contentConfigVO = service.saveFunctionGroup(node.get("functionGrpVal").toString().replaceAll("\"" , "").trim(), createdBy);
				ContentConfigVO unused = service.populateExistingFunctionGroup();
				contentConfigVO.setExistingFunctionGrpList(unused.getExistingFunctionGrpList());
				unused = null;
			} catch (JCTException e) {
				contentConfigVO = new ContentConfigVO();
				contentConfigVO.setStatusCode(StatusConstants.STATUS_FAILURE);
				contentConfigVO.setStatusDesc(this.messageSource.
						getMessage("exception.dao.user.profile.save", null, null));
				LOGGER.error(e.getLocalizedMessage());
			}
		}
		LOGGER.info("<<<<<< CententConfigController.saveFunctionGroup");
		return contentConfigVO;
	}
	/**
	 * Method saves new function group
	 * @param body
	 * @param request
	 * @return return status code only.
	 */
	@RequestMapping(value = ACTION_SAVE_JOB_LEVEL, method = RequestMethod.POST)
	@ResponseBody()
	public ContentConfigVO saveJobLevel(@RequestBody String body, HttpServletRequest request) 
			throws JCTException {
		LOGGER.info(">>>>>> CententConfigController.saveJobLevel");
		ContentConfigVO contentConfigVO = new ContentConfigVO();
		JsonNode node = CommonUtility.getNode(body);
		String jobLevel = node.get("jobLevelVal").toString().replaceAll("\"" , "").toUpperCase().trim();
		String jobLevel1 = node.get("jobLevelVal").toString().replaceAll("\"" , "").trim();
		String createdBy = node.get("createdBy").toString().replaceAll("\"" , "");
		String updationSuccess = service.validateExistenceJobLevel(jobLevel);
		if (updationSuccess.equals("success")) {
			contentConfigVO.setStatusCode(StatusConstants.ALREADY_EXISTS);
		} else {
			contentConfigVO.setStatusCode(StatusConstants.DOES_NOT_EXIST);
			try {
				contentConfigVO = service.saveJobLevel(jobLevel1, createdBy);
				ContentConfigVO unused = service.populateExistingJobLevel();
				contentConfigVO.setExistingJobLevelList(unused.getExistingJobLevelList());
				unused = null;
			} catch (JCTException e) {
				contentConfigVO = new ContentConfigVO();
				contentConfigVO.setStatusCode(StatusConstants.STATUS_FAILURE);
				contentConfigVO.setStatusDesc(this.messageSource.
						getMessage("exception.dao.user.profile.save", null, null));
				LOGGER.error(e.getLocalizedMessage());
			}
		}
		LOGGER.info("<<<<<< CententConfigController.saveJobLevel");
		return contentConfigVO;
	}
	/**
	 * Method populates the Region List.
	 * @param body
	 * @param request
	 * @return ContentConfigVO containing Region.
	 */
	@RequestMapping(value = ACTION_POPULATE_EXTNG_REGION, method = RequestMethod.POST)
	@ResponseBody()
	public ContentConfigVO populateExistingRegion(@RequestBody String body, HttpServletRequest request) {
		LOGGER.info(">>>>>> CententConfigController.populateExistingRegion");
		ContentConfigVO contentCongifVo = new ContentConfigVO();
		try {									
			contentCongifVo = service.populateExistingRegion();
		} catch (JCTException e) {
			contentCongifVo = new ContentConfigVO();
			contentCongifVo.setStatusCode(StatusConstants.STATUS_FAILURE);
			contentCongifVo.setStatusDesc(this.messageSource.getMessage("exception.dao.error", null, null));
			LOGGER.error(e.getLocalizedMessage());
		}
		LOGGER.info("<<<<<< CententConfigController.populateExistingRegion");
		return contentCongifVo;
	}
	/**
	 * Method saves new region
	 * @param body
	 * @param request
	 * @return return status code only.
	 */
	@RequestMapping(value = ACTION_SAVE_REGION, method = RequestMethod.POST)
	@ResponseBody()
	public ContentConfigVO saveRegion(@RequestBody String body, HttpServletRequest request) 
				 throws JCTException {
		LOGGER.info(">>>>>> CententConfigController.saveRegion");
		ContentConfigVO contentConfigVO = new ContentConfigVO();
		JsonNode node = CommonUtility.getNode(body);
		String regionName = node.get("regionNameVal").toString().replaceAll("\"" , "").toUpperCase().trim();
		String regionName1 = node.get("regionNameVal").toString().replaceAll("\"" , "").trim();
		String createdBy = node.get("createdBy").toString().replaceAll("\"" , "");
		String updationSuccess = service.validateExistenceRegion(regionName);
		if (updationSuccess.equals("success")) {
			contentConfigVO.setStatusCode(StatusConstants.ALREADY_EXISTS);
		} else {
			contentConfigVO.setStatusCode(StatusConstants.DOES_NOT_EXIST);
			try {
				contentConfigVO = service.saveRegion(regionName1, createdBy);
				ContentConfigVO unused = service.populateExistingRegion();
				contentConfigVO.setExistingRegionList(unused.getExistingRegionList());
				unused = null;
			} catch (JCTException e) {
				contentConfigVO = new ContentConfigVO();
				contentConfigVO.setStatusCode(StatusConstants.STATUS_FAILURE);
				contentConfigVO.setStatusDesc(this.messageSource.
						getMessage("exception.dao.user.profile.save", null, null));
				LOGGER.error(e.getLocalizedMessage());
			}
		}
		LOGGER.info("<<<<<< CententConfigController.saveRegion");
		return contentConfigVO;
	}
	/**
	 * Method update existing region data
	 * @param body
	 * @param request
	 * @return return status code only.
	 */
	@RequestMapping(value = ACTION_UPDATE_REGION, method = RequestMethod.POST)
	@ResponseBody()
	public ContentConfigVO updateRegion(@RequestBody String body, HttpServletRequest request) 
			throws JCTException {
		LOGGER.info(">>>>>> ContentConfigController.updateRegion");
		ContentConfigVO contentConfigVO = new ContentConfigVO();		
		JsonNode node = CommonUtility.getNode(body);
		String regionDesc = node.get("regionDescVal").toString().replaceAll("\"" , "").toUpperCase().trim();
		String regionDesc1 = node.get("regionDescVal").toString().replaceAll("\"" , "").trim();
		Integer tablePkId = Integer.parseInt(node.get("tablePkId").toString().replaceAll("\"" , ""));
		String updationSuccess = service.validateExistenceRegion(regionDesc);
		if (updationSuccess.equals("success")) {
			contentConfigVO.setStatusCode(StatusConstants.ALREADY_EXISTS);
		} else {
			contentConfigVO.setStatusCode(StatusConstants.DOES_NOT_EXIST);
			try {
				contentConfigVO = service.updateRegion(regionDesc1, tablePkId);
				ContentConfigVO unused = service.populateExistingRegion();
				contentConfigVO.setExistingRegionList(unused.getExistingRegionList());
				unused = null;
			} catch (JCTException e) {
				contentConfigVO = new ContentConfigVO();
				contentConfigVO.setStatusCode(StatusConstants.STATUS_FAILURE);
				contentConfigVO.setStatusDesc(this.messageSource.
						getMessage("exception.dao.user.group.update", null, null));
				LOGGER.error(e.getLocalizedMessage());
			}
		}
		LOGGER.info("<<<<<< ContentConfigController.updateRegion");
		return contentConfigVO;
	}
	/**
	 * Method to delete existing region data
	 * @param body
	 * @param request
	 * @return return status code only.
	 */
	@RequestMapping(value = ACTION_DELETE_REGION, method = RequestMethod.POST)
	@ResponseBody()
	public ContentConfigVO deleteRegion(@RequestBody String body, HttpServletRequest request) {
		LOGGER.info(">>>>>> ContentConfigController.deleteRegion");
		ContentConfigVO contentConfigVO = null;		
		JsonNode node = CommonUtility.getNode(body);		
		Integer tablePkId = Integer.parseInt(node.get("tablePkId").toString().replaceAll("\"" , ""));
		String regionDesc = node.get("regionDesc").toString().replaceAll("\"" , "");
		try {
			contentConfigVO = service.deleteRegion(regionDesc, tablePkId);
			ContentConfigVO unused = service.populateExistingRegion();
			contentConfigVO.setExistingRegionList(unused.getExistingRegionList());
			unused = null;
		} catch (JCTException e) {
			contentConfigVO = new ContentConfigVO();
			contentConfigVO.setStatusCode(StatusConstants.STATUS_FAILURE);
			contentConfigVO.setStatusDesc(this.messageSource.
					getMessage("exception.dao.user.group.update", null, null));
			LOGGER.error(e.getLocalizedMessage());
		}
		LOGGER.info("<<<<<< ContentConfigController.deleteRegion");
		return contentConfigVO;
	}
	/**
	 * Method populates the Mapping List.
	 * @param body
	 * @param request
	 * @return ContentConfigVO mapping values.
	 */
	@RequestMapping(value = ACTION_POPULATE_EXTNG_MAPPING, method = RequestMethod.POST)
	@ResponseBody()
	public ContentConfigVO populateExistingMapping(@RequestBody String body, HttpServletRequest request) {
		LOGGER.info(">>>>>> CententConfigController.populateExistingMapping");
		JsonNode node = CommonUtility.getNode(body);
		int profileId = Integer.parseInt(node.get("profileId").toString().replaceAll("\"" , ""));
		String relatedPage = node.get("relatedPage").toString().replaceAll("\"" , "");	
		ContentConfigVO contentCongifVo = new ContentConfigVO();
		try {						
			contentCongifVo.setUserProfileMap(this.service.populateUserProfile());			
			ContentConfigVO unused = this.service.populateExistingMapping(profileId, relatedPage);
			contentCongifVo.setExistingMappingList(unused.getExistingMappingList());			
			unused = null;
			contentCongifVo.setStatusCode(StatusConstants.STATUS_SUCCESS_WITH_RESULT);
		} catch (JCTException e) {
			contentCongifVo = new ContentConfigVO();
			contentCongifVo.setStatusCode(StatusConstants.STATUS_FAILURE);
			contentCongifVo.setStatusDesc(this.messageSource.getMessage("exception.dao.error",null, null));
			LOGGER.error(e.getLocalizedMessage());
		}
		LOGGER.info("<<<<<< CententConfigController.populateExistingMapping");
		return contentCongifVo;
	}
	/**
	 * Method saves new Strength
	 * @param body
	 * @param request
	 * @return return status code only.
	 */
	@RequestMapping(value = ACTION_SAVE_STRENGTH, method = RequestMethod.POST)
	@ResponseBody()
	public ContentConfigVO saveStrength(@RequestBody String body, HttpServletRequest request) 
			throws JCTException {
		LOGGER.info(">>>>>> CententConfigController.saveStrength");
		ContentConfigVO contentConfigVO = new ContentConfigVO();
		JobAttributeVO jobAttributeVO = null;
		JsonNode node = CommonUtility.getNode(body);
		String attrDescText = node.get("attrDescTextVal").toString().replaceAll("\"" , "").toUpperCase().trim();
		String attrDescText1 = node.get("attrDescTextVal").toString().replaceAll("\"" , "").trim();
		String userProfileValue = node.get("userProfileVal").toString().replaceAll("\"" , "");
		Integer userProfileId = Integer.parseInt(node.get("userProfileId").toString().replaceAll("\"" , ""));
		String attrFullDescText = node.get("attrFullDescText").toString().replaceAll("\"" , "").trim();
		//Integer attributeOrder = Integer.parseInt(node.get("attributeOrder").toString().replaceAll("\"" , ""));
		String createdBy = node.get("createdBy").toString().replaceAll("\"" , "");
		jobAttributeVO = new JobAttributeVO();
		jobAttributeVO.setJctJobAttributeName(attrDescText1);
		jobAttributeVO.setJctJobAttributeDesc(attrFullDescText);
		jobAttributeVO.setJctJobAttributeCode("STR");	
		//jobAttributeVO.setJctJobAttributeOrder(attributeOrder);
		jobAttributeVO.setUserProfileName(userProfileValue);
		jobAttributeVO.setUserProfileId(userProfileId);
		jobAttributeVO.setCreatedBy(createdBy);
		String updationSuccess = service.validateExistenceAttribute(attrDescText, userProfileId, "STR", "ADD");
		if (updationSuccess.equals("success")) {
			contentConfigVO.setStatusCode(StatusConstants.ALREADY_EXISTS);
		} else if (updationSuccess.equals("maximum")) {
			contentConfigVO.setStatusCode(StatusConstants.MAX_LIMIT);
		} else {
			contentConfigVO.setStatusCode(StatusConstants.DOES_NOT_EXIST);
			try {
				contentConfigVO = service.saveStrength(jobAttributeVO);
				ContentConfigVO unused = service.populateExistingMapping(userProfileId, "STR");
				contentConfigVO.setExistingMappingList(unused.getExistingMappingList());
				unused = null;
			} catch (JCTException e) {
				contentConfigVO = new ContentConfigVO();
				contentConfigVO.setStatusCode(StatusConstants.STATUS_FAILURE);
				contentConfigVO.setStatusDesc(this.messageSource.
						getMessage("exception.dao.user.group.save", null, null));
				LOGGER.error(e.getLocalizedMessage());
			}
		}
		LOGGER.info("<<<<<< CententConfigController.saveStrength");
		return contentConfigVO;
	}
	/**
	 * Method saves new Values
	 * @param body
	 * @param request
	 * @return return status code only.
	 */
	@RequestMapping(value = ACTION_SAVE_VALUE, method = RequestMethod.POST)
	@ResponseBody()
	public ContentConfigVO saveValue(@RequestBody String body, HttpServletRequest request) 
			throws JCTException {
		LOGGER.info(">>>>>> CententConfigController.saveValue");
		ContentConfigVO contentConfigVO = new ContentConfigVO();
		JobAttributeVO jobAttributeVO = null;
		JsonNode node = CommonUtility.getNode(body);
		String attrDescText = node.get("attrDescTextVal").toString().replaceAll("\"" , "").toUpperCase().trim();
		String attrDescText1 = node.get("attrDescTextVal").toString().replaceAll("\"" , "").trim();
		String userProfileValue = node.get("userProfileVal").toString().replaceAll("\"" , "");
		Integer userProfileId = Integer.parseInt(node.get("userProfileId").toString().replaceAll("\"" , ""));
		String attrFullDescText = node.get("attrFullDescText").toString().replaceAll("\"" , "").trim();
		//Integer attributeOrder = Integer.parseInt(node.get("attributeOrder").toString().replaceAll("\"" , ""));
		String createdBy = node.get("createdBy").toString().replaceAll("\"" , "");
		jobAttributeVO = new JobAttributeVO();
		jobAttributeVO.setJctJobAttributeName(attrDescText1);
		jobAttributeVO.setJctJobAttributeDesc(attrFullDescText);
		jobAttributeVO.setJctJobAttributeCode("VAL");	
		//jobAttributeVO.setJctJobAttributeOrder(attributeOrder);
		jobAttributeVO.setUserProfileName(userProfileValue);
		jobAttributeVO.setUserProfileId(userProfileId);
		jobAttributeVO.setCreatedBy(createdBy);
		String updationSuccess = service.validateExistenceAttribute(attrDescText, userProfileId, "VAL", "ADD");
		if (updationSuccess.equals("success")) {
			contentConfigVO.setStatusCode(StatusConstants.ALREADY_EXISTS);
		} else if (updationSuccess.equals("maximum")) {
			contentConfigVO.setStatusCode(StatusConstants.MAX_LIMIT);
		} else {
			contentConfigVO.setStatusCode(StatusConstants.DOES_NOT_EXIST);
			try {
				contentConfigVO = service.saveValue(jobAttributeVO);
				ContentConfigVO unused = service.populateExistingMapping(userProfileId, "VAL");
				contentConfigVO.setExistingMappingList(unused.getExistingMappingList());
				unused = null;
			} catch (JCTException e) {
				contentConfigVO = new ContentConfigVO();
				contentConfigVO.setStatusCode(StatusConstants.STATUS_FAILURE);
				contentConfigVO.setStatusDesc(this.messageSource.
						getMessage("exception.dao.user.group.save", null, null));
				LOGGER.error(e.getLocalizedMessage());
			}
		}
		LOGGER.info("<<<<<< CententConfigController.saveValue");
		return contentConfigVO;
	}
	/**
	 * Method saves new pssions
	 * @param body
	 * @param request
	 * @return return status code only.
	 */
	@RequestMapping(value = ACTION_SAVE_PASSION, method = RequestMethod.POST)
	@ResponseBody()
	public ContentConfigVO savePassion(@RequestBody String body, HttpServletRequest request) 
			throws JCTException{
		LOGGER.info(">>>>>> CententConfigController.savePassion");
		ContentConfigVO contentConfigVO = new ContentConfigVO();
		JobAttributeVO jobAttributeVO = null;
		JsonNode node = CommonUtility.getNode(body);
		String attrDescText = node.get("attrDescTextVal").toString().replaceAll("\"" , "").toUpperCase().trim();
		String attrDescText1 = node.get("attrDescTextVal").toString().replaceAll("\"" , "").trim();
		String userProfileValue = node.get("userProfileVal").toString().replaceAll("\"" , "");
		Integer userProfileId = Integer.parseInt(node.get("userProfileId").toString().replaceAll("\"" , ""));
		String attrFullDescText = node.get("attrFullDescText").toString().replaceAll("\"" , "").trim();
		//Integer attributeOrder = Integer.parseInt(node.get("attributeOrder").toString().replaceAll("\"" , ""));
		String createdBy = node.get("createdBy").toString().replaceAll("\"" , "");
		jobAttributeVO = new JobAttributeVO();
		jobAttributeVO.setJctJobAttributeName(attrDescText1);
		jobAttributeVO.setJctJobAttributeDesc(attrFullDescText);
		jobAttributeVO.setJctJobAttributeCode("PAS");		
		//jobAttributeVO.setJctJobAttributeOrder(attributeOrder);
		jobAttributeVO.setUserProfileName(userProfileValue);
		jobAttributeVO.setUserProfileId(userProfileId);
		jobAttributeVO.setCreatedBy(createdBy);
		String updationSuccess = service.validateExistenceAttribute(attrDescText, userProfileId, "PAS", "ADD");
		if (updationSuccess.equals("success")) {
			contentConfigVO.setStatusCode(StatusConstants.ALREADY_EXISTS);
		} else if (updationSuccess.equals("maximum")) {
			contentConfigVO.setStatusCode(StatusConstants.MAX_LIMIT);
		} else {
			contentConfigVO.setStatusCode(StatusConstants.DOES_NOT_EXIST);
			try {
				contentConfigVO = service.savePassion(jobAttributeVO);
				ContentConfigVO unused = service.populateExistingMapping(userProfileId, "PAS");
				contentConfigVO.setExistingMappingList(unused.getExistingMappingList());
				unused = null;
			} catch (JCTException e) {
				contentConfigVO = new ContentConfigVO();
				contentConfigVO.setStatusCode(StatusConstants.STATUS_FAILURE);
				contentConfigVO.setStatusDesc(this.messageSource.
						getMessage("exception.dao.user.group.save", null, null));
				LOGGER.error(e.getLocalizedMessage());
			}
		}
		LOGGER.info("<<<<<< CententConfigController.savePassion");
		return contentConfigVO;
	}
	/**
	 * Method update existing strengths
	 * @param body
	 * @param request
	 * @return return status code only.
	 */
	@RequestMapping(value = ACTION_UPDATE_STRENGTH, method = RequestMethod.POST)
	@ResponseBody()
	public ContentConfigVO updateStrength(@RequestBody String body, HttpServletRequest request) 
			throws JCTException {
		LOGGER.info(">>>>>> ContentConfigController.updateStrength");
		ContentConfigVO contentConfigVO = new ContentConfigVO();
		JobAttributeVO jobAttributeVO = null;
		JsonNode node = CommonUtility.getNode(body);
		//String atttrDesc = node.get("atttrDescVal").toString().replaceAll("\"" , "").toUpperCase().trim();
		String atttrDesc1 = node.get("atttrDescVal").toString().replaceAll("\"" , "").trim();
		String userProfileValue = node.get("userProfileVal").toString().replaceAll("\"" , "");
		Integer userProfileId = Integer.parseInt(node.get("userProfileId").toString().replaceAll("\"" , ""));
		Integer tablePkId = Integer.parseInt(node.get("tablePkId").toString().replaceAll("\"" , ""));
		String attrFullDesc = node.get("attrFullDesc").toString().replaceAll("\"" , "").trim();
		/***** ADDED FOR JCT PUBLIC VERSION *****/
		String chechBoxVal = node.get("chechBoxVal").toString().replaceAll("\"" , "").trim();
		/****** ENDED *****/
		jobAttributeVO = new JobAttributeVO();
		jobAttributeVO.setJctJobAttributeName(atttrDesc1);
		jobAttributeVO.setJctJobAttributeDesc(attrFullDesc);
		jobAttributeVO.setUserProfileName(userProfileValue);
		jobAttributeVO.setUserProfileId(userProfileId);
		jobAttributeVO.setPrimaryKeyVal(tablePkId);		
		contentConfigVO.setStatusCode(StatusConstants.DOES_NOT_EXIST);
			try {
				contentConfigVO = service.updateStrength(jobAttributeVO, null, chechBoxVal);
				ContentConfigVO unused = service.populateExistingMapping(userProfileId, "STR");
				contentConfigVO.setExistingMappingList(unused.getExistingMappingList());
				unused = null;
			} catch (JCTException e) {
				contentConfigVO = new ContentConfigVO();
				contentConfigVO.setStatusCode(StatusConstants.STATUS_FAILURE);
				contentConfigVO.setStatusDesc(this.messageSource.
						getMessage("exception.dao.user.group.update", null, null));
				LOGGER.error(e.getLocalizedMessage());
			}
		LOGGER.info("<<<<<< ContentConfigController.updateStrength");
		return contentConfigVO;
	}
	/**
	 * Method update existing values
	 * @param body
	 * @param request
	 * @return return status code only.
	 */
	@RequestMapping(value = ACTION_UPDATE_VALUE, method = RequestMethod.POST)
	@ResponseBody()
	public ContentConfigVO updateValue(@RequestBody String body, HttpServletRequest request) 
			throws JCTException {
		LOGGER.info(">>>>>> ContentConfigController.updateValue");
		ContentConfigVO contentConfigVO = new ContentConfigVO();
		JobAttributeVO jobAttributeVO = null;
		JsonNode node = CommonUtility.getNode(body);
		//String atttrDesc = node.get("atttrDescVal").toString().replaceAll("\"" , "").toUpperCase().trim();
		String atttrDesc1 = node.get("atttrDescVal").toString().replaceAll("\"" , "").trim();
		String userProfileValue = node.get("userProfileVal").toString().replaceAll("\"" , "");
		Integer userProfileId = Integer.parseInt(node.get("userProfileId").toString().replaceAll("\"" , ""));
		Integer tablePkId = Integer.parseInt(node.get("tablePkId").toString().replaceAll("\"" , ""));
		String attrFullDesc = node.get("attrFullDesc").toString().replaceAll("\"" , "").trim();
		/***** ADDED FOR JCT PUBLIC VERSION *****/
		String chechBoxVal = node.get("chechBoxVal").toString().replaceAll("\"" , "").trim();
		/****** ENDED *****/
		jobAttributeVO = new JobAttributeVO();
		jobAttributeVO.setJctJobAttributeName(atttrDesc1);
		jobAttributeVO.setJctJobAttributeDesc(attrFullDesc);
		//jobAttributeVO.setJctJobAttributeOrder(newAttrOrder);
		jobAttributeVO.setUserProfileName(userProfileValue);
		jobAttributeVO.setUserProfileId(userProfileId);
		jobAttributeVO.setPrimaryKeyVal(tablePkId);				
		contentConfigVO.setStatusCode(StatusConstants.DOES_NOT_EXIST);
		try {
			contentConfigVO = service.updateValue(jobAttributeVO, null, chechBoxVal);
			ContentConfigVO unused = service.populateExistingMapping(userProfileId, "VAL");
			contentConfigVO.setExistingMappingList(unused.getExistingMappingList());
			unused = null;
		} catch (JCTException e) {
			contentConfigVO = new ContentConfigVO();
			contentConfigVO.setStatusCode(StatusConstants.STATUS_FAILURE);
			contentConfigVO.setStatusDesc(this.messageSource.
					getMessage("exception.dao.user.group.update", null, null));
			LOGGER.error(e.getLocalizedMessage());
		}
		//}
		LOGGER.info("<<<<<< ContentConfigController.updateValue");
		return contentConfigVO;
	}
	/**
	 * Method update existing passions
	 * @param body
	 * @param request
	 * @return return status code only.
	 */
	@RequestMapping(value = ACTION_UPDATE_PASSION, method = RequestMethod.POST)
	@ResponseBody()
	public ContentConfigVO updatePassion(@RequestBody String body, HttpServletRequest request)
			throws JCTException {
		LOGGER.info(">>>>>> ContentConfigController.updatePassion");
		ContentConfigVO contentConfigVO = new ContentConfigVO();
		JobAttributeVO jobAttributeVO = null;
		//String updationSuccess = null;
		JsonNode node = CommonUtility.getNode(body);
		//String atttrDesc = node.get("atttrDescVal").toString().replaceAll("\"" , "").toUpperCase().trim();
		String atttrDesc1 = node.get("atttrDescVal").toString().replaceAll("\"" , "").trim();
		String userProfileValue = node.get("userProfileVal").toString().replaceAll("\"" , "");
		Integer userProfileId = Integer.parseInt(node.get("userProfileId").toString().replaceAll("\"" , ""));
		Integer tablePkId = Integer.parseInt(node.get("tablePkId").toString().replaceAll("\"" , ""));
		String attrFullDesc = node.get("attrFullDesc").toString().replaceAll("\"" , "").trim();		
		/***** ADDED FOR JCT PUBLIC VERSION *****/
		String chechBoxVal = node.get("chechBoxVal").toString().replaceAll("\"" , "").trim();
		/****** ENDED *****/
		
		jobAttributeVO = new JobAttributeVO();
		jobAttributeVO.setJctJobAttributeName(atttrDesc1);
		jobAttributeVO.setJctJobAttributeDesc(attrFullDesc);
		//jobAttributeVO.setJctJobAttributeOrder(newAttrOrder);
		jobAttributeVO.setUserProfileName(userProfileValue);
		jobAttributeVO.setUserProfileId(userProfileId);
		jobAttributeVO.setPrimaryKeyVal(tablePkId);				
		contentConfigVO.setStatusCode(StatusConstants.DOES_NOT_EXIST);
		try {
			contentConfigVO = service.updatePassion(jobAttributeVO, null, chechBoxVal);			
			ContentConfigVO unused = service.populateExistingMapping(userProfileId, "PAS");
			contentConfigVO.setExistingMappingList(unused.getExistingMappingList());
			unused = null;
		} catch (JCTException e) {
			contentConfigVO = new ContentConfigVO();
			contentConfigVO.setStatusCode(StatusConstants.STATUS_FAILURE);
			contentConfigVO.setStatusDesc(this.messageSource.
					getMessage("exception.dao.user.group.update", null, null));
			LOGGER.error(e.getLocalizedMessage());
		}
		//}
		LOGGER.info("<<<<<< ContentConfigController.updatePassion");
		return contentConfigVO;
	}
	/**
	 * Method deletes strength
	 * @param body
	 * @param request
	 * @return return status code only.
	 */
	@RequestMapping(value = ACTION_DELETE_STRENGTH, method = RequestMethod.POST)
	@ResponseBody()
	public ContentConfigVO deleteStrength(@RequestBody String body, HttpServletRequest request) {
		LOGGER.info(">>>>>> ContentConfigController.deleteStrength");
		ContentConfigVO contentConfigVO = null;
		JobAttributeVO jobAttributeVO = null;
		JsonNode node = CommonUtility.getNode(body);
		Integer tablePkId = Integer.parseInt(node.get("tablePkId").toString().replaceAll("\"" , ""));
		String attrDesc = node.get("attrDesc").toString().replaceAll("\"" , "");
		Integer profileId = Integer.parseInt(node.get("profileId").toString().replaceAll("\"" , ""));	
		jobAttributeVO = new JobAttributeVO();
		jobAttributeVO.setJctJobAttributeDesc(attrDesc);		
		jobAttributeVO.setPrimaryKeyVal(tablePkId);
		try {
			contentConfigVO = service.deleteStrength(jobAttributeVO, null);
			ContentConfigVO unused = service.populateExistingMapping(profileId, "STR");
			contentConfigVO.setExistingMappingList(unused.getExistingMappingList());
			unused = null;
		} catch (JCTException e) {
			contentConfigVO = new ContentConfigVO();
			contentConfigVO.setStatusCode(StatusConstants.STATUS_FAILURE);
			contentConfigVO.setStatusDesc(this.messageSource.
					getMessage("exception.dao.user.group.update", null, null));
			LOGGER.error(e.getLocalizedMessage());
		}
		LOGGER.info("<<<<<< ContentConfigController.deleteStrength");
		return contentConfigVO;		
	}
	/**
	 * Method deletes Value
	 * @param body
	 * @param request
	 * @return return status code only.
	 */
	@RequestMapping(value = ACTION_DELETE_VALUE, method = RequestMethod.POST)
	@ResponseBody()
	public ContentConfigVO deleteValue(@RequestBody String body, HttpServletRequest request) {
		LOGGER.info(">>>>>> ContentConfigController.deleteValue");
		ContentConfigVO contentConfigVO = null;
		JobAttributeVO jobAttributeVO = null;
		JsonNode node = CommonUtility.getNode(body);
		Integer tablePkId = Integer.parseInt(node.get("tablePkId").toString().replaceAll("\"" , ""));
		String attrDesc = node.get("attrDesc").toString().replaceAll("\"" , "");
		jobAttributeVO = new JobAttributeVO();
		jobAttributeVO.setJctJobAttributeDesc(attrDesc);		
		jobAttributeVO.setPrimaryKeyVal(tablePkId);
		try {
			contentConfigVO = service.deleteValue(jobAttributeVO, null);
			ContentConfigVO unused = service.populateExistingMapping(tablePkId, "VAL");
			contentConfigVO.setExistingMappingList(unused.getExistingMappingList());
			unused = null;
		} catch (JCTException e) {
			contentConfigVO = new ContentConfigVO();
			contentConfigVO.setStatusCode(StatusConstants.STATUS_FAILURE);
			contentConfigVO.setStatusDesc(this.messageSource.
					getMessage("exception.dao.user.group.update", null, null));
			LOGGER.error(e.getLocalizedMessage());
		}
		LOGGER.info("<<<<<< ContentConfigController.deleteValue");
		return contentConfigVO;		
	}
	/**
	 * Method deletes Passion
	 * @param body
	 * @param request
	 * @return return status code only.
	 */
	@RequestMapping(value = ACTION_DELETE_PASSION, method = RequestMethod.POST)
	@ResponseBody()
	public ContentConfigVO deletePassion(@RequestBody String body, HttpServletRequest request) {
		LOGGER.info(">>>>>> ContentConfigController.deletePassion");
		ContentConfigVO contentConfigVO = null;
		JobAttributeVO jobAttributeVO = null;
		JsonNode node = CommonUtility.getNode(body);
		Integer tablePkId = Integer.parseInt(node.get("tablePkId").toString().replaceAll("\"" , ""));
		String attrDesc = node.get("attrDesc").toString().replaceAll("\"" , "");
		jobAttributeVO = new JobAttributeVO();
		jobAttributeVO.setJctJobAttributeDesc(attrDesc);		
		jobAttributeVO.setPrimaryKeyVal(tablePkId);
		try {
			contentConfigVO = service.deletePassion(jobAttributeVO, null);
			ContentConfigVO unused = service.populateExistingMapping(tablePkId, "PAS");
			contentConfigVO.setExistingMappingList(unused.getExistingMappingList());
			unused = null;
		} catch (JCTException e) {
			contentConfigVO = new ContentConfigVO();
			contentConfigVO.setStatusCode(StatusConstants.STATUS_FAILURE);
			contentConfigVO.setStatusDesc(this.messageSource.
					getMessage("exception.dao.user.group.update", null, null));
			LOGGER.error(e.getLocalizedMessage());
		}
		LOGGER.info("<<<<<< ContentConfigController.deletePassion");
		return contentConfigVO;		
	}
	/**
	 * Method populates the User Profile List.
	 * @param body
	 * @param request
	 * @return ContentConfigVO containing user profiles.
	 */
	@RequestMapping(value = ACTION_POPULATE_USER_PROFILE_GENERAL, method = RequestMethod.POST)
	@ResponseBody()
	public ContentConfigVO populateUserProfileGeneral(@RequestBody String body, HttpServletRequest request) {
		LOGGER.info(">>>>>> ContentConfigController.populateUserProfileGeneral");
		ContentConfigVO contentConfigVO = new ContentConfigVO();
		try {
			contentConfigVO.setUserProfileMap(service.populateUserProfile());
			contentConfigVO.setStatusCode(StatusConstants.STATUS_SUCCESS_WITH_RESULT);
		} catch (JCTException e) {
			contentConfigVO = new ContentConfigVO();
			contentConfigVO.setStatusCode(StatusConstants.STATUS_FAILURE);
			contentConfigVO.setStatusDesc(this.messageSource.getMessage("exception.dao.error",null, null));
			LOGGER.error(e.getLocalizedMessage());
		}
		LOGGER.info("<<<<<< ContentConfigController.populateUserProfileGeneral");
		return contentConfigVO;
	}
	/**
	 * Method saves new instruction
	 * @param body
	 * @param request
	 * @return return status code only.
	 */
	@RequestMapping(value = ACTION_SAVE_INSTRUCTION, method = RequestMethod.POST)
	@ResponseBody()
	public ContentConfigVO saveInstruction(@RequestBody String body, HttpServletRequest request) {
		LOGGER.info(">>>>>> CententConfigController.saveInstruction");
		ContentConfigVO contentConfigVO = null;
		InstructionVO instructionVO = null;
		JsonNode node = CommonUtility.getNode(body);
		String insAreaText = node.get("insAreaTextVal").asText().replaceAll("\"" , "");
		String userProfileValue = node.get("userProfileVal").toString().replaceAll("\"" , "");
		Integer userProfileId = Integer.parseInt(node.get("userProfileId").toString().replaceAll("\"" , ""));
		String relatedPage = node.get("relatedPageText").toString().replaceAll("\"" , "");
		String createdBy = node.get("createdBy").toString().replaceAll("\"" , "");
		instructionVO = new InstructionVO();
		instructionVO.setJctInstructionDesc(insAreaText);
		instructionVO.setJctRelatedPageDesc(relatedPage);
		instructionVO.setUserProfileId(userProfileId);
		instructionVO.setUserProfileName(userProfileValue);
		instructionVO.setCreatedBy(createdBy);				
		try {
			contentConfigVO = service.saveInstruction(instructionVO);
		} catch (JCTException e) {
			contentConfigVO = new ContentConfigVO();
			contentConfigVO.setStatusCode(StatusConstants.STATUS_FAILURE);
			contentConfigVO.setStatusDesc(this.messageSource.
					getMessage("exception.dao.user.group.save", null, null));
			LOGGER.error(e.getLocalizedMessage());
		}
		LOGGER.info("<<<<<< CententConfigController.saveInstruction");
		return contentConfigVO;
	}
	/**
	 * Method populates the instruction
	 * @param body
	 * @param request
	 * @return ContentConfigVO containing instruction.
	 */
	@RequestMapping(value = ACTION_POPULATE_EXTNG_INSTRUCTION, method = RequestMethod.POST)
	@ResponseBody()
	public ContentConfigVO populateExistingInstruction(@RequestBody String body, HttpServletRequest request) {
		LOGGER.info(">>>>>> CententConfigController.populateExistingInstruction");
		ContentConfigVO contentConfigVO = null;
				JsonNode node = CommonUtility.getNode(body);		
		Integer userProfileId = Integer.parseInt(node.get("userProfileId").toString().replaceAll("\"" , ""));
		String relatedPage = node.get("relatedPageText").toString().replaceAll("\"" , "");					
		try {						
			contentConfigVO = service.populateExistingInstruction(userProfileId, relatedPage);
		} catch (JCTException e) {
			contentConfigVO = new ContentConfigVO();
			contentConfigVO.setStatusCode(StatusConstants.STATUS_FAILURE);
			contentConfigVO.setStatusDesc(this.messageSource.getMessage("exception.dao.error", null, null));
			LOGGER.error(e.getLocalizedMessage());
		}
		LOGGER.info("<<<<<< CententConfigController.populateExistingInstruction");
		return contentConfigVO;
	}
	/**
	 * Method upadtes instruction
	 * @param body
	 * @param request
	 * @return return status code only.
	 */
	@RequestMapping(value = ACTION_UPDATE_INSTRUCTION, method = RequestMethod.POST)
	@ResponseBody()
	public ContentConfigVO updateInstruction(@RequestBody String body, HttpServletRequest request) {
		LOGGER.info(">>>>>> CententConfigController.updateInstruction");
		ContentConfigVO contentConfigVO = new ContentConfigVO();
		JsonNode node = CommonUtility.getNode(body);
		String insAreaText = node.get("insAreaTextVal").asText().replaceAll("\"" , "");		
		Integer instructionId = Integer.parseInt(node.get("instructionId").toString().replaceAll("\"" , ""));				
		try {
			contentConfigVO = service.updateInstruction(instructionId, insAreaText);
		} catch (JCTException e) {
			contentConfigVO = new ContentConfigVO();
			contentConfigVO.setStatusCode(StatusConstants.STATUS_FAILURE);
			contentConfigVO.setStatusDesc(this.messageSource.
					getMessage("exception.dao.user.group.save", null, null));
			LOGGER.error(e.getLocalizedMessage());
		}
		LOGGER.info("<<<<<< CententConfigController.updateInstruction");
		return contentConfigVO;
	}
	
	
	/**
	 * Method populates the Reflection Question List.
	 * @param body
	 * @param request
	 * @return ContentConfigVO containing Questions.
	 */
	@RequestMapping(value = ACTION_POPULATE_MAIN_QTN, method = RequestMethod.POST)
	@ResponseBody()
	public ContentConfigVO populateMainQtn(@RequestBody String body, HttpServletRequest request) {
		LOGGER.info(">>>>>> CententConfigController.populateMainQtn");
		JsonNode node = CommonUtility.getNode(body);
		int profileId = Integer.parseInt(node.get("profileId").toString().replaceAll("\"" , ""));
		String relatedPage = node.get("relatedPage").toString().replaceAll("\"" , "");	
		ContentConfigVO contentCongifVo = new ContentConfigVO();
		try {						
			//contentCongifVo.setUserProfileMap(this.service.populateUserProfile());			
			ContentConfigVO unused = this.service.populateMainQtn(profileId, relatedPage);
			contentCongifVo.setMainQtnList(unused.getMainQtnList());
			unused = null;
			contentCongifVo.setStatusCode(StatusConstants.STATUS_SUCCESS_WITH_RESULT);
		} catch (JCTException e) {
			contentCongifVo = new ContentConfigVO();
			contentCongifVo.setStatusCode(StatusConstants.STATUS_FAILURE);
			contentCongifVo.setStatusDesc(this.messageSource.getMessage("exception.dao.error", null, null));
			LOGGER.error(e.getLocalizedMessage());
		}
		LOGGER.info("<<<<<< CententConfigController.populateMainQtn");
		return contentCongifVo;
	}
	

	/**
	 * Method populates the Sub Question List.
	 * @param body
	 * @param request
	 * @return ContentConfigVO containing Questions.
	 */
	@RequestMapping(value = ACTION_POPULATE_SUB_QTN, method = RequestMethod.POST)
	@ResponseBody()
	public ContentConfigVO populateSubQtn(@RequestBody String body, HttpServletRequest request) {
		LOGGER.info(">>>>>> CententConfigController.populateSubQtn");
		JsonNode node = CommonUtility.getNode(body);
		int profileId = Integer.parseInt(node.get("profileId").toString().replaceAll("\"" , ""));
		String relatedPage = node.get("relatedPage").toString().replaceAll("\"" , "");	
		String mainQuestion = node.get("mainQuestion").toString().replaceAll("\"" , "");	
		ContentConfigVO contentCongifVo = new ContentConfigVO();
		try {						
			//contentCongifVo.setUserProfileMap(this.service.populateUserProfile());			
			ContentConfigVO unused = this.service.populateSubQtn(profileId, relatedPage, mainQuestion);
			contentCongifVo.setSubQtnList(unused.getSubQtnList());
			unused = null;
			contentCongifVo.setStatusCode(StatusConstants.STATUS_SUCCESS_WITH_RESULT);
		} catch (JCTException e) {
			contentCongifVo = new ContentConfigVO();
			contentCongifVo.setStatusCode(StatusConstants.STATUS_FAILURE);
			contentCongifVo.setStatusDesc(this.messageSource.getMessage("exception.dao.error", null, null));
			LOGGER.error(e.getLocalizedMessage());
		}
		LOGGER.info("<<<<<< CententConfigController.populateSubQtn");
		return contentCongifVo;
	}
	
	/**
	 * Method save main question order
	 * @param body
	 * @param request
	 * @return return status code only.
	 */
	@RequestMapping(value = ACTION_SAVE_MAIN_QTN_ORDER, method = RequestMethod.POST)
	@ResponseBody()
	public ContentConfigVO saveMainQtnOrder(@RequestBody String body, HttpServletRequest request) throws JCTException{
		LOGGER.info(">>>>>> ContentConfigController.saveMainQtnOrder");
		ContentConfigVO contentConfigVO = new ContentConfigVO();
		JsonNode node = CommonUtility.getNode(body);
		String builder = node.get("builder").toString().replaceAll("\"" , "").trim();	
		Integer userProfile = Integer.parseInt(node.get("userProfile").toString().replaceAll("\"" , ""));	
		String relatedPage = node.get("relatedPage").toString().replaceAll("\"" , "").trim();	
			try {
				contentConfigVO = service.saveMainQtnOrder(builder, userProfile, relatedPage);
				ContentConfigVO unused = this.service.populateMainQtn(userProfile, relatedPage);
				contentConfigVO.setMainQtnList(unused.getMainQtnList());
				unused = null;
			} catch (JCTException e) {
				contentConfigVO = new ContentConfigVO();
				contentConfigVO.setStatusCode(StatusConstants.STATUS_FAILURE);
				contentConfigVO.setStatusDesc(this.messageSource.
						getMessage("exception.dao.user.group.update", null, null));
				LOGGER.error(e.getLocalizedMessage());
			}
		
		LOGGER.info("<<<<<< ContentConfigController.saveMainQtnOrder");
		return contentConfigVO;
	}
	
	/**
	 * Method save sub question order
	 * @param body
	 * @param request
	 * @return return status code only.
	 */
	@RequestMapping(value = ACTION_SAVE_SUB_QTN_ORDER, method = RequestMethod.POST)
	@ResponseBody()
	public ContentConfigVO saveSubQtnOrder(@RequestBody String body, HttpServletRequest request) throws JCTException{
		LOGGER.info(">>>>>> ContentConfigController.saveSubQtnOrder");
		ContentConfigVO contentConfigVO = new ContentConfigVO();
		JsonNode node = CommonUtility.getNode(body);
		String builder = node.get("builder").toString().replaceAll("\"" , "").trim();	
		Integer userProfile = Integer.parseInt(node.get("userProfile").toString().replaceAll("\"" , ""));	
		String relatedPage = node.get("relatedPage").toString().replaceAll("\"" , "").trim();
		String maninQtn = node.get("manipulativeQtn").toString().replaceAll("\"" , "").trim();
		
			try {
				contentConfigVO = service.saveSubQtnOrder(builder, userProfile, relatedPage, maninQtn);
				ContentConfigVO unused = this.service.populateSubQtn(userProfile, relatedPage, maninQtn);
				contentConfigVO.setSubQtnList(unused.getSubQtnList());
				unused = null;
			} catch (JCTException e) {
				contentConfigVO = new ContentConfigVO();
				contentConfigVO.setStatusCode(StatusConstants.STATUS_FAILURE);
				contentConfigVO.setStatusDesc(this.messageSource.
						getMessage("exception.dao.user.group.update", null, null));
				LOGGER.error(e.getLocalizedMessage());
			}
		
		LOGGER.info("<<<<<< ContentConfigController.saveSubQtnOrder");
		return contentConfigVO;
	}
	
	/**
	 * Method update existing function Group Order
	 * @param body
	 * @param request
	 * @return return status code only.
	 */
	@RequestMapping(value = ACTION_SAVE_FUNCTION_GROUP_ORDER, method = RequestMethod.POST)
	@ResponseBody()
	public ContentConfigVO saveFuncGrpOrder(@RequestBody String body, HttpServletRequest request) throws JCTException{
		LOGGER.info(">>>>>> ContentConfigController.saveFuncGrpOrder");
		ContentConfigVO contentConfigVO = new ContentConfigVO();		
		JsonNode node = CommonUtility.getNode(body);
		String builder = node.get("builder").toString().replaceAll("\"" , "").trim();			
			try {
				contentConfigVO = service.saveFuncGrpOrder(builder);
				ContentConfigVO unused = this.service.populateFunctionGroupByOrder();
				contentConfigVO.setExistingFunctionGrpList(unused.getExistingFunctionGrpList());
				unused = null;
			} catch (JCTException e) {
				contentConfigVO = new ContentConfigVO();
				contentConfigVO.setStatusCode(StatusConstants.STATUS_FAILURE);
				contentConfigVO.setStatusDesc(this.messageSource.
						getMessage("exception.dao.user.group.update", null, null));
				LOGGER.error(e.getLocalizedMessage());
			}
		
		LOGGER.info("<<<<<< ContentConfigController.saveFuncGrpOrder");
		return contentConfigVO;
	}
	
	/**
	 * Method populates the Function Group List by Order.
	 * @param body
	 * @param request
	 * @return ContentConfigVO containing Function Group.
	 */
	@RequestMapping(value = ACTION_POPULATE_EXTNG_FUNC_GRP_ORDER, method = RequestMethod.POST)
	@ResponseBody()
	public ContentConfigVO populateFunctionGroupByOrder(@RequestBody String body, HttpServletRequest request) {
		LOGGER.info(">>>>>> CententConfigController.populateFunctionGroupByOrder");
		ContentConfigVO contentCongifVo = new ContentConfigVO();
		try {											
			contentCongifVo = service.populateFunctionGroupByOrder();
		} catch (JCTException e) {
			contentCongifVo = new ContentConfigVO();
			contentCongifVo.setStatusCode(StatusConstants.STATUS_FAILURE);
			contentCongifVo.setStatusDesc(this.messageSource.getMessage("exception.dao.error", null, null));
			LOGGER.error(e.getLocalizedMessage());
		}
		LOGGER.info("<<<<<< CententConfigController.populateFunctionGroupByOrder");
		return contentCongifVo;
	}
	
	/**
	 * Method update existing job level Order
	 * @param body
	 * @param request
	 * @return return status code only.
	 */
	@RequestMapping(value = ACTION_SAVE_JOB_LEVEL_ORDER, method = RequestMethod.POST)
	@ResponseBody()
	public ContentConfigVO saveJobLevelOrder(@RequestBody String body, HttpServletRequest request) throws JCTException{
		LOGGER.info(">>>>>> ContentConfigController.saveJobLevelOrder");
		ContentConfigVO contentConfigVO = new ContentConfigVO();		
		JsonNode node = CommonUtility.getNode(body);
		String builder = node.get("builder").toString().replaceAll("\"" , "").trim();			
			try {
				contentConfigVO = service.saveJobLevelOrder(builder);
				ContentConfigVO unused = this.service.populateJobLevelByOrder();
				contentConfigVO.setExistingJobLevelList(unused.getExistingJobLevelList());
				unused = null;
			} catch (JCTException e) {
				contentConfigVO = new ContentConfigVO();
				contentConfigVO.setStatusCode(StatusConstants.STATUS_FAILURE);
				contentConfigVO.setStatusDesc(this.messageSource.
						getMessage("exception.dao.user.group.update", null, null));
				LOGGER.error(e.getLocalizedMessage());
			}
		
		LOGGER.info("<<<<<< ContentConfigController.saveJobLevelOrder");
		return contentConfigVO;
	}
	
	/**
	 * Method populates the job level List by Order.
	 * @param body
	 * @param request
	 * @return ContentConfigVO containing Function Group.
	 */
	@RequestMapping(value = ACTION_POPULATE_EXTNG_JOB_LEVEL_ORDER, method = RequestMethod.POST)
	@ResponseBody()
	public ContentConfigVO populateJobLevelByOrder(@RequestBody String body, HttpServletRequest request) {
		LOGGER.info(">>>>>> CententConfigController.populateJobLevelByOrder");
		ContentConfigVO contentCongifVo = new ContentConfigVO();
		try {											
			contentCongifVo = service.populateJobLevelByOrder();
		} catch (JCTException e) {
			contentCongifVo = new ContentConfigVO();
			contentCongifVo.setStatusCode(StatusConstants.STATUS_FAILURE);
			contentCongifVo.setStatusDesc(this.messageSource.getMessage("exception.dao.error", null, null));
			LOGGER.error(e.getLocalizedMessage());
		}
		LOGGER.info("<<<<<< CententConfigController.populateJobLevelByOrder");
		return contentCongifVo;
	}
	
	/**
	 * Method populates the attribute List by Order.
	 * @param body
	 * @param request
	 * @return ContentConfigVO containing attribute.
	 */
	@RequestMapping(value = ACTION_POPULATE_ATTRIBUTE_ORDER, method = RequestMethod.POST)
	@ResponseBody()
	public ContentConfigVO populateMappingList(@RequestBody String body, HttpServletRequest request) {
		LOGGER.info(">>>>>> CententConfigController.populateMappingList");
		ContentConfigVO contentCongifVo = new ContentConfigVO();
		JsonNode node = CommonUtility.getNode(body);
		int profileId = Integer.parseInt(node.get("profileId").toString().replaceAll("\"" , ""));
		String attrCode = node.get("attrCode").toString().replaceAll("\"" , "");		
		try {											
			contentCongifVo = service.populateMappingList(profileId, attrCode);
		} catch (JCTException e) {
			contentCongifVo = new ContentConfigVO();
			contentCongifVo.setStatusCode(StatusConstants.STATUS_FAILURE);
			contentCongifVo.setStatusDesc(this.messageSource.getMessage("exception.dao.error", null, null));
			LOGGER.error(e.getLocalizedMessage());
		}
		LOGGER.info("<<<<<< CententConfigController.populateMappingList");
		return contentCongifVo;
	}
	
	/**
	 * Method update attribute list order
	 * @param body
	 * @param request
	 * @return return status code only.
	 */
	@RequestMapping(value = ACTION_SAVE_ATTRIBUTE_ORDER, method = RequestMethod.POST)
	@ResponseBody()
	public ContentConfigVO saveAttributeOrder(@RequestBody String body, HttpServletRequest request) throws JCTException{
		LOGGER.info(">>>>>> ContentConfigController.saveAttributeOrder");
		ContentConfigVO contentConfigVO = new ContentConfigVO();		
		JsonNode node = CommonUtility.getNode(body);
		String builder = node.get("builder").toString().replaceAll("\"" , "").trim();	
		Integer userProfile = Integer.parseInt(node.get("userProfile").toString().replaceAll("\"" , ""));	
		String attrCode = node.get("attrCode").toString().replaceAll("\"" , "").trim();	
			try {
				contentConfigVO = service.saveAttributeOrder(builder, userProfile, attrCode);
				ContentConfigVO unused = this.service.populateMappingList(userProfile, attrCode);
				contentConfigVO.setExistingMappingList(unused.getExistingMappingList());
				unused = null;
			} catch (JCTException e) {
				contentConfigVO = new ContentConfigVO();
				contentConfigVO.setStatusCode(StatusConstants.STATUS_FAILURE);
				contentConfigVO.setStatusDesc(this.messageSource.
						getMessage("exception.dao.user.group.update", null, null));
				LOGGER.error(e.getLocalizedMessage());
			}
		
		LOGGER.info("<<<<<< ContentConfigController.saveAttributeOrder");
		return contentConfigVO;
	}
	
	/**
	 * Method update region order
	 * @param body
	 * @param request
	 * @return return status code only.
	 */
	@RequestMapping(value = ACTION_SAVE_REGION_ORDER, method = RequestMethod.POST)
	@ResponseBody()
	public ContentConfigVO saveRegionOrder(@RequestBody String body, HttpServletRequest request) throws JCTException{
		LOGGER.info(">>>>>> ContentConfigController.saveRegionOrder");
		ContentConfigVO contentConfigVO = new ContentConfigVO();		
		JsonNode node = CommonUtility.getNode(body);
		String builder = node.get("builder").toString().replaceAll("\"" , "").trim();	
			try {
				contentConfigVO = service.saveRegionOrder(builder);
				ContentConfigVO unused = this.service.populateExistingRegion();
				contentConfigVO.setExistingRegionList(unused.getExistingRegionList());
				unused = null;
			} catch (JCTException e) {
				contentConfigVO = new ContentConfigVO();
				contentConfigVO.setStatusCode(StatusConstants.STATUS_FAILURE);
				contentConfigVO.setStatusDesc(this.messageSource.
						getMessage("exception.dao.user.group.update", null, null));
				LOGGER.error(e.getLocalizedMessage());
			}
		
		LOGGER.info("<<<<<< ContentConfigController.saveRegionOrder");
		return contentConfigVO;
	}


	/******* THIS IS FOR JCT PUBLIC VERSION ********/
	/**
	 * Method populates the ONet Occupation List.
	 * @param body
	 * @param request
	 * @return ContentConfigVO ONet values.
	 */
	@RequestMapping(value = ACTION_POPULATE_ONET_OCCUPATION, method = RequestMethod.POST)
	@ResponseBody()
	public ContentConfigVO populateOnetOccupationList(@RequestBody String body, HttpServletRequest request) {
		LOGGER.info(">>>>>> CententConfigController.populateOnetOccupationList");
		ContentConfigVO contentCongifVo = new ContentConfigVO();
		try {									
			contentCongifVo = service.populateOnetDataList();
		} catch (JCTException e) {
			contentCongifVo = new ContentConfigVO();
			contentCongifVo.setStatusCode(StatusConstants.STATUS_FAILURE);
			contentCongifVo.setStatusDesc(this.messageSource.getMessage("exception.dao.error", null, null));
			LOGGER.error(e.getLocalizedMessage());
		}
		LOGGER.info("<<<<<< CententConfigController.populateOnetOccupationList");
		return contentCongifVo;
	}
	
	
	/**
	 * Method to fetch existing terms and condition by profile ID
	 * @param body
	 * @param request
	 * @return termsAndConditionVO values.
	 */
	@RequestMapping(value = ACTION_FETCH_TERMS_AND_CONDITIONS, method = RequestMethod.POST)
	@ResponseBody()
	public TermsAndConditionsVO fetchTermsAndConditions(@RequestBody String body, HttpServletRequest request) {
		LOGGER.info(">>>>>> CententConfigController.fetchTermsAndConditions");
		JsonNode node = CommonUtility.getNode(body);
		int userProfile = Integer.parseInt(node.get("userProfile").asText());	
		int userType = Integer.parseInt(node.get("userType").asText());
		TermsAndConditionsVO termsAndConditionsVO = new TermsAndConditionsVO();
		JctTermsAndConditions obj = new JctTermsAndConditions();
		try {		
			obj = service.fetchTermsAndCondition(userProfile, userType);
			if(null != obj){
				termsAndConditionsVO.setJctTcDescription(obj.getJctTcDescription());
				termsAndConditionsVO.setStatusCode(StatusConstants.STATUS_SUCCESS);				
			} else {
				termsAndConditionsVO.setStatusCode(StatusConstants.STATUS_FAILED);
				termsAndConditionsVO.setStatusDesc(this.messageSource.getMessage("termsAndCondition.not.found", null, null));				
			}			
		} catch (JCTException e) {
			termsAndConditionsVO.setStatusCode(StatusConstants.STATUS_FAILURE);
			LOGGER.error(e.getLocalizedMessage());
		}
		LOGGER.info("<<<<<< CententConfigController.fetchTermsAndConditions");
		return termsAndConditionsVO;
	}
	/**
	 * Method saves new Terms & conditions
	 * @param body
	 * @param request
	 * @return return status code only.
	 */
	@RequestMapping(value = ACTION_SAVE_TERMS_AND_CONDITIONS, method = RequestMethod.POST)
	@ResponseBody()
	public TermsAndConditionsVO saveAndUpdateTermAndCondition(@RequestBody String body, HttpServletRequest request) 
			throws JCTException {
		LOGGER.info(">>>>>> CententConfigController.saveAndUpdateTermAndCondition");
		TermsAndConditionsVO termsAndConditionsVO = new TermsAndConditionsVO();
		JsonNode node = CommonUtility.getNode(body);
		String tcText = node.get("tcText").toString().replaceAll("\"" , "");
		int userProfile = Integer.parseInt(node.get("userProfile").asText());
		int userType = Integer.parseInt(node.get("userType").asText());
		String createdBy = node.get("createdBy").toString().replaceAll("\"" , "");		
		String updationSuccess = service.saveTermsAndCondition(userProfile, userType, tcText.replaceAll("\n", "<br>"),createdBy);		
		if (updationSuccess.equals("success")) {
			termsAndConditionsVO.setStatusCode(StatusConstants.STATUS_SUCCESS_WITH_RESULT);
			termsAndConditionsVO.setStatusDesc(this.messageSource.getMessage("termsAndCondition.saveUpdate", null, null));
		} else {
			termsAndConditionsVO.setStatusCode(StatusConstants.STATUS_SUCCESS_WITH_RESULT);
			termsAndConditionsVO.setStatusDesc(this.messageSource.getMessage("termsAndCondition.saveUpdate", null, null));
		}
		LOGGER.info("<<<<<< CententConfigController.saveAndUpdateTermAndCondition");
		return termsAndConditionsVO;
	}
}