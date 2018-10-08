package com.vmware.jct.service.impl;

import java.io.BufferedInputStream;
import java.io.BufferedOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Date;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.TreeMap;

import org.apache.activemq.util.ByteArrayInputStream;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.context.MessageSource;
import org.springframework.stereotype.Service;
import org.springframework.transaction.annotation.Propagation;
import org.springframework.transaction.annotation.Transactional;

import com.vmware.jct.common.utility.CommonConstants;
import com.vmware.jct.common.utility.StatusConstants;
import com.vmware.jct.dao.ICommonDao;
import com.vmware.jct.dao.IContentServiceDAO;
import com.vmware.jct.dao.dto.FunctionDTO;
import com.vmware.jct.dao.dto.JobAttributeDTO;
import com.vmware.jct.dao.dto.LevelDTO;
import com.vmware.jct.dao.dto.OnetOccupationDTO;
import com.vmware.jct.dao.dto.QuestionearDTO;
import com.vmware.jct.dao.dto.RegionDTO;
import com.vmware.jct.dao.dto.UserProfileDTO;
import com.vmware.jct.exception.DAOException;
import com.vmware.jct.exception.JCTException;
import com.vmware.jct.model.JctFunction;
import com.vmware.jct.model.JctInstructionBar;
import com.vmware.jct.model.JctJobAttribute;
import com.vmware.jct.model.JctLevel;
import com.vmware.jct.model.JctQuestion;
import com.vmware.jct.model.JctRegion;
import com.vmware.jct.model.JctUserProfile;
import com.vmware.jct.model.JctTermsAndConditions;
import com.vmware.jct.service.IContentConfigService;
import com.vmware.jct.service.IStorageService;
import com.vmware.jct.service.vo.ContentConfigVO;
import com.vmware.jct.service.vo.InstructionVO;
import com.vmware.jct.service.vo.JobAttributeVO;
import com.vmware.jct.service.vo.QuestionearVO;

/**
 * 
 * <p><b>Class name:</b> ContentConfigServiceImpl.java</p>
 * <p><b>Author:</b> InterraIT</p>
 * <p><b>Purpose:</b> This is a ContentConfigServiceImpl class. This artifact is Business layer artifact.
 * ContentConfigServiceImpl implement IContentConfigService interface and override the following  methods.
 * -populateExistingRefQtn()
 * -populateUserProfile()
 * -saveRefQtn()
 * -updateRefQtn()
 * -deleteRefQtn()
 * -saveActionPlan()
 * -updateActionPlan()
 * -deleteActionPlan()
 * -populateExistingFunctionGroup()
 * -populateExistingJobLevel()
 * -saveFunctionGroup()
 * -saveJobLevel()
 * -populateExistingRegion()
 * -saveRegion()
 * -updateRegion()
 * -deleteRegion()
 * -populateExistingMapping()
 * -saveStrength()
 * -saveValue()
 * -savePassion()
 * -updateStrength()
 * -updateValue()
 * -updatePassion()
 * -deleteStrength()
 * -deleteValue()
 * -deletePassion()
 * -saveInstruction()
 * -populateExistingInstruction()
 * -validateExistenceRefQtn()
 * -validateExistenceActionPlan()
 * -validateExistenceFuncGrp()
 * -validateExistenceJobLevel()
 * -validateExistenceAttribute()
 * -validateExistenceRegion()
 * -updateInstruction()
 * -fetchTermsAndCondition()
 * </p>
 * <p><b>Description:</b> This class used to perform insert, modify ,fetch data from db and delete operation. Save, update, delete operations are performed in DataAccessObject class </p>
 * <p><b>Revision History:</b>
 * 	<li>InterraIT - 19/Jan/2014 - Implement Exception </li>
 * <li>InterraIT - 31/Jan/2014 - Implement comments, introduce Named query, Transactional attribute </li>
 * </p>
 */
@Service
public class ContentConfigServiceImpl implements IContentConfigService{
	
	@Autowired
	private IContentServiceDAO serviceDAO;
	
	@Autowired  
	private MessageSource messageSource;
	
	@Autowired
	private ICommonDao commonDaoImpl;
	
	@Autowired
	private IStorageService storageService;

	private static final Logger LOGGER = LoggerFactory.getLogger(ContentConfigServiceImpl.class);
	
	@Transactional(propagation = Propagation.REQUIRED)
	public ContentConfigVO populateExistingRefQtn(Integer profileId, String relatedPage) throws JCTException {
		LOGGER.info(">>>>>> ContentConfigServiceImpl.populateExistingUserGroup");
		StringBuilder builder = new StringBuilder("");
		ContentConfigVO vo = new ContentConfigVO();
		try {
			List<QuestionearDTO> dtoList = serviceDAO.populateExistingRefQtn(profileId, relatedPage);
			if (dtoList.size() > 0) {
				Iterator<QuestionearDTO> itr = dtoList.iterator();
				while (itr.hasNext()) {
					QuestionearDTO dto = (QuestionearDTO) itr.next();
					builder.append(dto.getJctQuestionsDesc());
					builder.append("~");
					builder.append(dto.getJctQuestionSubDesc());
					builder.append("~");
					builder.append(dto.getJctQuestionBsas());
					builder.append("~");
					builder.append(dto.getUserProfileDesc());
					builder.append("~");
					builder.append(dto.getJctQuestionsId());
					builder.append("~");
					builder.append(dto.getJctQuestionsSoftDelete());
					builder.append("~");
					builder.append(dto.getUserProfileId());
					builder.append("~");
					builder.append(dto.getJctQuestionsOrder());
					builder.append("~");
					builder.append(dto.getJctSubQuestionsOrder());
					builder.append("$$$");
				}
				vo.setExistingRefQtnList(builder.toString());
				vo.setStatusCode(StatusConstants.STATUS_SUCCESS_WITH_RESULT);
			} else {
				LOGGER.info("---- NO RECORDS FOUND -----");
				vo.setStatusCode(StatusConstants.STATUS_SUCCESS_WITH_NO_RESULT);
				vo.setStatusDesc(this.messageSource.getMessage("waring.no.user.profile", null, null));
			}
		} catch (DAOException e) {
			LOGGER.error("----"+e.getLocalizedMessage()+" ----");
			throw new JCTException(e.getLocalizedMessage());
		}
		LOGGER.info("<<<<<< ContentConfigServiceImpl.populateExistingUserGroup");
		return vo;
	}
	
	@Transactional(propagation = Propagation.REQUIRED)
	public Map<Integer, String> populateUserProfile() throws JCTException {
		LOGGER.info(">>>>>> ContentConfigServiceImpl.populateUserProfile");
		Map<Integer, String> userProfileMap = new TreeMap<Integer, String>();
			List<UserProfileDTO> dtoList = null;
			try {
				dtoList = serviceDAO.populateExistingUserProfileWithId();
				for (int index = 0; index < dtoList.size(); index++) {
					UserProfileDTO dto = (UserProfileDTO) dtoList.get(index);
					userProfileMap.put(dto.getUserProfileId(), dto.getUserProfileDesc());
				}
			} catch (DAOException e) {
				LOGGER.error("FAILED IN populateUserProfile --- "+e.getLocalizedMessage());
				throw new JCTException(e.getLocalizedMessage());
			}
		return userProfileMap;
	}
	
	
	@Transactional(propagation = Propagation.REQUIRED)
	public ContentConfigVO saveRefQtn(QuestionearVO questionearVO)
			throws JCTException {
		LOGGER.info(">>>>>> ContentConfigServiceImpl.saveRefQtn");		
		JctQuestion question = new JctQuestion();
		//JctUserGroup userGroup = new JctUserGroup();
		ContentConfigVO contentConfigVO = new ContentConfigVO();
		try {
			//int sunQtnLength = questionearVO.getNoOfSubQtn();			
			//for (int i = 0; i< sunQtnLength;i++){
				question.setJctQuestionsId(commonDaoImpl.generateKey("jct_questions"));
				question.setJctBsCreatedBy(questionearVO.getCreatedBy());
				question.setJctDsCreatedTs(new Date());
				question.setJctBsLastmodifiedBy(questionearVO.getCreatedBy());
				question.setJctBsLastmodifiedTs(new Date());
				question.setJctQuestionsSoftDelete(CommonConstants.ENTRY_NOT_DELETED);
				question.setJctQuestionsDesc(questionearVO.getQuestionsDesc());
				question.setJctQuestionSubDesc(questionearVO.getQuestionSubDesc());
				question.setJctProfilesDesc(questionearVO.getUserProfileName());
				//question.setJctQuestionsOrder(questionearVO.getQuestionOrder());
				//question.setJctQuestionsOrder(commonDaoImpl.generateOrder("BSQuestion", questionearVO.getUserProfileId(), questionearVO.getQuestionsDesc()));
				question.setJctQuestionsOrder(serviceDAO.generateOrderQtn("BS", questionearVO.getUserProfileId(), questionearVO.getQuestionsDesc()));
				question.setJctSubQuestionsOrder(questionearVO.getSubQuestionOrder());
				JctUserProfile profile = new JctUserProfile();
				profile.setJctUserProfile(questionearVO.getUserProfileId());
				question.setJctUserProfile(profile);
				question.setVersion(1);		
				question.setJctQuestionBsas("BS");
				String retString = serviceDAO.saveRefQtn(question);
				if (!retString.equals("success")) {
					contentConfigVO.setStatusCode(StatusConstants.STATUS_FAILURE);
					contentConfigVO.setStatusDesc(this.messageSource.
							getMessage("exception.dao.user.group.save", null, null));
				} else {
					contentConfigVO.setStatusCode(StatusConstants.STATUS_SUCCESS_WITH_RESULT);
					contentConfigVO.setStatusDesc(retString);
				}
				//}		
			} catch (DAOException e) {
			LOGGER.error("----"+e.getLocalizedMessage()+" ----");
			contentConfigVO.setStatusCode(StatusConstants.STATUS_FAILURE);
			throw new JCTException(e.getLocalizedMessage());
		}
		LOGGER.info("<<<<<< ManageUserServiceImpl.saveRefQtn");
		return contentConfigVO;
	}
	
	
	@Transactional(propagation = Propagation.REQUIRED)
	public ContentConfigVO updateRefQtn(QuestionearVO questionearVO, String dist, String chechBoxVal)
			throws JCTException {
		LOGGER.info(">>>>>> ContentConfigImpl.updateRefQtn");		
		ContentConfigVO contentConfigVO = new ContentConfigVO();
		String prevoiusSubQtn = "";
		String mainQuestion = "";
		try {
			JctQuestion question = serviceDAO.fetchRefQtn(questionearVO.getPrimaryKeyVal());
			if (question != null) {
				/***** ADDED FOR JCT PUBLIC VERSION *****/
				prevoiusSubQtn = question.getJctQuestionSubDesc();
				mainQuestion = question.getJctQuestionsDesc();
				/****** ENDED *****/
				question.setJctBsLastmodifiedTs(new Date());
				if (null == dist) {
					question.setJctQuestionsDesc(questionearVO.getQuestionsDesc());
					question.setJctQuestionSubDesc(questionearVO.getQuestionSubDesc());
					question.setJctProfilesDesc(questionearVO.getUserProfileName());
					JctUserProfile profile = new JctUserProfile();
					profile.setJctUserProfile(questionearVO.getUserProfileId());
					question.setJctUserProfile(profile);
				}
				if (null != dist){
					question.setJctQuestionsSoftDelete(0);
				}
				String retString = serviceDAO.updateRefQtn(question);
				if (!retString.equals("success")) {
					contentConfigVO.setStatusCode(StatusConstants.STATUS_FAILURE);
					contentConfigVO.setStatusDesc(this.messageSource.
							getMessage("exception.dao.user.group.update", null, null));
				} else {
					contentConfigVO.setStatusCode(StatusConstants.STATUS_SUCCESS_WITH_RESULT);
					contentConfigVO.setStatusDesc(this.messageSource.getMessage("reflecion.qtn.edit.sucess.msg", null, null));
				}
			} else {
				contentConfigVO.setStatusCode(StatusConstants.STATUS_FAILURE);
				contentConfigVO.setStatusDesc(this.messageSource.getMessage("exception.dao.qtn.update.fetch.delete", null, null));
			}
			/***** ADDED FOR JCT PUBLIC VERSION *****/
			if (chechBoxVal.equalsIgnoreCase("true")) {
				//To make global profile change
				contentConfigVO.setStatusDesc(this.messageSource.getMessage("reflecion.qtn.global.profile.edit.msg", null, null));
				serviceDAO.updateQuestionFrozen ("BS", mainQuestion, prevoiusSubQtn,
						questionearVO.getQuestionSubDesc());
			}
			/****** ENDED *****/
		} catch (DAOException e) {
			LOGGER.error("----"+e.getLocalizedMessage()+" ----");
			contentConfigVO.setStatusCode(StatusConstants.STATUS_FAILURE);
			throw new JCTException(e.getLocalizedMessage());
		}
		LOGGER.info("<<<<<< contentConfigImpl.updateRefQtn");
		return contentConfigVO;
	}
	
	
	@Transactional(propagation = Propagation.REQUIRED)
	public ContentConfigVO deleteRefQtn(QuestionearVO questionearVO, String dist)
			throws JCTException {
		LOGGER.info(">>>>>> ContentConfigImpl.deleteRefQtn");		
		ContentConfigVO contentConfigVO = new ContentConfigVO();
		try {
			JctQuestion question = serviceDAO.fetchRefQtn(questionearVO.getPrimaryKeyVal());
			//JctUserGroup userGroup = serviceDAO.fetchUserGroup(userGroupVO.getPrimaryKeyVal());
			question.setJctBsLastmodifiedTs(new Date());
			if (null == dist){
				question.setJctQuestionsDesc(questionearVO.getQuestionsDesc());
				question.setJctQuestionsSoftDelete(1);
			}
			if (null != dist){
				question.setJctQuestionsSoftDelete(0);
			}
			String retString = serviceDAO.deleteRefQtn(question);
			if (!retString.equals("success")) {
				contentConfigVO.setStatusCode(StatusConstants.STATUS_FAILURE);
				contentConfigVO.setStatusDesc(this.messageSource.
						getMessage("exception.dao.user.group.update", null, null));
			} else {
				contentConfigVO.setStatusCode(StatusConstants.STATUS_SUCCESS_WITH_RESULT);
				contentConfigVO.setStatusDesc(retString);
			}
		} catch (DAOException e) {
			LOGGER.error("----"+e.getLocalizedMessage()+" ----");
			contentConfigVO.setStatusCode(StatusConstants.STATUS_FAILURE);
			throw new JCTException(e.getLocalizedMessage());
		}
		LOGGER.info("<<<<<< contentConfigImpl.deleteRefQtn");
		return contentConfigVO;
	}
	
	
	@Transactional(propagation = Propagation.REQUIRED)
	public ContentConfigVO saveActionPlan(QuestionearVO questionearVO)
			throws JCTException {
		LOGGER.info(">>>>>> ContentConfigServiceImpl.saveActionPlan");
		JctQuestion question = new JctQuestion();
		ContentConfigVO contentConfigVO = new ContentConfigVO();
		try {
			//int sunQtnLength = questionearVO.getNoOfSubQtn();			
			//for (int i = 0; i< sunQtnLength;i++){
				question.setJctQuestionsId(commonDaoImpl.generateKey("jct_questions"));
				question.setJctBsCreatedBy(questionearVO.getCreatedBy());
				question.setJctDsCreatedTs(new Date());
				question.setJctBsLastmodifiedBy(questionearVO.getCreatedBy());
				question.setJctBsLastmodifiedTs(new Date());
				question.setJctQuestionsSoftDelete(CommonConstants.ENTRY_NOT_DELETED);
				question.setJctQuestionsDesc(questionearVO.getQuestionsDesc());
				question.setJctQuestionSubDesc(questionearVO.getQuestionSubDesc());
				question.setJctProfilesDesc(questionearVO.getUserProfileName());
				//question.setJctQuestionsOrder(questionearVO.getQuestionOrder());
				//question.setJctQuestionsOrder(commonDaoImpl.generateOrder("ASQuestion", questionearVO.getUserProfileId(), null));
				question.setJctQuestionsOrder(serviceDAO.generateOrderQtn("AS", questionearVO.getUserProfileId(), questionearVO.getQuestionsDesc()));
				question.setJctSubQuestionsOrder(questionearVO.getSubQuestionOrder());
				JctUserProfile profile = new JctUserProfile();
				profile.setJctUserProfile(questionearVO.getUserProfileId());
				question.setJctUserProfile(profile);
				question.setVersion(1);		
				question.setJctQuestionBsas("AS");
				String retString = serviceDAO.saveActionPlan(question);
				if (!retString.equals("success")) {
					contentConfigVO.setStatusCode(StatusConstants.STATUS_FAILURE);
					contentConfigVO.setStatusDesc(this.messageSource.
							getMessage("exception.dao.user.group.save", null, null));
				} else {
					contentConfigVO.setStatusCode(StatusConstants.STATUS_SUCCESS_WITH_RESULT);
					contentConfigVO.setStatusDesc(retString);
				}
				//}		
			} catch (DAOException e) {
			LOGGER.error("----"+e.getLocalizedMessage()+" ----");
			contentConfigVO.setStatusCode(StatusConstants.STATUS_FAILURE);
			throw new JCTException(e.getLocalizedMessage());
		}
		LOGGER.info("<<<<<< ContentConfigServiceImpl.saveActionPlan");
		return contentConfigVO;
	}
	
	@Transactional(propagation = Propagation.REQUIRED)
	public ContentConfigVO updateActionPlan(QuestionearVO questionearVO, String dist, String chechBoxVal) 
			throws JCTException {
		LOGGER.info(">>>>>> ContentConfigImpl.updateActionplan");					
		ContentConfigVO contentConfigVO = new ContentConfigVO();
		String prevoiusSubQtn = "";
		String mainQuestion = "";
		try {
			JctQuestion question = serviceDAO.fetchRefQtn(questionearVO.getPrimaryKeyVal());
			if (question != null) {
				/***** ADDED FOR JCT PUBLIC VERSION *****/
				prevoiusSubQtn = question.getJctQuestionSubDesc();
				mainQuestion = question.getJctQuestionsDesc();
				/****** ENDED *****/
				question.setJctBsLastmodifiedTs(new Date());
				if (null == dist) {
					question.setJctQuestionsDesc(questionearVO.getQuestionsDesc());
					question.setJctQuestionSubDesc(questionearVO.getQuestionSubDesc());
					question.setJctProfilesDesc(questionearVO.getUserProfileName());
					JctUserProfile profile = new JctUserProfile();
					profile.setJctUserProfile(questionearVO.getUserProfileId());
					question.setJctUserProfile(profile);
				}
				if (null != dist){
					question.setJctQuestionsSoftDelete(0);
				}
				String retString = serviceDAO.updateActionPlan(question);
				if (!retString.equals("success")) {
					contentConfigVO.setStatusCode(StatusConstants.STATUS_FAILURE);
					contentConfigVO.setStatusDesc(this.messageSource.
							getMessage("exception.dao.user.group.update", null, null));
				} else {
					contentConfigVO.setStatusCode(StatusConstants.STATUS_SUCCESS_WITH_RESULT);
					contentConfigVO.setStatusDesc(this.messageSource.getMessage("action.plan.edit.sucess.msg", null, null));
				}
			} else {
				contentConfigVO.setStatusCode(StatusConstants.STATUS_FAILURE);
				contentConfigVO.setStatusDesc(this.messageSource.getMessage("exception.dao.qtn.update.fetch.delete", null, null));
			}
			/***** ADDED FOR JCT PUBLIC VERSION *****/
			if (chechBoxVal.equalsIgnoreCase("true")) {
				//To make global profile change
				contentConfigVO.setStatusDesc(this.messageSource.getMessage("action.plan.global.profile.edit.msg", null, null));
				serviceDAO.updateQuestionFrozen ("AS", mainQuestion, prevoiusSubQtn,
						questionearVO.getQuestionSubDesc());
			}
			/****** ENDED *****/
		} catch (DAOException e) {
			LOGGER.error("----"+e.getLocalizedMessage()+" ----");
			contentConfigVO.setStatusCode(StatusConstants.STATUS_FAILURE);
			throw new JCTException(e.getLocalizedMessage());
		}
		LOGGER.info("<<<<<< contentConfigImpl.updateActionplan");
		return contentConfigVO;
	}
	
	
	@Transactional(propagation = Propagation.REQUIRED)
	public ContentConfigVO deleteActionPlan(QuestionearVO questionearVO, String dist)
			throws JCTException {
		LOGGER.info(">>>>>> ContentConfigImpl.deleteActionPlan");
		ContentConfigVO contentConfigVO = new ContentConfigVO();
		try {
			JctQuestion question = serviceDAO.fetchRefQtn(questionearVO.getPrimaryKeyVal());
			//JctUserGroup userGroup = serviceDAO.fetchUserGroup(userGroupVO.getPrimaryKeyVal());
			question.setJctBsLastmodifiedTs(new Date());
			if (null == dist){
				question.setJctQuestionsDesc(questionearVO.getQuestionsDesc());
				question.setJctQuestionsSoftDelete(1);
			}
			if (null != dist){
				question.setJctQuestionsSoftDelete(0);
			}
			String retString = serviceDAO.deleteActionPlan(question);
			if (!retString.equals("success")) {
				contentConfigVO.setStatusCode(StatusConstants.STATUS_FAILURE);
				contentConfigVO.setStatusDesc(this.messageSource.
						getMessage("exception.dao.user.group.update", null, null));
			} else {
				contentConfigVO.setStatusCode(StatusConstants.STATUS_SUCCESS_WITH_RESULT);
				contentConfigVO.setStatusDesc(retString);
			}
		} catch (DAOException e) {
			LOGGER.error("----"+e.getLocalizedMessage()+" ----");
			contentConfigVO.setStatusCode(StatusConstants.STATUS_FAILURE);
			throw new JCTException(e.getLocalizedMessage());
		}
		LOGGER.info("<<<<<< contentConfigImpl.deleteActionPlan");
		return contentConfigVO;
	}
	
	
	@Transactional(propagation = Propagation.REQUIRED)
	public ContentConfigVO populateExistingFunctionGroup() throws JCTException {
		LOGGER.info(">>>>>> ContentConfigServiceImpl.populateExistingFunctionGroup");
		StringBuilder builder = new StringBuilder("");
		ContentConfigVO vo = new ContentConfigVO();
		try {
			List<FunctionDTO> dtoList = serviceDAO.populateExistingFunctionGroup();
			if (dtoList.size() > 0){
				Iterator<FunctionDTO> itr = dtoList.iterator();
				while (itr.hasNext()) {
					FunctionDTO dto = (FunctionDTO) itr.next();					
					builder.append(dto.getJctFunctionName());					
					builder.append("~");
					builder.append(dto.getJctFunctionsOrder());					
					builder.append("$$$");
				}
				vo.setExistingFunctionGrpList(builder.toString());
				vo.setStatusCode(StatusConstants.STATUS_SUCCESS_WITH_RESULT);
			} else {
				LOGGER.info("---- NO RECORDS FOUND -----");
				vo.setStatusCode(StatusConstants.STATUS_SUCCESS_WITH_NO_RESULT);
				vo.setStatusDesc(this.messageSource.getMessage("waring.no.user.profile", null, null));
			}
		} catch (DAOException e) {
			LOGGER.error("----"+e.getLocalizedMessage()+" ----");
			throw new JCTException(e.getLocalizedMessage());
		}
		LOGGER.info("<<<<<< ContentConfigServiceImpl.populateExistingFunctionGroup");
		return vo;
	}
	
	@Transactional(propagation = Propagation.REQUIRED)
	public ContentConfigVO populateExistingJobLevel() throws JCTException {
		LOGGER.info(">>>>>> ContentConfigServiceImpl.populateExistingJobLevel");
		StringBuilder builder = new StringBuilder("");
		ContentConfigVO vo = new ContentConfigVO();
		try {
			List<LevelDTO> dtoList = serviceDAO.populateExistingJobLevel();
			if (dtoList.size() > 0) {
				Iterator<LevelDTO> itr = dtoList.iterator();
				while (itr.hasNext()) {
					LevelDTO dto = (LevelDTO) itr.next();					
					builder.append(dto.getJctLevelName());					
					builder.append("~");
					builder.append(dto.getJctLevelOrder());					
					builder.append("$$$");
				}
				vo.setExistingJobLevelList(builder.toString());
				vo.setStatusCode(StatusConstants.STATUS_SUCCESS_WITH_RESULT);
			} else {
				LOGGER.info("---- NO RECORDS FOUND -----");
				vo.setStatusCode(StatusConstants.STATUS_SUCCESS_WITH_NO_RESULT);
				vo.setStatusDesc(this.messageSource.getMessage("waring.no.user.profile", null, null));
			}
		} catch (DAOException e) {
			LOGGER.error("----"+e.getLocalizedMessage()+" ----");
			throw new JCTException(e.getLocalizedMessage());
		}
		LOGGER.info("<<<<<< ContentConfigServiceImpl.populateExistingJobLevel");
		return vo;
	}
	
	@Transactional(propagation = Propagation.REQUIRED)
	public ContentConfigVO saveFunctionGroup(String funcGroup, String createdBy) throws JCTException {
		LOGGER.info(">>>>>> ContentConfigServiceImpl.saveFunctionGroup");
		JctFunction jctFunctnGrp = new JctFunction();
		ContentConfigVO contentConfigVO = new ContentConfigVO();
		try {			
			jctFunctnGrp.setJctFunctionsId(commonDaoImpl.generateKey("jct_function_group"));
			jctFunctnGrp.setJctFunctionName(funcGroup);		
			jctFunctnGrp.setJctFunctionsDesc(funcGroup);
			/*jctFunctnGrp.setJctFunctionsOrder(commonDaoImpl.generateOrder("JctFunction", null, null));*/
			jctFunctnGrp.setJctFunctionsOrder(serviceDAO.generateOrderFunctionGrp());
			jctFunctnGrp.setJctBsCreatedBy(createdBy);
			jctFunctnGrp.setJctDsCreatedTs(new Date());
			jctFunctnGrp.setJctBsLastmodifiedBy(createdBy);
			jctFunctnGrp.setJctBsLastmodifiedTs(new Date());
			jctFunctnGrp.setJctFunctionsSoftDelete(CommonConstants.ENTRY_NOT_DELETED);
			jctFunctnGrp.setVersion(1);
			String retString = serviceDAO.saveFunctionGroup(jctFunctnGrp);
			if (!retString.equals("success")) {
				contentConfigVO.setStatusCode(StatusConstants.STATUS_FAILURE);
				contentConfigVO.setStatusDesc(this.messageSource.
						getMessage("exception.dao.user.profile.save", null, null));
			} else {
				contentConfigVO.setStatusCode(StatusConstants.STATUS_SUCCESS_WITH_RESULT);
				contentConfigVO.setStatusDesc(retString);
			}
		} catch (DAOException e) {
			LOGGER.error("----"+e.getLocalizedMessage()+" ----");
			contentConfigVO.setStatusCode(StatusConstants.STATUS_FAILURE);
			throw new JCTException(e.getLocalizedMessage());
		}
		LOGGER.info("<<<<<< ContentConfigServiceImpl.saveFunctionGroup");
		return contentConfigVO;
	}
	
	@Transactional(propagation = Propagation.REQUIRED)
	public ContentConfigVO saveJobLevel(String jobLevel, String createdBy) throws JCTException {
		LOGGER.info(">>>>>> ContentConfigServiceImpl.saveJobLevel");
		JctLevel jctJobLevel = new JctLevel();
		ContentConfigVO contentConfigVO = new ContentConfigVO();
		try {
			jctJobLevel.setJctLevelsId(commonDaoImpl.generateKey("jct_job_level"));
			jctJobLevel.setJctLevelName(jobLevel);		
			jctJobLevel.setJctLevelsDesc(jobLevel);
			//jctJobLevel.setJctLevelsOrder(commonDaoImpl.generateOrder("JctLevel", null, null));
			jctJobLevel.setJctLevelsOrder(serviceDAO.generateOrderLevel());
			jctJobLevel.setJctBsCreatedBy(createdBy);
			jctJobLevel.setJctDsCreatedTs(new Date());
			jctJobLevel.setJctBsLastmodifiedBy(createdBy);
			jctJobLevel.setJctBsLastmodifiedTs(new Date());
			jctJobLevel.setJctLevelsSoftDelete(CommonConstants.ENTRY_NOT_DELETED);
			jctJobLevel.setVersion(1);
			String retString = serviceDAO.saveJobLevel(jctJobLevel);
			if( !retString.equals("success")) {
				contentConfigVO.setStatusCode(StatusConstants.STATUS_FAILURE);
				contentConfigVO.setStatusDesc(this.messageSource.
						getMessage("exception.dao.user.profile.save", null, null));
			} else {
				contentConfigVO.setStatusCode(StatusConstants.STATUS_SUCCESS_WITH_RESULT);
				contentConfigVO.setStatusDesc(retString);
			}
		} catch (DAOException e) {
			LOGGER.error("----"+e.getLocalizedMessage()+" ----");
			contentConfigVO.setStatusCode(StatusConstants.STATUS_FAILURE);
			throw new JCTException(e.getLocalizedMessage());
		}
		LOGGER.info("<<<<<< ContentConfigServiceImpl.saveJobLevel");
		return contentConfigVO;
	}
	
	@Transactional(propagation = Propagation.REQUIRED)
	public ContentConfigVO populateExistingRegion() throws JCTException {
		LOGGER.info(">>>>>> ContentConfigServiceImpl.populateExistingRegion");
		StringBuilder builder = new StringBuilder("");
		ContentConfigVO vo = new ContentConfigVO();
		try {
			List<RegionDTO> dtoList = serviceDAO.populateExistingRegion();
			if (dtoList.size() > 0) {
				Iterator<RegionDTO> itr = dtoList.iterator();
				while (itr.hasNext()) {
					RegionDTO dto = (RegionDTO) itr.next();					
					builder.append(dto.getJctRegionId());					
					builder.append("~");
					builder.append(dto.getJctRegionName());					
					builder.append("$$$");
				}
				vo.setExistingRegionList(builder.toString());
				vo.setStatusCode(StatusConstants.STATUS_SUCCESS_WITH_RESULT);
			} else {
				LOGGER.info("---- NO RECORDS FOUND -----");
				vo.setStatusCode(StatusConstants.STATUS_SUCCESS_WITH_NO_RESULT);
				vo.setStatusDesc(this.messageSource.getMessage("waring.no.user.profile", null, null));
			}
		} catch (DAOException e) {
			LOGGER.error("----"+e.getLocalizedMessage()+" ----");
			throw new JCTException(e.getLocalizedMessage());
		}
		LOGGER.info("<<<<<< ContentConfigServiceImpl.populateExistingRegion");
		return vo;
	}
	
	@Transactional(propagation = Propagation.REQUIRED)
	public ContentConfigVO saveRegion(String regionName, String createdBy) throws JCTException {
		LOGGER.info(">>>>>> ContentConfigServiceImpl.saveRegion");
		JctRegion jctRegion = new JctRegion();
		ContentConfigVO contentConfigVO = new ContentConfigVO();
		try {
			jctRegion.setJctRegionId(commonDaoImpl.generateKey("jct_region"));
			jctRegion.setJctRegionName(regionName);		
			jctRegion.setJctregionDesc(regionName);
			jctRegion.setJctRegionOrder(serviceDAO.generateOrderRegion());
			jctRegion.setJctBsCreatedBy(createdBy);
			jctRegion.setJctDsCreatedTs(new Date());
			jctRegion.setJctBsLastmodifiedBy(createdBy);
			jctRegion.setJctBsLastmodifiedTs(new Date());
			jctRegion.setJctRegionSoftDelete(CommonConstants.ENTRY_NOT_DELETED);
			jctRegion.setVersion(1);
			String retString = serviceDAO.saveRegion(jctRegion);
			if (!retString.equals("success")) {
				contentConfigVO.setStatusCode(StatusConstants.STATUS_FAILURE);
				contentConfigVO.setStatusDesc(this.messageSource.
						getMessage("exception.dao.user.profile.save", null, null));
			} else {
				contentConfigVO.setStatusCode(StatusConstants.STATUS_SUCCESS_WITH_RESULT);
				contentConfigVO.setStatusDesc(retString);
			}
		} catch (DAOException e) {
			LOGGER.error("----"+e.getLocalizedMessage()+" ----");
			contentConfigVO.setStatusCode(StatusConstants.STATUS_FAILURE);
			throw new JCTException(e.getLocalizedMessage());
		}
		LOGGER.info("<<<<<< ContentConfigServiceImpl.saveRegion");
		return contentConfigVO;
	}
	
	@Transactional(propagation = Propagation.REQUIRED)
	public ContentConfigVO updateRegion(String regionName, Integer primaryKey) throws JCTException {
		LOGGER.info(">>>>>> ContentConfigServiceImpl.updateRegion");
		ContentConfigVO contentConfigVO = new ContentConfigVO();
		try {			
			JctRegion region = serviceDAO.fetchRegion(primaryKey);
			if (region != null) {
				region.setJctregionDesc(regionName);
				region.setJctRegionName(regionName);
				region.setJctRegionSoftDelete(0);
				String retString = serviceDAO.updateRegion(region);
				if (!retString.equals("success")) {
					contentConfigVO.setStatusCode(StatusConstants.STATUS_FAILURE);
					contentConfigVO.setStatusDesc(this.messageSource.
							getMessage("exception.dao.user.profile.save", null, null));
				} else {
					contentConfigVO.setStatusCode(StatusConstants.STATUS_SUCCESS_WITH_RESULT);
					contentConfigVO.setStatusDesc(retString);
				}
			} else {
				contentConfigVO.setStatusCode(StatusConstants.STATUS_FAILURE);
				contentConfigVO.setStatusDesc(this.messageSource.getMessage("exception.dao.regn.update.fetch.delete", null, null));
			}
				
		} catch (DAOException e) {
			LOGGER.error("----"+e.getLocalizedMessage()+" ----");
			contentConfigVO.setStatusCode(StatusConstants.STATUS_FAILURE);
			throw new JCTException(e.getLocalizedMessage());
		}
		LOGGER.info("<<<<<< ContentConfigServiceImpl.updateRegion");
		return contentConfigVO;
	}
	
	@Transactional(propagation = Propagation.REQUIRED)
	public ContentConfigVO deleteRegion(String regionName, Integer primaryKey) throws JCTException {
		LOGGER.info(">>>>>> ContentConfigServiceImpl.deleteRegion");
		ContentConfigVO contentConfigVO = new ContentConfigVO();
		try {			
			JctRegion region = serviceDAO.fetchRegion(primaryKey);
			region.setJctRegionSoftDelete(1);
			String retString = serviceDAO.deleteRegion(region);
			if (!retString.equals("success")) {
				contentConfigVO.setStatusCode(StatusConstants.STATUS_FAILURE);
				contentConfigVO.setStatusDesc(this.messageSource.
						getMessage("exception.dao.user.profile.save", null, null));
			} else {
				contentConfigVO.setStatusCode(StatusConstants.STATUS_SUCCESS_WITH_RESULT);
				contentConfigVO.setStatusDesc(retString);
			}		
		} catch (DAOException e) {
			LOGGER.error("----"+e.getLocalizedMessage()+" ----");
			contentConfigVO.setStatusCode(StatusConstants.STATUS_FAILURE);
			throw new JCTException(e.getLocalizedMessage());
		}
		LOGGER.info("<<<<<< ContentConfigServiceImpl.deleteRegion");
		return contentConfigVO;
	}
	
	@Transactional(propagation = Propagation.REQUIRED)
	public ContentConfigVO populateExistingMapping(Integer profileId, String relatedPage) throws JCTException {
		LOGGER.info(">>>>>> ContentConfigServiceImpl.populateExistingMapping");
		StringBuilder builder = new StringBuilder("");
		ContentConfigVO vo = new ContentConfigVO();
		try {
			List<JobAttributeDTO> dtoList = serviceDAO.populateExistingMapping(profileId, relatedPage);
			if (dtoList.size() > 0) {
				Iterator<JobAttributeDTO> itr = dtoList.iterator();
				while (itr.hasNext()) {
					JobAttributeDTO dto = (JobAttributeDTO) itr.next();
					builder.append(dto.getJctJobAttributeCode());
					builder.append("~");
					builder.append(dto.getJctJobAttributeName());
					builder.append("~");
					builder.append(dto.getJctJobAttributeId());
					builder.append("~");
					builder.append(dto.getJctJobAttributeSoftDelete());
					builder.append("~");
					builder.append(dto.getUserProfileDesc());
					builder.append("~");
					builder.append(dto.getJctJobAttributeDesc());
					builder.append("~");
					builder.append(dto.getJctJobAttributeOrder());
					builder.append("~");
					builder.append(dto.getUserProfileID());
					builder.append("$$$");
				}
				vo.setExistingMappingList(builder.toString());
				vo.setStatusCode(StatusConstants.STATUS_SUCCESS_WITH_RESULT);
			} else {
				LOGGER.info("---- NO RECORDS FOUND -----");
				vo.setStatusCode(StatusConstants.STATUS_SUCCESS_WITH_NO_RESULT);
				vo.setStatusDesc(this.messageSource.getMessage("waring.no.user.profile", null, null));
			}
		} catch (DAOException e) {
			LOGGER.error("----"+e.getLocalizedMessage()+" ----");
			throw new JCTException(e.getLocalizedMessage());
		}
		LOGGER.info("<<<<<< ContentConfigServiceImpl.populateExistingMapping");
		return vo;
	}
	
	@Transactional(propagation = Propagation.REQUIRED)
	public ContentConfigVO saveStrength(JobAttributeVO jobAttributeVO)
			throws JCTException {
		LOGGER.info(">>>>>> ContentConfigServiceImpl.saveStrength");
		JctJobAttribute jobAttribute = new JctJobAttribute();
		//JctUserGroup userGroup = new JctUserGroup();
		ContentConfigVO contentConfigVO = new ContentConfigVO();
		try {
			
			jobAttribute.setJctJobAttributeId(commonDaoImpl.generateKey("jct_job_attribute"));
			jobAttribute.setJctJobAttributeCreatedBy(jobAttributeVO.getCreatedBy());
			jobAttribute.setJctJobAttributeCreatedTs(new Date());
			jobAttribute.setJctJobAttributeLastmodifiedBy(jobAttributeVO.getCreatedBy());
			jobAttribute.setJctJobAttributeLastmodifiedTs(new Date());
			jobAttribute.setJctJobAttributeSoftDelete(CommonConstants.ENTRY_NOT_DELETED);			
			jobAttribute.setJctJobAttributeCode(jobAttributeVO.getJctJobAttributeCode());
			jobAttribute.setJctJobAttributeName(jobAttributeVO.getJctJobAttributeName());
			jobAttribute.setJctJobAttributeDesc(jobAttributeVO.getJctJobAttributeDesc());
			//jobAttribute.setJctJobAttributeOrder(jobAttributeVO.getJctJobAttributeOrder());
			//jobAttribute.setJctJobAttributeOrder(commonDaoImpl.generateOrder("Strength", jobAttributeVO.getUserProfileId(), null));
			jobAttribute.setJctJobAttributeOrder(serviceDAO.generateOrderAttr("STR", jobAttributeVO.getUserProfileId()));
			//JctJobAttribute jobAttribute = serviceDAO.fetchMapping(jobAttributeVO.getPrimaryKeyVal());
			
			jobAttribute.setJctProfilesDesc(jobAttributeVO.getUserProfileName());			
			JctUserProfile profile = new JctUserProfile();
			profile.setJctUserProfile(jobAttributeVO.getUserProfileId());
			jobAttribute.setJctUserProfile(profile);
			jobAttribute.setVersion(1);					
			String retString = serviceDAO.saveStrength(jobAttribute);
			if (!retString.equals("success")) {
				contentConfigVO.setStatusCode(StatusConstants.STATUS_FAILURE);
				contentConfigVO.setStatusDesc(this.messageSource.
						getMessage("exception.dao.user.group.save", null, null));
			} else {
				contentConfigVO.setStatusCode(StatusConstants.STATUS_SUCCESS_WITH_RESULT);
				contentConfigVO.setStatusDesc(retString);
			}
		} catch (DAOException e) {
			LOGGER.error("----"+e.getLocalizedMessage()+" ----");
			contentConfigVO.setStatusCode(StatusConstants.STATUS_FAILURE);
			throw new JCTException(e.getLocalizedMessage());
		}
		LOGGER.info("<<<<<< ContentConfigServiceImpl.saveStrength");
		return contentConfigVO;
	}
	
	@Transactional(propagation = Propagation.REQUIRED)
	public ContentConfigVO saveValue(JobAttributeVO jobAttributeVO)
			throws JCTException {
		LOGGER.info(">>>>>> ContentConfigServiceImpl.saveStrength");
		JctJobAttribute jobAttribute = new JctJobAttribute();
		//JctUserGroup userGroup = new JctUserGroup();
		ContentConfigVO contentConfigVO = new ContentConfigVO();
		try {			
			jobAttribute.setJctJobAttributeId(commonDaoImpl.generateKey("jct_job_attribute"));
			jobAttribute.setJctJobAttributeCreatedBy(jobAttributeVO.getCreatedBy());
			jobAttribute.setJctJobAttributeCreatedTs(new Date());
			jobAttribute.setJctJobAttributeLastmodifiedBy(jobAttributeVO.getCreatedBy());
			jobAttribute.setJctJobAttributeLastmodifiedTs(new Date());
			jobAttribute.setJctJobAttributeSoftDelete(CommonConstants.ENTRY_NOT_DELETED);			
			jobAttribute.setJctJobAttributeCode(jobAttributeVO.getJctJobAttributeCode());
			jobAttribute.setJctJobAttributeName(jobAttributeVO.getJctJobAttributeName());
			jobAttribute.setJctJobAttributeDesc(jobAttributeVO.getJctJobAttributeDesc());
			//jobAttribute.setJctJobAttributeOrder(jobAttributeVO.getJctJobAttributeOrder());
			//jobAttribute.setJctJobAttributeOrder(commonDaoImpl.generateOrder("Value", jobAttributeVO.getUserProfileId(), null));
			jobAttribute.setJctJobAttributeOrder(serviceDAO.generateOrderAttr("VAL", jobAttributeVO.getUserProfileId()));
			jobAttribute.setJctProfilesDesc(jobAttributeVO.getUserProfileName());			
			JctUserProfile profile = new JctUserProfile();
			profile.setJctUserProfile(jobAttributeVO.getUserProfileId());
			jobAttribute.setJctUserProfile(profile);
			jobAttribute.setVersion(1);					
			String retString = serviceDAO.saveValue(jobAttribute);
			if (!retString.equals("success")) {
				contentConfigVO.setStatusCode(StatusConstants.STATUS_FAILURE);
				contentConfigVO.setStatusDesc(this.messageSource.
						getMessage("exception.dao.user.group.save", null, null));
			} else {
				contentConfigVO.setStatusCode(StatusConstants.STATUS_SUCCESS_WITH_RESULT);
				contentConfigVO.setStatusDesc(retString);
			}
		} catch (DAOException e) {
			LOGGER.error("----"+e.getLocalizedMessage()+" ----");
			contentConfigVO.setStatusCode(StatusConstants.STATUS_FAILURE);
			throw new JCTException(e.getLocalizedMessage());
		}
		LOGGER.info("<<<<<< ContentConfigServiceImpl.saveStrength");
		return contentConfigVO;
	}
	
	@Transactional(propagation = Propagation.REQUIRED)
	public ContentConfigVO savePassion(JobAttributeVO jobAttributeVO)
			throws JCTException {
		LOGGER.info(">>>>>> ContentConfigServiceImpl.savePassion");
		JctJobAttribute jobAttribute = new JctJobAttribute();
		//JctUserGroup userGroup = new JctUserGroup();
		ContentConfigVO contentConfigVO = new ContentConfigVO();
		try {			
			jobAttribute.setJctJobAttributeId(commonDaoImpl.generateKey("jct_job_attribute"));
			jobAttribute.setJctJobAttributeCreatedBy(jobAttributeVO.getCreatedBy());
			jobAttribute.setJctJobAttributeCreatedTs(new Date());
			jobAttribute.setJctJobAttributeLastmodifiedBy(jobAttributeVO.getCreatedBy());
			jobAttribute.setJctJobAttributeLastmodifiedTs(new Date());
			jobAttribute.setJctJobAttributeSoftDelete(CommonConstants.ENTRY_NOT_DELETED);			
			jobAttribute.setJctJobAttributeCode(jobAttributeVO.getJctJobAttributeCode());
			jobAttribute.setJctJobAttributeName(jobAttributeVO.getJctJobAttributeName());
			jobAttribute.setJctJobAttributeDesc(jobAttributeVO.getJctJobAttributeDesc());
			//jobAttribute.setJctJobAttributeOrder(jobAttributeVO.getJctJobAttributeOrder());
			//jobAttribute.setJctJobAttributeOrder(commonDaoImpl.generateOrder("Passion", jobAttributeVO.getUserProfileId(), null));
			jobAttribute.setJctJobAttributeOrder(serviceDAO.generateOrderAttr("PAS", jobAttributeVO.getUserProfileId()));
			jobAttribute.setJctProfilesDesc(jobAttributeVO.getUserProfileName());			
			JctUserProfile profile = new JctUserProfile();
			profile.setJctUserProfile(jobAttributeVO.getUserProfileId());
			jobAttribute.setJctUserProfile(profile);
			jobAttribute.setVersion(1);					
			String retString = serviceDAO.savePassion(jobAttribute);
			if (!retString.equals("success")) {
				contentConfigVO.setStatusCode(StatusConstants.STATUS_FAILURE);
				contentConfigVO.setStatusDesc(this.messageSource.
						getMessage("exception.dao.user.group.save", null, null));
			} else {
				contentConfigVO.setStatusCode(StatusConstants.STATUS_SUCCESS_WITH_RESULT);
				contentConfigVO.setStatusDesc(retString);
			}
		} catch (DAOException e) {
			LOGGER.error("----"+e.getLocalizedMessage()+" ----");
			contentConfigVO.setStatusCode(StatusConstants.STATUS_FAILURE);
			throw new JCTException(e.getLocalizedMessage());
		}
		LOGGER.info("<<<<<< ContentConfigServiceImpl.savePassion");
		return contentConfigVO;
	}
	
	@Transactional(propagation = Propagation.REQUIRED)
	public ContentConfigVO updateStrength(JobAttributeVO jobAttributeVO, String dist, String chechBoxVal)
			throws JCTException {
		LOGGER.info(">>>>>> ContentConfigImpl.updateStrength");
		ContentConfigVO contentConfigVO = new ContentConfigVO();
		String prevoiusName = "";
		String prevoiusDesc = "";
		try {
			JctJobAttribute jobAttribute = serviceDAO.fetchMapping(jobAttributeVO.getPrimaryKeyVal());
			if (jobAttribute != null) {
				/***** ADDED FOR JCT PUBLIC VERSION *****/
				prevoiusName = jobAttribute.getJctJobAttributeName();
				prevoiusDesc = jobAttribute.getJctJobAttributeDesc();
				/****** ENDED *****/
				jobAttribute.setJctJobAttributeLastmodifiedTs(new Date());
				if (null == dist){
					jobAttribute.setJctJobAttributeName(jobAttributeVO.getJctJobAttributeName());
					jobAttribute.setJctJobAttributeDesc(jobAttributeVO.getJctJobAttributeDesc());
					jobAttribute.setJctProfilesDesc(jobAttributeVO.getUserProfileName());
					JctUserProfile profile = new JctUserProfile();
					profile.setJctUserProfile(jobAttributeVO.getUserProfileId());
					jobAttribute.setJctUserProfile(profile);
				}
				if (null != dist) {
					jobAttribute.setJctJobAttributeSoftDelete(0);
				}
				String retString = serviceDAO.updateStrength(jobAttribute);
				if (!retString.equals("success")) {
					contentConfigVO.setStatusCode(StatusConstants.STATUS_FAILURE);
					contentConfigVO.setStatusDesc(this.messageSource.
							getMessage("exception.dao.user.group.update", null, null));
				} else {
					contentConfigVO.setStatusCode(StatusConstants.STATUS_SUCCESS_WITH_RESULT);
					contentConfigVO.setStatusDesc(this.messageSource.getMessage("strength.edit.sucess.msg", null, null));
				}
			} else {
				contentConfigVO.setStatusCode(StatusConstants.STATUS_FAILURE);
				contentConfigVO.setStatusDesc(this.messageSource.getMessage("exception.dao.job.attr.update.fetch.delete", null, null));
			}
			/***** ADDED FOR JCT PUBLIC VERSION *****/
			if (chechBoxVal.equalsIgnoreCase("true")) {
				//To make global profile change
				contentConfigVO.setStatusDesc(this.messageSource.getMessage("strength.global.profile.edit.msg", null, null));
				serviceDAO.updateMappingFrozen ("STR", prevoiusName, 
						jobAttributeVO.getJctJobAttributeName(),prevoiusDesc, jobAttributeVO.getJctJobAttributeDesc());
			}
			/****** ENDED *****/
		} catch (DAOException e) {
			LOGGER.error("----"+e.getLocalizedMessage()+" ----");
			contentConfigVO.setStatusCode(StatusConstants.STATUS_FAILURE);
			throw new JCTException(e.getLocalizedMessage());
		}
		LOGGER.info("<<<<<< contentConfigImpl.updateStrength");
		return contentConfigVO;
	}
	
	@Transactional(propagation = Propagation.REQUIRED)
	public ContentConfigVO updateValue(JobAttributeVO jobAttributeVO, String dist, String chechBoxVal)
			throws JCTException {
		LOGGER.info(">>>>>> ContentConfigImpl.updateValue");	
		ContentConfigVO contentConfigVO = new ContentConfigVO();
		String prevoiusName = "";
		String prevoiusDesc = "";
		try {
			JctJobAttribute jobAttribute = serviceDAO.fetchMapping(jobAttributeVO.getPrimaryKeyVal());
			if (jobAttribute != null) {
				/***** ADDED FOR JCT PUBLIC VERSION *****/
				prevoiusName = jobAttribute.getJctJobAttributeName();
				prevoiusDesc = jobAttribute.getJctJobAttributeDesc();
				/****** ENDED *****/
				jobAttribute.setJctJobAttributeLastmodifiedTs(new Date());
				if (null == dist){
					jobAttribute.setJctJobAttributeName(jobAttributeVO.getJctJobAttributeName());
					jobAttribute.setJctJobAttributeDesc(jobAttributeVO.getJctJobAttributeDesc());
					jobAttribute.setJctProfilesDesc(jobAttributeVO.getUserProfileName());
					JctUserProfile profile = new JctUserProfile();
					profile.setJctUserProfile(jobAttributeVO.getUserProfileId());
					jobAttribute.setJctUserProfile(profile);
				}
				if (null != dist) {
					jobAttribute.setJctJobAttributeSoftDelete(0);
				}
				String retString = serviceDAO.updateValue(jobAttribute);
				if (!retString.equals("success")) {
					contentConfigVO.setStatusCode(StatusConstants.STATUS_FAILURE);
					contentConfigVO.setStatusDesc(this.messageSource.
							getMessage("exception.dao.user.group.update", null, null));
				} else {
					contentConfigVO.setStatusCode(StatusConstants.STATUS_SUCCESS_WITH_RESULT);
					contentConfigVO.setStatusDesc(this.messageSource.getMessage("value.edit.sucess.msg", null, null));
				}
			} else {
				contentConfigVO.setStatusCode(StatusConstants.STATUS_FAILURE);
				contentConfigVO.setStatusDesc(this.messageSource.getMessage("exception.dao.job.attr.update.fetch.delete", null, null));
			}
			/***** ADDED FOR JCT PUBLIC VERSION *****/
			if (chechBoxVal.equalsIgnoreCase("true")) {
				//To make global profile change
				contentConfigVO.setStatusDesc(this.messageSource.getMessage("value.global.profile.edit.msg", null, null));
				serviceDAO.updateMappingFrozen ("VAL", prevoiusName, 
						jobAttributeVO.getJctJobAttributeName(),prevoiusDesc, jobAttributeVO.getJctJobAttributeDesc());
			}
			/****** ENDED *****/
		} catch (DAOException e) {
			LOGGER.error("----"+e.getLocalizedMessage()+" ----");
			contentConfigVO.setStatusCode(StatusConstants.STATUS_FAILURE);
			throw new JCTException(e.getLocalizedMessage());
		}
		LOGGER.info("<<<<<< contentConfigImpl.updateValue");
		return contentConfigVO;
	}
	
	@Transactional(propagation = Propagation.REQUIRED)
	public ContentConfigVO updatePassion(JobAttributeVO jobAttributeVO, String dist, String chechBoxVal)
			throws JCTException {
		LOGGER.info(">>>>>> ContentConfigImpl.updatePassion");
		ContentConfigVO contentConfigVO = new ContentConfigVO();
		String prevoiusName = "";
		String prevoiusDesc = "";
		try {
			JctJobAttribute jobAttribute = serviceDAO.fetchMapping(jobAttributeVO.getPrimaryKeyVal());		
			if (jobAttribute != null) {
				jobAttribute.setJctJobAttributeLastmodifiedTs(new Date());
				/***** ADDED FOR JCT PUBLIC VERSION *****/
				prevoiusName = jobAttribute.getJctJobAttributeName();
				prevoiusDesc = jobAttribute.getJctJobAttributeDesc();
				/****** ENDED *****/
				if (null == dist) {
					jobAttribute.setJctJobAttributeName(jobAttributeVO.getJctJobAttributeName());
					jobAttribute.setJctJobAttributeDesc(jobAttributeVO.getJctJobAttributeDesc());
					jobAttribute.setJctProfilesDesc(jobAttributeVO.getUserProfileName());
					JctUserProfile profile = new JctUserProfile();
					profile.setJctUserProfile(jobAttributeVO.getUserProfileId());
					jobAttribute.setJctUserProfile(profile);
				}
				if (null != dist) {
					jobAttribute.setJctJobAttributeSoftDelete(0);
				}
				String retString = serviceDAO.updatePassion(jobAttribute);
				if (!retString.equals("success")) {
					contentConfigVO.setStatusCode(StatusConstants.STATUS_FAILURE);
					contentConfigVO.setStatusDesc(this.messageSource.
							getMessage("exception.dao.user.group.update", null, null));
				} else {
					contentConfigVO.setStatusCode(StatusConstants.STATUS_SUCCESS_WITH_RESULT);
					contentConfigVO.setStatusDesc(this.messageSource.getMessage("passion.edit.sucess.msg", null, null));
				}
			} else {
				contentConfigVO.setStatusCode(StatusConstants.STATUS_FAILURE);
				contentConfigVO.setStatusDesc(this.messageSource.getMessage("exception.dao.job.attr.update.fetch.delete", null, null));
			}
			/***** ADDED FOR JCT PUBLIC VERSION *****/
			if (chechBoxVal.equalsIgnoreCase("true")) {
				//To make global profile change
				contentConfigVO.setStatusDesc(this.messageSource.getMessage("passion.global.profile.edit.msg", null, null));
				serviceDAO.updateMappingFrozen ("PAS", prevoiusName, 
						jobAttributeVO.getJctJobAttributeName(),prevoiusDesc, jobAttributeVO.getJctJobAttributeDesc());
			}
			/****** ENDED *****/
		} catch (DAOException e) {
			LOGGER.error("----"+e.getLocalizedMessage()+" ----");
			contentConfigVO.setStatusCode(StatusConstants.STATUS_FAILURE);
			throw new JCTException(e.getLocalizedMessage());
		}
		LOGGER.info("<<<<<< contentConfigImpl.updatePassion");
		return contentConfigVO;
	}
	
	@Transactional(propagation = Propagation.REQUIRED)
	public ContentConfigVO deleteStrength(JobAttributeVO jobAttributeVO, String dist)
			throws JCTException {
		LOGGER.info(">>>>>> ContentConfigImpl.deleteStrength");
		ContentConfigVO contentConfigVO = new ContentConfigVO();
		try {
			JctJobAttribute jobAttribute = serviceDAO.fetchMapping(jobAttributeVO.getPrimaryKeyVal());
			
			//JctUserGroup userGroup = serviceDAO.fetchUserGroup(userGroupVO.getPrimaryKeyVal());
			jobAttribute.setJctJobAttributeLastmodifiedTs(new Date());
			if (null == dist){
				jobAttribute.setJctJobAttributeDesc(jobAttributeVO.getJctJobAttributeDesc());
				jobAttribute.setJctJobAttributeSoftDelete(1);
			}
			if (null != dist){
				jobAttribute.setJctJobAttributeSoftDelete(0);
			}
			String retString = serviceDAO.deleteStrength(jobAttribute);
			if (!retString.equals("success")) {
				contentConfigVO.setStatusCode(StatusConstants.STATUS_FAILURE);
				contentConfigVO.setStatusDesc(this.messageSource.
						getMessage("exception.dao.user.group.update", null, null));
			} else {
				contentConfigVO.setStatusCode(StatusConstants.STATUS_SUCCESS_WITH_RESULT);
				contentConfigVO.setStatusDesc(retString);
			}
		} catch (DAOException e) {
			LOGGER.error("----"+e.getLocalizedMessage()+" ----");
			contentConfigVO.setStatusCode(StatusConstants.STATUS_FAILURE);
			throw new JCTException(e.getLocalizedMessage());
		}
		LOGGER.info("<<<<<< contentConfigImpl.deleteStrength");
		return contentConfigVO;
	}
	
	@Transactional(propagation = Propagation.REQUIRED)
	public ContentConfigVO deleteValue(JobAttributeVO jobAttributeVO, String dist)
			throws JCTException {
		LOGGER.info(">>>>>> ContentConfigImpl.deleteValue");
		ContentConfigVO contentConfigVO = new ContentConfigVO();
		try {
			JctJobAttribute jobAttribute = serviceDAO.fetchMapping(jobAttributeVO.getPrimaryKeyVal());
			
			//JctUserGroup userGroup = serviceDAO.fetchUserGroup(userGroupVO.getPrimaryKeyVal());
			jobAttribute.setJctJobAttributeLastmodifiedTs(new Date());
			if (null == dist){
				jobAttribute.setJctJobAttributeDesc(jobAttributeVO.getJctJobAttributeDesc());
				jobAttribute.setJctJobAttributeSoftDelete(1);
			}
			if (null != dist){
				jobAttribute.setJctJobAttributeSoftDelete(0);
			}
			String retString = serviceDAO.deleteValue(jobAttribute);
			if (!retString.equals("success")) {
				contentConfigVO.setStatusCode(StatusConstants.STATUS_FAILURE);
				contentConfigVO.setStatusDesc(this.messageSource.
						getMessage("exception.dao.user.group.update", null, null));
			} else {
				contentConfigVO.setStatusCode(StatusConstants.STATUS_SUCCESS_WITH_RESULT);
				contentConfigVO.setStatusDesc(retString);
			}
		} catch (DAOException e) {
			LOGGER.error("----"+e.getLocalizedMessage()+" ----");
			contentConfigVO.setStatusCode(StatusConstants.STATUS_FAILURE);
			throw new JCTException(e.getLocalizedMessage());
		}
		LOGGER.info("<<<<<< contentConfigImpl.deleteValue");
		return contentConfigVO;
	}
	
	@Transactional(propagation = Propagation.REQUIRED)
	public ContentConfigVO deletePassion(JobAttributeVO jobAttributeVO, String dist)
			throws JCTException {
		LOGGER.info(">>>>>> ContentConfigImpl.deletePassion");
		ContentConfigVO contentConfigVO = new ContentConfigVO();
		try {
			JctJobAttribute jobAttribute = serviceDAO.fetchMapping(jobAttributeVO.getPrimaryKeyVal());
			
			//JctUserGroup userGroup = serviceDAO.fetchUserGroup(userGroupVO.getPrimaryKeyVal());
			jobAttribute.setJctJobAttributeLastmodifiedTs(new Date());
			if (null == dist){
				jobAttribute.setJctJobAttributeDesc(jobAttributeVO.getJctJobAttributeDesc());
				jobAttribute.setJctJobAttributeSoftDelete(1);
			}
			if (null != dist){
				jobAttribute.setJctJobAttributeSoftDelete(0);
			}
			String retString = serviceDAO.deletePassion(jobAttribute);
			if (!retString.equals("success")) {
				contentConfigVO.setStatusCode(StatusConstants.STATUS_FAILURE);
				contentConfigVO.setStatusDesc(this.messageSource.
						getMessage("exception.dao.user.group.update", null, null));
			} else {
				contentConfigVO.setStatusCode(StatusConstants.STATUS_SUCCESS_WITH_RESULT);
				contentConfigVO.setStatusDesc(retString);
			}
		} catch (DAOException e) {
			LOGGER.error("----"+e.getLocalizedMessage()+" ----");
			contentConfigVO.setStatusCode(StatusConstants.STATUS_FAILURE);
			throw new JCTException(e.getLocalizedMessage());
		}
		LOGGER.info("<<<<<< contentConfigImpl.deletePassion");
		return contentConfigVO;
	}
	
	
	@Transactional(propagation = Propagation.REQUIRED)
	public ContentConfigVO saveInstruction(InstructionVO instructionVO)
			throws JCTException {
		LOGGER.info(">>>>>> ContentConfigServiceImpl.saveInstruction");
		JctInstructionBar instruction = new JctInstructionBar();
		//JctUserGroup userGroup = new JctUserGroup();
		ContentConfigVO contentConfigVO = new ContentConfigVO();
		try {			
			instruction.setJctInstructionBarId(commonDaoImpl.generateKey("jct_instruction"));
			instruction.setJctInstructionBarCreatedBy(instructionVO.getCreatedBy());
			instruction.setJctInstructionBarCreatedTs(new Date());
			instruction.setJctInstructionBarLastmodifiedBy(instructionVO.getCreatedBy());
			instruction.setJctInstructionBarLastmodifiedTs(new Date());
			instruction.setJctInstructionBarSoftDelete(CommonConstants.ENTRY_NOT_DELETED);			
			instruction.setJctInstructionBarDesc(instructionVO.getJctInstructionDesc());
			instruction.setJctPageDetails(instructionVO.getJctRelatedPageDesc());
			instruction.setJctProfilesDesc(instructionVO.getUserProfileName());					
			JctUserProfile profile = new JctUserProfile();
			profile.setJctUserProfile(instructionVO.getUserProfileId());
			instruction.setJctUserProfile(profile);
			instruction.setVersion(1);		
			String retString = serviceDAO.saveInstruction(instruction);
			if (!retString.equals("success")) {
				contentConfigVO.setStatusCode(StatusConstants.STATUS_FAILURE);
				contentConfigVO.setStatusDesc(this.messageSource.
						getMessage("exception.dao.user.group.save", null, null));
			} else {
				contentConfigVO.setStatusCode(StatusConstants.STATUS_SUCCESS_WITH_RESULT);
				contentConfigVO.setStatusDesc(retString);
			}
		} catch (DAOException e) {
			LOGGER.error("----"+e.getLocalizedMessage()+" ----");
			contentConfigVO.setStatusCode(StatusConstants.STATUS_FAILURE);
			throw new JCTException(e.getLocalizedMessage());
		}
		LOGGER.info("<<<<<< ContentConfigServiceImpl.saveInstruction");
		return contentConfigVO;
	}
	
	
	@Transactional(propagation = Propagation.REQUIRED)
	public ContentConfigVO populateExistingInstruction(int userProfileId, String relatedPage) throws JCTException {
		LOGGER.info(">>>>>> ContentConfigServiceImpl.populateExistingRegion");
		//StringBuilder builder = new StringBuilder("");
		ContentConfigVO vo = new ContentConfigVO();
		try {
			JctInstructionBar instruction = serviceDAO.populateExistingInstruction(userProfileId, relatedPage);
			if( null != instruction) {
				//builder.append(instruction.getJctInstructionBarDesc());
				vo.setExistingInstructionId(instruction.getJctInstructionBarId());
				vo.setExistingInstruction(instruction.getJctInstructionBarDesc());
				vo.setExistingInstructionType(instruction.getJctInstructionType());
				vo.setExistingInstructionVideo(instruction.getJctVideoPath());
				vo.setStatusCode(StatusConstants.STATUS_SUCCESS_WITH_RESULT);
			}
			/*if (dtoList.size() > 0) {
				Iterator itr = dtoList.iterator();
				while (itr.hasNext()) {
					RegionDTO dto = (RegionDTO) itr.next();					
					builder.append(dto.getJctRegionId());					
					builder.append("~");
					builder.append(dto.getJctRegionName());					
					builder.append("$$$");
				}
				vo.setExistingRegionList(builder.toString());
				vo.setStatusCode(StatusConstants.STATUS_SUCCESS_WITH_RESULT);
			}*/ else {
				LOGGER.info("---- NO RECORDS FOUND -----");
				vo.setStatusCode(StatusConstants.STATUS_SUCCESS_WITH_NO_RESULT);
				vo.setStatusDesc(this.messageSource.getMessage("waring.no.user.profile", null, null));
			}
		} catch (DAOException e) {
			LOGGER.error("----"+e.getLocalizedMessage()+" ----");
			throw new JCTException(e.getLocalizedMessage());
		}
		LOGGER.info("<<<<<< ContentConfigServiceImpl.populateExistingRegion");
		return vo;
	}
	
	@Transactional(propagation=Propagation.REQUIRED)
	public String validateExistenceRefQtn(String refQtnDesc, int userProfileId)
			throws JCTException {
		LOGGER.info(">>>>>> ContentConfigServiceImpl.validateExistenceRefQtn");
		String statusMessage = "failure";
		try {
			statusMessage = serviceDAO.validateRefQtn(refQtnDesc, userProfileId);
		} catch (DAOException e) { 
			LOGGER.error("UNABLE TO FIND REF QTN: "+refQtnDesc+". FOR PROFILE: "+userProfileId);
			statusMessage = "failure";
		}
		LOGGER.info("<<<<<< ContentConfigServiceImpl.validateExistenceRefQtn");
		return statusMessage;
	}
	
	@Transactional(propagation=Propagation.REQUIRED)
	public String validateExistenceActionPlan(String refQtnDesc, String subQtnDesc , int userProfileId, String code, String action)
			throws JCTException {
		LOGGER.info(">>>>>> ContentConfigServiceImpl.validateExistenceActionPlan");
		String statusMessage = "failure";
		try {
			statusMessage = serviceDAO.validateActionPlan(refQtnDesc, subQtnDesc, userProfileId, code, action);
		} catch (DAOException e) {
			LOGGER.error("UNABLE TO FIND ACTION PLAN QTN: "+refQtnDesc+". SUB QTN :"+subQtnDesc+" FOR PROFILE: "+userProfileId);
			statusMessage = "failure";
		}
		LOGGER.info("<<<<<< ContentConfigServiceImpl.validateExistenceActionPlan");
		return statusMessage;
	}
	
	@Transactional(propagation=Propagation.REQUIRED)
	public String validateExistenceFuncGrp(String functionGrp)
			throws JCTException {
		LOGGER.info(">>>>>> ContentConfigServiceImpl.validateExistenceFuncGrp");
		String statusMessage = "failure";
		try {
			statusMessage = serviceDAO.validateFuncGrp(functionGrp);
		} catch (DAOException e) {
			LOGGER.error("UNABLE TO FIND FUNC GRP: "+functionGrp);
			statusMessage = "failure";
		}
		LOGGER.info("<<<< ContentConfigServiceImpl.validateExistenceFuncGrp");
		return statusMessage;
	}
	
	@Transactional(propagation=Propagation.REQUIRED)
	public String validateExistenceJobLevel(String jobLevel)
			throws JCTException {
		LOGGER.info(">>>>>> ContentConfigServiceImpl.validateExistenceJobLevel");
		String statusMessage = "failure";
		try {
			statusMessage = serviceDAO.validateJobLevel(jobLevel);
		} catch (DAOException e) {
			LOGGER.error("UNABLE TO FIND JOB LEVEL: "+jobLevel);
			statusMessage = "failure";
		}
		LOGGER.info("<<<< ContentConfigServiceImpl.validateExistenceJobLevel");
		return statusMessage;
	}
	
	@Transactional(propagation=Propagation.REQUIRED)
	public String validateExistenceAttribute(String attrDescText, int userProfileId, String attrCode, String action)
			throws JCTException {
		LOGGER.info(">>>>>> ContentConfigServiceImpl.validateExistenceAttribute");
		String statusMessage = "failure";
		try {
			statusMessage = serviceDAO.validateAttribute(attrDescText, userProfileId, attrCode, action);
		} catch (DAOException e) {
			LOGGER.error("UNABLE TO FIND ATTRIBUTE: "+attrDescText+". FOR PROFILE: "+userProfileId);
			statusMessage = "failure";
		}
		LOGGER.info("<<<< ContentConfigServiceImpl.validateExistenceAttribute");
		return statusMessage;
	}
	
	@Transactional(propagation=Propagation.REQUIRED)
	public String validateExistenceRegion(String regionName)
			throws JCTException {
		LOGGER.info(">>>>>> ContentConfigServiceImpl.validateExistenceRegion");
		String statusMessage = "failure";
		try {
			statusMessage = serviceDAO.validateRegion(regionName);
		} catch (DAOException e) {
			LOGGER.error("UNABLE TO FIND REGION: "+regionName);
			statusMessage = "failure";
		}
		LOGGER.info("<<<<<< ContentConfigServiceImpl.validateExistenceRegion");
		return statusMessage;
	}
	
	
	@Transactional(propagation = Propagation.REQUIRED)
	public ContentConfigVO updateInstruction(int primaryKey,
			String valueDesc) throws JCTException {
		LOGGER.info(">>>>>> ContentConfigServiceImpl.updateInstruction");
		ContentConfigVO vo = new ContentConfigVO();
		//MyAccountVO myAccountVO = new MyAccountVO();
		//String value = "firstname";
		try {			
			JctInstructionBar insBar = serviceDAO.fetchInstructionData(primaryKey);		
			insBar.setJctInstructionBarSoftDelete(1);
			//user.setJctUserSoftDelete(1);
			//try {
				JctInstructionBar insBarNew = new JctInstructionBar();
				insBarNew.setJctInstructionBarId(commonDaoImpl.generateKey("jct_instruction"));	
				insBarNew.setJctInstructionBarSoftDelete(0);			
				insBarNew.setJctInstructionBarDesc(valueDesc);
				
				
				insBarNew.setJctInstructionBarCreatedBy(insBar.getJctInstructionBarCreatedBy());
				insBarNew.setJctInstructionBarCreatedTs(new Date());
				insBarNew.setJctInstructionBarLastmodifiedBy(insBar.getJctInstructionBarLastmodifiedBy());
				insBarNew.setJctInstructionBarLastmodifiedTs(new Date());
				insBarNew.setJctPageDetails(insBar.getJctPageDetails());
				insBarNew.setJctProfilesDesc(insBar.getJctProfilesDesc());	
				insBarNew.setJctUserProfile(insBar.getJctUserProfile());								
				insBarNew.setVersion(1);	
				//String retString1 = serviceDAO.updateInstruction(insBar);
				serviceDAO.updateInstruction(insBar);
				String retString = serviceDAO.updateInstruction(insBarNew);
				vo.setExistingInstruction(insBarNew.getJctInstructionBarDesc());
				vo.setExistingInstructionId(insBarNew.getJctInstructionBarId());
				if (!retString.equals("success")) {					
					vo.setStatusCode(StatusConstants.STATUS_FAILURE);
					vo.setStatusDesc(this.messageSource.
							getMessage("exception.dao.user.group.save", null, null));
				} else {
					vo.setStatusCode(StatusConstants.STATUS_SUCCESS_WITH_RESULT);
					vo.setStatusDesc(retString);
				}			
			//} catch (Exception e){
			//	LOGGER.error(e.getLocalizedMessage());
			//}
		} catch (DAOException e) {
			LOGGER.error("----"+e.getLocalizedMessage()+" ----");
			vo.setStatusCode(StatusConstants.STATUS_FAILURE);
			throw new JCTException(e.getLocalizedMessage());
		}
		LOGGER.info("<<<<<< ContentConfigServiceImpl.updateInstruction");
		return vo;
	}
	
	@Transactional(propagation = Propagation.REQUIRED)
	public ContentConfigVO populateMainQtn(Integer profileId, String relatedPage) throws JCTException {
		LOGGER.info(">>>>>> ContentConfigServiceImpl.populateMainQtn");
		StringBuilder builder = new StringBuilder("");
		ContentConfigVO vo = new ContentConfigVO();
		try {
			List<QuestionearDTO> dtoList = serviceDAO.populateMainQtn(profileId, relatedPage);
			if (dtoList.size() > 0) {
				Iterator<QuestionearDTO> itr = dtoList.iterator();
				while (itr.hasNext()) {
					QuestionearDTO dto = (QuestionearDTO) itr.next();
					builder.append(dto.getJctQuestionsDesc());
					builder.append("~");										
					builder.append(dto.getJctQuestionsOrder());
					builder.append("$$$");
				}
				vo.setMainQtnList(builder.toString());
				vo.setStatusCode(StatusConstants.STATUS_SUCCESS_WITH_RESULT);
			} else {
				LOGGER.info("---- NO RECORDS FOUND -----");
				vo.setStatusCode(StatusConstants.STATUS_SUCCESS_WITH_NO_RESULT);
				vo.setStatusDesc(this.messageSource.getMessage("waring.no.user.profile", null, null));
			}
		} catch (DAOException e) {
			LOGGER.error("----"+e.getLocalizedMessage()+" ----");
			throw new JCTException(e.getLocalizedMessage());
		}
		LOGGER.info("<<<<<< ContentConfigServiceImpl.populateMainQtn");
		return vo;
	}
	
	@Transactional(propagation = Propagation.REQUIRED)
	public ContentConfigVO populateSubQtn(Integer profileId, String relatedPage, String mainQtn) throws JCTException {
		LOGGER.info(">>>>>> ContentConfigServiceImpl.populateSubQtn");
		StringBuilder builder = new StringBuilder("");
		ContentConfigVO vo = new ContentConfigVO();
		try {
			List<QuestionearDTO> dtoList = serviceDAO.populateSubQtn(profileId, relatedPage, mainQtn);
			if (dtoList.size() > 0) {
				Iterator<QuestionearDTO> itr = dtoList.iterator();
				while (itr.hasNext()) {
					QuestionearDTO dto = (QuestionearDTO) itr.next();
					builder.append(dto.getJctQuestionsDesc());
					builder.append("~");										
					builder.append(dto.getJctQuestionsOrder());
					builder.append("$$$");
				}
				vo.setSubQtnList(builder.toString());
				vo.setStatusCode(StatusConstants.STATUS_SUCCESS_WITH_RESULT);
			} else {
				LOGGER.info("---- NO RECORDS FOUND -----");
				vo.setStatusCode(StatusConstants.STATUS_SUCCESS_WITH_NO_RESULT);
				vo.setStatusDesc(this.messageSource.getMessage("waring.no.user.profile", null, null));
			}
		} catch (DAOException e) {
			LOGGER.error("----"+e.getLocalizedMessage()+" ----");
			throw new JCTException(e.getLocalizedMessage());
		}
		LOGGER.info("<<<<<< ContentConfigServiceImpl.populateSubQtn");
		return vo;
	}

	@Transactional(propagation = Propagation.REQUIRED)
	public ContentConfigVO saveMainQtnOrder(String mainQtnbuilder, Integer profileId,
			String relatedPage) throws JCTException {
		LOGGER.info(">>>>>> ContentConfigServiceImpl.saveMainQtnOrder");
		ContentConfigVO contentConfigVO = new ContentConfigVO();
		String retString = null;
		try {
			retString = serviceDAO.saveMainQtnOrder(mainQtnbuilder, profileId, relatedPage);
		} catch (DAOException e) {
			LOGGER.error("----"+e.getLocalizedMessage()+" ----");
			throw new JCTException(e.getLocalizedMessage());
		}
		
		if (!retString.equals("success")) {
			contentConfigVO.setStatusCode(StatusConstants.STATUS_FAILURE);
			contentConfigVO.setStatusDesc(this.messageSource.
					getMessage("exception.dao.user.group.update", null, null));
		} else {
			contentConfigVO.setStatusCode(StatusConstants.STATUS_SUCCESS_WITH_RESULT);
			contentConfigVO.setStatusDesc(retString);
		}
		LOGGER.info("<<<<<< ContentConfigServiceImpl.saveMainQtnOrder");
		return contentConfigVO;
	}

	@Transactional(propagation = Propagation.REQUIRED)
	public ContentConfigVO saveSubQtnOrder(String subQtnBuilder, Integer profileId,
			String relatedPage, String mainQtn) throws JCTException {
		LOGGER.info(">>>>>> ContentConfigServiceImpl.saveSubQtnOrder");
		ContentConfigVO contentConfigVO = new ContentConfigVO();
		String retString = null;
		try {
			retString = serviceDAO.saveSubQtnOrder(subQtnBuilder, profileId, relatedPage, mainQtn);
		} catch (DAOException e) {
			LOGGER.error("----"+e.getLocalizedMessage()+" ----");
			throw new JCTException(e.getLocalizedMessage());
		}
		
		if (!retString.equals("success")) {
			contentConfigVO.setStatusCode(StatusConstants.STATUS_FAILURE);
			contentConfigVO.setStatusDesc(this.messageSource.
					getMessage("exception.dao.user.group.update", null, null));
		} else {
			contentConfigVO.setStatusCode(StatusConstants.STATUS_SUCCESS_WITH_RESULT);
			contentConfigVO.setStatusDesc(retString);
		}
		LOGGER.info("<<<<<< ContentConfigServiceImpl.saveSubQtnOrder");
		return contentConfigVO;
	}

	@Transactional(propagation = Propagation.REQUIRED)
	public ContentConfigVO saveFuncGrpOrder(String builder) throws JCTException {
		LOGGER.info(">>>>>> ContentConfigServiceImpl.saveFuncGrpOrder");
		ContentConfigVO contentConfigVO = new ContentConfigVO();
		String retString = null;
		try {
			retString = serviceDAO.saveFuncGrpOrder(builder);
		} catch (DAOException e) {
			LOGGER.error("----"+e.getLocalizedMessage()+" ----");
			throw new JCTException(e.getLocalizedMessage());
		}
		
		if (!retString.equals("success")) {
			contentConfigVO.setStatusCode(StatusConstants.STATUS_FAILURE);
			contentConfigVO.setStatusDesc(this.messageSource.
					getMessage("exception.dao.user.group.update", null, null));
		} else {
			contentConfigVO.setStatusCode(StatusConstants.STATUS_SUCCESS_WITH_RESULT);
			contentConfigVO.setStatusDesc(retString);
		}
		LOGGER.info("<<<<<< ContentConfigServiceImpl.saveFuncGrpOrder");
		return contentConfigVO;
	}
	
	@Transactional(propagation = Propagation.REQUIRED)
	public ContentConfigVO populateFunctionGroupByOrder() throws JCTException {
		LOGGER.info(">>>>>> ContentConfigServiceImpl.populateFunctionGroupByOrder");
		StringBuilder builder = new StringBuilder("");
		ContentConfigVO vo = new ContentConfigVO();
		try {
			List<FunctionDTO> dtoList = serviceDAO.populateFunctionGroupByOrder();
			if (dtoList.size() > 0){
				Iterator<FunctionDTO> itr = dtoList.iterator();
				while (itr.hasNext()) {
					FunctionDTO dto = (FunctionDTO) itr.next();					
					builder.append(dto.getJctFunctionName());					
					builder.append("~");
					builder.append(dto.getJctFunctionsOrder());					
					builder.append("$$$");
				}
				vo.setExistingFunctionGrpList(builder.toString());
				vo.setStatusCode(StatusConstants.STATUS_SUCCESS_WITH_RESULT);
			} else {
				LOGGER.info("---- NO RECORDS FOUND -----");
				vo.setStatusCode(StatusConstants.STATUS_SUCCESS_WITH_NO_RESULT);
				vo.setStatusDesc(this.messageSource.getMessage("waring.no.user.profile", null, null));
			}
		} catch (DAOException e) {
			LOGGER.error("----"+e.getLocalizedMessage()+" ----");
			throw new JCTException(e.getLocalizedMessage());
		}
		LOGGER.info("<<<<<< ContentConfigServiceImpl.populateFunctionGroupByOrder");
		return vo;
	}
	
	@Transactional(propagation = Propagation.REQUIRED)
	public ContentConfigVO saveJobLevelOrder(String builder) throws JCTException {
		LOGGER.info(">>>>>> ContentConfigServiceImpl.saveJobLevelOrder");
		ContentConfigVO contentConfigVO = new ContentConfigVO();
		String retString = null;
		try {
			retString = serviceDAO.saveJobLevelOrder(builder);
		} catch (DAOException e) {
			LOGGER.error("----"+e.getLocalizedMessage()+" ----");
			throw new JCTException(e.getLocalizedMessage());
		}
		
		if (!retString.equals("success")) {
			contentConfigVO.setStatusCode(StatusConstants.STATUS_FAILURE);
			contentConfigVO.setStatusDesc(this.messageSource.
					getMessage("exception.dao.user.group.update", null, null));
		} else {
			contentConfigVO.setStatusCode(StatusConstants.STATUS_SUCCESS_WITH_RESULT);
			contentConfigVO.setStatusDesc(retString);
		}
		LOGGER.info("<<<<<< ContentConfigServiceImpl.saveJobLevelOrder");
		return contentConfigVO;
	}
	
	@Transactional(propagation = Propagation.REQUIRED)
	public ContentConfigVO populateJobLevelByOrder() throws JCTException {
		LOGGER.info(">>>>>> ContentConfigServiceImpl.populateJobLevelByOrder");
		StringBuilder builder = new StringBuilder("");
		ContentConfigVO vo = new ContentConfigVO();
		try {
			List<LevelDTO> dtoList = serviceDAO.populateJobLevelByOrder();
			if (dtoList.size() > 0) {
				Iterator<LevelDTO> itr = dtoList.iterator();
				while (itr.hasNext()) {
					LevelDTO dto = (LevelDTO) itr.next();					
					builder.append(dto.getJctLevelName());					
					builder.append("~");
					builder.append(dto.getJctLevelOrder());					
					builder.append("$$$");
				}
				vo.setExistingJobLevelList(builder.toString());
				vo.setStatusCode(StatusConstants.STATUS_SUCCESS_WITH_RESULT);
			} else {
				LOGGER.info("---- NO RECORDS FOUND -----");
				vo.setStatusCode(StatusConstants.STATUS_SUCCESS_WITH_NO_RESULT);
				vo.setStatusDesc(this.messageSource.getMessage("waring.no.user.profile", null, null));
			}
		} catch (DAOException e) {
			LOGGER.error("----"+e.getLocalizedMessage()+" ----");
			throw new JCTException(e.getLocalizedMessage());
		}
		LOGGER.info("<<<<<< ContentConfigServiceImpl.populateJobLevelByOrder");
		return vo;
	}
	
	@Transactional(propagation = Propagation.REQUIRED)
	public ContentConfigVO populateMappingList(Integer profileId, String attrCode) throws JCTException {
		LOGGER.info(">>>>>> ContentConfigServiceImpl.populateMappingList");
		StringBuilder builder = new StringBuilder("");
		ContentConfigVO vo = new ContentConfigVO();
		try {
			List<JobAttributeDTO> dtoList = serviceDAO.populateMappingListByOrder(profileId, attrCode);
			if (dtoList.size() > 0){
				Iterator<JobAttributeDTO> itr = dtoList.iterator();
				while (itr.hasNext()) {
					JobAttributeDTO dto = (JobAttributeDTO) itr.next();					
					builder.append(dto.getJctJobAttributeName());					
					builder.append("~");
					builder.append(dto.getJctJobAttributeOrder());					
					builder.append("$$$");
				}
				vo.setExistingMappingList(builder.toString());
				vo.setStatusCode(StatusConstants.STATUS_SUCCESS_WITH_RESULT);
			} else {
				LOGGER.info("---- NO RECORDS FOUND -----");
				vo.setStatusCode(StatusConstants.STATUS_SUCCESS_WITH_NO_RESULT);
				vo.setStatusDesc(this.messageSource.getMessage("waring.no.user.profile", null, null));
			}
		} catch (DAOException e) {
			LOGGER.error("----"+e.getLocalizedMessage()+" ----");
			throw new JCTException(e.getLocalizedMessage());
		}
		LOGGER.info("<<<<<< ContentConfigServiceImpl.populateMappingList");
		return vo;
	}

	@Transactional(propagation = Propagation.REQUIRED)
	public ContentConfigVO saveAttributeOrder(String builder,
			Integer profileId, String attrCode) throws JCTException {
		LOGGER.info(">>>>>> ContentConfigServiceImpl.saveAttributeOrder");
		ContentConfigVO contentConfigVO = new ContentConfigVO();
		String retString = null;
		try {
			retString = serviceDAO.saveAttributeOrder(builder, profileId, attrCode);
		} catch (DAOException e) {
			LOGGER.error("----"+e.getLocalizedMessage()+" ----");
			throw new JCTException(e.getLocalizedMessage());
		}
		
		if (!retString.equals("success")) {
			contentConfigVO.setStatusCode(StatusConstants.STATUS_FAILURE);
			contentConfigVO.setStatusDesc(this.messageSource.
					getMessage("exception.dao.user.group.update", null, null));
		} else {
			contentConfigVO.setStatusCode(StatusConstants.STATUS_SUCCESS_WITH_RESULT);
			contentConfigVO.setStatusDesc(retString);
		}
		LOGGER.info("<<<<<< ContentConfigServiceImpl.saveAttributeOrder");
		return contentConfigVO;
	}

	@Transactional(propagation = Propagation.REQUIRED)
	public ContentConfigVO saveRegionOrder(String builder) throws JCTException {
		LOGGER.info(">>>>>> ContentConfigServiceImpl.saveRegionOrder");
		ContentConfigVO contentConfigVO = new ContentConfigVO();
		String retString = null;
		try {
			retString = serviceDAO.saveRegionOrder(builder);
		} catch (DAOException e) {
			LOGGER.error("----"+e.getLocalizedMessage()+" ----");
			throw new JCTException(e.getLocalizedMessage());
		}
		
		if (!retString.equals("success")) {
			contentConfigVO.setStatusCode(StatusConstants.STATUS_FAILURE);
			contentConfigVO.setStatusDesc(this.messageSource.
					getMessage("exception.dao.user.group.update", null, null));
		} else {
			contentConfigVO.setStatusCode(StatusConstants.STATUS_SUCCESS_WITH_RESULT);
			contentConfigVO.setStatusDesc(retString);
		}
		LOGGER.info("<<<<<< ContentConfigServiceImpl.saveRegionOrder");
		return contentConfigVO;
	}

	/******* THIS IS FOR JCT PUBLIC VERSION ********/
	/**
	 * Method populates existing ONet data
	 * @param null
	 * @return ContentConfigVO
	 */
	@Transactional(propagation = Propagation.REQUIRED)
	public ContentConfigVO populateOnetDataList() throws JCTException {
		LOGGER.info(">>>>>> ContentConfigServiceImpl.populateOnetDataList");
		StringBuilder builder = new StringBuilder("");
		ContentConfigVO vo = new ContentConfigVO();
		try {
			List<OnetOccupationDTO> dtoList = serviceDAO.populateOnetDataList();
			if (dtoList.size() > 0) {
				Iterator<OnetOccupationDTO> itr = dtoList.iterator();
				while (itr.hasNext()) {
					OnetOccupationDTO dto = (OnetOccupationDTO) itr.next();					
					builder.append(dto.getJctOnetOccupationId());					
					builder.append("~");
					builder.append(dto.getJctOnetOccupationCode());	
					builder.append("~");
					builder.append(dto.getJctOnetOccupationTitle());	
					builder.append("~");
					builder.append(dto.getJctOnetOccupationDesc());	
					builder.append("$$$");
				}
				vo.setOnetOccupationList(builder.toString());
				vo.setStatusCode(StatusConstants.STATUS_SUCCESS_WITH_RESULT);
			} else {
				LOGGER.info("---- NO RECORDS FOUND -----");
				vo.setStatusCode(StatusConstants.STATUS_SUCCESS_WITH_NO_RESULT);
				vo.setStatusDesc(this.messageSource.getMessage("waring.no.onet.data", null, null));
			}
		} catch (DAOException e) {
			LOGGER.error("----"+e.getLocalizedMessage()+" ----");
			throw new JCTException(e.getLocalizedMessage());
		}
		LOGGER.info("<<<<<< ContentConfigServiceImpl.populateOnetDataList");
		return vo;
	}
	
	@Transactional(propagation = Propagation.REQUIRED)
	public List<OnetOccupationDTO> populateOnetDataListVo() throws JCTException {
		LOGGER.info(">>>>>> ContentConfigServiceImpl.populateOnetDataListVo");
		List<OnetOccupationDTO> dtoList = new ArrayList<OnetOccupationDTO>();
		ContentConfigVO vo = new ContentConfigVO();
		try {
			dtoList = serviceDAO.populateOnetDataList();
		} catch (DAOException e) {
			LOGGER.error("----"+e.getLocalizedMessage()+" ----");
			throw new JCTException(e.getLocalizedMessage());
		}
		return dtoList;
	}

	/**
	 * Method to fetch existing terms & condition
	 * @param userProfile
	 * @return ContentConfigVO
	 * @throws JCTException
	 */
	@Transactional(propagation = Propagation.REQUIRED)
	public JctTermsAndConditions fetchTermsAndCondition(int userProfile, int userType) throws JCTException {
		LOGGER.info("<<<<<< ContentConfigServiceImpl.fetchTermsAndCondition");
		JctTermsAndConditions obj = new JctTermsAndConditions(); 
		JctUserProfile jctUserProfile = new JctUserProfile();
		try {
			jctUserProfile = serviceDAO.getUserProfileByPk(userProfile);
			obj = serviceDAO.fetchTermsAndConditionDao(jctUserProfile.getJctUserProfile(), userType);
		} catch(Exception e){
			LOGGER.error("----"+e.getLocalizedMessage()+" ----");			
		}
		LOGGER.info("<<<<<< ContentConfigServiceImpl.fetchTermsAndCondition");
		return obj;
	}

	/**
	 * Method to add/Update terms & condition
	 * @param userProfile
	 * @param userType
	 * @return ContentConfigVO
	 * @throws JCTException
	 */
	@Transactional(propagation = Propagation.REQUIRED)
	public String saveTermsAndCondition(int userProfile, int userType,
			String tcText, String createdBy) throws JCTException {
		LOGGER.info("<<<<<< ContentConfigServiceImpl.saveTermsAndCondition");
		
		JctTermsAndConditions jctTermsAndConditions = new JctTermsAndConditions();
		JctTermsAndConditions existingTermsAndConditions = new JctTermsAndConditions();
		JctUserProfile jctUserProfile = null;
		String isSuccess = "fail";			
		try {
			jctUserProfile = serviceDAO.getUserProfileByPk(userProfile);			
			// fetch existing T&C for selected user profile
			existingTermsAndConditions = serviceDAO.fetchTermsAndConditionDao(jctUserProfile.getJctUserProfile(), userType);
			if(null != existingTermsAndConditions){	//	if T&C exist then hard delete it first, then add a new row
				serviceDAO.removeTermsAndCondition(existingTermsAndConditions);
			}
			jctTermsAndConditions.setJctTcDescription(tcText);
			jctTermsAndConditions.setJctUserProfile(jctUserProfile);
			jctTermsAndConditions.setJctTcUserType(userType);
			jctTermsAndConditions.setJctTcCreated_by(createdBy);
			jctTermsAndConditions.setJctTcCreatedOn(new Date());
			jctTermsAndConditions.setJctTcModifiedBy(createdBy);
			jctTermsAndConditions.setJctTcModifiedOn(new Date());				
			isSuccess = serviceDAO.saveTermsAndCondition(jctTermsAndConditions);					
		} catch (Exception e) {
			LOGGER.error("----"+e.getLocalizedMessage()+" ----");
		}		
		LOGGER.info("<<<<<< ContentConfigServiceImpl.saveTermsAndCondition");
		return isSuccess;
	}

	@Transactional(propagation = Propagation.REQUIRED)
	public String saveVideoInstruction(InstructionVO instructionVO)
			throws JCTException {
		String status = "success";

		LOGGER.info(">>>>>> ContentConfigServiceImpl.saveInstruction");
		JctInstructionBar instruction = new JctInstructionBar();
		ContentConfigVO contentConfigVO = new ContentConfigVO();
		try {
			// Fetch existing obj. Merge 
			JctInstructionBar obj = serviceDAO.populateExistingInstruction(instructionVO.getUserProfileId(), instructionVO.getJctRelatedPageDesc());
			if (obj != null) {
				obj.setJctInstructionBarSoftDelete(CommonConstants.ENTRY_SOFT_DELETED);
				serviceDAO.updateInstruction(obj);
			}
			String fileName = "BS_INSTRUCTION";
			Integer key = commonDaoImpl.generateKey("jct_instruction");
			instruction.setJctInstructionBarId(key);
			instruction.setJctInstructionBarCreatedBy(instructionVO.getCreatedBy());
			instruction.setJctInstructionBarCreatedTs(new Date());
			instruction.setJctInstructionBarLastmodifiedBy(instructionVO.getCreatedBy());
			instruction.setJctInstructionBarLastmodifiedTs(new Date());
			instruction.setJctInstructionBarSoftDelete(CommonConstants.ENTRY_NOT_DELETED);			
			instruction.setJctInstructionBarDesc(instructionVO.getJctInstructionDesc());
			instruction.setJctPageDetails(instructionVO.getJctRelatedPageDesc());
			instruction.setJctProfilesDesc(instructionVO.getUserProfileName());
			String baseVideoPath = this.messageSource.getMessage("video.path",null, null);
			String videoPath = baseVideoPath + "VIDEO-" + fileName+"-"+key+".mp4";
			if (instructionVO.getVideoFile().length != 0) {
				if (instructionVO.getJctRelatedPageDesc().equals("Mapping Yourself")) {
					fileName = "AS_MAP_INSTRUCTION";
				} else if (instructionVO.getJctRelatedPageDesc().equals("Creating After Diagram")){
					fileName = "AS_AFTER_INSTRUCTION";
				}
				instruction.setJctInstructionType("TEXTANDVIDEO");
				instruction.setJctVideoPath(videoPath);
//				instruction.setJctVideoPath("../../JCT_VIDEO/VIDEO-"+fileName+"-"+key+".mp4");
			} else {
				instruction.setJctInstructionType("TEXT");
			}
			JctUserProfile profile = new JctUserProfile();
			profile.setJctUserProfile(instructionVO.getUserProfileId());
			instruction.setJctUserProfile(profile);
			if (instructionVO.getVideoFile().length != 0) {
				status = "success";
				storageService.store(videoPath, new ByteArrayInputStream(instructionVO.getVideoFile()), "video/mp4");
//				boolean isSuccess;
//				isSuccess = saveFileOnDisk(instructionVO.getVideoFile(), instructionVO.getJctRelatedPageDesc(), fileName, key);
//				if (isSuccess) {
//					status = serviceDAO.saveInstruction(instruction);
//				} else {
//					status = "failure";
//				}
			} 
			status = serviceDAO.saveInstruction(instruction);
			if (!status.equals("success")) {
				status = "failure";
			}
		} catch (DAOException e) {
			LOGGER.error("----"+e.getLocalizedMessage()+" ----");
			contentConfigVO.setStatusCode(StatusConstants.STATUS_FAILURE);
			status = "failure";
			throw new JCTException(e.getLocalizedMessage());
		} catch (Exception e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
			status = "failure";
		}
		LOGGER.info("<<<<<< ContentConfigServiceImpl.saveInstruction");
		
		return status;
	}
	
	@Deprecated
	private boolean saveFileOnDisk (byte[] bytes, String fileType, String fileName, int videoFileId) throws IOException {
		boolean isSuccess = false;
//		File videoFile = new File("../webapps/JCT_VIDEO/VIDEO-" + fileName +
//   				"-" + videoFileId+ ".mp4");
		String videoPath = this.messageSource.getMessage("video.path",null, null);
		File videoFile = new File(videoPath + "JCT_VIDEO/VIDEO-" + fileName +
   				"-" + videoFileId+ ".mp4");
		/*File videoFile = new File("D:/JCT_SOFTWARE/apache-tomcat-6.0.44/webapps/JCT_VIDEO/VIDEO-" + fileName +
   				"-" + videoFileId+ ".mp4");*/
		BufferedOutputStream stream = new BufferedOutputStream(new FileOutputStream(videoFile));
		LOGGER.info("VIDEO FILE --> "+videoFile.getAbsolutePath());
		stream.write(bytes);
		stream.close();
		isSuccess = true;
		return isSuccess;
	}
}