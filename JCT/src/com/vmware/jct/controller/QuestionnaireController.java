package com.vmware.jct.controller;

import java.util.ArrayList;
import java.util.Calendar;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.StringTokenizer;

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
import com.vmware.jct.dao.dto.JobAttributeDTO;
import com.vmware.jct.exception.JCTException;
import com.vmware.jct.service.IAfterSketchService;
import com.vmware.jct.service.IAuthenticatorService;
import com.vmware.jct.service.IJobAttributeFrozenService;
import com.vmware.jct.service.IQuestionnaireService;
import com.vmware.jct.service.vo.AfterSketchOneVO;
import com.vmware.jct.service.vo.QuestionnaireVO;
import com.vmware.jct.service.vo.ReturnVO;
/**
 * 
 * <p><b>Class name:</b> QuestionnaireController.java</p>
 * <p><b>Author:</b> InterraIt</p>
 * <p><b>Purpose:</b> This class acts as a controller for Questionnaire.
 * QuestionnaireController extends BasicController  and has following  methods.
 * -saveAnswers()
 * <p><b>Description:</b> This class basically intercept all questionnaire service calls and 
 * then delegated the it to the service layer for further actions. </p>
 * <p><b>Copyrights:</b> 	All rights reserved by Interra IT and should only be used for its internal application development.</p>
 * <p><b>Created Date:</b> 23/Jan/2014</p>
 * <p><b>Revision History:</b>
 * 	<li>InterraIt - 04/Feb/2014 - Implemented comments</li>
 * </p>
 */
@Controller
@RequestMapping(value="/questionnaire")
public class QuestionnaireController extends BasicController{
	
	@Autowired
	private IQuestionnaireService iQuestionnaireService;
	
	@Autowired
	private IAfterSketchService afterSketchService;
	
	@Autowired
	private IAuthenticatorService authenticatorService;
	
	@Autowired
	private IJobAttributeFrozenService attributeFrozenService;

	@Autowired  
	private MessageSource messageSource;
	
	private static final Logger LOGGER = LoggerFactory.getLogger(QuestionnaireController.class);
	
	/**
	 * Method authenticates the user and returns to the specified 
	 * page.
	 * @return UserVO
	 */
	@RequestMapping(value = ACTION_SAVE_ANSWERS, method = RequestMethod.POST)
	@ResponseBody()
	public AfterSketchOneVO saveAnswers(@RequestBody String body,
			HttpServletRequest request) {
		LOGGER.info(">>>> QuestionnaireController.saveAnswers");
		AfterSketchOneVO afterSketchOneVO=new AfterSketchOneVO();
		Calendar startTime = Calendar.getInstance();
		JsonNode node = CommonUtility.getNode(body);
		List<QuestionnaireVO> qtnrVOList = new ArrayList<QuestionnaireVO>();
		String jobReferenceNumber = node.get("jobRefNo").asText();
		int rowId = Integer.parseInt(node.get("rowId").asText());
		String myPlan = node.get("myAns").asText();
		String user = node.get("user").asText();
		int profileId = Integer.parseInt(node.get("profileId").toString().replaceAll("\"", ""));
		String linkClicked = node.get("linkClicked").asText();
		String isCompleted = node.get("isCompleted").asText();
		String isPrevious = node.get("isPrevious").asText();
		/** THIS IS FOR PUBLIC VERSION **/
		int jctUserId = Integer.parseInt(node.get("jctUserId").asText());
		/********************************/
		StringTokenizer token = new StringTokenizer(myPlan, ")(");
		while (token.hasMoreTokens()) {
			
			String totalToken = (String) token.nextToken();
			StringTokenizer innerTokenizer = new StringTokenizer(totalToken, "@@@");
			while (innerTokenizer.hasMoreTokens()) {
				String question = innerTokenizer.nextToken().toString();
				String answerString = innerTokenizer.nextToken().toString();
				StringTokenizer answerToken = new StringTokenizer(answerString, "#");
				while (answerToken.hasMoreTokens()) {
					String subQtnAndAns = (String) answerToken.nextToken();
					StringTokenizer finalToken = new StringTokenizer(subQtnAndAns, "~");
					while (finalToken.hasMoreElements()) {
						String subQtns = (String) finalToken.nextToken();
						try {
							String answers = (String) finalToken.nextToken();
							String[] answerArr = answers.split("`");
							for (int answerIndex = 0; answerIndex < 1; answerIndex++) {
								
								QuestionnaireVO answerVO = new QuestionnaireVO();
								/** THIS IS FOR PUBLIC VERSION **/
								answerVO.setJctUserId(jctUserId);
								/********************************/
								answerVO.setJobReferenceNumber(jobReferenceNumber);
								answerVO.setQuestionDescription(question);
								answerVO.setAnswerDescription(answerArr[answerIndex].replaceAll("-_-", "#").replaceAll(":=:", "(").replaceAll(";=;", ")").replaceAll(":-:", "@").replaceAll(";_;", "?"));
								answerVO.setQtnSubDesc(subQtns);
								answerVO.setCreatedBy(user);
								qtnrVOList.add(answerVO);
							}
						} catch (ArrayIndexOutOfBoundsException ex) {
							LOGGER.info("ArrayIndexOutOfBoundsException..."+ex.getMessage());
						} catch (Exception ex) {
							LOGGER.info("Exception..."+ex.getMessage());
						}
					}
				}
			}
		}
		try{
			//actionPlanService.saveActinoPlans(actionPlanVOList, jobReferenceNumber);
			iQuestionnaireService.saveOrUpdateAnswers(qtnrVOList, jobReferenceNumber, isCompleted);
			
			List<JobAttributeDTO> jctJobAttributeList=null;
			
			Map<Integer, String> strengthMap=new HashMap<Integer, String>();
			Map<Integer, String> passionMap=new HashMap<Integer, String>();
			Map<Integer, String> valueMap=new HashMap<Integer, String>();
			try {
				/** THIS IS FOR PUBLIC VERSION **/
				//jctJobAttributeList=iQuestionnaireService.fetchJobAttribute(profileId);
				jctJobAttributeList = attributeFrozenService.fetchJobAttribute(profileId, user, jctUserId);
				/********************************/
				int j=0;
				for(int i=1;i<=jctJobAttributeList.size();i++){
					if(jctJobAttributeList.get(j).getJctJobAttributeCode().equals(CommonConstants.JOB_ATTRIBUTE_STRENGTH)){
						strengthMap.put(jctJobAttributeList.get(j).jctJobAttributeId, jctJobAttributeList.get(j).getJctJobAttributeName()+"+++"+jctJobAttributeList.get(j).getJctJobAttributeDesc().replace(",", "~~"));
					}else if(jctJobAttributeList.get(j).getJctJobAttributeCode().equals(CommonConstants.JOB_ATTRIBUTE_PASSION)){
						passionMap.put(jctJobAttributeList.get(j).jctJobAttributeId, jctJobAttributeList.get(j).getJctJobAttributeName()+"+++"+jctJobAttributeList.get(j).getJctJobAttributeDesc().replace(",", "~~"));
					}else{
						valueMap.put(jctJobAttributeList.get(j).jctJobAttributeId, jctJobAttributeList.get(j).getJctJobAttributeName()+"+++"+jctJobAttributeList.get(j).getJctJobAttributeDesc().replace(",", "~~"));
					}
					j++;
					
				}
				afterSketchOneVO.setStrengthMap(strengthMap);
				afterSketchOneVO.setPassionMap(passionMap);
				afterSketchOneVO.setValueMap(valueMap);
				
				afterSketchOneVO.setAfterSketchCheckedEle(afterSketchService.fetchCheckedElements(jobReferenceNumber, linkClicked)); //CheckedElements
				afterSketchOneVO.setAfterSkPageOneTotalJson(afterSketchService.fetchASOneJson(jobReferenceNumber, linkClicked));// json
				afterSketchOneVO.setPassionCount(afterSketchService.getPassionCount(jobReferenceNumber));
				afterSketchOneVO.setValueCount(afterSketchService.getValueCount(jobReferenceNumber));
				afterSketchOneVO.setStrengthCount(afterSketchService.getStrengthCount(jobReferenceNumber));
				
				
				String str = afterSketchService.getCheckedItems(jobReferenceNumber, linkClicked);
				StringTokenizer strToken = new StringTokenizer(str, "#");
				while(strToken.hasMoreTokens()){
					afterSketchOneVO.setCheckedStrength(strToken.nextToken().toString());
					afterSketchOneVO.setCheckedPassion(strToken.nextToken().toString());
					afterSketchOneVO.setCheckedValue(strToken.nextToken().toString());
				}
			} catch (JCTException e) {
				LOGGER.error(e.getLocalizedMessage());
			}
			/*long diff = Calendar.getInstance().getTimeInMillis() - startTime.getTimeInMillis();
			LOGGER.info("-------------- Saving Questionnaire Took: "+diff+" Milli seconds ------------------");
			LOGGER.info("<<<< QuestionnaireController.saveAnswers");
			return afterSketchOneVO;*/
			//authenticatorService.updateLoginInfo(jobReferenceNumber, rowId, "/user/view/Questionaire.jsp");
			if (isPrevious.equals("Y")) {
				authenticatorService.updateLoginInfo(jobReferenceNumber, rowId, "/user/view/beforeSketch.jsp");
			} else {
				authenticatorService.updateLoginInfo(jobReferenceNumber, rowId, "/user/view/afterSketch1.jsp");
			}
			
		} catch(JCTException exp){
			LOGGER.error("ERROR IN SAVING QUESTIONNAIRE: "+exp.getLocalizedMessage());
		}
		long diff = Calendar.getInstance().getTimeInMillis() - startTime.getTimeInMillis();
		LOGGER.info("-------------- Saving Questionnaire Took: "+diff+" Milli seconds ------------------");
		LOGGER.info("<<<< ActionPlanController.saveActionPlan");
		return afterSketchOneVO;
	
	}
	/**
	 * Method fetches questionnaire list
	 * @param body
	 * @param request
	 * @return ReturnVO
	 */
	@RequestMapping(value = ACTION_FETCH_QUESTIONNAIRE, method = RequestMethod.POST)
	@ResponseBody()
	public ReturnVO fetchQuestionnaire(@RequestBody String body, HttpServletRequest request) {
		LOGGER.info(">>>> QuestionnaireController.fetchQuestionnaire");
		Calendar startTime = Calendar.getInstance();
		JsonNode node = CommonUtility.getNode(body);
		ReturnVO vo = new ReturnVO();
		String jobReferenceNo = node.get("jobReferenceNo").asText();
		int profileId = Integer.parseInt(node.get("profileId").asText());
		try {
			vo.setQuestionnaireList(iQuestionnaireService.fetQuestionnaireDynamic(jobReferenceNo, profileId));
		} catch (JCTException e) {
			LOGGER.error("UNABLE TO FETCH QUESTIONNAIRE LIST: "+e.getLocalizedMessage());
		}
		long diff = Calendar.getInstance().getTimeInMillis() - startTime.getTimeInMillis();
		LOGGER.info("-------------- Fetching Questionnaire Took: "+diff+" Milli seconds ------------------");
		LOGGER.info("<<<< QuestionnaireController.fetchQuestionnaire");
		return vo;
	}
}
