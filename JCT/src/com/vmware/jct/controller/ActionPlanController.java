package com.vmware.jct.controller;

import java.util.ArrayList;
import java.util.Calendar;
import java.util.List;
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

import com.vmware.jct.common.utility.CommonUtility;
import com.vmware.jct.exception.JCTException;
import com.vmware.jct.service.IActionPlanService;
import com.vmware.jct.service.IAfterSketchService;
import com.vmware.jct.service.IAuthenticatorService;
import com.vmware.jct.service.IConfirmationService;
import com.vmware.jct.service.IGeneratePDFService;
import com.vmware.jct.service.vo.ActionPlanVO;
import com.vmware.jct.service.vo.ReturnVO;

/**
 * 
 * <p><b>Class name:</b> ActionPlanController.java</p>
 * <p><b>Author:</b> InterraIT</p>
 * <p><b>Purpose:</b> This class acts as a controller for Action Plan..
 * ActionPlanController extends BasicController  and has following  methods.
 * -saveActionPlan()
 * <p><b>Description:</b> This class acts as a controller for Action Plan. </p>
 * <p><b>Copyrights:</b> 	All rights reserved by Interra IT and should only be used for its internal application development.</p>
 * <p><b>Created Date:</b> 20/Feb/2014</p>
 * <p><b>Revision History:</b>
 * 	<li>6/May/2014: Changed the saving of action plan to dynamic.</li>
 * </p>
 */
@Controller
@RequestMapping(value="/actionPlan")
public class ActionPlanController extends BasicController{
	@Autowired  
	private MessageSource messageSource;
	
	@Autowired
	private IAfterSketchService afterSketchService;
	
	@Autowired
	private IConfirmationService confirmationService;
	
	@Autowired
	private IAuthenticatorService authenticatorService;
	
	@Autowired
	private IActionPlanService actionPlanService;
	
	@Autowired
	private IGeneratePDFService pdfService;
	
	private static final Logger LOGGER = LoggerFactory.getLogger(ActionPlanController.class);
	
	/**
	 * Method saves the action plan and then generate the PDF
	 * @param body
	 * @param request
	 * @return ReturnVO
	 */
	@RequestMapping(value = ACTION_SAVE_ACTION_PLAN, method = RequestMethod.POST)
	@ResponseBody()
	public ReturnVO saveActionPlan(@RequestBody String body, HttpServletRequest request) {
		LOGGER.info(">>>> ActionPlanController.saveActionPlan");
		ReturnVO vo = new ReturnVO();
		Calendar startTime = Calendar.getInstance();
		JsonNode node = CommonUtility.getNode(body);
		List<ActionPlanVO> actionPlanVOList = new ArrayList<ActionPlanVO>();
		String jobReferenceNumber = node.get("jobRefNo").asText();
		String myPlan = node.get("myPlan").asText();
		String linkClicked = node.get("linkClicked").asText();
		String pdfRed = node.get("pdfReqs").asText();
		String user = node.get("user").asText();
		int rowId = Integer.parseInt(node.get("rowId").asText());
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
							for (int answerIndex = 0; answerIndex < 3; answerIndex++) {
								ActionPlanVO planVO = new ActionPlanVO();
								/** THIS IS FOR PUBLIC VERSION **/
								planVO.setJctUserId(jctUserId);
								/********************************/
								planVO.setJobRefNo(jobReferenceNumber);
								planVO.setUserModified(user);
								planVO.setQuestionDesc(question);
								planVO.setAnswerDesc(answerArr[answerIndex].replaceAll("-_-", "#").replaceAll(":=:", "(").replaceAll(";=;", ")").replaceAll(":-:", "@").replaceAll(";_;", "?"));
								planVO.setQuestionSubDesc(subQtns);
								actionPlanVOList.add(planVO);
							}
						} catch (ArrayIndexOutOfBoundsException ex) {
							LOGGER.info("ArrayIndexOutOfBoundsException..."+ex.getMessage());
						}catch (Exception ex) {
							LOGGER.info("Exception..."+ex.getMessage());
						}
					}
				}
			}
		}
		try{
			actionPlanService.saveActinoPlans(actionPlanVOList, jobReferenceNumber, linkClicked);
			//Generate the PDF
			String pdfPath = this.messageSource.getMessage("pdf.path",null, null);
			if (pdfRed.equals("Y")) {
				String pdfGeneration = pdfService.generatePdf(jobReferenceNumber, user, pdfPath, request, jctUserId);
				LOGGER.info("PDF GENERATION STATUS: "+pdfGeneration+" FOR JOB REFERENCE NOS: "+jobReferenceNumber);
			}
			//Check if the user has completed at least once then the new result page is to be shown
			int nosOfCompletion = actionPlanService.getNumberOfCompletion(jobReferenceNumber);
			vo.setUrl("resultPageFinal.jsp");
			
			if (nosOfCompletion == 3 || nosOfCompletion == 0) {
				vo.setUrl("resultPage.jsp");
			} else {
				// Fetch the demographic information and formulate the message
				String message = this.messageSource.getMessage("label.demo.details.confirmation",null, null);
				String demoInfo = authenticatorService.getDemographicInformation(user);
				//vo.setDemographicPopupMsg(message + demoInfo);
				vo.setDemographicPopupMsg(demoInfo);
			}
			int completionCount = confirmationService.getExcersizeCompletionCount(jobReferenceNumber);
			String page = "/user/view/resultPage.jsp";
			if (completionCount > 0) {
				page = "/user/view/resultPageFinal.jsp";
			}
			//authenticatorService.updateLoginInfo(jobReferenceNumber, rowId, "/user/view/actionPlan.jsp");
			authenticatorService.updateLoginInfo(jobReferenceNumber, rowId, page);
			
			
		} catch(JCTException exp){
			LOGGER.error("ERROR IN SAVING ACTION PLAN / PDF GENERATION: "+exp.getLocalizedMessage());
		}
		long diff = Calendar.getInstance().getTimeInMillis() - startTime.getTimeInMillis();
		LOGGER.info("-------------- Saving Action Plan Took: "+diff+" Milli seconds ------------------");
		LOGGER.info("<<<< ActionPlanController.saveActionPlan");
		return vo;
	}
	/**
	 * Method saves the action plan and then generate the PDF
	 * @param body
	 * @param request
	 * @return ReturnVO
	 */
	@RequestMapping(value = ACTION_FETCH_ACTION_PLAN, method = RequestMethod.POST)
	@ResponseBody()
	public ReturnVO fetchActionPlan(@RequestBody String body, HttpServletRequest request) {
		LOGGER.info(">>>> ActionPlanController.fetchActionPlan");
		Calendar startTime = Calendar.getInstance();
		JsonNode node = CommonUtility.getNode(body);
		ReturnVO vo = new ReturnVO();
		String jobReferenceNo = node.get("jobReferenceNo").asText();
		int profileId = Integer.parseInt(node.get("profileId").asText());
		try {
			//List returnList = afterSketchService.fetActionPlanDynamic(jobReferenceNo, profileId);
			vo.setActionPlanList(afterSketchService.fetActionPlanDynamic(jobReferenceNo, profileId));
		} catch (JCTException e) {
			LOGGER.error("UNABLE TO FETCH ACTION PLAN LIST: "+e.getLocalizedMessage());
		}
		long diff = Calendar.getInstance().getTimeInMillis() - startTime.getTimeInMillis();
		LOGGER.info("-------------- Fetching Action Plan Took: "+diff+" Milli seconds ------------------");
		LOGGER.info("<<<< ActionPlanController.fetchActionPlan");
		return vo;
	}
}
