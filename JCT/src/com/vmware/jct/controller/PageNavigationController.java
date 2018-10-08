package com.vmware.jct.controller;

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
import com.vmware.jct.model.JctUser;
import com.vmware.jct.service.IAfterSketchService;
import com.vmware.jct.service.IAuthenticatorService;
import com.vmware.jct.service.IBeforeSketchService;
import com.vmware.jct.service.IConfirmationService;
import com.vmware.jct.service.IJobAttributeFrozenService;
import com.vmware.jct.service.IQuestionnaireService;
import com.vmware.jct.service.vo.NavigationVO;
import com.vmware.jct.service.vo.PopulateInstructionVO;
import com.vmware.jct.service.vo.UserLoginInfoVO;
import com.vmware.jct.service.vo.UserVO;

/**
 * 
 * <p><b>Class name:</b> PageNavigationController.java</p>
 * <p><b>Author:</b> InterraIt</p>
 * <p><b>Purpose:</b> This class acts as a controller for page navigation activities.
 * PageNavigationController extends BasicController  and has following  methods.
 * -goPrevious()
 * -populateVO()
 * -populateActionPlan()
 * -populateMappingYourself()
 *-populateAfterDiagram()
 *-populateInstruction()
 * <p><b>Description:</b> This class basically intercept all page navigation service calls and 
 * then delegated the it to the service layer for further actions. </p>
 * <p><b>Copyrights:</b> 	All rights reserved by Interra IT and should only be used for its internal application development.</p>
 * <p><b>Created Date:</b> 27/Jan/2014</p>
 * <p><b>Revision History:</b>
 * 	<li>InterraIt - 04/Feb/2014 - Implemented comments</li>
 * </p>
 */
@Controller
@RequestMapping(value="/navigation")
public class PageNavigationController extends BasicController {
	
	@Autowired
	private IBeforeSketchService beforeSketchService;
	
	@Autowired
	private IQuestionnaireService iQuestionnaireService;
	
	@Autowired
	private IJobAttributeFrozenService attributeFrozenService;
	
	@Autowired
	private IAfterSketchService afterSketchService;
	
	@Autowired
	private IAuthenticatorService authenticatorService;
	
	@Autowired
	private IConfirmationService confirmationService;
	
	@Autowired  
	private MessageSource messageSource;
	
	private static final Logger LOGGER = LoggerFactory.getLogger(PageNavigationController.class);
	
	@RequestMapping(value = ACTION_GO_PREVIOUS, method = RequestMethod.POST)
	@ResponseBody()
	private NavigationVO goPrevious(@RequestBody String body,
			HttpServletRequest request){
		LOGGER.info(">>>> PageNavigationController.goPrevious");
		JsonNode node = CommonUtility.getNode(body);
		NavigationVO vo = new NavigationVO();
		String pageToGo = node.get("navigationSeq").toString().replaceAll("\"", ""); //1, 2 ,3, 4, 5
		String jobReferenceNo = node.get("jobReferenceNos").toString().replaceAll("\"", "");
		int profileId = Integer.parseInt(node.get("profileId").toString().replaceAll("\"", ""));
		int rowId = Integer.parseInt(node.get("rowId").asText());
		String linkClicked = node.get("linkClicked").asText();
		try {
			//Check if completed
			int count = authenticatorService.getIsCompleted(jobReferenceNo);
			if(count > 0){
				vo.setIsCompleted(1);
			} else {
				vo.setIsCompleted(0);
			}
			
			vo = populateVO(pageToGo, jobReferenceNo, profileId, linkClicked, vo, rowId);
		} catch (JCTException e) {
			//e.printStackTrace();
			LOGGER.error(e.getLocalizedMessage());
			
		}
		LOGGER.info("<<<< PageNavigationController.goPrevious");
		return vo;
	}
	private NavigationVO populateVO(String pageToGo, String jobReferenceNo, 
			int profileId, String linkClicked, NavigationVO vo, int rowId) throws JCTException {
		LOGGER.info(">>>> PageNavigationController.populateVO");
		//NavigationVO vo = new NavigationVO();
		if (linkClicked.equals("Y")) {
			int status = confirmationService.updateLinkStatus(jobReferenceNo);
			LOGGER.info("STATUS : "+status);
		} 
		if(pageToGo.equalsIgnoreCase("1")) {
			if (linkClicked.equals("FE")) { // Coming from BS Edit page
				vo.setBsJson(beforeSketchService.getJsonFromTemp(jobReferenceNo, linkClicked));
				vo.setInitialJson(beforeSketchService.getInitialJsonFromTemp(jobReferenceNo, linkClicked));
			} else {
				vo.setBsJson(beforeSketchService.getJson(jobReferenceNo, linkClicked));
				vo.setInitialJson(beforeSketchService.getInitialJson(jobReferenceNo, linkClicked));
			}
		} else if(pageToGo.equalsIgnoreCase("2")){ //questionnaire
			vo.setJctBaseString(beforeSketchService.fetchBase64String(jobReferenceNo));
		} else if(pageToGo.equalsIgnoreCase("3")){ //after sketch 1
			// Get the userId corresponding to the row number
			int userId = authenticatorService.getUserId(rowId);
			vo = populateMappingYourself(vo, jobReferenceNo, profileId, linkClicked, rowId, userId);
		} else if(pageToGo.equalsIgnoreCase("4")){ //after sketch 2
			vo = populateAfterDiagram(vo, jobReferenceNo, linkClicked);
		} else if(pageToGo.equalsIgnoreCase("5")){ //action plan
			vo = populateActionPlan(vo, jobReferenceNo, profileId, linkClicked);
			authenticatorService.updateLoginInfo(jobReferenceNo, rowId, "/user/view/actionPlan.jsp");
		} else if (pageToGo.equalsIgnoreCase("6")) {
			authenticatorService.updateLoginInfo(jobReferenceNo, rowId, "/user/view/actionPlan.jsp");
		}
		LOGGER.info("<<<< PageNavigationController.populateVO");
		return vo;
	}
	private NavigationVO populateActionPlan(NavigationVO vo,
			String jobReferenceNo, int profileId, String linkClicked) throws JCTException{
		LOGGER.info(">>>> PageNavigationController.populateActionPlan");
		//Fetch the snapshot url of the aftersketch diagram
		vo.setSnapShot(afterSketchService.getFinalPageSnapShot(jobReferenceNo, linkClicked));
		LOGGER.info("<<<< PageNavigationController.populateActionPlan");
		return vo;
	}
	/**
	 * Method populates the Mapping yourself screen (3rd screen)
	 * @param userVO
	 * @param jobRefNo
	 * @return
	 * @throws JCTException
	 */
	private NavigationVO populateMappingYourself(NavigationVO navVO, String jobRefNo,
			int profileId, String linkClicked, int rowId, int jctUserId) throws JCTException{
		LOGGER.info(">>>> PageNavigationController.populateMappingYourself");
		
		//Populate the check box
		List<JobAttributeDTO> jctJobAttributeList=null;
		Map<Integer, String> strengthMap=new HashMap<Integer, String>();
		Map<Integer, String> passionMap=new HashMap<Integer, String>();
		Map<Integer, String> valueMap=new HashMap<Integer, String>();
		
		/** THIS IS FOR PUBLIC VERSION **/
		//jctJobAttributeList = iQuestionnaireService.fetchJobAttribute(profileId);
		//Fetch the corresponding user with the rowid from login info table
		
		jctJobAttributeList = attributeFrozenService.fetchJobAttribute(profileId, authenticatorService.getEmailIdByRowNos(rowId), jctUserId);
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
		navVO.setStrengthMap(strengthMap);
		navVO.setPassionMap(passionMap);
		navVO.setValueMap(valueMap);
		navVO.setAfterSketchCheckedEle(afterSketchService.fetchCheckedElements(jobRefNo, linkClicked)); //CheckedElements
		navVO.setAfterSkPageOneTotalJson(afterSketchService.fetchASOneJson(jobRefNo, linkClicked));// json
		navVO.setPassionCount(afterSketchService.getPassionCount(jobRefNo));
		navVO.setValueCount(afterSketchService.getValueCount(jobRefNo));
		navVO.setStrengthCount(afterSketchService.getStrengthCount(jobRefNo));
		
		String str = afterSketchService.getCheckedItems(jobRefNo, linkClicked);
		StringTokenizer strToken = new StringTokenizer(str, "#");
		while(strToken.hasMoreTokens()){
			navVO.setCheckedStrength(strToken.nextToken().toString());
			navVO.setCheckedPassion(strToken.nextToken().toString());
			navVO.setCheckedValue(strToken.nextToken().toString());
		}
		LOGGER.info("<<<< PageNavigationController.populateMappingYourself");
		return navVO;
	}
	
	private NavigationVO populateAfterDiagram(NavigationVO navVO, 
			String jobRefNo, String linkClicked) throws JCTException {
		LOGGER.info(">>>> PageNavigationController.populateAfterDiagram");
		
		//Fetch total_json_object [string / value / strength json]
		navVO.setPageOneJson(afterSketchService.fetchASOneJson(jobRefNo, linkClicked));
		//Fetch total_json_add_task
		navVO.setTotalJsonAddTask(beforeSketchService.getJson(jobRefNo, linkClicked));
		//Set div_initial_value
		navVO.setDivInitialValue(afterSketchService.getFinalPageInitialJson(jobRefNo, linkClicked));
		//fetch the total json of after task
		navVO.setAfterSketch2TotalJsonObj(afterSketchService.fetchASFinalPageJson(jobRefNo, linkClicked));
		String str = afterSketchService.getDivHeightAndWidth(jobRefNo, linkClicked);
		StringTokenizer token = new StringTokenizer(str, "#");
		String height = "";
		String width = "";
		try{
		while(token.hasMoreTokens()){
			height = (String)token.nextToken();
			width = (String)token.nextToken();
		}
		}	catch(Exception ex){
			LOGGER.error("---------------------------->WIDTH AND HEIGHT ISSUE AFTER DRAGGING"+ex.getLocalizedMessage());
		}
		navVO.setDivHeight(height);
		navVO.setDivWidth(width);
		navVO.setAsTotalCount(afterSketchService.getAsTotalCount(jobRefNo, linkClicked));
		navVO.setIsCompleted(0);
		LOGGER.info("<<<< PageNavigationController.populateAfterDiagram");
		return navVO;
	}
	
	@RequestMapping(value = ACTION_POPULATE_INSTRUCTION, method = RequestMethod.POST)
	@ResponseBody()
	public NavigationVO populateInstruction(@RequestBody String body,
			HttpServletRequest request){
		LOGGER.info(">>>> PageNavigationController.populateInstruction");
		
		JsonNode node = CommonUtility.getNode(body);
		NavigationVO vo = new NavigationVO();
		String relatedPage = node.get("relatedPage").toString().replaceAll("\"", "");
		int profileId = Integer.parseInt(node.get("profileId").toString().replaceAll("\"", ""));
		try {
			vo = authenticatorService.populateInstruction(profileId, relatedPage);
		} catch (JCTException e) {
			LOGGER.error(e.getLocalizedMessage());
		}
		LOGGER.info("<<<< PageNavigationController.populateInstruction");
		return vo;
	}
	/**
	 * Method restarts the whole excersize all over again
	 * @param body
	 * @param request
	 * @return NavigationVO
	 */
	@RequestMapping(value = ACTION_RESTART_EXCERSIZE, method = RequestMethod.POST)
	@ResponseBody()
	public UserVO restartExcersize(@RequestBody String body,
			HttpServletRequest request){
		LOGGER.info(">>>> PageNavigationController.restartExcersize");
		JsonNode node = CommonUtility.getNode(body);
		String jobReferenceNo = node.get("jobReferenceNo").asText();
		String emailId = node.get("emailId").asText();
		String distinction = node.get("distinction").asText();
		UserVO userVO = new UserVO();
		userVO.setEmail(emailId);
		try {
			//Initially set the status as failed
			userVO.setStatusCode(CommonConstants.FAILURE);
			JctUser user = (JctUser) authenticatorService.checkUser(userVO).get(0);
			//Call function
			String status = confirmationService.restartExcersize(jobReferenceNo, distinction);
			LOGGER.info("STATUS FOR RESTART EXCERSIZE FOR JOB REFERENCE NOS: "+jobReferenceNo+" IS: "+status+". ");
			if (status.equals("success")) {
				//prepare the blank data
				userVO = prepareBeforeSketchData(user, userVO, jobReferenceNo);

				//Insert into login info table
				userVO = insertLoginInfo(userVO, jobReferenceNo, user.getJctUserId(), "/user/view/beforeSketch.jsp");
				userVO.setStatusCode(CommonConstants.SUCCESSFUL);
			}
		} catch (JCTException ex) {
			LOGGER.error("SOME THING FAILED WHILE RESTARTING THE EXCERSIZE: "+ex.getLocalizedMessage());
			userVO.setStatusDesc(this.messageSource.getMessage("error.restart.failed",null, null));
		}
		LOGGER.info("<<<< PageNavigationController.restartExcersize");
		return userVO;
	}
	private UserVO prepareBeforeSketchData(JctUser user, UserVO userVO, String jobReferenceNo) {
		userVO.setProfileId(user.getJctUserDetails().getJctUserDetailsProfileId());
		userVO.setGroupId(user.getJctUserDetails().getJctUserDetailsGroupId());
		userVO.setJobTitle(user.getJctUserDetails().getJctUserDetailsLevels());
		userVO.setFirstName(user.getJctUserDetails().getJctUserDetailsFirstName());
		userVO.setBsJson("none");
		userVO.setInitialJson("none");
		try {
			userVO.setTotalTime(confirmationService.getTotalTime(user.getJctUserId()));
		} catch (JCTException e) {
			userVO.setTotalTime("00:00:00");
		}
		userVO.setUrl("/user/view/beforeSketch.jsp");
		userVO.setIsCompleted(0);
		userVO.setLastPage("none");
		userVO.setJobRefNo(jobReferenceNo);
		userVO.setPageSequence(1);
		userVO.setIsCompleted(0);
		return userVO;
	}
	private UserVO insertLoginInfo(UserVO userVO, String jobRefNo, int userId, String pageInfo) throws JCTException{
		LOGGER.info(">>>> AuthenticatingContoller.insertLoginInfo");
		UserLoginInfoVO loginInfoVO = new UserLoginInfoVO();
		loginInfoVO.setJobReferenceNo(jobRefNo);
		loginInfoVO.setUserId(userId);
		loginInfoVO.setPageInfo(pageInfo);
		String status = authenticatorService.saveLoginInfo(loginInfoVO);
		LOGGER.info("---------> LOGIN INFO INSERTED FOR RESTART EXCERSIZE: "+status+" FOR JOBREFERENCE NUMBER: "+jobRefNo);
		if(status.contains("success")){
			userVO.setIdentifier(status.substring(status.indexOf("#")+1, status.length()));
		} else{
			userVO.setIdentifier("0");
		}
		LOGGER.info("<<<< AuthenticatingContoller.insertLoginInfo");
		return userVO;
	}
/*Populate PopUp*/
	@RequestMapping(value = ACTION_POPULATE_POPUP, method = RequestMethod.POST)
	@ResponseBody()
	public PopulateInstructionVO populatePopUp(@RequestBody String body,
			HttpServletRequest request){
		LOGGER.info(">>>> PageNavigationController.populatePopUp");
		JsonNode node = CommonUtility.getNode(body);
		PopulateInstructionVO vo = new PopulateInstructionVO();
		/*int pageSequence = Integer.parseInt(node.get("pageSequence").toString().replaceAll("\"", ""));*/
		String relatedPage = node.get("relatedPage").toString().replaceAll("\"", "");
		int profileId = Integer.parseInt(node.get("profileId").toString().replaceAll("\"", ""));
		try {
			vo = authenticatorService.populatePopUp(profileId,relatedPage);
		} catch (JCTException e) {
			LOGGER.error(e.getLocalizedMessage());
		}
		LOGGER.info("<<<< PageNavigationController.populatePopUp");
		return vo;
	}
}
