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
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.RequestBody;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestMethod;
import org.springframework.web.bind.annotation.ResponseBody;

import com.vmware.jct.common.utility.CommonUtility;
import com.vmware.jct.common.utility.EncoderDecoder;
import com.vmware.jct.exception.JCTException;
import com.vmware.jct.service.IAfterSketchService;
import com.vmware.jct.service.IAuthenticatorService;
import com.vmware.jct.service.IBeforeSketchService;
import com.vmware.jct.service.vo.JctBeforeSketchDetailsVO;
import com.vmware.jct.service.vo.JctBeforeSketchHeaderVO;
import com.vmware.jct.service.vo.ReturnVO;
/**
 * 
 * <p><b>Class name:</b> BeforeSketchController.java</p>
 * <p><b>Author:</b> InterraIt</p>
 * <p><b>Purpose:</b> This class acts as a controller for before sketch activities.
 * BeforeSketchController extends BasicController  and has following  methods.
 * -storeBeforeSketch()
 * <p><b>Description:</b> This class basically intercept all the before sketch related service calls and 
 * then delegated the it to the service layer for further actions. </p>
 * <p><b>Copyrights:</b> 	All rights reserved by Interra IT and should only be used for its internal application development.</p>
 * <p><b>Created Date:</b> 20/Jan/2014</p>
 * <p><b>Revision History:</b>
 * 	<li>InterraIt - 04/Feb/2014 - Implemented comments</li>
 * </p>
 */
@Controller
public class BeforeSketchController extends BasicController{
	
	@Autowired
	private IBeforeSketchService beforeSketchService;
	
	@Autowired
	private IAuthenticatorService authenticatorService;
	
	@Autowired
	private IAfterSketchService afterSketchService;
	
	private static final Logger LOGGER = LoggerFactory.getLogger(BeforeSketchController.class);
	/**
	 * Method stores the before sketch diagrams.
	 * @return String
	 */
	@RequestMapping(value = ACTION_STORE_BEFORE_SKETCH, method = RequestMethod.POST)
	@ResponseBody()
    public ReturnVO storeBeforeSketch(@RequestBody String body, HttpServletRequest request) {
		LOGGER.info(">>>> BeforeSketchController.storeBeforeSketch");
		Calendar startTime = Calendar.getInstance();
		ReturnVO returnVO = new ReturnVO();
		JsonNode node = CommonUtility.getNode(body);
		String createdBy = node.get("createdBy").toString().replaceAll("\"", ""); //username is the email id
		String jSon = node.get("jsonString").toString();//.replaceAll("\"", "");
		String objectList = jSon.replace("\"", "");
		String jobReferenceNumber = node.get("jobReferenceString").toString().replaceAll("\"", "");
		String jobTitle = node.get("jobTitle").toString().replaceAll("\"", "");
		String initialJsonVal = node.get("initialJsonVal").toString();//.replaceAll("\"", "");
		String isEdit = node.get("isEdits").asText();
		/** THIS IS FOR PUBLIC VERSION **/
		int jctUserId = Integer.parseInt(node.get("jctUserId").asText());
		/********************************/
		//String linkClicked = node.get("linkClicked").asText();
		int timeTaken = Integer.parseInt(node.get("totalTimes").toString().replaceAll("\"", ""));
		int profileId = Integer.parseInt(node.get("profileId").toString().replaceAll("\"", ""));
		int rowId = Integer.parseInt(node.get("rowId").asText());
		String objectList1 = objectList.replaceAll("},", "#");
		String objectList2 = objectList1.replaceAll("}]", "#");
		String objectList3 = objectList2.substring(2, objectList2.length());
		
		//Create Parent VO Object
		JctBeforeSketchHeaderVO parentVO = new JctBeforeSketchHeaderVO();
		/** THIS IS FOR PUBLIC VERSION **/
		parentVO.setJctUserId(jctUserId);
		/********************************/
		parentVO.setJobReferenceNumber(jobReferenceNumber);
		parentVO.setTotalJson(jSon);
		String file = node.get("snapShot").toString();
		byte[] imageInByte = EncoderDecoder.getImageBytes(file
				.substring(file.indexOf(",") + 1, file.length()));
		parentVO.setSnapShot(imageInByte);
		parentVO.setJobTitle(jobTitle);
		parentVO.setInitialJsonVal(initialJsonVal);
		parentVO.setCreatedBy(createdBy);
		parentVO.setImageBase64String(file);
		parentVO.setTimeTaken(timeTaken);
		//Create Child VO and put it in list
		List<JctBeforeSketchDetailsVO> childList = new ArrayList<JctBeforeSketchDetailsVO>();
		StringTokenizer token = new StringTokenizer(objectList3,"#");
		while(token.hasMoreTokens()){
			String innerToken = token.nextToken().toString();
			if(!innerToken.contains("none")){
				JctBeforeSketchDetailsVO childObj = new JctBeforeSketchDetailsVO();
				
				//FOR ID
				String[] idSplitArray = innerToken.toString().split(",");
				String idSplit = idSplitArray[0];
				StringTokenizer idTok = new StringTokenizer(idSplit, ":");
				while(idTok.hasMoreTokens()){
					idTok.nextToken();
					String idVal = idTok.nextToken().toString();
					childObj.setTaskId(Integer.parseInt(idVal.substring(idVal.indexOf("_")+1, idVal.length())));
				}
				
				childObj.setJobReferenceNumber(jobReferenceNumber);
				String[] split = innerToken.split("energyValue:");
				childObj.setEnergy(Integer.parseInt(split[1].substring(0, split[1].indexOf(","))));
				
				childObj.setPosition("");
				
				String[] task = innerToken.split("areaValue:");
				childObj.setTaskDesc(task[1].substring(0, task[1].indexOf(",")).replaceAll("`", ",").replaceAll("-_-", "#").replaceAll(";_;", "?"));
				
				childObj.setCreatedBy(createdBy);
				childList.add(childObj);
			}
		}
		try {
			beforeSketchService.saveOrUpdateBeforeSketch(parentVO, childList, isEdit);
			
			//Update jct_next_page in login info with BS
			authenticatorService.updateNextPage(jobReferenceNumber, "BS");
			
			if(isEdit.equals("N")) {
				returnVO.setList(beforeSketchService.fetchQuestionAnswer(jobReferenceNumber, profileId));
			} else {
				// populate the after sketch edit page..
				returnVO = populateAfterDiagram(returnVO, jobReferenceNumber, "Y");
			}
			if (!isEdit.equals("Y")) {
				//authenticatorService.updateLoginInfo(jobReferenceNumber, rowId, "/user/view/beforeSketch.jsp");
				authenticatorService.updateLoginInfo(jobReferenceNumber, rowId, "/user/view/Questionaire.jsp");
			}
		} catch (JCTException e) {
			LOGGER.error("ERROR IN SAVING BEFORE SKETCH DIAGRAM:"+e.getLocalizedMessage());
		}
		
		long diff = Calendar.getInstance().getTimeInMillis() - startTime.getTimeInMillis();
		LOGGER.info("-------------- Saving Before Sketch Took: "+diff+" Milli seconds ------------------");
		LOGGER.info(">>>> BeforeSketchController.storeBeforeSketch");
		return returnVO;
	}
	private ReturnVO populateAfterDiagram(ReturnVO returnVO, String jobRefNo, String linkClicked) throws JCTException {
		LOGGER.info(">>>> BeforeSketchController.populateAfterDiagram");
		
		returnVO.setPageOneJson(afterSketchService.fetchASOneJsonForEdit(jobRefNo, linkClicked));
		returnVO.setTotalJsonAddTask(beforeSketchService.getJsonForEdit(jobRefNo, linkClicked));
		returnVO.setDivInitialValue(afterSketchService.getFinalPageInitialJsonForEdit(jobRefNo, linkClicked));
		returnVO.setAfterSketch2TotalJsonObj(afterSketchService.fetchASFinalPageJsonForEdit(jobRefNo, linkClicked));
		String str = afterSketchService.getDivHeightAndWidthForEdit(jobRefNo, linkClicked);
		
		StringTokenizer token = new StringTokenizer(str, "#");
		String height = "";
		String width = "";
		try {
			while (token.hasMoreTokens()) {
				height = (String) token.nextToken();
				width = (String) token.nextToken();
			}
		} catch (Exception ex) {
			LOGGER.error("---------------------------->WIDTH AND HEIGHT ISSUE AFTER DRAGGING");
		}
		returnVO.setDivHeight(height);
		returnVO.setDivWidth(width);
		returnVO.setAsTotalCount(afterSketchService.getAsTotalCountForEdit(jobRefNo, linkClicked));
		LOGGER.info("<<<< BeforeSketchController.populateAfterDiagram");
		return returnVO;
	}
}
