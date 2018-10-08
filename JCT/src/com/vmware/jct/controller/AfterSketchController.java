package com.vmware.jct.controller;

import java.util.ArrayList;
import java.util.Calendar;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.NoSuchElementException;
import java.util.StringTokenizer;
import java.util.regex.Pattern;

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
import com.vmware.jct.common.utility.EncoderDecoder;
import com.vmware.jct.exception.JCTException;
import com.vmware.jct.service.IAfterSketchService;
import com.vmware.jct.service.IAuthenticatorService;
import com.vmware.jct.service.IBeforeSketchService;
import com.vmware.jct.service.vo.JctAfterSketchFinalPageVO;
import com.vmware.jct.service.vo.JctAfterSketchHeaderVO;
import com.vmware.jct.service.vo.JctAfterSketchPageOneVO;
import com.vmware.jct.service.vo.ReturnVO;
/**
 * 
 * <p><b>Class name:</b> AfterSketchController.java</p>
 * <p><b>Author:</b> InterraIT</p>
 * <p><b>Purpose:</b> This class acts as a controller for after sketch activities.
 * AfterSketchController extends BasicController  and has following  methods.
 * -saveAfterSketchBasic()
 * <p><b>Description:</b> This class basically intercept all the after sketch related service calls and 
 * then delegated the it to the service layer for further actions. </p>
 * <p><b>Copyrights:</b> 	All rights reserved by Interra IT and should only be used for its internal application development.</p>
 * <p><b>Created Date:</b> 27/Jan/2014</p>
 * <p><b>Revision History:</b>
 * 	<li>InterraIT - 04/Feb/2014 - Implemented comments</li>
 * </p>
 */
@Controller
public class AfterSketchController extends BasicController {
	@Autowired
	private IAfterSketchService afterSketchService;
	
	@Autowired
	private IBeforeSketchService beforeSketchService;
	
	@Autowired
	private IAuthenticatorService authenticatorService;
	
	@Autowired  
	private MessageSource messageSource;
	
	private static final Logger LOGGER = LoggerFactory.getLogger(AfterSketchController.class);
	
	/**
	 * Method saves the Mapping yourself
	 * @param body
	 * @param request
	 * @return ReturnVO
	 */
	@RequestMapping(value = ACTION_SAVE_AFTER_SKETCH_1, method = RequestMethod.POST)
	@ResponseBody()
	public ReturnVO saveAfterSketchBasic(@RequestBody String body,
			HttpServletRequest request) {
		LOGGER.info(">>>>   AfterSketchController.saveAfterSketchBasic");
		Calendar startTime = Calendar.getInstance();
		ReturnVO returnVO = new ReturnVO();
		JsonNode node = CommonUtility.getNode(body);
		String jobReferenceNo = node.get("jobReferenceNo").toString().replaceAll("\"", "");
		int rowId = Integer.parseInt(node.get("rowId").asText());
		String jSon = node.get("jsonStringFirstPage").toString();
		String createdBy = node.get("createdBy").toString().replaceAll("\"", "");
		String idChecked = node.get("idChecked").toString().replaceAll("\"", "");
		String jobTitle = node.get("jobTitle").toString().replaceAll("\"", "");
		String checkedElements = node.get("checkedElems").toString();
		int passionCount = Integer.parseInt(node.get("passionCount").toString().replaceAll("\"", ""));
		int valueCount = Integer.parseInt(node.get("valueCount").toString().replaceAll("\"", ""));
		int strengthCount = Integer.parseInt(node.get("strengthCount").toString().replaceAll("\"", ""));
		String checkedPassion = node.get("checked_passion").toString().replaceAll("\"", "");
		String checkedStrength = node.get("checked_strength").toString().replaceAll("\"", "");
		String checkedValue = node.get("checked_value").toString().replaceAll("\"", "");
		int timeSpent = Integer.parseInt(node.get("timeSpents").toString().replaceAll("\"", ""));
		String linkClicked = node.get("linkClicked").asText();
		String isPrevious = node.get("isPrevious").asText();
		/** THIS IS FOR PUBLIC VERSION **/
		int jctUserId = Integer.parseInt(node.get("jctUserId").asText());
		/********************************/
		//For breaking down
		String jSonString = node.get("jsonStringFirstPage").toString().replaceAll("\"", "");
		String jSonString1 = jSonString.replaceAll("},", "#");
		String jSonString2 = jSonString1.replaceAll("}]", "#");
		String jSonString3 = jSonString2.substring(2, jSonString2.length());
		
		
		//Create header
		JctAfterSketchHeaderVO parent = new JctAfterSketchHeaderVO();
		/** THIS IS FOR PUBLIC VERSION **/
		parent.setJctUserId(jctUserId);
		/********************************/
		parent.setJobReferenceNo(jobReferenceNo);
		
		parent.setJctPageOneJson(jSon);
		
		parent.setJctCreatedBy(createdBy);
		parent.setJctLastModifiedBy(createdBy);
		parent.setJctJobTitle(jobTitle);
		parent.setJctPageOneCheckedElements(checkedElements);
		parent.setCheckedPassion(checkedPassion);
		parent.setCheckedStrength(checkedStrength);
		parent.setCheckedValue(checkedValue);
		parent.setJctPageOneTimeSpent(timeSpent);
		parent.setJctFinalPageInitialJSonValue(node.get("divInitialVal").toString());
		//create the child list
		List<JctAfterSketchPageOneVO> childList = new ArrayList<JctAfterSketchPageOneVO>();
		StringTokenizer token = new StringTokenizer(jSonString3,"#");
		List<String> dummyList = new ArrayList<String>();
		while(token.hasMoreTokens()){
			String value = token.nextToken().toString();
			if(!value.contains("display:none") && !value.contains("display: none") && !value.contains("display : none") && !value.contains("display :none")){
				dummyList.add(value);
			}
		}
		List<Integer> idList = new ArrayList<Integer>();
		StringTokenizer idToken = new StringTokenizer(idChecked,"#");
		while(idToken.hasMoreTokens()){
			idList.add(Integer.parseInt(idToken.nextToken().toString()));
		}
		
		for(int index=0; index<dummyList.size(); index++){
			JctAfterSketchPageOneVO child = new JctAfterSketchPageOneVO();
			child.setJobReferenceNo(jobReferenceNo);
			String[] split = dummyList.get(index).toString().split("divId:");
			StringBuilder elementCode = new StringBuilder("");
			if(split[1].contains("Strength")){
				elementCode.append(CommonConstants.JOB_ATTRIBUTE_STRENGTH);
			} else if(split[1].contains("Passion")){
				elementCode.append(CommonConstants.JOB_ATTRIBUTE_PASSION);
			} else {
				elementCode.append(CommonConstants.JOB_ATTRIBUTE_VALUE);
			}
			child.setElementId(idList.get(index));
			child.setElementCode(elementCode.toString());
			child.setPosition("");
			
			//set the count
			child.setPassionCount(passionCount);
			child.setValueCount(valueCount);
			child.setStrengthCount(strengthCount);
			childList.add(child);
		}
		try {
			//Check if AS json already present for job reference number...
			returnVO.setAfterSketch2TotalJsonObj(afterSketchService.fetchASFinalPageJson(jobReferenceNo, linkClicked));
			String str = afterSketchService.getDivHeightAndWidth(jobReferenceNo, linkClicked);
			StringTokenizer sizeToken = new StringTokenizer(str, "#");
			String height = "";
			String width = "";
			try{
				while(sizeToken.hasMoreTokens()){
					height = (String)sizeToken.nextToken();
					width = (String)sizeToken.nextToken();
				}
				}	catch(Exception ex){
					LOGGER.error("---------------------------->WIDTH AND HEIGHT ISSUE AFTER DRAGGING");
			}
			
			returnVO.setDivHeight(height);
			returnVO.setDivWidth(width);
			returnVO.setAsTotalCount(afterSketchService.getAsTotalCount(jobReferenceNo, linkClicked));
			String plottingJson = afterSketchService.saveOrUpdatePageOnde(parent, childList, linkClicked);
			returnVO.setPageOneJson(plottingJson);
			//For population of new after sketch page 2
			returnVO.setBsJson(beforeSketchService.getJson(jobReferenceNo, linkClicked));
			String initialJson = afterSketchService.getFinalPageInitialJson(jobReferenceNo, linkClicked);
			/*if ((null != afterSketchService.getFinalPageInitialJson(jobReferenceNo, linkClicked)) 
					&& (!afterSketchService.getFinalPageInitialJson(jobReferenceNo, linkClicked).equals("null"))
					&& (!afterSketchService.getFinalPageInitialJson(jobReferenceNo, linkClicked).equals(""))) {*/
			if ((null != initialJson) && (!initialJson.equals("null")) && (!initialJson.equals(""))) {
				returnVO.setInitialJson(afterSketchService.getFinalPageInitialJson(jobReferenceNo, linkClicked));
			} else {
				returnVO.setInitialJson(beforeSketchService.getInitialJson(jobReferenceNo, linkClicked));
			}
			
			//authenticatorService.updateLoginInfo(jobReferenceNo, rowId, "/user/view/afterSketch1.jsp");
			if (isPrevious.equals("Next")) {
				authenticatorService.updateLoginInfo(jobReferenceNo, rowId, "/user/view/afterSketch2.jsp"); // When going next
			} else if (isPrevious.equals("Prev")) {
				authenticatorService.updateLoginInfo(jobReferenceNo, rowId, "/user/view/Questionaire.jsp"); // When going previos
			} else {
				authenticatorService.updateLoginInfo(jobReferenceNo, rowId, "/user/view/afterSketch1.jsp"); // When Logging out
			}
		} catch (JCTException e) {
			LOGGER.error("Exception occured at outer try..."+e.getLocalizedMessage());
		}
		long diff = Calendar.getInstance().getTimeInMillis() - startTime.getTimeInMillis();
		LOGGER.info("-------------- Saving Mapping Yourself Took: "+diff+" Milli seconds ------------------");
		LOGGER.info("<<<<   AfterSketchController.saveAfterSketchBasic");
		return returnVO;
	}
	/**
	 * Method stores the After Sketch Final Diagram.
	 * @param body
	 * @param request
	 * @return ReturnVO
	 */
	@RequestMapping(value = ACTION_SAVE_AFTER_SKETCH_2, method = RequestMethod.POST)
	@ResponseBody()
	public ReturnVO saveAfterSketchDiagram(@RequestBody String body,
			HttpServletRequest request) {
		LOGGER.info(">>>> AfterSketchController.saveAfterSketchDiagram");
		Calendar startTime = Calendar.getInstance();
		ReturnVO returnVO = new ReturnVO();
		JsonNode node = CommonUtility.getNode(body);
		String jSon = node.get("jsonString").toString();
		String objectList = jSon.replace("\"", "");
		String jobReferenceNo = node.get("jobReferenceString").toString().replaceAll("\"", "");
		String roleJson = node.get("roleJson").toString().replaceAll("\"", "");
		String isEdit = node.get("isEdits").asText();
		int rowId = Integer.parseInt(node.get("rowId").asText());
		String initialSave = node.get("initialSave").asText();
		String linkClicked = node.get("linkClicked").asText();
		int timeSpent = Integer.parseInt(node.get("timeSpents").toString().replaceAll("\"", ""));
		String isPrevious = node.get("isPrevious").asText();
		/** THIS IS FOR PUBLIC VERSION **/
		int jctUserId = Integer.parseInt(node.get("jctUserId").asText());
		String elementOutsideRoleJson = node.get("elementOutsideRoleJson").toString().replaceAll("\"", "");
		/********************************/
		String positionJson = node.get("positionJson").asText().replaceAll("\"", "").replaceAll("},", "#").replaceAll("}]", "#");
		Map<String, String> positionMap = new HashMap<String, String> ();
		StringTokenizer positionToken = new StringTokenizer(positionJson, "#");
		while (positionToken.hasMoreTokens()) {
			String item = (String) positionToken.nextToken();
			if (!item.contains("cloneOval")) {
				String[] itemArr = item.split(",");
				String id = (String) itemArr[0].split(":")[1];
				String top = (String) itemArr [2].split(":")[1];
				String left = (String) itemArr [1].split(":")[1];
				positionMap.put(id, top+"/"+left);
			}
		}
		
		//StringTokenizer mapToken = new StringTokenizer(positionJson, ",");
		String objectList1 = objectList.replaceAll("},", "#");
		String objectList2 = objectList1.replaceAll("}]", "#");
		String objectList3 = objectList2.substring(2, objectList2.length());
		
		
		String roleJson1 = roleJson.replaceAll("},", "```~``~~``````-_-");
		String roleJson2 = roleJson1.replaceAll("}]", "```~``~~``````-_-");
		String roleJson3 = roleJson2.substring(2, roleJson2.length());
		
		// For things which are not in role
		String noRoleJson1 = elementOutsideRoleJson.replaceAll("},", "```~``~~``````-_-");
		String noRoleJson2 = noRoleJson1.replaceAll("}]", "```~``~~``````-_-");
		String noRoleJson3 = noRoleJson2.substring(2, noRoleJson2.length());
		
		String file = node.get("snapShot").toString();
		byte[] imageInByte = EncoderDecoder.getImageBytes(file.substring(file.indexOf(",") + 1, file.length()));
		
		String divHeight = node.get("divHeight").toString().replaceAll("\"", "");
		String divWidth = node.get("divWidth").toString().replaceAll("\"", "");
		String isCompleted = node.get("isCompleteds").toString().replaceAll("\"", "");
		//Create header
		JctAfterSketchHeaderVO parent = new JctAfterSketchHeaderVO();
		/** THIS IS FOR PUBLIC VERSION **/
		parent.setJctUserId(jctUserId);
		/********************************/
		parent.setJobReferenceNo(jobReferenceNo); //No need. Already there in db and we need to update
		parent.setJctFinalPageJson(jSon);
		parent.setJctFinalPageSnapshot(imageInByte);
		parent.setJctFinalPageTimeSpent(timeSpent);
		parent.setJctFinalPageSnapshotString(file);
		//set the parent div height and width
		parent.setDivHeight(divHeight);
		parent.setDivWidth(divWidth);
		//setting the total count
		parent.setAsTotalCount(Integer.parseInt(node.get("totalCount").toString().replaceAll("\"", "")));
		
		//set the final page initial json val
		parent.setJctFinalPageInitialJSonValue(node.get("divInitialVal").toString());
		
		//create the child list
		List<JctAfterSketchFinalPageVO> childList = new ArrayList<JctAfterSketchFinalPageVO>();
		StringTokenizer token = new StringTokenizer(objectList3,"#");
		List<String> elementList = new ArrayList<String>(); // this will have elements only which are visible
		while(token.hasMoreTokens()){
			String value = token.nextToken().toString();
			if(!(value.contains("display:none")) && !(value.contains("display : none")) && !(value.contains("display: none")) && !(value.contains("cloneOval"))){
				elementList.add(value);
			}
		}
		
		String patternString = "(```~``~~``````-_-)";
		Pattern pattern = Pattern.compile(patternString);
		
		String[] roleSplit = pattern.split(roleJson3);
		List<String> roleJsonList = new ArrayList<String>(); // this will have elements only which are visible
		for(String element : roleSplit){
			if(!(element.contains("display:none")) && !(element.contains("display: none")) && !(element.contains("display : none")) && !(element.contains("display :none"))){
				roleJsonList.add(element);
			}
		}
		
		String[] noRoleSplit = pattern.split(noRoleJson3);
		List<String> noRoleJsonList = new ArrayList<String>(); // this will have elements only which are visible
		for(String element : noRoleSplit){
			if(!(element.contains("display:none")) && !(element.contains("display: none")) && !(element.contains("display : none")) && !(element.contains("display :none"))){
				noRoleJsonList.add(element);
			}
		}
		
		//Every item appearing in dummy list will be a child. Moreover, if it appears in 'n' roles (roleJson),
		//'n' child will be created.
		for(int elementIndex=0; elementIndex<elementList.size(); elementIndex++) {
			String element = elementList.get(elementIndex).toString().substring(elementList.get(elementIndex).indexOf("divId:")+6, elementList.get(elementIndex).indexOf(","));
			String[] split = elementList.get(elementIndex).toString().split(",");
			String energySplit = split[3].toString(); 
			String idSplit = split[0].toString();
			String taskDescSplit = "";
			if( elementList.get(elementIndex).toString().contains("divImg")){
				taskDescSplit = split[6].toString();
			} 
			String taskAdditionalDescSplit = "";
			if( elementList.get(elementIndex).toString().contains("divImg")){
				taskAdditionalDescSplit = split[8].toString();
			}
			int id = 0;
			int taskEnergyVal = 0;
			//in how many role this element falls
			for (int roleIndex=0; roleIndex<roleJsonList.size(); roleIndex++) {
				StringTokenizer commaToken = new StringTokenizer(roleJsonList.get(roleIndex).toString(), ",");
				while (commaToken.hasMoreTokens()) {
					String innerSplitStr = (String) commaToken.nextToken();
					if (innerSplitStr.contains("ElementId")) {
						String elementIdStr = innerSplitStr.split(":")[1].toString();
						if (elementIdStr.trim().equals(element.trim())) {
							 //create a child obj
							JctAfterSketchFinalPageVO child = new JctAfterSketchFinalPageVO();
							child.setJobRefNo(jobReferenceNo);
							if (element.contains("divStrengthClone")) {
								child.setElementCode(CommonConstants.JOB_ATTRIBUTE_STRENGTH);
								child.setElementDesc(getElementDesc(elementList.get(elementIndex).toString()));
								//get the position from the map
								String positionValue = positionMap.get(elementIdStr);
								child.setPosition(positionValue);
							} else if(element.contains("divValueClone")){
								child.setElementCode(CommonConstants.JOB_ATTRIBUTE_VALUE);
								child.setElementDesc(getElementDesc(elementList.get(elementIndex).toString()));
								String positionValue = positionMap.get(elementIdStr);
								child.setPosition(positionValue);
							} else if(element.contains("divPassionClone")){
								child.setElementCode(CommonConstants.JOB_ATTRIBUTE_PASSION);
								child.setElementDesc(getElementDesc(elementList.get(elementIndex).toString()));
								String positionValue = positionMap.get(elementIdStr);
								child.setPosition(positionValue);
							} else {
								child.setElementCode("");
								child.setElementDesc("");
							}
							
							//child.setPosition("");
							String role = "";
							StringTokenizer str = new StringTokenizer(roleJsonList.get(roleIndex).toString(), ",");
							while(str.hasMoreTokens()){
								str.nextToken();
								String roleVal = str.nextToken();
								role = roleVal.substring(roleVal.indexOf(":")+1, roleVal.length()).replaceAll(";_;", "?").replaceAll("-_-", ",");
								break;
							}
							StringTokenizer energySplitTok = new StringTokenizer(energySplit, ":");
							while(energySplitTok.hasMoreTokens()){
								energySplitTok.nextToken();
								try{
									taskEnergyVal = Integer.parseInt(energySplitTok.nextToken().toString());
								}catch(NumberFormatException s){
									LOGGER.error("NumberFormatException.."+s.getLocalizedMessage());
									taskEnergyVal = 0;
								}
								break;
							}
							StringTokenizer idTok = new StringTokenizer(idSplit, ":");
							while(idTok.hasMoreTokens()){
								idTok.nextToken();
								String idVal = idTok.nextToken().toString();
								if(idVal.contains("divImg")){
									id = Integer.parseInt(idVal.substring(idVal.indexOf("_")+1, idVal.length()));
									//get the position from the map
									String positionValue = positionMap.get(idVal);
									child.setPosition(positionValue);
								} else { 
									id = 0;
								}
								break;
							}
							//"areaValue":"xaXAS: hghghghg : bnhgjhggh"
							String[] taskDescValSplit = taskDescSplit.split(":");
							if (taskDescValSplit.length-1 > 0) {
								StringBuilder sb = new StringBuilder(""); 
								for (int index = 1; index < taskDescValSplit.length; index++) { //index = 1 as index 0 is the key and we want value
									if (taskDescValSplit.length == 1) { // only 1 value
										child.setTaskDescription(taskDescValSplit[1].replaceAll("`", ",").replaceAll("-_-", "#").replaceAll(";_;", "?"));
									} else {
										sb.append(taskDescValSplit[index]);
										if(taskDescValSplit.length-1 != index) {
											sb.append(" : ");
										}
									}
									if (sb.toString().contains("position")) {
										sb = new StringBuilder("");
									}
									child.setTaskDescription(sb.toString().replaceAll("`", ",").replaceAll("-_-", "#").replaceAll(";_;", "?"));
								}
							} else {
								child.setTaskDescription("");
							}
							
							if(taskAdditionalDescSplit == ""){
								child.setAdditionalDesc("");
							} else {
								StringTokenizer addDescTok = new StringTokenizer(taskAdditionalDescSplit, ":");
								while(addDescTok.hasMoreTokens()){
									addDescTok.nextToken();
									try{
										child.setAdditionalDesc(addDescTok.nextToken().toString().replaceAll("`", ",").replaceAll("-_-", "#").replaceAll(";_;", "?"));
									}catch(NoSuchElementException npe){
										child.setAdditionalDesc("");
										
									}
									break;
								}
							}
							child.setTaskEnergy(taskEnergyVal);
							child.setRoleDescription(role);
							child.setTaskId(id);
							childList.add(child);
						}
					}
				}
			}
			
			for (int noRoleIndex=0; noRoleIndex<noRoleJsonList.size(); noRoleIndex++) {
				StringTokenizer commaToken = new StringTokenizer(noRoleJsonList.get(noRoleIndex).toString(), ",");
				while (commaToken.hasMoreTokens()) {
					String innerSplitStr = (String) commaToken.nextToken();
					if (innerSplitStr.contains("ElementId")) {
						String elementIdStr = innerSplitStr.split(":")[1].toString();
						if (elementIdStr.trim().equals(element.trim())) {
							 //create a child obj
							JctAfterSketchFinalPageVO child = new JctAfterSketchFinalPageVO();
							child.setJobRefNo(jobReferenceNo);
							if (element.contains("divStrengthClone")) {
								child.setElementCode(CommonConstants.JOB_ATTRIBUTE_STRENGTH);
								child.setElementDesc(getElementDesc(elementList.get(elementIndex).toString()));
								//get the position from the map
								String positionValue = positionMap.get(elementIdStr);
								child.setPosition(positionValue);
							} else if(element.contains("divValueClone")){
								child.setElementCode(CommonConstants.JOB_ATTRIBUTE_VALUE);
								child.setElementDesc(getElementDesc(elementList.get(elementIndex).toString()));
								String positionValue = positionMap.get(elementIdStr);
								child.setPosition(positionValue);
							} else if(element.contains("divPassionClone")){
								child.setElementCode(CommonConstants.JOB_ATTRIBUTE_PASSION);
								child.setElementDesc(getElementDesc(elementList.get(elementIndex).toString()));
								String positionValue = positionMap.get(elementIdStr);
								child.setPosition(positionValue);
							} else {
								child.setElementCode("");
								child.setElementDesc("");
							}
							
							//child.setPosition("");
							String role = "";
							StringTokenizer str = new StringTokenizer(noRoleJsonList.get(noRoleIndex).toString(), ",");
							while(str.hasMoreTokens()){
								str.nextToken();
								String roleVal = str.nextToken();
								role = roleVal.substring(roleVal.indexOf(":")+1, roleVal.length()).replaceAll(";_;", "?").replaceAll("-_-", ",");
								break;
							}
							StringTokenizer energySplitTok = new StringTokenizer(energySplit, ":");
							while(energySplitTok.hasMoreTokens()){
								energySplitTok.nextToken();
								try{
									taskEnergyVal = Integer.parseInt(energySplitTok.nextToken().toString());
								}catch(NumberFormatException s){
									LOGGER.error("NumberFormatException.."+s.getLocalizedMessage());
									taskEnergyVal = 0;
								}
								break;
							}
							StringTokenizer idTok = new StringTokenizer(idSplit, ":");
							while(idTok.hasMoreTokens()){
								idTok.nextToken();
								String idVal = idTok.nextToken().toString();
								if(idVal.contains("divImg")){
									id = Integer.parseInt(idVal.substring(idVal.indexOf("_")+1, idVal.length()));
									//get the position from the map
									String positionValue = positionMap.get(idVal);
									child.setPosition(positionValue);
								} else { 
									id = 0;
								}
								break;
							}
							//"areaValue":"xaXAS: hghghghg : bnhgjhggh"
							String[] taskDescValSplit = taskDescSplit.split(":");
							if (taskDescValSplit.length-1 > 0) {
								StringBuilder sb = new StringBuilder(""); 
								for (int index = 1; index < taskDescValSplit.length; index++) { //index = 1 as index 0 is the key and we want value
									if (taskDescValSplit.length == 1) { // only 1 value
										child.setTaskDescription(taskDescValSplit[1].replaceAll("`", ",").replaceAll("-_-", "#").replaceAll(";_;", "?"));
									} else {
										sb.append(taskDescValSplit[index]);
										if(taskDescValSplit.length-1 != index) {
											sb.append(" : ");
										}
									}
									if (sb.toString().contains("position")) {
										sb = new StringBuilder("");
									}
									child.setTaskDescription(sb.toString().replaceAll("`", ",").replaceAll("-_-", "#").replaceAll(";_;", "?"));
								}
							} else {
								child.setTaskDescription("");
							}
							
							if(taskAdditionalDescSplit == ""){
								child.setAdditionalDesc("");
							} else {
								StringTokenizer addDescTok = new StringTokenizer(taskAdditionalDescSplit, ":");
								while(addDescTok.hasMoreTokens()){
									addDescTok.nextToken();
									try{
										child.setAdditionalDesc(addDescTok.nextToken().toString().replaceAll("`", ",").replaceAll("-_-", "#").replaceAll(";_;", "?"));
									}catch(NoSuchElementException npe){
										child.setAdditionalDesc("");
										
									}
									break;
								}
							}
							child.setTaskEnergy(taskEnergyVal);
							child.setRoleDescription(role);
							child.setTaskId(id);
							/*if (!childList.contains(child)) {
								childList.add(child);
							}*/
							if (childList.contains(child) == false) {
								childList.add(child);
							}
						}
					}
				}
			}
		}
		try {
			afterSketchService.saveAfterSketchFinalPage(parent, childList, isCompleted, isEdit, linkClicked);
			//Update jct_next_page in login info with AS
			authenticatorService.updateNextPage(jobReferenceNo, "AS");
			if (!isEdit.equals("Y")) {
				//authenticatorService.updateLoginInfo(jobReferenceNo, rowId, "/user/view/afterSketch2.jsp");
				if (initialSave.equals("Y")) {
					authenticatorService.updateLoginInfo(jobReferenceNo, rowId, "/user/view/afterSketch2.jsp");
				} else {
					if (isPrevious.equals("Y")) {
						authenticatorService.updateLoginInfo(jobReferenceNo, rowId, "/user/view/afterSketch1.jsp");
					} else {
						authenticatorService.updateLoginInfo(jobReferenceNo, rowId, "/user/view/actionPlan.jsp");
					}
				}
			}
		} catch (JCTException e) {
			LOGGER.error("ERROR IN SAVING AFTER SKETCH DIADRAM: "+e.getLocalizedMessage());
		}
		
		/********************** NOTE **********************/
		//Directly returning the snapshot string as it need not
		//be fetched from db. But if login directly to the action plan page
		//we need to fetch it from db.
		returnVO.setSnapShot(node.get("snapShot").toString());
		returnVO.setIsCompleted(1); //Already completed... for other pages, fetch it from DB
		long diff = Calendar.getInstance().getTimeInMillis() - startTime.getTimeInMillis();
		LOGGER.info("-------------- Saving Putting It Together Took: "+diff+" Milli seconds ------------------");
		LOGGER.info("<<<<  AfterSketchController.saveAfterSketchDiagram");
		return returnVO;
	}
	
	private String getElementDesc(String element){
		LOGGER.info(">>>> AfterSketchController.getElementDesc");
		String value = "";
		StringTokenizer str = new StringTokenizer(element, ",");
		while(str.hasMoreTokens()){
			str.nextToken();
			String[] val = str.nextToken().toString().split(":");
			value = val[1];
			break;
		}
		LOGGER.info("<<<<  AfterSketchController.getElementDesc");
		return value;
	}
}