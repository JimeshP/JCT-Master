package com.vmware.jct.service.vo;

import java.util.List;
import java.util.Map;
/**
 * 
 * <p><b>Class name:</b> NavigationVO.java</p>
 * <p><b>Author:</b> InterraIT</p>
 * <p><b>Purpose:</b> This class acts as a data transfer object.
 * <p><b>Description:</b> This class acts as a data transfer object.. </p>
 * <p><b>Copyrights:</b> 	All rights reserved by Interra IT and should only be used for its internal application development.</p>
 * <p><b>Created Date:</b> 20/Feb/2014</p>
 * <p><b>Revision History:</b>
 * 	<li>6/May/2014: Changed the saving of action plan to dynamic.</li>
 * </p>
 */
public class NavigationVO {
	private String bsJson;
	private String initialJson;
	private List list;
	private String jctBaseString;
	private Map strengthMap;
	private Map passionMap;
	private Map valueMap;
	private String afterSketchCheckedEle;
	private String afterSkPageOneTotalJson;
	private int valueCount;
	private int passionCount;
	private int strengthCount;
	private String pageOneJson;
	private String totalJsonAddTask;
	private String divInitialValue;
	private String afterSketch2TotalJsonObj;
	private String divHeight;
	private String divWidth;
	private String checkedPassion;
	private String checkedStrength;
	private String checkedValue;
	private String existing;
	private String snapShot;
	private int isCompleted;
	private int asTotalCount;
	private int statusCode;
	private String instructionDesc;
	private List<List> questionnaireList;
	private String videoUrl;
	public List<List> getQuestionnaireList() {
		return questionnaireList;
	}
	public void setQuestionnaireList(List<List> questionnaireList) {
		this.questionnaireList = questionnaireList;
	}
	
	public int getAsTotalCount() {
		return asTotalCount;
	}
	public void setAsTotalCount(int asTotalCount) {
		this.asTotalCount = asTotalCount;
	}
	
	public int getIsCompleted() {
		return isCompleted;
	}
	public void setIsCompleted(int isCompleted) {
		this.isCompleted = isCompleted;
	}
	public String getExisting() {
		return existing;
	}
	public void setExisting(String existing) {
		this.existing = existing;
	}
	public String getSnapShot() {
		return snapShot;
	}
	public void setSnapShot(String snapShot) {
		this.snapShot = snapShot;
	}
	public String getCheckedPassion() {
		return checkedPassion;
	}
	public void setCheckedPassion(String checkedPassion) {
		this.checkedPassion = checkedPassion;
	}
	public String getCheckedStrength() {
		return checkedStrength;
	}
	public void setCheckedStrength(String checkedStrength) {
		this.checkedStrength = checkedStrength;
	}
	public String getCheckedValue() {
		return checkedValue;
	}
	public void setCheckedValue(String checkedValue) {
		this.checkedValue = checkedValue;
	}
	public String getDivHeight() {
		return divHeight;
	}
	public void setDivHeight(String divHeight) {
		this.divHeight = divHeight;
	}
	public String getDivWidth() {
		return divWidth;
	}
	public void setDivWidth(String divWidth) {
		this.divWidth = divWidth;
	}
	public String getPageOneJson() {
		return pageOneJson;
	}
	public void setPageOneJson(String pageOneJson) {
		this.pageOneJson = pageOneJson;
	}
	public String getTotalJsonAddTask() {
		return totalJsonAddTask;
	}
	public void setTotalJsonAddTask(String totalJsonAddTask) {
		this.totalJsonAddTask = totalJsonAddTask;
	}
	public String getDivInitialValue() {
		return divInitialValue;
	}
	public void setDivInitialValue(String divInitialValue) {
		this.divInitialValue = divInitialValue;
	}
	public String getAfterSketch2TotalJsonObj() {
		return afterSketch2TotalJsonObj;
	}
	public void setAfterSketch2TotalJsonObj(String afterSketch2TotalJsonObj) {
		this.afterSketch2TotalJsonObj = afterSketch2TotalJsonObj;
	}
	public int getValueCount() {
		return valueCount;
	}
	public void setValueCount(int valueCount) {
		this.valueCount = valueCount;
	}
	public int getPassionCount() {
		return passionCount;
	}
	public void setPassionCount(int passionCount) {
		this.passionCount = passionCount;
	}
	public int getStrengthCount() {
		return strengthCount;
	}
	public void setStrengthCount(int strengthCount) {
		this.strengthCount = strengthCount;
	}
	public String getAfterSketchCheckedEle() {
		return afterSketchCheckedEle;
	}
	public void setAfterSketchCheckedEle(String afterSketchCheckedEle) {
		this.afterSketchCheckedEle = afterSketchCheckedEle;
	}
	public String getAfterSkPageOneTotalJson() {
		return afterSkPageOneTotalJson;
	}
	public void setAfterSkPageOneTotalJson(String afterSkPageOneTotalJson) {
		this.afterSkPageOneTotalJson = afterSkPageOneTotalJson;
	}
	public Map getStrengthMap() {
		return strengthMap;
	}
	public void setStrengthMap(Map strengthMap) {
		this.strengthMap = strengthMap;
	}
	public Map getPassionMap() {
		return passionMap;
	}
	public void setPassionMap(Map passionMap) {
		this.passionMap = passionMap;
	}
	public Map getValueMap() {
		return valueMap;
	}
	public void setValueMap(Map valueMap) {
		this.valueMap = valueMap;
	}
	public String getBsJson() {
		return bsJson;
	}
	public void setBsJson(String bsJson) {
		this.bsJson = bsJson;
	}
	public String getInitialJson() {
		return initialJson;
	}
	public void setInitialJson(String initialJson) {
		this.initialJson = initialJson;
	}
	public List getList() {
		return list;
	}
	public void setList(List list) {
		this.list = list;
	}
	public String getJctBaseString() {
		return jctBaseString;
	}
	public void setJctBaseString(String jctBaseString) {
		this.jctBaseString = jctBaseString;
	}
	public int getStatusCode() {
		return statusCode;
	}
	public void setStatusCode(int statusCode) {
		this.statusCode = statusCode;
	}
	public String getInstructionDesc() {
		return instructionDesc;
	}
	public void setInstructionDesc(String instructionDesc) {
		this.instructionDesc = instructionDesc;
	}
	public String getVideoUrl() {
		return videoUrl;
	}
	public void setVideoUrl(String videoUrl) {
		this.videoUrl = videoUrl;
	}
	
}
