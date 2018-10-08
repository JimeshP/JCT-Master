package com.vmware.jct.service.vo;

import java.util.List;

import com.vmware.jct.dao.dto.OccupationListDTO;
import com.vmware.jct.dao.dto.QuestionearDTO;
/**
 * 
 * <p><b>Class name:</b> ReturnVO.java</p>
 * <p><b>Author:</b> InterraIT</p>
 * <p><b>Purpose:</b> This class acts as a data transfer object.
 * <p><b>Description:</b> This class acts as a data transfer object.. </p>
 * <p><b>Copyrights:</b> 	All rights reserved by Interra IT and should only be used for its internal application development.</p>
 * <p><b>Created Date:</b> 20/Feb/2014</p>
 * <p><b>Revision History:</b>
 * 	<li>6/May/2014: Changed the saving of action plan to dynamic.</li>
 * </p>
 */
public class ReturnVO {
	private String statusCode;
	private String statusMsg;
	private String statusDesc;
	private String pathFinder;
	private String userName;
	private String email;
	private List list;
	private String bsJson;
	private String initialJson;
	private List<QuestionearDTO> qtnList;
	private String existing;
	private String snapShot;
	private byte[] test;
	private int isCompleted;
	private String afterSketch2TotalJsonObj;
	private String divHeight;
	private String divWidth;
	private int asTotalCount;
	private String url;
	private List<List> actionPlanList;
	private List<List> questionnaireList;
	private List<OccupationListDTO> dtoList;
	
	//After Sketch Edit specific
	private String pageOneJson;
	private String totalJsonAddTask;
	private String divInitialValue;
	private String demographicPopupMsg;
	private int profileId;
	
	private String fName;
	private String lName;
	
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
	public List<List> getQuestionnaireList() {
		return questionnaireList;
	}
	public void setQuestionnaireList(List<List> questionnaireList) {
		this.questionnaireList = questionnaireList;
	}
	public List<List> getActionPlanList() {
		return actionPlanList;
	}
	public void setActionPlanList(List<List> actionPlanList) {
		this.actionPlanList = actionPlanList;
	}
	/**
	 * @return the url
	 */
	public String getUrl() {
		return url;
	}
	/**
	 * @param url the url to set
	 */
	public void setUrl(String url) {
		this.url = url;
	}
	public int getAsTotalCount() {
		return asTotalCount;
	}
	public void setAsTotalCount(int asTotalCount) {
		this.asTotalCount = asTotalCount;
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
	public String getAfterSketch2TotalJsonObj() {
		return afterSketch2TotalJsonObj;
	}
	public void setAfterSketch2TotalJsonObj(String afterSketch2TotalJsonObj) {
		this.afterSketch2TotalJsonObj = afterSketch2TotalJsonObj;
	}
	public int getIsCompleted() {
		return isCompleted;
	}
	public void setIsCompleted(int isCompleted) {
		this.isCompleted = isCompleted;
	}
	public byte[] getTest() {
		return test;
	}
	public void setTest(byte[] test) {
		this.test = test;
	}
	public String getSnapShot() {
		return snapShot;
	}
	public void setSnapShot(String snapShot) {
		this.snapShot = snapShot;
	}
	public String getExisting() {
		return existing;
	}
	public void setExisting(String existing) {
		this.existing = existing;
	}
	public List<QuestionearDTO> getQtnList() {
		return qtnList;
	}
	public void setQtnList(List<QuestionearDTO> qtnList) {
		this.qtnList = qtnList;
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
	public String getEmail() {
		return email;
	}
	public void setEmail(String email) {
		this.email = email;
	}
	public String getUserName() {
		return userName;
	}
	public void setUserName(String userName) {
		this.userName = userName;
	}
	public String getStatusCode() {
		return statusCode;
	}
	public void setStatusCode(String statusCode) {
		this.statusCode = statusCode;
	}
	public String getStatusMsg() {
		return statusMsg;
	}
	public void setStatusMsg(String statusMsg) {
		this.statusMsg = statusMsg;
	}
	public String getStatusDesc() {
		return statusDesc;
	}
	public void setStatusDesc(String statusDesc) {
		this.statusDesc = statusDesc;
	}
	public String getPathFinder() {
		return pathFinder;
	}
	public void setPathFinder(String pathFinder) {
		this.pathFinder = pathFinder;
	}
	public String getDemographicPopupMsg() {
		return demographicPopupMsg;
	}
	public void setDemographicPopupMsg(String demographicPopupMsg) {
		this.demographicPopupMsg = demographicPopupMsg;
	}
	public List<OccupationListDTO> getDtoList() {
		return dtoList;
	}
	public void setDtoList(List<OccupationListDTO> dtoList) {
		this.dtoList = dtoList;
	}
	public int getProfileId() {
		return profileId;
	}
	public void setProfileId(int profileId) {
		this.profileId = profileId;
	}
	public String getfName() {
		return fName;
	}
	public void setfName(String fName) {
		this.fName = fName;
	}
	public String getlName() {
		return lName;
	}
	public void setlName(String lName) {
		this.lName = lName;
	}	
}
