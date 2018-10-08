package com.vmware.jct.service.vo;

import java.util.Date;
import java.util.List;
import java.util.Map;
/**
 * 
 * <p><b>Class name:</b> UserVO.java</p>
 * <p><b>Author:</b> InterraIT</p>
 * <p><b>Purpose:</b> This class acts as a value object. Transfer data from presentation layer to controller and vice versa
 * <p><b>Description:</b>This class acts as a value object. Transfer data from presentation layer to controller and vice versa </p>
 * <p><b>Copyrights:</b> 	All rights reserved by Interra IT and should only be used for its internal application development.</p>
 * <p><b>Created Date:</b> 20/Feb/2014</p>
 * <p><b>Revision History:</b>
 * 	<li>6/May/2014: Changed the saving of action plan to dynamic.</li>
 * </p>
 */
public class UserVO {
	private String firstName;
	private String lastName;
	private String userName;
	private String password;
	private String email;
	private String gender;
	private Date dateOfBirth;
	private byte[] profilePicture;
	private int profileId;
	private String statusMsg;
	private String statusCode;
	private String statusDesc;
	private String url;
	private String jctBaseString;
	private String location;
	private String jobTitle;
	private String afterSketchDivIniVal;
	private String afterSketchCheckedEle;
	private String afterSkPageOneTotalJson;
	private Map strengthMap;
	private Map passionMap;
	private Map valueMap;
	private int valueCount;
	private int passionCount;
	private int strengthCount;
	private Map region;
	private Map functionGroup;
	private Map jobLevel;
	
	private String supervise;
	private String tenure;
	private String functionGrpSelected;
	private String jobLevelSelected;
	
	private String afterSketch2TotalJsonObj;
	private String totalJsonAddTask;
	private String divInitialValue;
	private String pageOneJson;
	
	private String divHeight;
	private String divWidth;
	
	private String checkedPassion;
	private String checkedStrength;
	private String checkedValue;
	
	private String existing;
	private String snapShot;
	
	private String menu;
	private int roleId;
	private String unauthMsg;
	private String interfaceTitleStr;
	private String customerId;;
	
	private int activeYn;
	private int jctUserId;
	// survey questions list
	private String surveyTextList;
	private String surveyRadioList;
	private String surveyDropdownList;
	private String surveyCheckboxList;
	
	private Date expiryDate;
	private String userGroup;
	private String userProfile;
	private int statusCodeInt;
	
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
	/**
	 * @return the strengthMap
	 */
	public Map getStrengthMap() {
		return strengthMap;
	}
	/**
	 * @param strengthMap the strengthMap to set
	 */
	public void setStrengthMap(Map strengthMap) {
		this.strengthMap = strengthMap;
	}
	/**
	 * @return the passionMap
	 */
	public Map getPassionMap() {
		return passionMap;
	}
	/**
	 * @param passionMap the passionMap to set
	 */
	public void setPassionMap(Map passionMap) {
		this.passionMap = passionMap;
	}
	/**
	 * @return the valueMap
	 */
	public Map getValueMap() {
		return valueMap;
	}
	/**
	 * @param valueMap the valueMap to set
	 */
	public void setValueMap(Map valueMap) {
		this.valueMap = valueMap;
	}
	
	
	
	public String getAfterSkPageOneTotalJson() {
		return afterSkPageOneTotalJson;
	}
	public void setAfterSkPageOneTotalJson(String afterSkPageOneTotalJson) {
		this.afterSkPageOneTotalJson = afterSkPageOneTotalJson;
	}
	public String getAfterSketchCheckedEle() {
		return afterSketchCheckedEle;
	}
	public void setAfterSketchCheckedEle(String afterSketchCheckedEle) {
		this.afterSketchCheckedEle = afterSketchCheckedEle;
	}
	public String getAfterSketchDivIniVal() {
		return afterSketchDivIniVal;
	}
	public void setAfterSketchDivIniVal(String afterSketchDivIniVal) {
		this.afterSketchDivIniVal = afterSketchDivIniVal;
	}
	//Only for display
	private String jctPreString;
	
	//Only for reset mechanism
	private String initialPassword;
	
	//For page navigation on ui side
	private int pageSequence;
	
	//
	private List list;
	
	/************ BEFORE SKETCH SPECIFIC ***************/
	//For new / existing job reference no
	private String jobRefNo;
	//For total time
	private String totalTime;
	//For user beforeSketchJSon
	private String bsJson;
	
	public String getInitialJson() {
		return initialJson;
	}
	public void setInitialJson(String initialJson) {
		this.initialJson = initialJson;
	}
	private String initialJson;
	
	//Page info id
	private String identifier;
	
	public String getFirstName() {
		return firstName;
	}
	public void setFirstName(String firstName) {
		this.firstName = firstName;
	}
	public String getLastName() {
		return lastName;
	}
	public void setLastName(String lastName) {
		this.lastName = lastName;
	}
	public String getGender() {
		return gender;
	}
	public void setGender(String gender) {
		this.gender = gender;
	}
	public Date getDateOfBirth() {
		return dateOfBirth;
	}
	public void setDateOfBirth(Date dateOfBirth) {
		this.dateOfBirth = dateOfBirth;
	}
	public byte[] getProfilePicture() {
		return profilePicture;
	}
	public void setProfilePicture(byte[] profilePicture) {
		this.profilePicture = profilePicture;
	}
	
	public String getUrl() {
		return url;
	}
	public void setUrl(String url) {
		this.url = url;
	}
	public String getPassword() {
		return password;
	}
	public void setPassword(String password) {
		this.password = password;
	}
	public String getStatusMsg() {
		return statusMsg;
	}
	public void setStatusMsg(String statusMsg) {
		this.statusMsg = statusMsg;
	}
	public String getUserName() {
		return userName;
	}
	public void setUserName(String userName) {
		this.userName = userName;
	}
	public String getEmail() {
		return email;
	}
	public void setEmail(String email) {
		this.email = email;
	}
	public String getJctBaseString() {
		return jctBaseString;
	}
	public void setJctBaseString(String jctBaseString) {
		this.jctBaseString = jctBaseString;
	}
	public String getJctPreString() {
		return jctPreString;
	}
	public void setJctPreString(String jctPreString) {
		this.jctPreString = jctPreString;
	}
	
	public String getInitialPassword() {
		return initialPassword;
	}
	public void setInitialPassword(String initialPassword) {
		this.initialPassword = initialPassword;
	}
	public String getStatusCode() {
		return statusCode;
	}
	public void setStatusCode(String statusCode) {
		this.statusCode = statusCode;
	}
	public String getStatusDesc() {
		return statusDesc;
	}
	public void setStatusDesc(String statusDesc) {
		this.statusDesc = statusDesc;
	}
	
	public String getJobRefNo() {
		return jobRefNo;
	}
	public void setJobRefNo(String jobRefNo) {
		this.jobRefNo = jobRefNo;
	}
	
	public String getTotalTime() {
		return totalTime;
	}
	public void setTotalTime(String totalTime) {
		this.totalTime = totalTime;
	}
	
	public String getBsJson() {
		return bsJson;
	}
	public void setBsJson(String bsJson) {
		this.bsJson = bsJson;
	}
	
	public String getLocation() {
		return location;
	}
	public void setLocation(String location) {
		this.location = location;
	}
	public String getJobTitle() {
		return jobTitle;
	}
	public void setJobTitle(String jobTitle) {
		this.jobTitle = jobTitle;
	}
	
	public int getPageSequence() {
		return pageSequence;
	}
	public void setPageSequence(int pageSequence) {
		this.pageSequence = pageSequence;
	}
	
	public String getIdentifier() {
		return identifier;
	}
	public void setIdentifier(String identifier) {
		this.identifier = identifier;
	}
	
	public List getList() {
		return list;
	}
	public void setList(List list) {
		this.list = list;
	}
	/**
	 * @return the region
	 */
	public Map getRegion() {
		return region;
	}
	/**
	 * @param region the region to set
	 */
	public void setRegion(Map region) {
		this.region = region;
	}
	/**
	 * @return the functionGroup
	 */
	public Map getFunctionGroup() {
		return functionGroup;
	}
	/**
	 * @param functionGroup the functionGroup to set
	 */
	public void setFunctionGroup(Map functionGroup) {
		this.functionGroup = functionGroup;
	}
	/**
	 * @return the jobLevel
	 */
	public Map getJobLevel() {
		return jobLevel;
	}
	/**
	 * @param jobLevel the jobLevel to set
	 */
	public void setJobLevel(Map jobLevel) {
		this.jobLevel = jobLevel;
	}
	
	public String getSupervise() {
		return supervise;
	}
	public void setSupervise(String supervise) {
		this.supervise = supervise;
	}
	public String getTenure() {
		return tenure;
	}
	public void setTenure(String tenure) {
		this.tenure = tenure;
	}
	public String getFunctionGrpSelected() {
		return functionGrpSelected;
	}
	public void setFunctionGrpSelected(String functionGrpSelected) {
		this.functionGrpSelected = functionGrpSelected;
	}
	public String getJobLevelSelected() {
		return jobLevelSelected;
	}
	public void setJobLevelSelected(String jobLevelSelected) {
		this.jobLevelSelected = jobLevelSelected;
	}
	
	public String getAfterSketch2TotalJsonObj() {
		return afterSketch2TotalJsonObj;
	}
	public void setAfterSketch2TotalJsonObj(String afterSketch2TotalJsonObj) {
		this.afterSketch2TotalJsonObj = afterSketch2TotalJsonObj;
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
	
	public String getPageOneJson() {
		return pageOneJson;
	}
	public void setPageOneJson(String pageOneJson) {
		this.pageOneJson = pageOneJson;
	}
	public String getMenu() {
		return menu;
	}
	public void setMenu(String menu) {
		this.menu = menu;
	}
	public int getRoleId() {
		return roleId;
	}
	public void setRoleId(int roleId) {
		this.roleId = roleId;
	}
	public String getUnauthMsg() {
		return unauthMsg;
	}
	public void setUnauthMsg(String unauthMsg) {
		this.unauthMsg = unauthMsg;
	}
	public String getInterfaceTitleStr() {
		return interfaceTitleStr;
	}
	public void setInterfaceTitleStr(String interfaceTitleStr) {
		this.interfaceTitleStr = interfaceTitleStr;
	}
	public String getCustomerId() {
		return customerId;
	}
	public void setCustomerId(String customerId) {
		this.customerId = customerId;
	}
	public int getActiveYn() {
		return activeYn;
	}
	public void setActiveYn(int activeYn) {
		this.activeYn = activeYn;
	}
	public String getSurveyTextList() {
		return surveyTextList;
	}
	public void setSurveyTextList(String surveyTextList) {
		this.surveyTextList = surveyTextList;
	}
	public String getSurveyRadioList() {
		return surveyRadioList;
	}
	public void setSurveyRadioList(String surveyRadioList) {
		this.surveyRadioList = surveyRadioList;
	}
	public String getSurveyDropdownList() {
		return surveyDropdownList;
	}
	public void setSurveyDropdownList(String surveyDropdownList) {
		this.surveyDropdownList = surveyDropdownList;
	}
	public String getSurveyCheckboxList() {
		return surveyCheckboxList;
	}
	public void setSurveyCheckboxList(String surveyCheckboxList) {
		this.surveyCheckboxList = surveyCheckboxList;
	}
	public int getJctUserId() {
		return jctUserId;
	}
	public void setJctUserId(int jctUserId) {
		this.jctUserId = jctUserId;
	}
	public Date getExpiryDate() {
		return expiryDate;
	}
	public void setExpiryDate(Date expiryDate) {
		this.expiryDate = expiryDate;
	}
	public String getUserGroup() {
		return userGroup;
	}
	public void setUserGroup(String userGroup) {
		this.userGroup = userGroup;
	}
	public String getUserProfile() {
		return userProfile;
	}
	public void setUserProfile(String userProfile) {
		this.userProfile = userProfile;
	}
	public int getStatusCodeInt() {
		return statusCodeInt;
	}
	public void setStatusCodeInt(int statusCodeInt) {
		this.statusCodeInt = statusCodeInt;
	}
	public int getProfileId() {
		return profileId;
	}
	public void setProfileId(int profileId) {
		this.profileId = profileId;
	}	
}
