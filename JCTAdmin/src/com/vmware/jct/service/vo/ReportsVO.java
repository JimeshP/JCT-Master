package com.vmware.jct.service.vo;

import java.util.List;
import java.util.Map;
/**
 * 
 * <p><b>Class name:</b> ReportsVO.java</p>
 * <p><b>Author:</b> InterraIT</p>
 * <p><b>Purpose:</b> This class acts as a value object. Transfer data from presentation layer to controller and vice versa
 * <p><b>Description:</b>This class acts as a value object. Transfer data from presentation layer to controller and vice versa </p>
 * <p><b>Copyrights:</b> 	All rights reserved by Interra IT and should only be used for its internal application development.</p>
 * <p><b>Created Date:</b> 20/Feb/2014</p>
 * <p><b>Revision History:</b>
 * 	<li>6/May/2014: Changed the saving of action plan to dynamic.</li>
 * </p>
 */
public class ReportsVO {
	private Map<String, String> functionGroup;
	private Map<String, String> jobLevel;
	private Map<Integer, String> userGroupMap;
	private String reportList;
	private int totalCount;
	private int previousDisabled;
	private int nextDisabled;
	private String headerDisplayText;
	private String tableHeaderString;
	private String strengthString;
	private String passionString;
	private String valueString;
	
	//General status
	private int statusCode;
	private String statusDesc;
	
	private Map<Integer, String> userProfile;
	private String reportInstructionList;
	private String reportRefQtnList;
	private String reportActionPlanList;
	private String reportValueList;
	private String reportStrengthList;
	private String reportPassioList;
	private String reportProfileDescList;
	private Map<Integer, String> userNameMap;
	
	private List<AfterSketchReportVO> list;
	private List<BeforeSketchReportVO> bsList;
	private List<BeforeSketchToAfterSketchReportVO> bsToAsList;
	
	/**
	 * 
	 * @return
	 */
	public Map<String, String> getFunctionGroup() {
		return functionGroup;
	}
	/**
	 * 
	 * @param functionGroup
	 */
	public void setFunctionGroup(Map<String, String> functionGroup) {
		this.functionGroup = functionGroup;
	}
	/**
	 * 
	 * @return
	 */
	public Map<String, String> getJobLevel() {
		return jobLevel;
	}
	/**
	 * 
	 * @param jobLevel
	 */
	public void setJobLevel(Map<String, String> jobLevel) {
		this.jobLevel = jobLevel;
	}
	/**
	 * 
	 * @return
	 */
	public int getStatusCode() {
		return statusCode;
	}
	/**
	 * 
	 * @param statusCode
	 */
	public void setStatusCode(int statusCode) {
		this.statusCode = statusCode;
	}
	/**
	 * 
	 * @return
	 */
	public String getStatusDesc() {
		return statusDesc;
	}
	/**
	 * 
	 * @param statusDesc
	 */
	public void setStatusDesc(String statusDesc) {
		this.statusDesc = statusDesc;
	}
	/**
	 * 
	 * @return
	 */
	public String getReportList() {
		return reportList;
	}
	/**
	 * 
	 * @param reportList
	 */
	public void setReportList(String reportList) {
		this.reportList = reportList;
	}
	/**
	 * 
	 * @return
	 */
	public int getTotalCount() {
		return totalCount;
	}
	/**
	 * 
	 * @param totalCount
	 */
	public void setTotalCount(int totalCount) {
		this.totalCount = totalCount;
	}
	/**
	 * 
	 * @return
	 */
	public int getPreviousDisabled() {
		return previousDisabled;
	}
	/**
	 * 
	 * @param previousDisabled
	 */
	public void setPreviousDisabled(int previousDisabled) {
		this.previousDisabled = previousDisabled;
	}
	/**
	 * 
	 * @return
	 */
	public int getNextDisabled() {
		return nextDisabled;
	}
	/**
	 * 
	 * @param nextDisabled
	 */
	public void setNextDisabled(int nextDisabled) {
		this.nextDisabled = nextDisabled;
	}
	/**
	 * 
	 * @return
	 */
	public String getHeaderDisplayText() {
		return headerDisplayText;
	}
	/**
	 * 
	 * @param headerDisplayText
	 */
	public void setHeaderDisplayText(String headerDisplayText) {
		this.headerDisplayText = headerDisplayText;
	}
	/**
	 * 
	 * @return
	 */
	public String getTableHeaderString() {
		return tableHeaderString;
	}
	/**
	 * 
	 * @param tableHeaderString
	 */
	public void setTableHeaderString(String tableHeaderString) {
		this.tableHeaderString = tableHeaderString;
	}
	/**
	 * @return the userGroupMap
	 */
	public Map<Integer, String> getUserGroupMap() {
		return userGroupMap;
	}
	/**
	 * @param userGroupMap the userGroupMap to set
	 */
	public void setUserGroupMap(Map<Integer, String> userGroupMap) {
		this.userGroupMap = userGroupMap;
	}
	/**
	 * @return the strengthString
	 */
	public String getStrengthString() {
		return strengthString;
	}
	/**
	 * @param strengthString the strengthString to set
	 */
	public void setStrengthString(String strengthString) {
		this.strengthString = strengthString;
	}
	/**
	 * @return the passionString
	 */
	public String getPassionString() {
		return passionString;
	}
	/**
	 * @param passionString the passionString to set
	 */
	public void setPassionString(String passionString) {
		this.passionString = passionString;
	}
	/**
	 * @return the valueString
	 */
	public String getValueString() {
		return valueString;
	}
	/**
	 * @param valueString the valueString to set
	 */
	public void setValueString(String valueString) {
		this.valueString = valueString;
	}
	/**
	 * 
	 * @return
	 */
	public Map<Integer, String> getUserProfile() {
		return userProfile;
	}
	/**
	 * 
	 * @param userProfile
	 */
	public void setUserProfile(Map<Integer, String> userProfile) {
		this.userProfile = userProfile;
	}
	/**
	 * 
	 * @return
	 */
	public String getReportInstructionList() {
		return reportInstructionList;
	}
	/**
	 * 
	 * @param reportInstructionList
	 */
	public void setReportInstructionList(String reportInstructionList) {
		this.reportInstructionList = reportInstructionList;
	}
	/**
	 * 
	 * @return
	 */
	public String getReportRefQtnList() {
		return reportRefQtnList;
	}
	/**
	 * 
	 * @param reportRefQtnList
	 */
	public void setReportRefQtnList(String reportRefQtnList) {
		this.reportRefQtnList = reportRefQtnList;
	}
	/**
	 * 
	 * @return
	 */
	public String getReportActionPlanList() {
		return reportActionPlanList;
	}
	/**
	 * 
	 * @param reportActionPlanList
	 */
	public void setReportActionPlanList(String reportActionPlanList) {
		this.reportActionPlanList = reportActionPlanList;
	}
	/**
	 * 
	 * @return
	 */
	public String getReportValueList() {
		return reportValueList;
	}
	/**
	 * 
	 * @param reportValueList
	 */
	public void setReportValueList(String reportValueList) {
		this.reportValueList = reportValueList;
	}
	/**
	 * 
	 * @return
	 */
	public String getReportStrengthList() {
		return reportStrengthList;
	}
	/**
	 * 
	 * @param reportStrengthList
	 */
	public void setReportStrengthList(String reportStrengthList) {
		this.reportStrengthList = reportStrengthList;
	}
	/**
	 * 
	 * @return
	 */
	public String getReportPassioList() {
		return reportPassioList;
	}
	/**
	 * 
	 * @param reportPassioList
	 */
	public void setReportPassioList(String reportPassioList) {
		this.reportPassioList = reportPassioList;
	}
	/**
	 * 
	 * @return
	 */
	public String getReportProfileDescList() {
		return reportProfileDescList;
	}
	/**
	 * 
	 * @param reportProfileDescList
	 */
	public void setReportProfileDescList(String reportProfileDescList) {
		this.reportProfileDescList = reportProfileDescList;
	}
	/**
	 * 
	 * @return
	 */
	public Map<Integer, String> getUserNameMap() {
		return userNameMap;
	}
	/**
	 * 
	 * @param userNameMap
	 */
	public void setUserNameMap(Map<Integer, String> userNameMap) {
		this.userNameMap = userNameMap;
	}
	public List<AfterSketchReportVO> getList() {
		return list;
	}
	public void setList(List<AfterSketchReportVO> list) {
		this.list = list;
	}
	public List<BeforeSketchReportVO> getBsList() {
		return bsList;
	}
	public void setBsList(List<BeforeSketchReportVO> bsList) {
		this.bsList = bsList;
	}
	public List<BeforeSketchToAfterSketchReportVO> getBsToAsList() {
		return bsToAsList;
	}
	public void setBsToAsList(List<BeforeSketchToAfterSketchReportVO> bsToAsList) {
		this.bsToAsList = bsToAsList;
	}
}