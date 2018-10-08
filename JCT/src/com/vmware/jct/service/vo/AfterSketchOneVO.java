/**
 * 
 */
package com.vmware.jct.service.vo;

import java.util.Map;

/**
 * 
 * <p><b>Class name:</b> AfterSketchOneVO.java</p>
 * <p><b>Author:</b> InterraIT</p>
 * <p><b>Purpose:</b> This class acts as a value object.
 * <p><b>Description:</b> This class acts as a value object .. </p>
 * <p><b>Copyrights:</b> 	All rights reserved by Interra IT and should only be used for its internal application development.</p>
 * <p><b>Created Date:</b> 20/Feb/2014</p>
 * <p><b>Revision History:</b>
 * 	<li>6/May/2014: Changed the saving of action plan to dynamic.</li>
 * </p>
 */
public class AfterSketchOneVO {
	
	private Map strengthMap;
	private Map passionMap;
	private Map valueMap;
	
	//If already answered
	private String afterSketchCheckedEle;
	private String afterSkPageOneTotalJson;
	
	private int valueCount;
	private int passionCount;
	private int strengthCount;
	
	private String checkedPassion;
	private String checkedStrength;
	private String checkedValue;
	
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

}
