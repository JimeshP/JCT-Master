package com.vmware.jct.service.vo;
/**
 * 
 * <p><b>Class name:</b> JctAfterSketchHeaderVO.java</p>
 * <p><b>Author:</b> InterraIT</p>
 * <p><b>Purpose:</b> This class acts as a value object.
 * <p><b>Description:</b> This class acts as a value object .. </p>
 * <p><b>Copyrights:</b> 	All rights reserved by Interra IT and should only be used for its internal application development.</p>
 * <p><b>Created Date:</b> 20/Feb/2014</p>
 * <p><b>Revision History:</b>
 * 	<li>6/May/2014: Changed the saving of action plan to dynamic.</li>
 * </p>
 */
public class JctAfterSketchHeaderVO {
	private String jobReferenceNo;
	private String jctPageOneJson;
	private String jctFinalPageJson;
	private int jctPageOneTimeSpent;
	private int jctFinalPageTimeSpent;
	private String jctLastModifiedBy;
	private String jctCreatedBy;
	private String jctJobTitle;
	private byte[] jctFinalPageSnapshot;
	private String jctFinalPageInitialJSonValue;
	private String jctFinalPageSnapshotString;
	private String jctPageOneCheckedElements;
	private String divHeight;
	private String divWidth;
	private String checkedPassion;
	private String checkedStrength;
	private String checkedValue;
	private int asTotalCount;
	private int jctUserId;
	
	public int getAsTotalCount() {
		return asTotalCount;
	}
	public void setAsTotalCount(int asTotalCount) {
		this.asTotalCount = asTotalCount;
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
	public String getJobReferenceNo() {
		return jobReferenceNo;
	}
	public void setJobReferenceNo(String jobReferenceNo) {
		this.jobReferenceNo = jobReferenceNo;
	}
	public String getJctPageOneJson() {
		return jctPageOneJson;
	}
	public void setJctPageOneJson(String jctPageOneJson) {
		this.jctPageOneJson = jctPageOneJson;
	}
	public String getJctFinalPageJson() {
		return jctFinalPageJson;
	}
	public void setJctFinalPageJson(String jctFinalPageJson) {
		this.jctFinalPageJson = jctFinalPageJson;
	}
	public int getJctPageOneTimeSpent() {
		return jctPageOneTimeSpent;
	}
	public void setJctPageOneTimeSpent(int jctPageOneTimeSpent) {
		this.jctPageOneTimeSpent = jctPageOneTimeSpent;
	}
	public int getJctFinalPageTimeSpent() {
		return jctFinalPageTimeSpent;
	}
	public void setJctFinalPageTimeSpent(int jctFinalPageTimeSpent) {
		this.jctFinalPageTimeSpent = jctFinalPageTimeSpent;
	}
	public String getJctLastModifiedBy() {
		return jctLastModifiedBy;
	}
	public void setJctLastModifiedBy(String jctLastModifiedBy) {
		this.jctLastModifiedBy = jctLastModifiedBy;
	}
	public String getJctCreatedBy() {
		return jctCreatedBy;
	}
	public void setJctCreatedBy(String jctCreatedBy) {
		this.jctCreatedBy = jctCreatedBy;
	}
	public String getJctJobTitle() {
		return jctJobTitle;
	}
	public void setJctJobTitle(String jctJobTitle) {
		this.jctJobTitle = jctJobTitle;
	}
	public byte[] getJctFinalPageSnapshot() {
		return jctFinalPageSnapshot;
	}
	public void setJctFinalPageSnapshot(byte[] jctFinalPageSnapshot) {
		this.jctFinalPageSnapshot = jctFinalPageSnapshot;
	}
	public String getJctFinalPageInitialJSonValue() {
		return jctFinalPageInitialJSonValue;
	}
	public void setJctFinalPageInitialJSonValue(String jctFinalPageInitialJSonValue) {
		this.jctFinalPageInitialJSonValue = jctFinalPageInitialJSonValue;
	}
	public String getJctFinalPageSnapshotString() {
		return jctFinalPageSnapshotString;
	}
	public void setJctFinalPageSnapshotString(String jctFinalPageSnapshotString) {
		this.jctFinalPageSnapshotString = jctFinalPageSnapshotString;
	}
	public String getJctPageOneCheckedElements() {
		return jctPageOneCheckedElements;
	}
	public void setJctPageOneCheckedElements(String jctPageOneCheckedElements) {
		this.jctPageOneCheckedElements = jctPageOneCheckedElements;
	}
	public int getJctUserId() {
		return jctUserId;
	}
	public void setJctUserId(int jctUserId) {
		this.jctUserId = jctUserId;
	}
}
