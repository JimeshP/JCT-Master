package com.vmware.jct.service.vo;

public class PopulateInstructionVO {	
	private String popupType;
	private int statusCode;
	private String instructionVideo;
	private String instructionTextBeforeVideo;
	private String instructionTextAfterVideo;
		
	public String getPopupType() {
		return popupType;
	}
	public void setPopupType(String popupType) {
		this.popupType = popupType;
	}
	public int getStatusCode() {
		return statusCode;
	}
	public void setStatusCode(int statusCode) {
		this.statusCode = statusCode;
	}
	public String getInstructionVideo() {
		return instructionVideo;
	}
	public void setInstructionVideo(String instructionVideo) {
		this.instructionVideo = instructionVideo;
	}
	public String getInstructionTextBeforeVideo() {
		return instructionTextBeforeVideo;
	}
	public void setInstructionTextBeforeVideo(String instructionTextBeforeVideo) {
		this.instructionTextBeforeVideo = instructionTextBeforeVideo;
	}
	public String getInstructionTextAfterVideo() {
		return instructionTextAfterVideo;
	}
	public void setInstructionTextAfterVideo(String instructionTextAfterVideo) {
		this.instructionTextAfterVideo = instructionTextAfterVideo;
	}
}
