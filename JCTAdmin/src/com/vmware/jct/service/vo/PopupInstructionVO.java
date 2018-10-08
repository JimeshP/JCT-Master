package com.vmware.jct.service.vo;

import java.util.Map;

public class PopupInstructionVO {
	private int statusCode;
	private String statusDesc;
	private Map<Integer, String> userProfileMap;
	private String videoLink;
	private String instructionBeforeVideo;
	private String instructionAfterVideo;
	
	public int getStatusCode() {
		return statusCode;
	}
	public void setStatusCode(int statusCode) {
		this.statusCode = statusCode;
	}
	public String getStatusDesc() {
		return statusDesc;
	}
	public void setStatusDesc(String statusDesc) {
		this.statusDesc = statusDesc;
	}
	public Map<Integer, String> getUserProfileMap() {
		return userProfileMap;
	}
	public void setUserProfileMap(Map<Integer, String> userProfileMap) {
		this.userProfileMap = userProfileMap;
	}
	public String getVideoLink() {
		return videoLink;
	}
	public void setVideoLink(String videoLink) {
		this.videoLink = videoLink;
	}
	public String getInstructionBeforeVideo() {
		return instructionBeforeVideo;
	}
	public void setInstructionBeforeVideo(String instructionBeforeVideo) {
		this.instructionBeforeVideo = instructionBeforeVideo;
	}
	public String getInstructionAfterVideo() {
		return instructionAfterVideo;
	}
	public void setInstructionAfterVideo(String instructionAfterVideo) {
		this.instructionAfterVideo = instructionAfterVideo;
	}
}