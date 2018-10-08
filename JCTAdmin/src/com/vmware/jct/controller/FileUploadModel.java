package com.vmware.jct.controller;

import org.springframework.web.multipart.MultipartFile;

public class FileUploadModel {

	private MultipartFile file;

	public MultipartFile getFile() {
		return file;
	}

	public void setFile(MultipartFile file) {
		this.file = file;
	}
	
	/**
	 * FOR VIDEO ONLY
	 */
	private String hiddenFileName;
	private String fileType;
	private String inputUserProfileInpt;
	private String inputTextBeforeVideo;
	private String inputTextAfterVideo;

	public String getHiddenFileName() {
		return hiddenFileName;
	}

	public void setHiddenFileName(String hiddenFileName) {
		this.hiddenFileName = hiddenFileName;
	}

	public String getFileType() {
		return fileType;
	}

	public void setFileType(String fileType) {
		this.fileType = fileType;
	}

	public String getInputUserProfileInpt() {
		return inputUserProfileInpt;
	}

	public void setInputUserProfileInpt(String inputUserProfileInpt) {
		this.inputUserProfileInpt = inputUserProfileInpt;
	}

	public String getInputTextBeforeVideo() {
		return inputTextBeforeVideo;
	}

	public void setInputTextBeforeVideo(String inputTextBeforeVideo) {
		this.inputTextBeforeVideo = inputTextBeforeVideo;
	}

	public String getInputTextAfterVideo() {
		return inputTextAfterVideo;
	}

	public void setInputTextAfterVideo(String inputTextAfterVideo) {
		this.inputTextAfterVideo = inputTextAfterVideo;
	}
}
