package com.vmware.jct.controller;

import org.springframework.web.multipart.MultipartFile;

public class InstuctionUploadForm {
	private MultipartFile file;

	public MultipartFile getFile() {
		return file;
	}

	public void setFile(MultipartFile file) {
		this.file = file;
	}
	
	private String hiddenFileName;
	private String hiddenProfileInput;
	private String hiddenProfileValInput;
	private String hiddenPageInput;
	private String editor1;
	private String createdBy;
	
	public String getHiddenFileName() {
		return hiddenFileName;
	}

	public void setHiddenFileName(String hiddenFileName) {
		this.hiddenFileName = hiddenFileName;
	}

	

	public String getEditor1() {
		return editor1;
	}

	public void setEditor1(String editor1) {
		this.editor1 = editor1;
	}

	public String getHiddenProfileInput() {
		return hiddenProfileInput;
	}

	public void setHiddenProfileInput(String hiddenProfileInput) {
		this.hiddenProfileInput = hiddenProfileInput;
	}

	public String getHiddenPageInput() {
		return hiddenPageInput;
	}

	public void setHiddenPageInput(String hiddenPageInput) {
		this.hiddenPageInput = hiddenPageInput;
	}

	public String getHiddenProfileValInput() {
		return hiddenProfileValInput;
	}

	public void setHiddenProfileValInput(String hiddenProfileValInput) {
		this.hiddenProfileValInput = hiddenProfileValInput;
	}

	public String getCreatedBy() {
		return createdBy;
	}

	public void setCreatedBy(String createdBy) {
		this.createdBy = createdBy;
	}

	
	
	
}
