package com.vmware.jct.service.vo;

import java.util.List;

import com.vmware.jct.dao.dto.SearchPdfDTO;

public class PdfSearchVO {
	
	String statuscode;
	String statusMsg;
	List<SearchPdfDTO> dto;
	
	public String getStatuscode() {
		return statuscode;
	}
	public void setStatuscode(String statuscode) {
		this.statuscode = statuscode;
	}
	public String getStatusMsg() {
		return statusMsg;
	}
	public void setStatusMsg(String statusMsg) {
		this.statusMsg = statusMsg;
	}
	public List<SearchPdfDTO> getDto() {
		return dto;
	}
	public void setDto(List<SearchPdfDTO> dto) {
		this.dto = dto;
	}
}