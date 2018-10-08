package com.vmware.jct.dao.dto;

public class OccupationListDTO {
	private String code;
	private String title;
	private String desc;
	public OccupationListDTO(String code,String title, String desc) {
		this.code = code;
		this.title = title;
		this.desc = desc;
	}
	public String getCode() {
		return code;
	}
	public void setCode(String code) {
		this.code = code;
	}
	public String getDesc() {
		return desc;
	}
	public void setDesc(String desc) {
		this.desc = desc;
	}
	public String getTitle() {
		return title;
	}
	public void setTitle(String title) {
		this.title = title;
	}
	
}
