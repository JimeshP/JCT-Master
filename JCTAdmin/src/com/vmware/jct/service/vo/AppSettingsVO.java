package com.vmware.jct.service.vo;

public class AppSettingsVO {
	private String message;
	private int statusCode;
	
	private String headerColor;
	private String footerColor;
	private String subHeaderColor;
	private String instructionPanelColor;
	private String instructionPanelTextColor;
	private String image;
	
	public String getMessage() {
		return message;
	}
	public void setMessage(String message) {
		this.message = message;
	}
	public int getStatusCode() {
		return statusCode;
	}
	public void setStatusCode(int statusCode) {
		this.statusCode = statusCode;
	}
	/**
	 * @return the headerColor
	 */
	public String getHeaderColor() {
		return headerColor;
	}
	/**
	 * @param headerColor the headerColor to set
	 */
	public void setHeaderColor(String headerColor) {
		this.headerColor = headerColor;
	}
	/**
	 * @return the footerColor
	 */
	public String getFooterColor() {
		return footerColor;
	}
	/**
	 * @param footerColor the footerColor to set
	 */
	public void setFooterColor(String footerColor) {
		this.footerColor = footerColor;
	}
	/**
	 * @return the subHeaderColor
	 */
	public String getSubHeaderColor() {
		return subHeaderColor;
	}
	/**
	 * @param subHeaderColor the subHeaderColor to set
	 */
	public void setSubHeaderColor(String subHeaderColor) {
		this.subHeaderColor = subHeaderColor;
	}
	/**
	 * @return the instructionPanelColor
	 */
	public String getInstructionPanelColor() {
		return instructionPanelColor;
	}
	/**
	 * @param instructionPanelColor the instructionPanelColor to set
	 */
	public void setInstructionPanelColor(String instructionPanelColor) {
		this.instructionPanelColor = instructionPanelColor;
	}
	/**
	 * @return the instructionPanelTextColor
	 */
	public String getInstructionPanelTextColor() {
		return instructionPanelTextColor;
	}
	/**
	 * @param instructionPanelTextColor the instructionPanelTextColor to set
	 */
	public void setInstructionPanelTextColor(String instructionPanelTextColor) {
		this.instructionPanelTextColor = instructionPanelTextColor;
	}
	/**
	 * @return the image
	 */
	public String getImage() {
		return image;
	}
	/**
	 * @param imageInByte the imageInByte to set
	 */
	public void setImage(String image) {
		this.image = image;
	}
}
