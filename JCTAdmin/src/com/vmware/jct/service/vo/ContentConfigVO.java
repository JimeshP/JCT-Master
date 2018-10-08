package com.vmware.jct.service.vo;

import java.util.Map;
/**
 * 
 * <p><b>Class name:</b> ContentConfigVO.java</p>
 * <p><b>Author:</b> InterraIT</p>
 * <p><b>Purpose:</b> This class acts as a value object. Transfer data from presentation layer to controller and vice versa
 * <p><b>Description:</b>This class acts as a value object. Transfer data from presentation layer to controller and vice versa </p>
 * <p><b>Copyrights:</b> 	All rights reserved by Interra IT and should only be used for its internal application development.</p>
 * <p><b>Created Date:</b> 20/Feb/2014</p>
 * <p><b>Revision History:</b>
 * 	<li>6/May/2014: Changed the saving of action plan to dynamic.</li>
 * </p>
 */

public class ContentConfigVO {
	private String existingUserProfileList;
	private String existingRefQtnList;
	private int statusCode;
	private String statusDesc;
	private Map<Integer, String> userProfileMap;
	private String noOfSubQtn;
	private String existingFunctionGrpList;
	private String existingJobLevelList;
	private String existingRegionList;
	private String existingMappingList;
	private String existingInstruction;
	private int existingInstructionId;
	private String mainQtnList;
	private String subQtnList;
	private String onetOccupationList;
	
	private String existingInstructionType;
	private String existingInstructionVideo;

	/**
	 * @return the existingUserProfileList
	 */
	
	public String getExistingUserProfileList() {
		return this.existingUserProfileList;
	}
	
	/**
	 * @param existingUserProfileList the existingUserProfileList to set
	 */
	public void setExistingUserProfileList(String existingUserProfileList) {
		this.existingUserProfileList = existingUserProfileList;
	}
		
	/**
	 * @return the existingRefQtnList
	 */	
	public String getExistingRefQtnList() {
		return this.existingRefQtnList;
	}
	
	/**
	 * @param existingRefQtnList the existingRefQtnList to set
	 */
	public void setExistingRefQtnList(String existingRefQtnList) {
		this.existingRefQtnList = existingRefQtnList;
	}
	
	/**
	 * @return the satusCode
	 */
	public int getStatusCode() {
		return this.statusCode;
	}
	
	/**
	 * @param satusCode the satusCode to set
	 */
	public void setStatusCode(int statusCode) {
		this.statusCode = statusCode;
	}
	
	/**
	 * @return the statusDesc
	 */
	public String getStatusDesc() {
		return this.statusDesc;
	}
	
	/**
	 * @param statusDesc the statusDesc to set
	 */
	public void setStatusDesc(String statusDesc) {
		this.statusDesc = statusDesc;
	}

	/**
	 * @return the userProfileMap
	 */
	public Map<Integer, String> getUserProfileMap() {
		return userProfileMap;
	}

	/**
	 * @param userProfileMap the userProfileMap to set
	 */
	public void setUserProfileMap(Map<Integer, String> userProfileMap) {
		this.userProfileMap = userProfileMap;
	}

	/**
	 * @param noOfSubQtn the noOfSubQtn to set
	 */
	public String getNoOfSubQtn() {
		return noOfSubQtn;
	}

	/**
	 * @return the noOfSubQtn
	 */
	public void setNoOfSubQtn(String noOfSubQtn) {
		this.noOfSubQtn = noOfSubQtn;
	}

	/**
	 * @return the existingFunctionGrpList
	 */	
	public String getExistingFunctionGrpList() {
		return existingFunctionGrpList;
	}

	/**
	 * @param existingFunctionGrpList the existingFunctionGrpList to set
	 */
	public void setExistingFunctionGrpList(String existingFunctionGrpList) {
		this.existingFunctionGrpList = existingFunctionGrpList;
	}

	/**
	 * @return the existingJobLevelList
	 */	
	public String getExistingJobLevelList() {
		return existingJobLevelList;
	}

	/**
	 * @param existingJobLevelList the existingJobLevelList to set
	 */
	public void setExistingJobLevelList(String existingJobLevelList) {
		this.existingJobLevelList = existingJobLevelList;
	}

	/**
	 * @return the existingRegionList
	 */	
	public String getExistingRegionList() {
		return existingRegionList;
	}

	/**
	 * @param existingRegionList the existingRegionList to set
	 */
	public void setExistingRegionList(String existingRegionList) {
		this.existingRegionList = existingRegionList;
	}

	/**
	 * @return the existingMappingList
	 */	
	public String getExistingMappingList() {
		return existingMappingList;
	}

	/**
	 * @param existingMappingList the existingMappingList to set
	 */
	public void setExistingMappingList(String existingMappingList) {
		this.existingMappingList = existingMappingList;
	}

	/**
	 * 
	 * @return
	 */
	public String getExistingInstruction() {
		return existingInstruction;
	}

	/**
	 * 
	 * @param existingInstruction
	 */
	public void setExistingInstruction(String existingInstruction) {
		this.existingInstruction = existingInstruction;
	}

	/**
	 * 
	 * @return
	 */
	public int getExistingInstructionId() {
		return existingInstructionId;
	}

	/**
	 * 
	 * @param existingInstructionId
	 */
	public void setExistingInstructionId(int existingInstructionId) {
		this.existingInstructionId = existingInstructionId;
	}
	/**
	 * 
	 * @return
	 */
	public String getMainQtnList() {
		return mainQtnList;
	}
	/**
	 * 
	 * @param mainQtnList
	 */
	public void setMainQtnList(String mainQtnList) {
		this.mainQtnList = mainQtnList;
	}
	/**
	 * 
	 * @return
	 */
	public String getSubQtnList() {
		return subQtnList;
	}
	/**
	 * 
	 * @param subQtnList
	 */
	public void setSubQtnList(String subQtnList) {
		this.subQtnList = subQtnList;
	}
	/**
	 * 
	 * @return
	 */
	public String getOnetOccupationList() {
		return onetOccupationList;
	}
	/**
	 * 
	 * @param onetOccupationList
	 */
	public void setOnetOccupationList(String onetOccupationList) {
		this.onetOccupationList = onetOccupationList;
	}

	public String getExistingInstructionType() {
		return existingInstructionType;
	}

	public void setExistingInstructionType(String existingInstructionType) {
		this.existingInstructionType = existingInstructionType;
	}

	public String getExistingInstructionVideo() {
		return existingInstructionVideo;
	}

	public void setExistingInstructionVideo(String existingInstructionVideo) {
		this.existingInstructionVideo = existingInstructionVideo;
	}
}
