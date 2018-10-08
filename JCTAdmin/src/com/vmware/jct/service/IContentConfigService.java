package com.vmware.jct.service;

import java.util.List;
import java.util.Map;

import com.vmware.jct.dao.dto.OnetOccupationDTO;
import com.vmware.jct.exception.JCTException;
import com.vmware.jct.model.JctTermsAndConditions;
import com.vmware.jct.service.vo.ContentConfigVO;
import com.vmware.jct.service.vo.InstructionVO;
import com.vmware.jct.service.vo.JobAttributeVO;
import com.vmware.jct.service.vo.QuestionearVO;
/**
 * 
 * <p><b>Class name:</b>IContentConfigService.java</p>
 * <p><b>@author:</b> InterraIT </p>
 * <p><b>Purpose:</b> 		This Service interface represents a IContentConfigService interface
 * <p><b>Description:</b> 	This will have all methods where we implement all business logic and all service layer </p>
 * <p><b>Copyrights:</b> 	All rights reserved by InterraIT and should only be used for its internal application development.</p>
 *  <pre>
 * <b>Revision History:</b>
 * InterraIT - 31/Jan/2014 - Implement exception through out the application
 * </pre>
 */
public interface IContentConfigService {
	/**
	 * Method populate reflection question
	 * @return ContentConfigVO
	 * @throws JCTException
	 */
	public ContentConfigVO populateExistingRefQtn(Integer profileId, String relatedPage) throws JCTException;
	/**
	 * Method populates user profile
	 * @return Map
	 * @throws JCTException
	 */
	public Map<Integer, String> populateUserProfile() throws JCTException;
	/**
	 * Method saves new reflection question
	 * @param questionearVO
	 * @return ContentConfigVO
	 * @throws JCTException
	 */
	public ContentConfigVO saveRefQtn(QuestionearVO questionearVO) throws JCTException;
	/**
	 * Method updates existing reflection question
	 * @param questionearVO
	 * @param dist
	 * @return ContentConfigVO
	 * @throws JCTException
	 */
	public ContentConfigVO updateRefQtn(QuestionearVO questionearVO, String dist, String chechBoxVal) throws JCTException;
	/**
	 * Method deletes existing reflection question
	 * @param questionearVO
	 * @param dist
	 * @return ContentConfigVO
	 * @throws JCTException
	 */
	public ContentConfigVO deleteRefQtn(QuestionearVO questionearVO, String dist) throws JCTException;
	/**
	 * Method saves action plan
	 * @param questionearVO
	 * @return ContentConfigVO
	 * @throws JCTException
	 */
	public ContentConfigVO saveActionPlan(QuestionearVO questionearVO) throws JCTException;	
	/**
	 * Method updates existing action plan
	 * @param questionearVO
	 * @param dist
	 * @return ContentConfigVO
	 * @throws JCTException
	 */
	public ContentConfigVO updateActionPlan(QuestionearVO questionearVO, String dist, String chechBoxVal) throws JCTException;
	/**
	 * Method deletes existing action plan
	 * @param questionearVO
	 * @param dist
	 * @return ContentConfigVO
	 * @throws JCTException
	 */
	public ContentConfigVO deleteActionPlan(QuestionearVO questionearVO, String dist) throws JCTException;
	/**
	 * Method populates existing function group
	 * @return ContentConfigVO
	 * @throws JCTException
	 */
	public ContentConfigVO populateExistingFunctionGroup() throws JCTException;
	/**
	 * Method populates existing job level
	 * @return ContentConfigVO
	 * @throws JCTException
	 */
	public ContentConfigVO populateExistingJobLevel() throws JCTException;
	/**
	 * Method saves new function group
	 * @param functionGrp
	 * @param createdBy
	 * @return ContentConfigVO
	 * @throws JCTException
	 */
	public ContentConfigVO saveFunctionGroup(String functionGrp, String createdBy) throws JCTException;	
	/**
	 * Method saves new job level
	 * @param jobLevel
	 * @param createdBy
	 * @return ContentConfigVO
	 * @throws JCTException
	 */
	public ContentConfigVO saveJobLevel(String jobLevel, String createdBy) throws JCTException;	
	/**
	 * Method populates existing region
	 * @return ContentConfigVO
	 * @throws JCTException
	 */
	public ContentConfigVO populateExistingRegion() throws JCTException;
	/**
	 * Method saves new Region
	 * @param regionName
	 * @param createdBy
	 * @return ContentConfigVO
	 * @throws JCTException
	 */
	public ContentConfigVO saveRegion(String regionName, String createdBy) throws JCTException;	
	/**
	 * Method updates existing region
	 * @param regionName
	 * @param tablePkId
	 * @return ContentConfigVO
	 * @throws JCTException
	 */
	public ContentConfigVO updateRegion(String regionName, Integer tablePkId) throws JCTException;	
	/**
	 * Method deletes existing region
	 * @param regionName
	 * @param tablePkId
	 * @return ContentConfigVO
	 * @throws JCTException
	 */
	public ContentConfigVO deleteRegion(String regionName, Integer tablePkId) throws JCTException;	
	/**
	 * Method populates existing job attribute mappings
	 * @return ContentConfigVO
	 * @throws JCTException
	 */
	public ContentConfigVO populateExistingMapping(Integer profileId, String relatedPage) throws JCTException;
	/**
	 * Method saves new strength (attribute mapping)
	 * @param jobAttributeVO
	 * @return ContentConfigVO
	 * @throws JCTException
	 */
	public ContentConfigVO saveStrength(JobAttributeVO jobAttributeVO) throws JCTException;
	/**
	 * Method saves new value (attribute mapping)
	 * @param jobAttributeVO
	 * @return ContentConfigVO
	 * @throws JCTException
	 */
	public ContentConfigVO saveValue(JobAttributeVO jobAttributeVO) throws JCTException;
	/**
	 * Method saves new passion (attribute mapping)
	 * @param jobAttributeVO
	 * @return ContentConfigVO
	 * @throws JCTException
	 */
	public ContentConfigVO savePassion(JobAttributeVO jobAttributeVO) throws JCTException;	
	/**
	 * Method updates existing strength (attribute mapping)
	 * @param jobAttributeVO
	 * @param chechBoxVal
	 * @return ContentConfigVO
	 * @throws JCTException
	 */
	public ContentConfigVO updateStrength(JobAttributeVO jobAttributeVO, String dist, String chechBoxVal) throws JCTException;
	/**
	 * Method updates existing value (attribute mapping)
	 * @param jobAttributeVO
	 * @param chechBoxVal
	 * @return ContentConfigVO
	 * @throws JCTException
	 */
	public ContentConfigVO updateValue(JobAttributeVO jobAttributeVO, String dist, String chechBoxVal) throws JCTException;
	/**
	 * Method updates existing passion (attribute mapping)
	 * @param jobAttributeVO
	 * @param chechBoxVal
	 * @return ContentConfigVO
	 * @throws JCTException
	 */
	public ContentConfigVO updatePassion(JobAttributeVO jobAttributeVO, String dist, String chechBoxVal) throws JCTException;
	/**
	 * Method deletes existing strength (attribute mapping)
	 * @param jobAttributeVO
	 * @return ContentConfigVO
	 * @throws JCTException
	 */
	public ContentConfigVO deleteStrength(JobAttributeVO jobAttributeVO, String dist) throws JCTException;
	/**
	 * Method deletes existing value (attribute mapping)
	 * @param jobAttributeVO
	 * @return ContentConfigVO
	 * @throws JCTException
	 */
	public ContentConfigVO deleteValue(JobAttributeVO jobAttributeVO, String dist) throws JCTException;
	/**
	 * Method deletes existing passion (attribute mapping)
	 * @param jobAttributeVO
	 * @return ContentConfigVO
	 * @throws JCTException
	 */
	public ContentConfigVO deletePassion(JobAttributeVO jobAttributeVO, String dist) throws JCTException;
	/**
	 * Method saves new instruction data
	 * @param jobAttributeVO
	 * @return ContentConfigVO
	 * @throws JCTException
	 */
	public ContentConfigVO saveInstruction(InstructionVO instructionVO) throws JCTException;	
	/**
	 * Method updates existing instruction data
	 * @param instructionId
	 * @param insAreaText
	 * @return ContentConfigVO
	 * @throws JCTException
	 */
	public ContentConfigVO updateInstruction(int instructionId, String insAreaText) throws JCTException;	
	/**
	 * Method fetches existing information data based on user group and user profile
	 * @param userProfileId
	 * @param relatedPage
	 * @return ContentConfigVO
	 * @throws JCTException
	 */
	public ContentConfigVO populateExistingInstruction(int userProfileId, String relatedPage) throws JCTException;	
	/**
	 * Method validates existing reflection question
	 * @param userGroup
	 * @param userProfile
	 * @return String - status
	 * @throws JCTException
	 */
	public String validateExistenceRefQtn(String userGroup,int userProfile) throws JCTException;
	/**
	 * Method validates existing action plan
	 * @param userGroup
	 * @param userProfile
	 * @return String - status
	 * @throws JCTException
	 */
	public String validateExistenceActionPlan(String refQtnDesc, String subQtnDesc, int userProfileId, String code, String action) throws JCTException;
	/**
	 * Method validates existing function group
	 * @param userGroup
	 * @param userProfile
	 * @return String - status
	 * @throws JCTException
	 */
	public String validateExistenceFuncGrp(String functionGrp) throws JCTException;
	/**
	 * Method validates existing job level
	 * @param userGroup
	 * @param userProfile
	 * @return String - status
	 * @throws JCTException
	 */
	public String validateExistenceJobLevel(String jobLevel) throws JCTException;
	/**
	 * Method validates existing job attributes
	 * @param userGroup
	 * @param userProfile
	 * @return String - status
	 * @throws JCTException
	 */
	public String validateExistenceAttribute(String attrDescText,int userProfile,String attrCode, String action) throws JCTException;
	/**
	 * Method validates existing region
	 * @param userGroup
	 * @param userProfile
	 * @return String - status
	 * @throws JCTException
	 */
	public String validateExistenceRegion(String regionName) throws JCTException;
	/**
	 * Method populate main question
	 * @return ContentConfigVO
	 * @throws JCTException
	 */
	public ContentConfigVO populateMainQtn(Integer profileId, String relatedPage) throws JCTException;
	/**
	 * Method populate sub question
	 * @return ContentConfigVO
	 * @throws JCTException
	 */
	public ContentConfigVO populateSubQtn(Integer profileId, String relatedPage, String mainQtn) throws JCTException;
	/**
	 * Method save main question by order
	 * @return ContentConfigVO
	 * @throws JCTException
	 */
	public ContentConfigVO saveMainQtnOrder(String mainQtn, Integer profileId, String relatedPage) throws JCTException;
	/**
	 * Method save sub question by order
	 * @return ContentConfigVO
	 * @throws JCTException
	 */
	public ContentConfigVO saveSubQtnOrder(String subQtn, Integer profileId, String relatedPage, String mainQtn) throws JCTException;
	/**
	 * Method populate function Group
	 * @return ContentConfigVO
	 * @throws JCTException
	 */
	public ContentConfigVO saveFuncGrpOrder(String builder) throws JCTException;
	/**
	 * Method populates existing function group by order
	 * @return ContentConfigVO
	 * @throws JCTException
	 */
	public ContentConfigVO populateFunctionGroupByOrder() throws JCTException;	
	/**
	 * Method populate job level
	 * @return ContentConfigVO
	 * @throws JCTException
	 */
	public ContentConfigVO saveJobLevelOrder(String builder) throws JCTException;
	/**
	 * Method populates existing job level by order
	 * @return ContentConfigVO
	 * @throws JCTException
	 */
	public ContentConfigVO populateJobLevelByOrder() throws JCTException;	
	/**
	 * Method populates existing attributes by order
	 * @return ContentConfigVO
	 * @throws JCTException
	 */
	public ContentConfigVO populateMappingList(Integer profileId, String attrCode) throws JCTException;	
	/**
	 * Method populate attribute
	 * @return ContentConfigVO
	 * @throws JCTException
	 */
	public ContentConfigVO saveAttributeOrder(String builder, Integer profileId, String attrCode) throws JCTException;
	/**
	 * Method populate region
	 * @return ContentConfigVO
	 * @throws JCTException
	 */
	public ContentConfigVO saveRegionOrder(String builder) throws JCTException;
	
	/******* THIS IS FOR JCT PUBLIC VERSION ********/
	/**
	 * Method populates existing ONet data
	 * @return ContentConfigVO
	 * @throws JCTException
	 */
	public ContentConfigVO populateOnetDataList() throws JCTException;
	
	/**
	 * Method to fetch existing terms & condition
	 * @param userProfile
	 * @param userType
	 * @return ContentConfigVO
	 * @throws JCTException
	 */
	public JctTermsAndConditions fetchTermsAndCondition(int userProfile, int userType) throws JCTException;
	/**
	 * Method to fetch existing terms & condition
	 * @param userProfile
	 * @param userType
	 * @return ContentConfigVO
	 * @throws JCTException
	 */
	public String saveTermsAndCondition(int userProfile, int userType, String tcText, String createdBy) throws JCTException;
	
	public List<OnetOccupationDTO> populateOnetDataListVo() throws JCTException;
	
	public String saveVideoInstruction(InstructionVO instructionVO) throws JCTException;
}
