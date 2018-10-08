package com.vmware.jct.dao;

import java.util.List;

import com.vmware.jct.dao.dto.FunctionDTO;
import com.vmware.jct.dao.dto.JobAttributeDTO;
import com.vmware.jct.dao.dto.LevelDTO;
import com.vmware.jct.dao.dto.OnetOccupationDTO;
import com.vmware.jct.dao.dto.QuestionearDTO;
import com.vmware.jct.dao.dto.RegionDTO;
import com.vmware.jct.dao.dto.UserProfileDTO;
import com.vmware.jct.exception.DAOEntityExistsException;
import com.vmware.jct.exception.DAOException;
import com.vmware.jct.exception.DAONullException;
import com.vmware.jct.exception.JCTException;
import com.vmware.jct.model.JctFunction;
import com.vmware.jct.model.JctInstructionBar;
import com.vmware.jct.model.JctJobAttribute;
import com.vmware.jct.model.JctLevel;
import com.vmware.jct.model.JctQuestion;
import com.vmware.jct.model.JctRegion;
import com.vmware.jct.model.JctTermsAndConditions;
import com.vmware.jct.model.JctUserProfile;

/**
 * 
 * <p><b>Class name:</b>IContentServiceDAO.java</p>
 * <p><b>@author:</b> InterraIT </p>
 * <p><b>Purpose:</b> 		This DAO interface represents a IAuthenticatorDAO interface
 * <p><b>Description:</b> 	This will have all methods mapped to the database and fetching data from database to be used for Content service. </p>
 * <p><b>Note:</b> 			This methods implemented by DAOImpl classes and execute the mapped queries and get the records.</p>  
 * <p><b>Copyrights:</b> 	All rights reserved by InterraIT and should only be used for its internal application development.</p>
 *  <pre>
 * <b>Revision History:</b>
 * <li> </li>
 * </pre>
 */
public interface IContentServiceDAO {
	/**
	 * Method returns existing reflection questions in list of 
	 * questionearDTO
	 * @return list of Questionear dto
	 * @throws DAOException
	 */
	public List<QuestionearDTO> populateExistingRefQtn(Integer profileId, String relatedPage) throws DAOException;
	/**
	 * Method returns list of user profile dtos
	 * @return list of user profile dto
	 * @throws DAOException
	 */
	public List<UserProfileDTO> populateExistingUserProfileWithId() throws DAOException;
	/**
	 * Method saves new reflection question
	 * @param question
	 * @return status
	 * @throws DAOException
	 */
	public String saveRefQtn(JctQuestion question) throws DAOException;
	/**
	 * Method fetches existing reflection question
	 * @param pkId
	 * @return Jct question entity
	 * @throws DAOException
	 */
	public JctQuestion fetchRefQtn(int pkId) throws DAOException;
	/**
	 * Method updates existing reflection question
	 * @param question
	 * @return status
	 * @throws DAOException
	 */
	public String updateRefQtn(JctQuestion question) throws DAOException;
	/**
	 * Method deletes existing reflection question
	 * @param question
	 * @return status
	 * @throws DAOException
	 */
	public String deleteRefQtn(JctQuestion question) throws DAOException;
	/**
	 * Method saves new action plan
	 * @param question
	 * @return status
	 * @throws DAOException
	 */
	public String saveActionPlan(JctQuestion question) throws DAOException;
	/**
	 * Method updates existing action plan
	 * @param question
	 * @return status
	 * @throws DAOException
	 */
	public String updateActionPlan(JctQuestion question) throws DAOException;
	/**
	 * Method deletes existing action plan
	 * @param question
	 * @return status
	 * @throws DAOException
	 */
	public String deleteActionPlan(JctQuestion question) throws DAOException;
	/**
	 * Method fetches existing function groups
	 * @return list of function group dto
	 * @throws DAOException
	 */
	public List<FunctionDTO> populateExistingFunctionGroup() throws DAOException;
	/**
	 * Method populates existing job level
	 * @return list of LevelDTO
	 * @throws DAOException
	 */
	public List<LevelDTO> populateExistingJobLevel() throws DAOException;
	/**
	 * Method saves new function group
	 * @param functionGrp
	 * @return status
	 * @throws DAOException
	 */
	public String saveFunctionGroup(JctFunction functionGrp) throws DAOException;
	/**
	 * Method saves new job level
	 * @param jobLevel
	 * @return status
	 * @throws DAOException
	 */
	public String saveJobLevel(JctLevel jobLevel) throws DAOException;
	/**
	 * Method populates existing region
	 * @return list of RegionDTO
	 * @throws DAOException
	 */
	public List<RegionDTO> populateExistingRegion() throws DAOException;
	/**
	 * Method saves new region
	 * @param regionName
	 * @return status
	 * @throws DAOException
	 */
	public String saveRegion(JctRegion regionName) throws DAOException;
	/**
	 * Method fetch existing region
	 * @param pkId
	 * @return JctRegion entity
	 * @throws DAOException
	 */
	public JctRegion fetchRegion(int pkId) throws DAOException;
	/**
	 * Method updates the existing region
	 * @param region
	 * @return status
	 * @throws DAOException
	 */
	public String updateRegion(JctRegion region) throws DAOException;
	/**
	 * Method deletes existing region
	 * @param region
	 * @return status
	 * @throws DAOException
	 */
	public String deleteRegion(JctRegion region) throws DAOException;
	/**
	 * Method populates existing mapping attributes
	 * @return List of JobAttributeDTO
	 * @throws DAOException
	 */
	public List<JobAttributeDTO> populateExistingMapping(Integer profileId, String relatedPage) throws DAOException;
	/**
	 * Method saves new strength
	 * @param jobAttribute
	 * @return status
	 * @throws DAOException
	 */
	public String saveStrength(JctJobAttribute jobAttribute) throws DAOException;
	/**
	 * Method saves new value
	 * @param jobAttribute
	 * @return status
	 * @throws DAOException
	 */
	public String saveValue(JctJobAttribute jobAttribute) throws DAOException;
	/**
	 * Method saves new passion
	 * @param jobAttribute
	 * @return status
	 * @throws DAOException
	 */
	public String savePassion(JctJobAttribute jobAttribute) throws DAOException;
	/**
	 * Method updates existing Strength
	 * @param jobAttribute
	 * @return status
	 * @throws DAOException
	 */
	public String updateStrength(JctJobAttribute jobAttribute) throws DAOException;
	/**
	 * Method updates existing value
	 * @param jobAttribute
	 * @return status
	 * @throws DAOException
	 */
	public String updateValue(JctJobAttribute jobAttribute) throws DAOException;
	/**
	 * Method updates existing passion
	 * @param jobAttribute
	 * @return
	 * @throws DAOException
	 */
	public String updatePassion(JctJobAttribute jobAttribute) throws DAOException;
	/**
	 * Method deletes existing strength
	 * @param jobAttribute
	 * @return status
	 * @throws DAOException
	 */
	public String deleteStrength(JctJobAttribute jobAttribute) throws DAOException;
	/**
	 * Method deletes existing value
	 * @param jobAttribute
	 * @return status
	 * @throws DAOException
	 */
	public String deleteValue(JctJobAttribute jobAttribute) throws DAOException;
	/**
	 * Method deletes existing passion
	 * @param jobAttribute
	 * @return status
	 * @throws DAOException
	 */
	public String deletePassion(JctJobAttribute jobAttribute) throws DAOException;
	/**
	 * Method fetches mapping by primary key
	 * @param pkId
	 * @return JctJobAttribute entity
	 * @throws DAOException
	 */
	public JctJobAttribute fetchMapping(int pkId) throws DAOException;
	/**
	 * Method saves new instruction
	 * @param instruction
	 * @return status
	 * @throws DAOException
	 */
	public String saveInstruction(JctInstructionBar instruction) throws DAOException;
	/**
	 * Method fetches existing instruction data
	 * @param primaryKey
	 * @return JctInstructionBar entity
	 * @throws DAOException
	 */
	public JctInstructionBar fetchInstructionData(Integer primaryKey) throws DAOException;
	/**
	 * Method updates existing Instruction
	 * @param instruction
	 * @return status
	 * @throws DAOException
	 */
	public String updateInstruction(JctInstructionBar instruction) throws DAOException;
	/**
	 * Method fetches existing instruction based on profile id and related page
	 * @param userProfileId
	 * @param relatedPage
	 * @return JctInstructionBar entity
	 * @throws DAOException
	 */
	public JctInstructionBar populateExistingInstruction(int userProfileId, String relatedPage) throws DAOException;
	/**
	 * Method validates reflection question
	 * @param refQtnDesc
	 * @param userProfileId
	 * @return status
	 * @throws DAOException
	 */
	public String validateRefQtn(String refQtnDesc, int userProfileId) throws DAOException;
	/**
	 * Method validates action plan
	 * @param refQtnDesc
	 * @param subQtnDesc
	 * @param userProfileId
	 * @return status
	 * @throws DAOException
	 */
	public String validateActionPlan(String refQtnDesc, String subQtnDesc, int userProfileId, String code, String action) throws DAOException;
	/**
	 * Method validates function group
	 * @param functionGrp
	 * @return status
	 * @throws DAOException
	 */
	public String validateFuncGrp(String functionGrp) throws DAOException;
	/**
	 * Method validates job level
	 * @param jobLevel
	 * @return status
	 * @throws DAOException
	 */
	public String validateJobLevel(String jobLevel) throws DAOException;
	/**
	 * Method validates job attributes
	 * @param attrDescText
	 * @param userProfileId
	 * @param attrCode
	 * @param action
	 * @return status
	 * @throws DAOException
	 */
	public String validateAttribute(String attrDescText, int userProfileId, String attrCode, String action) throws DAOException;
	/**
	 * Method validates region
	 * @param regionName
	 * @return status
	 * @throws DAOException
	 */
	public String validateRegion(String regionName) throws DAOException;
	/**
	 * Method fetches existing reflection question
	 * @param pkId
	 * @return Jct question entity
	 * @throws DAOException
	 */
	public JctQuestion fetchRefQtnByQtnOrder(int profileId, int qtnOrder, String page) throws DAOException;
	/**
	 * Method fetches mapping by profile Id
	 * @param profileId
	 * @param attrOrder
	 * @param page
	 * @return JctJobAttribute entity
	 * @throws DAOException
	 */
	public JctJobAttribute fetchMappingByOrder(int profileId, int attrOrder, String page) throws DAOException;
	/**
	 * 
	 * @param It will take string  object as a input parameter
	 * @return unique order number
	 * @throws DAOException
	 */
	public Integer generateOrderAttr(String attrCode, Integer profileId) throws DAOException;
	/**
	 * 	
	 * @param null
	 * @return unique order no
	 * @throws DAOException
	 */
	public Integer generateOrderFunctionGrp() throws DAOException;
	/**
	 * 	
	 * @param null
	 * @return unique order no
	 * @throws DAOException
	 */
	public Integer generateOrderLevel() throws DAOException;
	/**
	 * 	
	 * @param It will take string  object as a input parameter
	 * @return unique order no
	 * @throws DAOException
	 */
	public Integer generateOrderQtn(String qtnCode, Integer profileId, String qtnDesc) throws DAOException;
	/**
	 * Method returns existing main questions in list of 
	 * questionearDTO
	 * @return list of Questionear dto
	 * @throws DAOException
	 */
	public List<QuestionearDTO> populateMainQtn(Integer profileId, String relatedPage) throws DAOException;
	/**
	 * Method returns existing sub questions in list of 
	 * questionearDTO
	 * @return list of Questionear dto
	 * @throws DAOException
	 */
	public List<QuestionearDTO> populateSubQtn(Integer profileId, String relatedPage, String subQtn) throws DAOException;
	/**
	 * Method returns existing sub questions in list of 
	 * questionearDTO
	 * @return list of Questionear dto
	 * @throws DAOException
	 */
	public String saveSubQtnOrder(String subQtnBuilder, Integer profileId, String relatedPage, String mainQtn) throws DAOException;
	/**
	 * Method returns existing sub questions in list of 
	 * questionearDTO
	 * @return list of Questionear dto
	 * @throws DAOException
	 */
	public String saveMainQtnOrder(String mainQtnBuilder, Integer profileId, String relatedPage) throws DAOException;
	/**
	 * Method returns existing function Group in list of 
	 * functionDTO
	 * @return list of Function Group dto
	 * @throws DAOException
	 */
	public String saveFuncGrpOrder(String builder) throws DAOException;
	/**
	 * Method fetches existing function groups by order
	 * @return list of function group dto
	 * @throws DAOException
	 */
	public List<FunctionDTO> populateFunctionGroupByOrder() throws DAOException;
	/**
	 * Method returns existing job level list of 
	 * jobLevelDTO
	 * @return list of Function Group dto
	 * @throws DAOException
	 */
	public String saveJobLevelOrder(String builder) throws DAOException;
	/**
	 * Method fetches existing job level by order
	 * @return list of function group dto
	 * @throws DAOException
	 */
	public List<LevelDTO> populateJobLevelByOrder() throws DAOException;
	/**
	 * Method fetches existing attribute by order
	 * @return list of attribute dto
	 * @throws DAOException
	 */
	public List<JobAttributeDTO> populateMappingListByOrder(Integer profileId,  String arrtCode) throws DAOException;
	/**
	 * Method returns existing sub questions in list of 
	 * questionearDTO
	 * @return list of Questionear dto
	 * @throws DAOException
	 */
	public String saveAttributeOrder(String builder, Integer profileId, String attrCode) throws DAOException;
	/**
	 * 	
	 * @param null
	 * @return unique order no
	 * @throws DAOException
	 */
	public Integer generateOrderRegion() throws DAOException;
	/**
	 * Method returns existing region list of 
	 * jobLevelDTO
	 * @return list of region dto
	 * @throws DAOException
	 */
	public String saveRegionOrder(String builder) throws DAOException;
	
	/******* THIS IS FOR JCT PUBLIC VERSION ********/
	/**
	 * Method populates existing ONet Occupation data
	 * @return list of OnetOccupationDTO
	 * @throws DAOException
	 */
	public List<OnetOccupationDTO> populateOnetDataList() throws DAOException;
	
	/**
	 * Method updates existing Strength for global profile change
	 * @param code
	 * @param prevoiusName
	 * @param newName
	 * @param prevoiusDesc
	 * @param newDesc
	 * @throws DAOException
	 */
	public void updateMappingFrozen (String code, String prevoiusName, String newName, String prevoiusDesc, String newDesc) throws DAOException;
	
	/**
	 * Method updates existing reflection question and action plan for global profile change
	 * @param code
	 * @param mainQuestion
	 * @param prevoiusSubQtn
	 * @param newSubQtn
	 * @throws DAOException
	 */
	public void updateQuestionFrozen (String code, String mainQuestion, String prevoiusSubQtn, String newSubQtn) throws DAOException;
	/**
	 * Method to fetch existing terms & condition
	 * @param userProfile
	 * @param userType
	 * @return ContentConfigVO
	 * @throws JCTException
	 */
	public JctTermsAndConditions fetchTermsAndConditionDao(int userProfile, int userType) throws JCTException;

	/**
	 * Method to save terms & condition
	 * @param userProfile
	 * @return ContentConfigVO
	 * @throws DAOException 
	 * @throws DAOEntityExistsException 
	 * @throws DAONullException 
	 * @throws JCTException
	 */	
	public String saveTermsAndCondition(JctTermsAndConditions jctTermsAndConditions) throws DAONullException, DAOEntityExistsException, DAOException;
	/**
	 * Method to update terms & condition
	 * @param userProfile
	 * @throws DAOException 
	 * @throws DAOEntityExistsException 
	 * @throws DAONullException 
	 * @throws JCTException
	 */	
	public void removeTermsAndCondition(JctTermsAndConditions jctTermsAndConditions) throws DAONullException, DAOEntityExistsException, DAOException;
	/**
	 * get user profile by pk
	 * @param userProfile
	 * @return JctuserProfile
	 * */
	public JctUserProfile getUserProfileByPk(int userProfile) throws DAONullException, DAOEntityExistsException, DAOException;
}