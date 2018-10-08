package com.vmware.jct.dao.impl;

import java.util.List;

import org.hibernate.SessionFactory;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.scheduling.annotation.Async;
import org.springframework.stereotype.Repository;
import org.springframework.transaction.annotation.Propagation;
import org.springframework.transaction.annotation.Transactional;

import com.vmware.jct.dao.DataAccessObject;
import com.vmware.jct.dao.IContentServiceDAO;
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
 * <p><b>Class name:</b> ContentConfigDAOImpl.java</p>
 * <p><b>Author:</b> InterraIT</p>
 * <p><b>Purpose:</b> This is a ContentConfigDAOImpl class. This artifact is Persistence Manager layer artifact.
 * ContentConfigDAOImpl implement IContentServiceDAO interface and override the following  methods.
 * -populateExistingRefQtn()
 * -populateExistingUserProfileWithId()
 * -saveRefQtn()
 * -updateRefQtn()
 * -fetchRefQtn()
 * -deleteRefQtn()
 * -saveActionPlan()
 * -updateActionPlan()
 * -deleteActionPlan()
 * -populateExistingFunctionGroup()
 * -populateExistingJobLevel()
 * -saveFunctionGroup()
 * -saveJobLevel()
 * -populateExistingRegion()
 * -saveRegion()
 * -fetchRegion()
 * -updateRegion()
 * -deleteRegion()
 * -populateExistingMapping()
 * -saveStrength()
 * -saveValue()
 * -savePassion()
 * -updateStrength()
 * -updateValue()
 * -updatePassion()
 * -fetchMapping()
 * -deleteStrength()
 * -deleteValue()
 * -deletePassion()
 * -saveInstruction()
 * -populateExistingInstruction()
 * -validateRefQtn()
 * -validateActionPlan()
 * -validateFuncGrp()
 * -validateJobLevel()
 * -validateAttribute()
 * -validateRegion()
 * -fetchInstructionData()
 * -updateInstruction()
 * </p>
 * <p><b>Description:</b> This class used to perform insert, modify ,fetch data from db and delete operation. Save, update, delete operations are performed in DataAccessObject class </p>
 * <p><b>Revision History:</b>
 * 	<li>InterraIT - 19/Jan/2014 - Implement Exception </li>
 * <li>InterraIT  - 31/Jan/2014 - Implement comments, introduce Named query, Transactional attribute </li>
 * <li>InterraIT - 16/May/2014 - Updated comments </li>
 * </p>
 */
@Repository
public class ContentConfigDAOImpl extends DataAccessObject implements IContentServiceDAO {

	@Autowired
	private SessionFactory sessionFactory;
	
	private static final Logger LOGGER = LoggerFactory.getLogger(ContentConfigDAOImpl.class);
	
	/**
	 * 
	 */
	@Transactional(propagation = Propagation.REQUIRED)
	public List<QuestionearDTO> populateExistingRefQtn(Integer profileId, String relatedPage)
			throws DAOException {
		LOGGER.info(">>>>>> ContentConfigDAOImpl.populateExistingRefQtn");
		List<QuestionearDTO> list = null;
		if(0 == profileId) {
			list = sessionFactory.getCurrentSession().getNamedQuery("fetchAllRefQtn").
					setParameter("relatedPage", relatedPage).list();
		} else {
			list = sessionFactory.getCurrentSession().getNamedQuery("fetchAllRefQtnByProfileId").
					setParameter("profileId", profileId).setParameter("relatedPage", relatedPage).list();
		}
		LOGGER.info("<<<<<< ContentConfigDAOImpl.populateExistingRefQtn");
		return list;
	}
	/**
	 * 
	 */
	@Transactional(propagation = Propagation.REQUIRED)
	public List<UserProfileDTO> populateExistingUserProfileWithId() throws DAOException {
		LOGGER.info(">>>>>> ContentConfigDAOImpl.populateExistingUserProfileWithId Starts");
		List<UserProfileDTO> list = sessionFactory.getCurrentSession().getNamedQuery("fetchAllUserProfileWithId").list();
		LOGGER.info("<<<<<< ContentConfigDAOImpl.populateExistingUserProfileWithId Ends");
		return list;
	}
	/**
	 * 
	 */
	@Transactional(propagation = Propagation.REQUIRED)
	public String saveRefQtn(JctQuestion question) throws DAOException {
		LOGGER.info(">>>>>> ContentConfigDAOImpl.saveRefQtn");
		return save(question);
	}
	/**
	 * 
	 */
	@Transactional(propagation = Propagation.REQUIRED)
	public String updateRefQtn(JctQuestion question) throws DAOException {
		LOGGER.info(">>>>>> ContentConfigDAOImpl.updateRefQtn");
		return update(question);
	}
	/**
	 * 
	 */
	@Transactional(propagation = Propagation.REQUIRED)
	public JctQuestion fetchRefQtn(int pkId) throws DAOException {
		LOGGER.info(">>>>>> ContentConfigDAOImpl.fetchRefQtn");
		JctQuestion qtn = null;
		try {
			qtn = (JctQuestion) sessionFactory.getCurrentSession().getNamedQuery("fetchRefQtn")
					.setParameter("pkId", pkId).list().get(0);
		} catch (java.lang.IndexOutOfBoundsException exp) {
			LOGGER.info("JCT QUESTION NOT FETCHED WITH PRIMARY KEY: "+pkId+" AND SOFT DELETE: 0");
		}
		return qtn;
	}
	/**
	 * 
	 */
	@Transactional(propagation = Propagation.REQUIRED)
	public String deleteRefQtn(JctQuestion question) throws DAOException {
		LOGGER.info(">>>>>> ContentConfigDAOImpl.deleteRefQtn");
		String returnStr = update(question);			
		List<QuestionearDTO> totalSubQtn = sessionFactory.getCurrentSession().
				getNamedQuery("fetchTotalSubQtn").
				setParameter("mainQtnDesc", question.getJctQuestionsDesc()).
				setParameter("profileId", question.getJctUserProfile().getJctUserProfile())
				.setParameter("relatedPage", "BS").list();
		if (totalSubQtn.size() != 0) {
			int subQtnOrder = question.getJctSubQuestionsOrder();
			List<QuestionearDTO> existingList = sessionFactory.getCurrentSession().
					getNamedQuery("fetchQuestionBySubOrder").
					setParameter("subQtnOrder", question.getJctSubQuestionsOrder()).
					setParameter("mainQtnOrder", question.getJctQuestionsOrder()).
					setParameter("mainQtnDesc", question.getJctQuestionsDesc()).
					setParameter("profileId", question.getJctUserProfile().getJctUserProfile())
					.setParameter("relatedPage", "BS").list();
			
			if (existingList.size() != 0) {
				for (int i = 0; i < existingList.size(); i++) {
					JctQuestion existingQtn = (JctQuestion) sessionFactory.getCurrentSession().
							getNamedQuery("fetchRefQtn")
							.setParameter("pkId", existingList.get(i).
							getJctQuestionsId()).list().get(0);
					existingQtn.setJctSubQuestionsOrder(subQtnOrder + i );
					String returnStr1 = update(existingQtn);	
				}
			}
		} else {
			int qtnOrder = question.getJctQuestionsOrder();
			List<QuestionearDTO> existingMainQtnList = sessionFactory.getCurrentSession().
					getNamedQuery("fetchMainQtnByOrder").
					setParameter("mainQtnOrder", question.getJctQuestionsOrder()).
					setParameter("profileId", question.getJctUserProfile().getJctUserProfile())
					.setParameter("relatedPage", "BS").list();
			
			if (existingMainQtnList.size() != 0) {
				for (int i = 0; i < existingMainQtnList.size(); i++) {
					JctQuestion existingMainQtn = (JctQuestion) sessionFactory.getCurrentSession().
							getNamedQuery("fetchRefQtn").
							setParameter("pkId", existingMainQtnList.get(i).getJctQuestionsId()).
							list().get(0);
					existingMainQtn.setJctQuestionsOrder(qtnOrder + i );
					String returnStr1 = update(existingMainQtn);	
				}
			}
		}
		return returnStr;
	}
	/**
	 * 
	 */
	@Transactional(propagation = Propagation.REQUIRED)
	public String saveActionPlan(JctQuestion question) throws DAOException {
		LOGGER.info(">>>>>> ContentConfigDAOImpl.actionPlan");
		return save(question);
	}
	/**
	 * 	
	 */
	@Transactional(propagation = Propagation.REQUIRED)
	public String updateActionPlan(JctQuestion question) throws DAOException {
		LOGGER.info(">>>>>> ContentConfigDAOImpl.updateActionPlan");
		return update(question);
	}
	/**
	 * 
	 */
	@Transactional(propagation = Propagation.REQUIRED)
	public String deleteActionPlan(JctQuestion question) throws DAOException {
		LOGGER.info(">>>>>> ContentConfigDAOImpl.deleteActionPlan");
		String returnStr = update(question);			
		List<QuestionearDTO> totalSubQtn = sessionFactory.getCurrentSession().
				getNamedQuery("fetchTotalSubQtn").
				setParameter("mainQtnDesc", question.getJctQuestionsDesc()).
				setParameter("profileId", question.getJctUserProfile().getJctUserProfile())
				.setParameter("relatedPage", "AS").list();
		if (totalSubQtn.size() != 0) {
			int subQtnOrder = question.getJctSubQuestionsOrder();
			List<QuestionearDTO> existingList = sessionFactory.getCurrentSession().
					getNamedQuery("fetchQuestionBySubOrder").
					setParameter("subQtnOrder", question.getJctSubQuestionsOrder()).
					setParameter("mainQtnOrder", question.getJctQuestionsOrder()).
					setParameter("mainQtnDesc", question.getJctQuestionsDesc()).
					setParameter("profileId", question.getJctUserProfile().getJctUserProfile())
					.setParameter("relatedPage", "AS").list();
			
			if (existingList.size() != 0) {
				for (int i = 0; i < existingList.size(); i++) {
					JctQuestion existingQtn = (JctQuestion) sessionFactory.getCurrentSession().
							getNamedQuery("fetchRefQtn")
							.setParameter("pkId", existingList.get(i).
							getJctQuestionsId()).list().get(0);
					existingQtn.setJctSubQuestionsOrder(subQtnOrder + i );
					String returnStr1 = update(existingQtn);	
				}
			}
		} else {
			int qtnOrder = question.getJctQuestionsOrder();
			List<QuestionearDTO> existingMainQtnList = sessionFactory.getCurrentSession().
					getNamedQuery("fetchMainQtnByOrder").
					setParameter("mainQtnOrder", question.getJctQuestionsOrder()).
					setParameter("profileId", question.getJctUserProfile().getJctUserProfile())
					.setParameter("relatedPage", "AS").list();
			
			if (existingMainQtnList.size() != 0) {
				for (int i = 0; i < existingMainQtnList.size(); i++) {
					JctQuestion existingMainQtn = (JctQuestion) sessionFactory.getCurrentSession().
							getNamedQuery("fetchRefQtn").
							setParameter("pkId", existingMainQtnList.get(i).getJctQuestionsId()).
							list().get(0);
					existingMainQtn.setJctQuestionsOrder(qtnOrder + i );
					String returnStr1 = update(existingMainQtn);	
				}
			}
		}	
		return returnStr;
	}
	/**
	 * 	
	 */
	@Transactional(propagation = Propagation.REQUIRED)
	public List<FunctionDTO> populateExistingFunctionGroup()
			throws DAOException {
		LOGGER.info(">>>>>> ContentConfigDAOImpl.populateExistingFunctionGroup");
		List<FunctionDTO> list = sessionFactory.getCurrentSession().getNamedQuery("fetchFunction").list();
		LOGGER.info("<<<<<< ContentConfigDAOImpl.populateExistingFunctionGroup");
		return list;
	}
	/**
	 * 
	 */
	@Transactional(propagation = Propagation.REQUIRED)
	public List<LevelDTO> populateExistingJobLevel()
			throws DAOException {
		LOGGER.info(">>>>>> ContentConfigDAOImpl.populateExistingJobLevel");
		List<LevelDTO> list = sessionFactory.getCurrentSession().getNamedQuery("fetchJobLevel").list();
		LOGGER.info("<<<<<< ContentConfigDAOImpl.populateExistingJobLevel");
		return list;
	}
	/**
	 * 
	 */
	@Transactional(propagation = Propagation.REQUIRED)
	public String saveFunctionGroup(JctFunction functionGrp) throws DAOException {
		LOGGER.info(">>>>>> ContentConfigDAOImpl.functionGrp");
		return save(functionGrp);
	}
	/**
	 * 
	 */
	@Transactional(propagation = Propagation.REQUIRED)
	public String saveJobLevel(JctLevel jobLevel) throws DAOException {
		LOGGER.info(">>>>>> ContentConfigDAOImpl.saveJobLevel");
		return save(jobLevel);
	}
	/**
	 * 
	 */
	@Transactional(propagation = Propagation.REQUIRED)
	public List<RegionDTO> populateExistingRegion()
			throws DAOException {
		LOGGER.info(">>>>>> ContentConfigDAOImpl.populateExistingRegion");
		List<RegionDTO> list = sessionFactory.getCurrentSession().getNamedQuery("fetchRegion").list();
		LOGGER.info("<<<<<< ContentConfigDAOImpl.populateExistingRegion");
		return list;
	}
	/**
	 * 
	 */
	@Transactional(propagation = Propagation.REQUIRED)
	public String saveRegion(JctRegion regionName) throws DAOException {
		LOGGER.info(">>>>>> ContentConfigDAOImpl.saveRegion");
		return save(regionName);
	}
	/**
	 * 
	 */
	@Transactional(propagation = Propagation.REQUIRED)
	public JctRegion fetchRegion(int pkId) throws DAOException {
		LOGGER.info(">>>>>> ContentConfigDAOImpl.fetchRegion");
		JctRegion region = null;
		try {
			region = (JctRegion) sessionFactory.getCurrentSession().getNamedQuery("fetchRegionById")
					.setParameter("pkId", pkId).list().get(0);
		}  catch (java.lang.IndexOutOfBoundsException exp) {
			LOGGER.info("REGION NOT FETCHED WITH PRIMARY KEY: "+pkId+" AND SOFT DELETE: 0");
		}
		return region;
	}
	/**
	 * 
	 */
	@Transactional(propagation = Propagation.REQUIRED)
	public String updateRegion(JctRegion region) throws DAOException {
		LOGGER.info(">>>>>> ContentConfigDAOImpl.updateRegion");
		return update(region);
	}
	/**
	 * 
	 */
	@Transactional(propagation = Propagation.REQUIRED)
	public String deleteRegion(JctRegion region) throws DAOException {
		LOGGER.info(">>>>>> ContentConfigDAOImpl.deleteRegion");
		//return update(region);
		String returnStr = update(region);		
		int regionOrder = region.getJctRegionOrder();	
		List<RegionDTO> existingList = sessionFactory.getCurrentSession().
				getNamedQuery("fetchRegionBybOrder").
				setParameter("regionOrder", region.getJctRegionOrder()).list();
			
		if (existingList.size() != 0) {
			for (int i = 0; i < existingList.size(); i++) {
				JctRegion existingVal = (JctRegion) sessionFactory.getCurrentSession().
						getNamedQuery("fetchRegionById")
						.setParameter("pkId", existingList.get(i).
						getJctRegionId()).list().get(0);
				existingVal.setJctRegionOrder(regionOrder + i );
				String returnStr1 = update(existingVal);	
			}
		}		
		return returnStr;
	}
	/**
	 * 
	 */
	@Transactional(propagation = Propagation.REQUIRED)
	public List<JobAttributeDTO> populateExistingMapping(Integer profileId, String relatedPage)
			throws DAOException {
		LOGGER.info(">>>>>> ContentConfigDAOImpl.populateExistingMapping");
		//List<JobAttributeDTO> list = sessionFactory.getCurrentSession().getNamedQuery("fetchAllJobAttribute").list();	
		List<JobAttributeDTO> list = null;
		if(0 == profileId) {
			list = sessionFactory.getCurrentSession().getNamedQuery("fetchAllJobAttribute").
					setParameter("relatedPage", relatedPage).list();
		} else {
			list = sessionFactory.getCurrentSession().getNamedQuery("fetchAllAttributeByProfileId").
					setParameter("profileId", profileId).setParameter("relatedPage", relatedPage).list();
		}
		
		LOGGER.info("<<<<<< ContentConfigDAOImpl.populateExistingMapping");
		return list;
	}
	/**
	 * 
	 */
	@Transactional(propagation = Propagation.REQUIRED)
	public String saveStrength(JctJobAttribute jobAttribute) throws DAOException {
		LOGGER.info(">>>>>> ContentConfigDAOImpl.saveStrength");
		return save(jobAttribute);
	}
	/**
	 * 
	 */
	@Transactional(propagation = Propagation.REQUIRED)
	public String saveValue(JctJobAttribute jobAttribute) throws DAOException {
		LOGGER.info(">>>>>> ContentConfigDAOImpl.saveValue");
		return save(jobAttribute);
	}
	/**
	 * 
	 */
	@Transactional(propagation = Propagation.REQUIRED)
	public String savePassion(JctJobAttribute jobAttribute) throws DAOException {
		LOGGER.info(">>>>>> ContentConfigDAOImpl.savePassion");
		return save(jobAttribute);
	}
	/**
	 * 
	 */
	@Transactional(propagation = Propagation.REQUIRED)
	public String updateStrength(JctJobAttribute jobAttribute) throws DAOException {
		LOGGER.info(">>>>>> ContentConfigDAOImpl.updateStrength");
		return update(jobAttribute);
	}
	/**
	 * 
	 */
	@Transactional(propagation = Propagation.REQUIRED)
	public String updateValue(JctJobAttribute jobAttribute) throws DAOException {
		LOGGER.info(">>>>>> ContentConfigDAOImpl.updateValue");
		return update(jobAttribute);
	}
	/**
	 * 
	 */
	@Transactional(propagation = Propagation.REQUIRED)
	public String updatePassion(JctJobAttribute jobAttribute) throws DAOException {
		LOGGER.info(">>>>>> ContentConfigDAOImpl.updatePassion");
		return update(jobAttribute);
	}
	/**
	 * 
	 */
	@Transactional(propagation = Propagation.REQUIRED)
	public JctJobAttribute fetchMapping(int pkId) throws DAOException {
		LOGGER.info(">>>>>> ContentConfigDAOImpl.fetchMapping");
		JctJobAttribute attr = null;
		try {
			attr = (JctJobAttribute) sessionFactory.getCurrentSession().getNamedQuery("fetchJobAttributeById")
					.setParameter("pkId", pkId).list().get(0);
		} catch (java.lang.IndexOutOfBoundsException exp) {
			LOGGER.info("JCT ATTRIBUTE NOT FETCHED WITH PRIMARY KEY: "+pkId+" AND SOFT DELETE: 0");
		}
		return attr;
	}
	/**
	 * 
	 */
	@Transactional(propagation = Propagation.REQUIRED)
	public String deleteStrength(JctJobAttribute jobAttribute) throws DAOException {
		LOGGER.info(">>>>>> ContentConfigDAOImpl.deleteStrength");
		String returnStr = update(jobAttribute);		
		int attrOrder = jobAttribute.getJctJobAttributeOrder();	
		List<JobAttributeDTO> existingList = sessionFactory.getCurrentSession().
				getNamedQuery("fetchAttrBybOrder").
				setParameter("attrOrder", jobAttribute.getJctJobAttributeOrder()).
				setParameter("profileId", jobAttribute.getJctUserProfile().getJctUserProfile())
				.setParameter("attrCode", "STR").list();
			
		if (existingList.size() != 0) {
			for (int i = 0; i < existingList.size(); i++) {
				JctJobAttribute existingQtn = (JctJobAttribute) sessionFactory.getCurrentSession().
						getNamedQuery("fetchJobAttributeById")
						.setParameter("pkId", existingList.get(i).
						getJctJobAttributeId()).list().get(0);
				existingQtn.setJctJobAttributeOrder(attrOrder + i );
				String returnStr1 = update(existingQtn);	
			}
		}		
		return returnStr;
	}
	/**
	 * 
	 */
	@Transactional(propagation = Propagation.REQUIRED)
	public String deleteValue(JctJobAttribute jobAttribute) throws DAOException {
		LOGGER.info(">>>>>> ContentConfigDAOImpl.deleteValue");
		String returnStr = update(jobAttribute);		
		int attrOrder = jobAttribute.getJctJobAttributeOrder();	
		List<JobAttributeDTO> existingList = sessionFactory.getCurrentSession().
				getNamedQuery("fetchAttrBybOrder").
				setParameter("attrOrder", jobAttribute.getJctJobAttributeOrder()).
				setParameter("profileId", jobAttribute.getJctUserProfile().getJctUserProfile())
				.setParameter("attrCode", "VAL").list();
			
		if (existingList.size() != 0) {
			for (int i = 0; i < existingList.size(); i++) {
				JctJobAttribute existingQtn = (JctJobAttribute) sessionFactory.getCurrentSession().
						getNamedQuery("fetchJobAttributeById")
						.setParameter("pkId", existingList.get(i).
						getJctJobAttributeId()).list().get(0);
				existingQtn.setJctJobAttributeOrder(attrOrder + i );
				String returnStr1 = update(existingQtn);	
			}
		}
		return returnStr;
	}
	/**
	 * 
	 */
	@Transactional(propagation = Propagation.REQUIRED)
	public String deletePassion(JctJobAttribute jobAttribute) throws DAOException {
		LOGGER.info(">>>>>> ContentConfigDAOImpl.deletePassion");
		String returnStr = update(jobAttribute);		
		int attrOrder = jobAttribute.getJctJobAttributeOrder();	
		List<JobAttributeDTO> existingList = sessionFactory.getCurrentSession().
				getNamedQuery("fetchAttrBybOrder").
				setParameter("attrOrder", jobAttribute.getJctJobAttributeOrder()).
				setParameter("profileId", jobAttribute.getJctUserProfile().getJctUserProfile())
				.setParameter("attrCode", "PAS").list();
			
		if (existingList.size() != 0) {
			for (int i = 0; i < existingList.size(); i++) {
				JctJobAttribute existingQtn = (JctJobAttribute) sessionFactory.getCurrentSession().
						getNamedQuery("fetchJobAttributeById")
						.setParameter("pkId", existingList.get(i).
						getJctJobAttributeId()).list().get(0);
				existingQtn.setJctJobAttributeOrder(attrOrder + i );
				String returnStr1 = update(existingQtn);	
			}
		}
		return returnStr;
	}
	/**
	 * 
	 */
	@Transactional(propagation = Propagation.REQUIRED)
	public String saveInstruction(JctInstructionBar instruction) throws DAOException {
		LOGGER.info(">>>>>> ContentConfigDAOImpl.saveInstruction");
		return save(instruction);
	}
	/**
	 * 
	 */
	@Transactional(propagation = Propagation.REQUIRED)
	public JctInstructionBar populateExistingInstruction(int userProfileId, String relatedPage)
			throws DAOException {
		LOGGER.info(">>>>>> ContentConfigDAOImpl.populateExistingInstruction");
		JctInstructionBar object = null;
		List<JctInstructionBar> listObj = sessionFactory.getCurrentSession()
				.getNamedQuery("fetchInstructionById")
				.setParameter("userProfileId", userProfileId)
				.setParameter("relatedPage", relatedPage).list();
		if (listObj.size() > 0) {
			object = (JctInstructionBar) listObj.get(0);
		} 
		LOGGER.info("<<<<<<< ContentConfigDAOImpl.populateExistingInstruction");
		return object;				
	}
	/**
	 * 
	 */
	@Transactional(propagation=Propagation.REQUIRED)
	public String validateRefQtn(String refQtnDesc, int userProfileId) throws DAOException {
		LOGGER.info(">>>>>> ContentConfigDAOImpl.validateRefQtn");
		List<JctQuestion> list = null;
		String status = "failure";
		try{
			list = sessionFactory.getCurrentSession().getNamedQuery("fetchRefQtnByDesc")
					.setParameter("refQtnDesc", refQtnDesc).setParameter("userProfileId", userProfileId).list();
			if(list.size() == 0) {
				status = "failure";
			} else {
				status = "success";
			}			
		}catch(Exception ez){
			status = "failure";
			LOGGER.error(ez.getLocalizedMessage());
		}
		LOGGER.info("<<<<<<< ContentConfigDAOImpl.validateRefQtn");
		return status;
	}
	/**
	 * 
	 */
	@Transactional(propagation=Propagation.REQUIRED)
	public String validateActionPlan(String refQtnDesc, String subQtnDesc, int userProfileId, String code, String action) throws DAOException {
		LOGGER.info(">>>>>> ContentConfigDAOImpl.validateActionPlan");
		List<JctQuestion> list = null;
		List<QuestionearDTO> list1 = null;
		String status = "failure";
		try{
			list = sessionFactory.getCurrentSession().getNamedQuery("fetchActionPlanByDesc")
					.setParameter("refQtnDesc", refQtnDesc).setParameter("subQtnDesc", subQtnDesc).setParameter("userProfileId", userProfileId).setParameter("code", code).list();
			list1 = sessionFactory.getCurrentSession().getNamedQuery("fetchAllExistingQtn")
					.setParameter("code", code).setParameter("userProfileId", userProfileId).list();
			if(list.size() == 0) {
				//status = "failure";
				if(action.equalsIgnoreCase("ADD")){
					if(list1.size() >= 6){
						status = "maximum";
					} else {
						status = "failure";
					}	
				} else {
					status = "failure";
				}	
			} else {
				status = "success";
			}			
		}catch(Exception ez){
			status = "failure";
			LOGGER.error(ez.getLocalizedMessage());
		}
		LOGGER.info("<<<<<<< ContentConfigDAOImpl.validateActionPlan");
		return status;
	}
	/**
	 * 
	 */
	@Transactional(propagation=Propagation.REQUIRED)
	public String validateFuncGrp(String functionGrp) throws DAOException {
		LOGGER.info(">>>>>> ContentConfigDAOImpl.validateFuncGrp");
		List<JctFunction> list = null;
		String status = "failure";
		try{
			list = sessionFactory.getCurrentSession().getNamedQuery("fetchFuncGrpByDesc")
					.setParameter("functionGrp", functionGrp).list();
			if(list.size() == 0) {
				status = "failure";
			} else {
				status = "success";
			}			
		}catch(Exception ez){
			status = "failure";
			LOGGER.error(ez.getLocalizedMessage());
		}
		LOGGER.info("<<<<<<< ContentConfigDAOImpl.validateFuncGrp");
		return status;
	}
	/**
	 * 
	 */
	@Transactional(propagation=Propagation.REQUIRED)
	public String validateJobLevel(String jobLevel) throws DAOException {
		LOGGER.info(">>>>>> ContentConfigDAOImpl.validateJobLevel");
		List<JctLevel> list = null;
		String status = "failure";
		try{
			list = sessionFactory.getCurrentSession().getNamedQuery("fetchJobLevelByDesc")
					.setParameter("jobLevel", jobLevel).list();
			if(list.size() == 0) {
				status = "failure";
			} else {
				status = "success";
			}			
		}catch(Exception ez){
			status = "failure";
			LOGGER.error(ez.getLocalizedMessage());
		}
		LOGGER.info("<<<<<<< ContentConfigDAOImpl.validateJobLevel");
		return status;
	}
	/**
	 * 
	 */
	@Transactional(propagation=Propagation.REQUIRED)
	public String validateAttribute(String attrDescText, int userProfileId, String attrCode, String action) throws DAOException {
		LOGGER.info(">>>>>> ContentConfigDAOImpl.validateAttribute");
		List<JctJobAttribute> list = null;
		List<JctJobAttribute> list1 = null;
		String status = "failure";
		try{
			list = sessionFactory.getCurrentSession().getNamedQuery("fetchJobAttributeByDesc")
					.setParameter("attrDescText", attrDescText).setParameter("userProfileId", userProfileId).setParameter("attrCode", attrCode).list();
			list1 = sessionFactory.getCurrentSession().getNamedQuery("fetchExistingJobAttribute")
					.setParameter("userProfileId", userProfileId).setParameter("attrCode", attrCode).list();
			if(list.size() == 0) {
				if(action.equalsIgnoreCase("ADD")){
					if(list1.size() >= 24){
						status = "maximum";
					} else {
						status = "failure";
					}	
				} else {
					status = "failure";
				}							
			} else {
				status = "success";
			}			
		}catch(Exception ez){
			status = "failure";
			LOGGER.error(ez.getLocalizedMessage());
		}
		LOGGER.info("<<<<<<< ContentConfigDAOImpl.validateAttribute");
		return status;
	}
	/**
	 * 
	 */
	@Transactional(propagation=Propagation.REQUIRED)
	public String validateRegion(String regionName) throws DAOException {
		LOGGER.info(">>>>>> ContentConfigDAOImpl.validateRegion");
		List<JctRegion> list = null;
		String status = "failure";
		try{
			list = sessionFactory.getCurrentSession().getNamedQuery("fetchRegionByDesc")
					.setParameter("regionName", regionName).list();
			if(list.size() == 0) {
				status = "failure";
			} else {
				status = "success";
			}			
		}catch(Exception ez){
			status = "failure";
			LOGGER.error(ez.getLocalizedMessage());
		}
		LOGGER.info("<<<<<<< ContentConfigDAOImpl.validateRegion");
		return status;
	}
	/**
	 * 
	 */
	@Transactional(propagation = Propagation.REQUIRED)
	public JctInstructionBar fetchInstructionData(Integer pk)
			throws DAOException {
		LOGGER.info(">>>>>> ContentConfigDAOImpl.JctInstructionBar");				
		return (JctInstructionBar) sessionFactory.getCurrentSession().
				getNamedQuery("fetchInstructionByPk")
				.setParameter("pk", pk).list().get(0);		
	}
	/**
	 * 
	 */
	@Transactional(propagation = Propagation.REQUIRED)
	public String updateInstruction(JctInstructionBar instruction) throws DAOException {
		LOGGER.info(">>>>>> ContentConfigDAOImpl.updateInstruction");
		return update(instruction);
	}
	
	/**
	 * 
	 */
	@Transactional(propagation = Propagation.REQUIRED)
	public JctQuestion fetchRefQtnByQtnOrder(int profileId, int qtnOrder, String page) throws DAOException {
		LOGGER.info(">>>>>> ContentConfigDAOImpl.fetchRefQtn");
		return (JctQuestion) sessionFactory.getCurrentSession().getNamedQuery("fetchRefQtnByOrder")
				.setParameter("profileId", profileId).setParameter("qtnOrder", qtnOrder).setParameter("page", page).list().get(0);
	}
	
	/**
	 * 
	 */
	@Transactional(propagation = Propagation.REQUIRED)
	public JctJobAttribute fetchMappingByOrder(int profileId, int attrOrder, String page) throws DAOException {
		LOGGER.info(">>>>>> ContentConfigDAOImpl.fetchMappingByOrder");
		return (JctJobAttribute) sessionFactory.getCurrentSession().getNamedQuery("fetchJobAttributeByOrder")
				.setParameter("profileId", profileId).setParameter("attrOrder", attrOrder).setParameter("page", page).list().get(0);
	}
	/**
	 * 
	 */
	@Transactional(propagation = Propagation.REQUIRED)
	public Integer generateOrderAttr(String attrCode, Integer profileId) 
			throws DAOException {
		LOGGER.info(">>>>>> ContentConfigDAOImpl.generateOrder");
		Integer order = null;
		Integer value = (Integer) sessionFactory.getCurrentSession().getNamedQuery("fetchOrder")
				.setParameter("profileId", profileId).setParameter("attrCode", attrCode).uniqueResult();
		if( null == value) {
			order = 1;
		} else {
			order = value + 1;
		}
		return order;
	}
	/**
	 * 
	 */
	@Transactional(propagation = Propagation.REQUIRED)
	public Integer generateOrderFunctionGrp() throws DAOException {
		LOGGER.info(">>>>>> ContentConfigDAOImpl.generateOrderFunctionGrp");
		Integer order = null;
		Integer value = (Integer) sessionFactory.getCurrentSession().getNamedQuery("fetchOrderFG").uniqueResult();
		if( null == value) {
			order = 1;
		} else {
			order = value + 1;
		}
		return order;
	}
	/**
	 * 
	 */
	@Transactional(propagation = Propagation.REQUIRED)
	public Integer generateOrderLevel() throws DAOException {
		LOGGER.info(">>>>>> ContentConfigDAOImpl.generateOrderLevel");
		Integer order = null;
		Integer value = (Integer) sessionFactory.getCurrentSession().getNamedQuery("fetchOrderLevel").uniqueResult();
		if( null == value) {
			order = 1;
		} else {
			order = value + 1;
		}
		return order;
	}
	/**
	 * 
	 */
	@Transactional(propagation = Propagation.REQUIRED)
	public Integer generateOrderQtn(String qtnCode, Integer profileId,
			String qtnDesc) throws DAOException {
		LOGGER.info(">>>>>> ContentConfigDAOImpl.generateOrderQtn");
		Integer order = null;
		Integer value = (Integer) sessionFactory.getCurrentSession().getNamedQuery("fetchOrderQtn")
				.setParameter("profileId", profileId).setParameter("qtnCode", qtnCode).setParameter("qtnDesc", qtnDesc).uniqueResult();
		if( null == value) {
			order = 1;
		} else {
			order = value + 1;
		}
		return order;
	}	
	
	/**
	 * 
	 */
	@Transactional(propagation = Propagation.REQUIRED)
	public List<QuestionearDTO> populateMainQtn(Integer profileId, String relatedPage)
			throws DAOException {
		LOGGER.info(">>>>>> ContentConfigDAOImpl.populateMainQtn");
		List<QuestionearDTO> list = null;	
		list = sessionFactory.getCurrentSession().getNamedQuery("fetchMainQtnList").
				setParameter("profileId", profileId).setParameter("relatedPage", relatedPage).list();		
		LOGGER.info("<<<<<< ContentConfigDAOImpl.populateMainQtn");
		return list;
	}
	
	/**
	 * 
	 */
	@Transactional(propagation = Propagation.REQUIRED)
	public List<QuestionearDTO> populateSubQtn(Integer profileId, String relatedPage, String mainQtn)
			throws DAOException {
		LOGGER.info(">>>>>> ContentConfigDAOImpl.populateSubQtn");
		List<QuestionearDTO> list = null;	
		list = sessionFactory.getCurrentSession().getNamedQuery("fetchSubQtnList").
				setParameter("profileId", profileId).setParameter("relatedPage", relatedPage).setParameter("mainQtn", mainQtn).list();		
		LOGGER.info("<<<<<< ContentConfigDAOImpl.populateSubQtn");
		return list;
	}
	/**
	 * 
	 */
	@Transactional(propagation = Propagation.REQUIRED)
	public String saveMainQtnOrder(String mainQtnBuilder,
			Integer profileId, String relatedPage) throws DAOException {
		LOGGER.info(">>>>>> ContentConfigDAOImpl.saveMainQtnOrder");
		String returnStr = "";
		List<JctQuestion> list = null;
		String[] splitQtnDesc = mainQtnBuilder.split("~~");
		int sunQtnLength = splitQtnDesc.length;	
		for (int i = 0; i < sunQtnLength; i++) {
			
			List<QuestionearDTO> existingList = sessionFactory.getCurrentSession().getNamedQuery("mainQtnList").
					setParameter("profileId", profileId).setParameter("relatedPage", relatedPage).
					setParameter("mainQtnDesc", splitQtnDesc[i]).list();
			
			if (existingList.size() != 0) {
				for (int j = 0; j < existingList.size(); j++) {
					JctQuestion existingQtn = (JctQuestion) sessionFactory.getCurrentSession().getNamedQuery("fetchToUpdateMainOrder")
							.setParameter("profileId", profileId).setParameter("relatedPage", relatedPage).
							setParameter("mainQtnDesc", splitQtnDesc[i]).list().get(j);
					existingQtn.setJctQuestionsOrder(i+1);
					returnStr = update(existingQtn);
				}
			}
		}
		LOGGER.info("<<<<<< ContentConfigDAOImpl.saveMainQtnOrder");
		return returnStr;
	}
	/**
	 * 
	 */
	@Transactional(propagation = Propagation.REQUIRED)
	public String saveSubQtnOrder(String subQtnBuilder, Integer profileId,
			String relatedPage, String mainQtn) throws DAOException {
		LOGGER.info(">>>>>> ContentConfigDAOImpl.saveSubQtnOrder");
		String returnStr = "";		
		String[] splitQtnDesc = subQtnBuilder.split("~~");
		int sunQtnLength = splitQtnDesc.length;	
		for (int i = 0; i < sunQtnLength; i++) {
			JctQuestion existingQtn = (JctQuestion) sessionFactory.getCurrentSession().getNamedQuery("fetchToUpdateSubOrder")
					.setParameter("profileId", profileId).setParameter("relatedPage", relatedPage)
					.setParameter("mainQtn", mainQtn).setParameter("subQtnDesc", splitQtnDesc[i]).list().get(0);
			existingQtn.setJctSubQuestionsOrder(i+1);
			returnStr = update(existingQtn);
		}
		LOGGER.info("<<<<<< ContentConfigDAOImpl.saveSubQtnOrder");
		return returnStr;
	}
	/**
	 * 
	 */
	@Transactional(propagation = Propagation.REQUIRED)
	public String saveFuncGrpOrder(String builder) throws DAOException {
		LOGGER.info(">>>>>> ContentConfigDAOImpl.saveFuncGrpOrder");
		String returnStr = "";		
		String[] splitDesc = builder.split("~~");
		int sunQtnLength = splitDesc.length;	
		for (int i = 0; i < sunQtnLength; i++) {
			JctFunction existingVal = (JctFunction) sessionFactory.getCurrentSession().getNamedQuery("fetchToUpdateFG")
					.setParameter("functionGroup", splitDesc[i]).list().get(0);
			existingVal.setJctFunctionsOrder(i+1);
			returnStr = update(existingVal);
		}
		LOGGER.info("<<<<<< ContentConfigDAOImpl.saveFuncGrpOrder");
		return returnStr;
	}
	
	/**
	 * 	
	 */
	@Transactional(propagation = Propagation.REQUIRED)
	public List<FunctionDTO> populateFunctionGroupByOrder()
			throws DAOException {
		LOGGER.info(">>>>>> ContentConfigDAOImpl.populateFunctionGroupByOrder");
		List<FunctionDTO> list = sessionFactory.getCurrentSession().getNamedQuery("fetchFunctionByOrder").list();
		LOGGER.info("<<<<<< ContentConfigDAOImpl.populateFunctionGroupByOrder");
		return list;
	}
	
	/**
	 * 
	 */
	@Transactional(propagation = Propagation.REQUIRED)
	public String saveJobLevelOrder(String builder) throws DAOException {
		LOGGER.info(">>>>>> ContentConfigDAOImpl.saveJobLevelOrder");
		String returnStr = "";		
		String[] splitDesc = builder.split("~~");
		int sunQtnLength = splitDesc.length;	
		for (int i = 0; i < sunQtnLength; i++) {
			JctLevel existingVal = (JctLevel) sessionFactory.getCurrentSession().getNamedQuery("fetchToUpdateJL")
					.setParameter("jobLevel", splitDesc[i]).list().get(0);
			existingVal.setJctLevelsOrder(i+1);
			returnStr = update(existingVal);
		}
		LOGGER.info("<<<<<< ContentConfigDAOImpl.saveJobLevelOrder");
		return returnStr;
	}
	/**
	 * 
	 */
	@Transactional(propagation = Propagation.REQUIRED)
	public List<LevelDTO> populateJobLevelByOrder()
			throws DAOException {
		LOGGER.info(">>>>>> ContentConfigDAOImpl.populateJobLevelByOrder");
		List<LevelDTO> list = sessionFactory.getCurrentSession().getNamedQuery("fetchJobLevelByOrder").list();
		LOGGER.info("<<<<<< ContentConfigDAOImpl.populateJobLevelByOrder");
		return list;
	}
	/**
	 * 
	 */
	@Transactional(propagation = Propagation.REQUIRED)
	public List<JobAttributeDTO> populateMappingListByOrder(Integer profileId, String arrtCode)
			throws DAOException {
		LOGGER.info(">>>>>> ContentConfigDAOImpl.populateMappingListByOrder");
		List<JobAttributeDTO> list = sessionFactory.getCurrentSession().
				getNamedQuery("fetchAttributeByOrder").
				setParameter("profileId", profileId).
				setParameter("arrtCode", arrtCode).list();
		LOGGER.info("<<<<<< ContentConfigDAOImpl.populateMappingListByOrder");
		return list;
	}
	/**
	 * 
	 */
	@Transactional(propagation = Propagation.REQUIRED)
	public String saveAttributeOrder(String builder, Integer profileId,
			String attrCode) throws DAOException {
		LOGGER.info(">>>>>> ContentConfigDAOImpl.saveAttributeOrder");
		String returnStr = "";		
		String[] splitDesc = builder.split("~~");
		int sunQtnLength = splitDesc.length;	
		for (int i = 0; i < sunQtnLength; i++) {
			JctJobAttribute existingVal = (JctJobAttribute) sessionFactory.
					getCurrentSession().getNamedQuery("fetchToUpdateAttr").
					setParameter("profileId", profileId).
					setParameter("attrCode", attrCode).
					setParameter("attrName", splitDesc[i]).list().get(0);
			existingVal.setJctJobAttributeOrder(i+1);
			returnStr = update(existingVal);
		}
		LOGGER.info("<<<<<< ContentConfigDAOImpl.saveAttributeOrder");
		return returnStr;
	}
	
	/**
	 * 
	 */
	@Transactional(propagation = Propagation.REQUIRED)
	public Integer generateOrderRegion() throws DAOException {
		LOGGER.info(">>>>>> ContentConfigDAOImpl.generateOrderRegion");
		// TODO Auto-generated method stub
		Integer order = null;
		Integer value = (Integer) sessionFactory.getCurrentSession().getNamedQuery("fetchOrderRG").uniqueResult();
		if( null == value) {
			order = 1;
		} else {
			order = value + 1;
		}
		return order;
	}
	/**
	 * 
	 */
	@Transactional(propagation = Propagation.REQUIRED)
	public String saveRegionOrder(String builder) throws DAOException {
		LOGGER.info(">>>>>> ContentConfigDAOImpl.saveRegionOrder");
		String returnStr = "";		
		String[] splitDesc = builder.split("~~");
		int sunQtnLength = splitDesc.length;	
		String regionName = "";
		for (int i = 0; i < sunQtnLength; i++) {
			regionName = splitDesc[i];
			regionName = regionName.replace("&amp;", "&");
			JctRegion existingVal = (JctRegion) sessionFactory.getCurrentSession().getNamedQuery("fetchToUpdateRG")
					.setParameter("regionName", regionName).list().get(0);
			existingVal.setJctRegionOrder(i+1);
			returnStr = update(existingVal);
		}
		LOGGER.info("<<<<<< ContentConfigDAOImpl.saveRegionOrder");
		return returnStr;
	}
	
	/******* THIS IS FOR JCT PUBLIC VERSION ********/
	/**
	 * Method populates existing ONet Occupation data
	 * @return list of OnetOccupationDTO
	 */
	@Transactional(propagation = Propagation.REQUIRED)
	public List<OnetOccupationDTO> populateOnetDataList() throws DAOException {
		LOGGER.info(">>>>>> ContentConfigDAOImpl.populateOnetDataList");
		List<OnetOccupationDTO> list = sessionFactory.getCurrentSession().getNamedQuery("fetchAllOnetData").list();
		LOGGER.info("<<<<<< ContentConfigDAOImpl.populateOnetDataList");
		return list;
	}


	/******* THIS IS FOR JCT PUBLIC VERSION ********/
	@Transactional(propagation = Propagation.REQUIRES_NEW)
	@Async
	public void updateMappingFrozen(String code, String prevoiusName,
			String newName, String prevoiusDesc, String newDesc)
			throws DAOException {
		LOGGER.info(">>>>>> ContentConfigDAOImpl.updateMappingFrozen");
		sessionFactory.getCurrentSession().createSQLQuery("SELECT jct_global_profile_change_mapping_attrs" +
				"('"+code+"', '"+prevoiusName+"', '"+newName+"', '"+prevoiusDesc+"', '"+newDesc+"')").uniqueResult();
	}

	/******* THIS IS FOR JCT PUBLIC VERSION ********/
	@Transactional(propagation = Propagation.REQUIRES_NEW)
	@Async
	public void updateQuestionFrozen(String code, String mainQuestion,
			String prevoiusSubQtn, String newSubQtn) throws DAOException {
		LOGGER.info(">>>>>> ContentConfigDAOImpl.updateQuestionFrozen");
		sessionFactory.getCurrentSession().createSQLQuery("SELECT jct_global_profile_change_question" +
				"('"+code+"', '"+mainQuestion+"', '"+prevoiusSubQtn+"', '"+newSubQtn+"')").uniqueResult();
		
	}
	/******* THIS IS FOR JCT PUBLIC VERSION ********/
	@Transactional(propagation = Propagation.REQUIRES_NEW)
	public JctTermsAndConditions fetchTermsAndConditionDao(int userProfile, int userType)
			throws JCTException {
		LOGGER.info(">>>>>> ContentConfigDAOImpl.fetchTermsAndConditionDao");
		return (JctTermsAndConditions)sessionFactory.getCurrentSession().getNamedQuery("fetchTcByProfileId")
			.setParameter("userProfile", userProfile)
			.setParameter("userType", userType)
			.uniqueResult();
	}
	/******* THIS IS FOR JCT PUBLIC VERSION ********/
	@Transactional(propagation = Propagation.REQUIRES_NEW)
	public String saveTermsAndCondition(JctTermsAndConditions jctTermsAndConditions) 
			throws DAONullException, DAOEntityExistsException, DAOException {
		LOGGER.info(">>>>>> ContentConfigDAOImpl.saveTermsAndCondition");		
		return save(jctTermsAndConditions);
	}
	/******* THIS IS FOR JCT PUBLIC VERSION ********/
	@Transactional(propagation = Propagation.REQUIRES_NEW)
	public void removeTermsAndCondition(JctTermsAndConditions existingTermsAndConditions) 
			throws DAONullException, DAOEntityExistsException, DAOException {
		LOGGER.info(">>>>>> ContentConfigDAOImpl.saveTermsAndCondition");		
		Delete(existingTermsAndConditions);
	}
	
	/******* THIS IS FOR JCT PUBLIC VERSION ********/
	@Transactional(propagation = Propagation.REQUIRES_NEW)
	public JctUserProfile getUserProfileByPk(int userProfile)
			throws DAONullException, DAOEntityExistsException, DAOException {
		LOGGER.info(">>>>>> ContentConfigDAOImpl.saveTermsAndCondition");
		return (JctUserProfile)sessionFactory.getCurrentSession().getNamedQuery("fetchDefaultProfileObjByPk")
				.setParameter("pk", userProfile).uniqueResult();	
	}
}
