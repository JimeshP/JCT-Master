package com.vmware.jct.service.impl;

import java.util.Date;
import java.util.Iterator;
import java.util.List;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;
import org.springframework.transaction.annotation.Propagation;
import org.springframework.transaction.annotation.Transactional;

import com.vmware.jct.dao.IJobAttributeFrozenDAO;
import com.vmware.jct.dao.IQuestionnaireDAO;
import com.vmware.jct.dao.dto.JobAttributeDTO;
import com.vmware.jct.exception.DAOException;
import com.vmware.jct.exception.JCTException;
import com.vmware.jct.model.JctJobAttributeFrozen;
import com.vmware.jct.service.IJobAttributeFrozenService;

/**
 * 
 * <p><b>Class name:</b> JobAttributeFrozenServiceImpl.java</p>
 * <p><b>Author:</b> InterraIT</p>
 * <p><b>Purpose:</b> This is a JobAttributeFrozenServiceImpl class. This artifact is Business layer artifact.
 * JobAttributeFrozenServiceImpl implement IJobAttributeFrozenService interface and override the following  methods.
 * -freezeContent(String jrNo, String createdFor, int profileId) 
 * </p>
 * <p><b>Description:</b> This class used to perform insert, modify ,fetch data from db and delete operation. Save, update, delete operations are performed in DataAccessObject class </p>
 * <p><b>Revision History:</b>
 * 	<li>InterraIT - 13/Oct/2014 - Introduced the class </li>
 * </p>
 */
@Service
public class JobAttributeFrozenServiceImpl implements IJobAttributeFrozenService {
	
	@Autowired
	private IQuestionnaireDAO questionnaireDAO;
	
	@Autowired
	private IJobAttributeFrozenDAO frozenDAO;
	
	private static final Logger LOGGER = LoggerFactory.getLogger(JobAttributeFrozenServiceImpl.class);
	
	/**
	 * Method freezes the job attribute content.
	 * @param job refarance no, created for and profile id
	 * @throws JCTException
	 * @return String
	 */
	@Transactional(propagation=Propagation.REQUIRED)
	public String freezeContent(String jrNo, String createdFor, int profileId, int userId)
			throws JCTException {
		String msg = "failure";
		// Get all the job attributes corresponding to the profile from master table (jct_job_attribute)
		List<JobAttributeDTO> jctJobAttributeList = null;
		try{					
			// Delete all the job attributes in the frozen table which matches jrNo, createdFor and profileId
			Long listCount = questionnaireDAO.getExistingFrozenAttrsCount(userId, createdFor, profileId);
			if (listCount.intValue() == 0) {
				jctJobAttributeList = questionnaireDAO.getFetchJobAttribute(profileId);
				//Iterate over the list and save it in the database.
				Iterator<JobAttributeDTO> itr = jctJobAttributeList.iterator();
				while (itr.hasNext()) {
					JobAttributeDTO dto = (JobAttributeDTO) itr.next();
					JctJobAttributeFrozen frozen = new JctJobAttributeFrozen();
					frozen.setJctJobAttributeCode(dto.getJctJobAttributeCode());
					frozen.setJctJobAttributeName(dto.getJctJobAttributeName());
					frozen.setJctJobAttributeProfileId(profileId);
					frozen.setJctJobAttributeProfileDesc(dto.getJctProfileDesc());
					frozen.setJctJobAttributeDesc(dto.getJctJobAttributeDesc());
					frozen.setJctJobAttributeOrder(dto.getJctJobAttributeOrder());
					frozen.setJctJobAttributeCreatedFor(createdFor);
					frozen.setJctJobAttributeCreatedBy(createdFor);
					frozen.setJctJobAttributeCreatedTs(new Date());
					frozen.setJctJobAttributemodifiedBy(createdFor);
					frozen.setJctJobAttributemodifiedTs(new Date());
					frozen.setJctJobAttributeUserId(userId);
					// Save the entity
					frozenDAO.saveFrozenAttributes(frozen);
					msg = "success";
				}
			} else {
				LOGGER.info("FROZEN CONTENT ALREADY EXIST FOR JOB ATTRIBUTE FOR PROFILE ID: "+profileId+", AND USER ID: "+userId);
				msg = "success";
			}
		} catch(DAOException daoException) {
			LOGGER.error(daoException.getLocalizedMessage());
		}
		return msg;
	}

	/**
	 *<p><b>Description:</b> This method fetches the job attributes which will be shown in mapping page.</p>
	 * @param profile id
	 * @return List<JobAttributeDTO>
	 * @throws JCTException
	 * 
	 */
	 public List<JobAttributeDTO> fetchJobAttribute (int profileId, String userName, int userId) throws JCTException {
		 LOGGER.info(">>>> QuestionnaireServiceImpl.fetchJobAttribute");
		List<JobAttributeDTO> jctJobAttributeList=null;
		try{					
			jctJobAttributeList = frozenDAO.getFrozenJobAttribute(profileId, userName, userId);
		}catch(DAOException daoException){
			LOGGER.error(daoException.getLocalizedMessage());
		}
		LOGGER.info("<<<< QuestionnaireServiceImpl.fetchJobAttribute");
		return jctJobAttributeList; 
	 }
}
