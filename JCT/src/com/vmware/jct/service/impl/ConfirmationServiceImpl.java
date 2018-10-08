package com.vmware.jct.service.impl;

import java.util.Calendar;
import java.util.Date;
import java.util.List;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;
import org.springframework.transaction.annotation.Propagation;
import org.springframework.transaction.annotation.Transactional;

import com.vmware.jct.common.utility.CommonUtility;
import com.vmware.jct.dao.IActionPlanDAO;
import com.vmware.jct.dao.IAfterSketchDAO;
import com.vmware.jct.dao.IAuthenticatorDAO;
import com.vmware.jct.dao.IBeforeSketchDAO;
import com.vmware.jct.dao.ICommonDao;
import com.vmware.jct.dao.IConfirmationDAO;
import com.vmware.jct.exception.DAOException;
import com.vmware.jct.exception.JCTException;
import com.vmware.jct.model.JctCompletionStatus;
import com.vmware.jct.model.JctUserLoginInfo;
import com.vmware.jct.service.IConfirmationService;

@Service
public class ConfirmationServiceImpl implements IConfirmationService{
	
	@Autowired
	private IConfirmationDAO confirmationDAO;
	
	@Autowired
	private ICommonDao commonDaoImpl;
	
	@Autowired
	private IBeforeSketchDAO beforeSketchDAO;
	
	@Autowired
	private IActionPlanDAO actionPlanDAO;
	
	@Autowired
	private IAfterSketchDAO afterSketchDAO;
	
	@Autowired
	private IAuthenticatorDAO authenticatorDAO;
	private static final Logger LOGGER = LoggerFactory.getLogger(ConfirmationServiceImpl.class);
	/**
	 *<p><b>Description:</b> This method call  function updateStatus  and update soft delete and status in different transaction table.</p>
	 * @param job reference number
	 * @return String
	 * @throws JCTException
	 * 
	 */
	@Transactional(propagation=Propagation.REQUIRED)
	public String addDiagWithoutEdits(String jobReferenceNo)
			throws JCTException {
		LOGGER.info(">>>> ConfirmationServiceImpl.addDiagWithoutEdits");
		String result = "success";
		try{
			result = confirmationDAO.addDiagWithoutEdits(jobReferenceNo);
			completionStatusData(jobReferenceNo, 1);
			updateUserLoginInfo(jobReferenceNo, 1);
		} catch(DAOException ex){
			result = "failure";
			LOGGER.error(ex.getLocalizedMessage());
			throw new JCTException();
		}
		LOGGER.info("<<<< ConfirmationServiceImpl.addDiagWithoutEdits");
		return result;
	}
	/**
	 *<p><b>Description:</b> This method call  function updateStatus  and update soft delete and status in different transaction table. The 
	 *diagram will not be searchable till the user complets</p>
	 * @param job reference number
	 * @return String
	 * @throws JCTException
	 * 
	 */
	@Transactional(propagation=Propagation.REQUIRED)
	public String addEdits(String jobReferenceNo)
			throws JCTException {
		LOGGER.info(">>>> ConfirmationServiceImpl.addEdits");
		String result = "success";
		try{
			//Update soft delete status to 1 for all existing entry of jobRefNo
			confirmationDAO.updateDeleteStatusForEdit(jobReferenceNo);
			result = confirmationDAO.addEdits(jobReferenceNo);
			//completionStatusData(jobReferenceNo);
			updateUserLoginInfo(jobReferenceNo, 2);
		} catch(DAOException ex){
			result = "failure";
			LOGGER.error(ex.getLocalizedMessage());
			throw new JCTException();
		}
		LOGGER.info("<<<< ConfirmationServiceImpl.addEdits");
		return result;
	}
	/**
	 *<p><b>Description:</b> This method updates the user login info table.</p>
	 * @param job reference number
	 * @return String
	 * @throws JCTException
	 * 
	 */
	@Transactional(propagation=Propagation.REQUIRED)
	private void updateUserLoginInfo(String jobReferenceNo, int dist) {
		LOGGER.info(">>>> ConfirmationServiceImpl.updateUserLoginInfo");
		try {
			int maxId = (int)authenticatorDAO.getMaxLoginInfoId(jobReferenceNo);
			//if (dist == 1) { // 1 - Add diags without edits
				authenticatorDAO.updateLoginInfo(maxId, "/user/view/finalPage.jsp");
			//} else {
				//Check if the user has completed at least once then the new result page is to be shown
				int nosOfCompletion = actionPlanDAO.getLastCompletionStatus(jobReferenceNo);
				/*String page = "/user/view/resultPage.jsp"; 
				if (nosOfCompletion > 0) {
					page = "/user/view/resultPageFinal.jsp";
				} 
				authenticatorDAO.updateLoginInfo(maxId, page);*/
			//}
			LOGGER.info("<<<< ConfirmationServiceImpl.updateUserLoginInfo");
		} catch(DAOException dao) {
			LOGGER.error(dao.getLocalizedMessage());
		}
	}
	/**
	 *<p><b>Description:</b> This method disables the diagrams and these diagrams will not be made searchable.</p>
	 * @param job reference number
	 * @return String
	 * @throws JCTException
	 * 
	 */
	@Transactional(propagation=Propagation.REQUIRED)
	public int disableStatus(String jobReferenceNo) throws JCTException {
		LOGGER.info(">>>> ConfirmationServiceImpl.disableStatus");
		int result = 0;
		try{
			result = confirmationDAO.disableStatus(jobReferenceNo);
			completionStatusData(jobReferenceNo, 3);
			updateUserLoginInfo(jobReferenceNo, 3);
		} catch(DAOException ex){
			LOGGER.error(ex.getLocalizedMessage());
			throw new JCTException();
		}
		LOGGER.info("<<<< ConfirmationServiceImpl.disableStatus");
		return result;
	}
	
	private void completionStatusData(String jctJobReferenceNo, int decision) throws JCTException {
		LOGGER.info(">>>> ConfirmationServiceImpl.completionStatusData");
		try {
			JctCompletionStatus entity = new JctCompletionStatus();
			entity.setJctStartDate(new Date());
			entity.setJctEndDate(new Date());
			entity.setJctJobReferenceNo(jctJobReferenceNo);
			entity.setJctOptionSelected(decision);
			confirmationDAO.saveOrUpdateBeforeSketch(entity);
			LOGGER.info("<<<< ConfirmationServiceImpl.completionStatusData");
		} catch(DAOException dao) {
			LOGGER.error(dao.getLocalizedMessage());
		}
	}
	/**
	 *<p><b>Description:</b> This method closes the present entry and opens up the previous one.</p>
	 * @param job reference number
	 * @return String
	 * @throws JCTException
	 * 
	 */
	@Transactional(propagation=Propagation.REQUIRED)
	public String updatePrevious(String jobReferenceNo) throws JCTException {
		String result = "failure";
		try {
			int status = confirmationDAO.keepOriginal(jobReferenceNo);
			LOGGER.info("KEEP ORIGINAL UPDATE RESULT: "+status);
			completionStatusData(jobReferenceNo, 4);
			updateUserLoginInfo(jobReferenceNo, 4);
			result = "success";
		} catch (DAOException daoEx) {
			LOGGER.error("---- ERROR ----: "+daoEx.getLocalizedMessage());
		}
		return result;
	}
	/**
	 *<p><b>Description:</b> This method updates the soft delete and status.</p>
	 * @param job reference number
	 * @return String
	 * @throws JCTException
	 * 
	 */
	@Transactional(propagation=Propagation.REQUIRED)
	public int updateLinkStatus(String jobReferenceNo) throws JCTException {
		LOGGER.info(">>>> ConfirmationServiceImpl.updateLinkStatus");
		int status = 0;
		try {
			status = confirmationDAO.updateLinkStatus(jobReferenceNo);
		} catch (DAOException e) {
			LOGGER.error("--------------- FAILURE ------------------");
		}
		LOGGER.info("<<<< ConfirmationServiceImpl.updateLinkStatus");
		return status;
	}
	/**
	 *<p><b>Description:</b> This method restarts the whole excersize.</p>
	 * @param job reference number
	 * @return String
	 * @throws JCTException
	 * 
	 */
	@Transactional(propagation=Propagation.REQUIRED)
	public String restartExcersize(String jobReferenceNo, String distinction) throws JCTException {
		LOGGER.info(">>>> ConfirmationServiceImpl.restartExcersize");
		String status = "failure";
		try {
			int restarted = confirmationDAO.restartExcersize(jobReferenceNo, distinction);
			if (restarted == 0) {
				status = "success";
			}
		} catch (DAOException daoEx) {
			LOGGER.error("---- ERROR ----: "+daoEx.getLocalizedMessage());
		}
		LOGGER.info("<<<< ConfirmationServiceImpl.restartExcersize");
		return status;
	}
	/**
	 *<p><b>Description:</b> This method fetches the count of completion.</p>
	 * @param job reference number
	 * @return int
	 * @throws JCTException
	 * 
	 */
	@Transactional(propagation=Propagation.REQUIRED)
	public int getExcersizeCompletionCount (String jobRefNo) throws JCTException {
		int count = 0;
		try {
			List<JctCompletionStatus> list= confirmationDAO.getExcersizeCompletionCount(jobRefNo);
			count = list.size();
		} catch (DAOException daoEx) {
			LOGGER.error("---- ERROR ----: "+daoEx.getLocalizedMessage());
		}
		return count;
	}
	@Transactional(propagation=Propagation.REQUIRED)
	public String getTotalTime(int userId) throws JCTException {
		String time = "00:00:00";
		try {
			long totalMillis = 0L;
			List<JctUserLoginInfo> list= confirmationDAO.getTotalTime(userId);
			for (JctUserLoginInfo obj : list) {
				Calendar startCal = Calendar.getInstance();
				startCal.setTime(obj.getJctStartTime());
				
				Calendar endCal = Calendar.getInstance();
				endCal.setTime(obj.getJctEndTime());
				totalMillis += (endCal.getTimeInMillis() - startCal.getTimeInMillis());
			}
			time = CommonUtility.convertMillis(totalMillis);
			
		} catch (DAOException daoEx) {
			LOGGER.error("---- ERROR ----: "+daoEx.getLocalizedMessage());
		}
		return time;
	}
}
