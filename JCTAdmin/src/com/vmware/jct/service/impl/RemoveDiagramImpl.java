package com.vmware.jct.service.impl;

import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.context.MessageSource;
import org.springframework.stereotype.Service;
import org.springframework.transaction.annotation.Propagation;
import org.springframework.transaction.annotation.Transactional;

import com.vmware.jct.common.utility.MailNotificationHelper;
import com.vmware.jct.common.utility.StatusConstants;
import com.vmware.jct.dao.IManageServiceDAO;
import com.vmware.jct.dao.IRemoveDiagramDAO;
import com.vmware.jct.dao.dto.RemoveDiagramDTO;
import com.vmware.jct.exception.DAOException;
import com.vmware.jct.exception.JCTException;
import com.vmware.jct.exception.MailingException;
import com.vmware.jct.model.JctUser;
import com.vmware.jct.service.IRemoveDiagram;
import com.vmware.jct.service.vo.RemoveDiagramVO;
import com.vmware.jct.service.vo.RemoveDiagramVOList;
import com.vmware.jct.service.vo.UserVO;
/**
 * 
 * <p><b>Class name:</b> RemoveDiagramImpl.java</p>
 * <p><b>Author:</b> InterraIT</p>
 * <p><b>Purpose:</b> This is a RemoveDiagramImpl class. This artifact is Business layer artifact.
 * RemoveDiagramImpl implement IRemoveDiagram interface and override the following  methods.
 * -getDiagrams(String emailId))
 * </p>
 * <p><b>Description:</b> This class used to perform insert, modify ,fetch data from db and delete operation. Save, update, delete operations are performed in DataAccessObject class </p>
 * <p><b>Revision History:</b>
 * 	<li>InterraIT - 06/Oct/2014 - Introduced the class </li>
 * </p>
 */
@Service
public class RemoveDiagramImpl implements IRemoveDiagram {

	@Autowired
	private IRemoveDiagramDAO diagramDAO;
	
	@Autowired  
	private MessageSource messageSource;
	
	@Autowired
	private MailNotificationHelper mailer;
	
	private static final Logger LOGGER = LoggerFactory.getLogger(RemoveDiagramImpl.class);
	
	@Transactional(propagation = Propagation.REQUIRED)
	public RemoveDiagramVOList getDiagrams(String emailId) throws JCTException {
		LOGGER.info(">>>>>>>> RemoveDiagramImpl.getDiagrams");
		RemoveDiagramVOList list = new RemoveDiagramVOList();	
		RemoveDiagramVO diagramVO = null;//new RemoveDiagramVO();
		List<RemoveDiagramVO> voList = new ArrayList<RemoveDiagramVO>();
		String userName = "";
		try {
			List<RemoveDiagramDTO> dtoList = diagramDAO.getDiagrams(emailId);
			if (dtoList.size() > 0) {
				Iterator<RemoveDiagramDTO> itr = dtoList.iterator();
				while (itr.hasNext()) {
					RemoveDiagramDTO dto = (RemoveDiagramDTO) itr.next();
					diagramVO = new RemoveDiagramVO();
					diagramVO.setRowId(dto.getRowId());
					diagramVO.setBeforeSketchBaseString(dto.getBeforeSketchBaseString());
					diagramVO.setAftereSketchBaseString(dto.getAftereSketchBaseString());
					userName = diagramDAO.fetchUserName(dto.getJctJobrefNo());
					diagramVO.setUserName(userName);
					voList.add(diagramVO);
				}
				list.setStatusCode(StatusConstants.STATUS_SUCCESS_WITH_RESULT);
			} else {
				diagramVO = new RemoveDiagramVO();
				diagramVO.setRowId(0);
				diagramVO.setBeforeSketchBaseString("NO-DATA");
				diagramVO.setAftereSketchBaseString("NO-DATA");
				diagramVO.setStatusCode(StatusConstants.STATUS_SUCCESS_WITH_NO_RESULT);
				voList.add(diagramVO);
				list.setStatusCode(StatusConstants.STATUS_SUCCESS_WITH_NO_RESULT);
			}
		} catch (DAOException dao) {
			LOGGER.error("UNABLE TO FETCH THE BEFORE SKETCH AND AFTER SKETCH DIAGRAM: "+dao.getLocalizedMessage());
			diagramVO = new RemoveDiagramVO();
			diagramVO.setRowId(0);
			diagramVO.setBeforeSketchBaseString("FAILED");
			diagramVO.setAftereSketchBaseString("FAILED");
			diagramVO.setStatusCode(StatusConstants.STATUS_FAILURE);
			diagramVO.setMessage(this.messageSource.getMessage("msg.diag.not.fetched", null, null));
			voList.add(diagramVO);
			list.setStatusCode(StatusConstants.STATUS_FAILURE);
		}
		list.setDiagramList(voList);
		LOGGER.info("<<<<<<<< RemoveDiagramImpl.getDiagrams");
		return list;
	}

	@Transactional(propagation = Propagation.REQUIRED)
	public String removeDiagram(RemoveDiagramVOList vo) throws JCTException {
		LOGGER.info(">>>>>>>> RemoveDiagramImpl.removeDiagram");
		String status = "failure";
		int execStatus = 0;
		try {
			if(vo.getRowId() != 0) {
				execStatus =  diagramDAO.removeDiagram(vo);
			} else{
				String checkedValue = vo.getRowIdSetString();
				String[] rowId = checkedValue.split("``");
				for(int outerIndex = 1; outerIndex < rowId.length; outerIndex++){	
					vo.setRowId(Integer.parseInt(rowId[outerIndex]));
					execStatus =  diagramDAO.removeDiagram(vo);
				}
			}		
			if (execStatus == 1) {
				status = "success";
			}
		} catch (DAOException dao) {
			LOGGER.error("UNABLE TO DELETE THE BEFORE SKETCH AND AFTER SKETCH DIAGRAM: "+dao.getLocalizedMessage());
		}
		LOGGER.info(">>>>>>>> RemoveDiagramImpl.removeDiagram");
		return status;
	}

	@Transactional(propagation = Propagation.REQUIRED)
	public RemoveDiagramVOList populateAllDiagram(int recordIndex)
			throws JCTException {
		LOGGER.info(">>>>>>>> RemoveDiagramImpl.populateAllDiagram");
		RemoveDiagramVOList list = new RemoveDiagramVOList();			
		RemoveDiagramVO diagramVO = null;//new RemoveDiagramVO();
		List<RemoveDiagramVO> voList = new ArrayList<RemoveDiagramVO>();
		try {
			List<RemoveDiagramDTO> dtoList = diagramDAO.populateAllDiagram(recordIndex);
			String userName = "";
			if (dtoList.size() > 0) {
				Iterator<RemoveDiagramDTO> itr = dtoList.iterator();
				while (itr.hasNext()) {
					RemoveDiagramDTO dto = (RemoveDiagramDTO) itr.next();
					diagramVO = new RemoveDiagramVO();
					diagramVO.setRowId(dto.getRowId());
					diagramVO.setBeforeSketchBaseString(dto.getBeforeSketchBaseString());
					diagramVO.setAftereSketchBaseString(dto.getAftereSketchBaseString());
					userName = diagramDAO.fetchUserName(dto.getJctJobrefNo());
					diagramVO.setUserName(userName);
					diagramVO.setJctCreatedTs(dto.getJctCreatedTs());
					voList.add(diagramVO);
				}
				list.setStatusCode(StatusConstants.STATUS_SUCCESS_WITH_RESULT);
			} else {
				diagramVO = new RemoveDiagramVO();
				diagramVO.setRowId(0);
				diagramVO.setBeforeSketchBaseString("NO-DATA");
				diagramVO.setAftereSketchBaseString("NO-DATA");
				diagramVO.setStatusCode(StatusConstants.STATUS_SUCCESS_WITH_NO_RESULT);
				voList.add(diagramVO);
				list.setStatusCode(StatusConstants.STATUS_SUCCESS_WITH_NO_RESULT);
			}
		} catch (DAOException dao) {
			LOGGER.error("UNABLE TO FETCH THE BEFORE SKETCH AND AFTER SKETCH DIAGRAM: "+dao.getLocalizedMessage());
			diagramVO = new RemoveDiagramVO();
			diagramVO.setRowId(0);
			diagramVO.setBeforeSketchBaseString("FAILED");
			diagramVO.setAftereSketchBaseString("FAILED");
			diagramVO.setStatusCode(StatusConstants.STATUS_FAILURE);
			diagramVO.setMessage(this.messageSource.getMessage("msg.diag.not.fetched", null, null));
			voList.add(diagramVO);
			list.setStatusCode(StatusConstants.STATUS_FAILURE);
		}
		list.setDiagramList(voList);
		LOGGER.info("<<<<<<<< RemoveDiagramImpl.populateAllDiagram");
		return list;
	}

	@Transactional(propagation = Propagation.REQUIRED)
	public Long fetchAllDiagramCount() throws JCTException {
		LOGGER.info(">>>>>>>> RemoveDiagramImpl.fetchAllDiagramCount");
		long count = 0;
		try {
			count = diagramDAO.fetchAllDiagramCount();
		} catch (DAOException dao) {
			LOGGER.error("UNABLE TO FETCH THE BEFORE SKETCH AND AFTER SKETCH DIAGRAM: "+dao.getLocalizedMessage());
		}
		LOGGER.info(">>>>>>>> RemoveDiagramImpl.fetchAllDiagramCount");
		return count;
	}

	@Transactional(propagation = Propagation.REQUIRED)
	public UserVO searchUserForRefundRequest(String emailId)
			throws JCTException {
		LOGGER.info(">>>>>>>> RemoveDiagramImpl.searchUserForRefundRequest");
		UserVO vo = new UserVO();
		JctUser obj = null;
		try {
			obj = diagramDAO.searchUserForRefundRequest(emailId);
			if(obj != null){
				vo.setUserName(obj.getJctUserName());
				vo.setUserGroup(obj.getJctUserDetails().getJctUserDetailsGroupName());
				vo.setUserProfile(obj.getJctUserDetails().getJctUserDetailsProfileName());
				vo.setExpiryDate(obj.getJctAccountExpirationDate());
				vo.setJctUserId(obj.getJctUserId());
				vo.setStatusCode("success");
			} else {
				vo.setStatusCode("failure");
			}
					
			//count = diagramDAO.fetchAllDiagramCount();
		} catch (DAOException dao) {
			LOGGER.error("UNABLE TO FETCH THE BEFORE SKETCH AND AFTER SKETCH DIAGRAM: "+dao.getLocalizedMessage());
		}
		LOGGER.info(">>>>>>>> RemoveDiagramImpl.searchUserForRefundRequest");
		return vo;
	}

	@Transactional(propagation = Propagation.REQUIRED)
	public String disableUserForRefundRequest(int userId, String emailId) throws JCTException {
		LOGGER.info(">>>>>>>> RemoveDiagramImpl.searchUserForRefundRequest");
		String status = "failure";		
		try {	
			boolean updatedRowUser = diagramDAO.disableUserForRefundRequestByUserId(userId);			
			if(updatedRowUser){
				status = "success";	
				try {
					mailer.userDisable(emailId);	//	send mail
				}
				catch (MailingException e) {
					LOGGER.error("UNABLE TO SEND MAIL AFTER DISABLE USER. EMAIL: "+emailId+" REASON: "+e.getLocalizedMessage());
				}
			}
		} catch (DAOException dao) {
			LOGGER.error("UNABLE TO DELETE USER: "+dao.getLocalizedMessage());
		} 
		LOGGER.info(">>>>>>>> RemoveDiagramImpl.searchUserForRefundRequest");
		return status;
	}

}