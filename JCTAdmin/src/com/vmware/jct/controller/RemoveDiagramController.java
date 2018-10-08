package com.vmware.jct.controller;

import javax.servlet.http.HttpServletRequest;

import org.codehaus.jackson.JsonNode;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.context.MessageSource;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.RequestBody;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestMethod;
import org.springframework.web.bind.annotation.ResponseBody;

import com.vmware.jct.common.utility.CommonUtility;
import com.vmware.jct.common.utility.StatusConstants;
import com.vmware.jct.exception.JCTException;
import com.vmware.jct.service.IRemoveDiagram;
import com.vmware.jct.service.vo.RemoveDiagramVO;
import com.vmware.jct.service.vo.RemoveDiagramVOList;
import com.vmware.jct.service.vo.ReturnVO;
import com.vmware.jct.service.vo.UserVO;

/**
 * 
 * <p><b>Class name:</b> RemoveDiagramController.java</p>
 * <p><b>Author:</b> InterraIT</p>
 * <p><b>Purpose:</b> This class acts as a controller for removing the shared diagram on Admin Side..
 * RemoveDiagramController extends BasicController  and has following  methods.
 * -searchUserDiagram()
 * <p><b>Description:</b> This class acts as a controller for removing the shared diagram on Admin Side. </p>
 * <p><b>Copyrights:</b>All rights reserved by Interra IT and should only be used for its internal 
 * application development.</p>
 * <p><b>Created Date:</b> 06/Nov/2014</p>
 * <p><b>Revision History:</b>
 * 	<li></li>
 * </p>
 */
@Controller
@RequestMapping(value = "/adminRemoveDiagram")
public class RemoveDiagramController  extends BasicController {

	@Autowired
	private IRemoveDiagram service;
	
	@Autowired  
	private MessageSource messageSource;
	
	private static final Logger LOGGER = LoggerFactory.getLogger(RemoveDiagramController.class);
	/**
	 * Method helps the administrator to search the diagrams which are made available by the users.
	 * @param body
	 * @param request
	 * @return
	 */
	@RequestMapping(value = ACTION_SEARCH_SHARED_DIAGRAM, method = RequestMethod.POST)
	@ResponseBody()
	public RemoveDiagramVOList searchUserDiagram(@RequestBody String body, HttpServletRequest request) {
		LOGGER.info(">>>>>>>> RemoveDiagramController.searchUserDiagram");
		JsonNode node = CommonUtility.getNode(body);
		String emailId = node.get("emailId").asText();
		RemoveDiagramVOList list = null;
		try {
			list = service.getDiagrams(emailId);
		} catch (JCTException ex) {
			LOGGER.error("UNABLE TO FETCH: "+ex.getLocalizedMessage());
			list = new RemoveDiagramVOList();
			list.setStatusCode(StatusConstants.STATUS_FAILURE);
		}
		LOGGER.info("<<<<<<<< RemoveDiagramController.searchUserDiagram");
		return list;
	}
	/**
	 * Method helps the administrator to remove the diagrams which are made available by the users.
	 * @param body
	 * @param request
	 * @return
	 */
	@RequestMapping(value = ACTION_REMOVE_SHARED_DIAGRAM, method = RequestMethod.POST)
	@ResponseBody()
	public RemoveDiagramVOList removeDiagram(@RequestBody String body, HttpServletRequest request) {
		LOGGER.info(">>>>>>>> RemoveDiagramController.removeDiagram");
		JsonNode node = CommonUtility.getNode(body);
		String emailId = node.get("emailId").asText();
		int rowId = node.get("rowId").asInt();
		//String rowIdSet = node.get("rowIdSet").asText();
		String rowIdSet = node.get("rowIdSet").toString().replaceAll("\"" , "");
		RemoveDiagramVOList vo = new RemoveDiagramVOList();
		vo.setStatusCode(StatusConstants.STATUS_FAILURE);
		String status = "failure";
		try {
			vo.setEmailId(emailId);
			vo.setRowId(rowId);
			vo.setRowIdSetString(rowIdSet);
			status = service.removeDiagram(vo);
			if (status.equals("success")) {
				vo.setStatusCode(StatusConstants.STATUS_SUCCESS);
				vo.setStatusMessage(this.messageSource.getMessage("diag.soft.delete.success",null, null));
			}
		} catch (JCTException ex) {
			LOGGER.error("UNABLE TO FETCH: "+ex.getLocalizedMessage());
			vo.setStatusMessage(this.messageSource.getMessage("diag.soft.delete.failure",null, null));
		}
		LOGGER.info("<<<<<<<< RemoveDiagramController.removeDiagram");
		return vo;
	}
	
	/**
	 * Method helps the administrator to fetch all the diagrams.
	 * @param body
	 * @param request
	 * @return
	 */
	@RequestMapping(value = ACTION_POPULATE_ALL_DIAGRAM, method = RequestMethod.POST)
	@ResponseBody()
	public RemoveDiagramVOList populateAllDiagram(@RequestBody String body, HttpServletRequest request) {
		LOGGER.info(">>>>>>>> RemoveDiagramController.populateAllDiagram");
		JsonNode node = CommonUtility.getNode(body);
		int recordIndex = Integer.parseInt(node.get("recordIndex").toString().replaceAll("\"" , ""));
		RemoveDiagramVOList list = null;
		try {
			list = service.populateAllDiagram(recordIndex);
		} catch (JCTException ex) {
			LOGGER.error("UNABLE TO FETCH: "+ex.getLocalizedMessage());
			list = new RemoveDiagramVOList();
			list.setStatusCode(StatusConstants.STATUS_FAILURE);
		}
		LOGGER.info("<<<<<<<< RemoveDiagramController.populateAllDiagram");
		return list;
	}
	
	/**
	 * Method helps the administrator to fetch the count of all the diagrams.
	 * @param body
	 * @param request
	 * @return
	 */
	@RequestMapping(value = ACTION_POPULATE_DIAGRAM_COUNT, method = RequestMethod.POST)
	@ResponseBody()
	public RemoveDiagramVO fetchAllDiagramCount(@RequestBody String body, HttpServletRequest request) {
		LOGGER.info(">>>>>>>> RemoveDiagramController.populateAllDiagram");
		RemoveDiagramVO vo = new RemoveDiagramVO();
		long count = 0;
		try {
			count = service.fetchAllDiagramCount();
			vo.setTotalDiagramCount(count);
			vo.setStatusCode(StatusConstants.STATUS_SUCCESS);
		} catch (JCTException ex) {
			LOGGER.error("UNABLE TO FETCH: "+ex.getLocalizedMessage());
			vo = new RemoveDiagramVO();
			vo.setStatusCode(StatusConstants.STATUS_FAILURE);
		}
		LOGGER.info("<<<<<<<< RemoveDiagramController.populateAllDiagram");
		return vo;
	}
	
	/**
	 * Method helps the administrator to search the user details based on user name.
	 * @param body
	 * @param request
	 * @return
	 */
	@RequestMapping(value = ACTION_SEARCH_USER_REFUND_REQUEST, method = RequestMethod.POST)
	@ResponseBody()
	public UserVO searchUserForRefundRequest(@RequestBody String body, HttpServletRequest request) {
		LOGGER.info(">>>>>>>> RemoveDiagramController.searchUserForRefundRequest");
		JsonNode node = CommonUtility.getNode(body);
		String emailId = node.get("emailId").asText();
		UserVO vo = new UserVO();
		try {
			vo = service.searchUserForRefundRequest(emailId);
			if(vo.getStatusCode().equals("success")){
				vo.setStatusCodeInt(StatusConstants.STATUS_SUCCESS);
				vo.setStatusDesc("success");				
			} else {
				vo.setStatusCodeInt(StatusConstants.STATUS_FAILURE);
				vo.setStatusDesc("failure");				
			}			
		} catch (JCTException ex) {
			LOGGER.error("UNABLE TO FETCH: "+ex.getLocalizedMessage());
			vo.setStatusCodeInt(StatusConstants.STATUS_FAILURE);
			vo.setStatusDesc("failure");
		}
		LOGGER.info("<<<<<<<< RemoveDiagramController.searchUserForRefundRequest");
		return vo;
	}
	
	/**
	 * Method helps the administrator to fetch all the diagrams.
	 * @param body
	 * @param request
	 * @return
	 */
	@RequestMapping(value = ACTION_DISABLE_USER_REFUND_REQUEST, method = RequestMethod.POST)
	@ResponseBody()
	public ReturnVO disableUserForRefundRequest(@RequestBody String body, HttpServletRequest request) {
		LOGGER.info(">>>>>>>> RemoveDiagramController.disableUserForRefundRequest");
		JsonNode node = CommonUtility.getNode(body);
		int userId = Integer.parseInt(node.get("userId").toString().replaceAll("\"" , ""));
		String emailId = node.get("emailId").asText();
		ReturnVO vo = new ReturnVO();
		String status = "failure";
		try {
			status = service.disableUserForRefundRequest(userId,emailId);
			if(status.equalsIgnoreCase("success")) {
				vo.setStatusCodeInt(StatusConstants.STATUS_SUCCESS);
				vo.setStatusDesc("success");
			}
		} catch (JCTException ex) {
			LOGGER.error("UNABLE TO FETCH: "+ex.getLocalizedMessage());
			vo.setStatusCodeInt(StatusConstants.STATUS_SUCCESS);
			vo.setStatusDesc("success");
		}
		LOGGER.info("<<<<<<<< RemoveDiagramController.disableUserForRefundRequest");
		return vo;
	}
}