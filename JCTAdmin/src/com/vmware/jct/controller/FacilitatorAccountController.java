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
import com.vmware.jct.service.IFacilitatorAccountService;
import com.vmware.jct.service.vo.MyAccountVO;

/**
 * 
 * <p><b>Class name:</b> FacilitatorAccountController.java</p>
 * <p><b>Author:</b> InterraIT</p>
 * <p><b>Purpose:</b> This class acts as a controller for all facilitator account on facilitator Side..
 * FacilitatorAccountController extends BasicController  and has following  methods.
 * -populateExistingFacilitatorData()
 * -updateUserNameFacilitator()
 * -updateEmailIdFacilitator()
 * -updatePasswordFacilitator()
 * <p><b>Description:</b> This class acts as a controller for Facilitator Side. </p>
 * <p><b>Copyrights:</b>All rights reserved by Interra IT and should only be used for its internal 
 * application development.</p>
 * <p><b>Created Date:</b> 01/Oct/2014</p>
 * <p><b>Revision History:</b>
 * 	<li></li>
 * </p>
 */
@Controller
@RequestMapping(value = "/facilitatorAccount")
public class FacilitatorAccountController extends BasicController {
	
	@Autowired
	private IFacilitatorAccountService service;
	
	@Autowired  
	private MessageSource messageSource;
	
	private static final Logger LOGGER = LoggerFactory.getLogger(FacilitatorAccountController.class);
	/**
	 * Method populates the facilitator data.
	 * @param body
	 * @param request
	 * @return myAccountVO
	 */
	@RequestMapping(value = ACTION_POPULATE_EXTNG_FACILITATOR_DATA, method = RequestMethod.POST)
	@ResponseBody()
	public MyAccountVO populateExistingFacilitatorData(@RequestBody String body, HttpServletRequest request) {
		LOGGER.info(">>>>>> FacilitatorAccountController.populateExistingFacilitatorData");
		MyAccountVO myAccountVO = new MyAccountVO();	
		JsonNode node = CommonUtility.getNode(body);
		String facilitatorEmail = node.get("facilitatorEmail").asText();
		try {						
			myAccountVO = service.populateExistingFacilitatorData(facilitatorEmail);
		} catch (JCTException e) {
			myAccountVO = new MyAccountVO();
			myAccountVO.setStatusCode(StatusConstants.STATUS_FAILURE);
			myAccountVO.setStatusDesc(this.messageSource.getMessage("exception.dao.error", null, null));
			LOGGER.error(e.getLocalizedMessage());
		}
		LOGGER.info("<<<<<< FacilitatorAccountController.populateExistingFacilitatorData");
		return myAccountVO;
	}
	/**
	 * Method update user name for facilitator
	 * @param body
	 * @param request
	 * @return myAccountVO
	 */
	@RequestMapping(value = ACTION_UPDATE_USER_NAME_FACILITATOR, method = RequestMethod.POST)
	@ResponseBody()
	public MyAccountVO updateUserNameFacilitator(@RequestBody String body, HttpServletRequest request) {
		LOGGER.info(">>>>>> FacilitatorAccountController.updateUserNameFacilitator");
		MyAccountVO myAccountVO = new MyAccountVO();
		JsonNode node = CommonUtility.getNode(body);
		String valueDesc = node.get("valueDesc").toString().replaceAll("\"" , "");
		Integer tablePkId = Integer.parseInt(node.get("tablePkId").toString().replaceAll("\"" , ""));
		String facilitatorEmail = node.get("facilitatorEmail").asText();
		try {
			myAccountVO = service.updateUserNameFacilitator(valueDesc, tablePkId, facilitatorEmail);
			myAccountVO = service.populateExistingFacilitatorData(valueDesc);
			myAccountVO.setStatusDesc(this.messageSource.
					getMessage("success.msg.user.name.update", null, null));
		} catch (JCTException e) {
			myAccountVO = new MyAccountVO();
			myAccountVO.setStatusCode(StatusConstants.STATUS_FAILURE);
			myAccountVO.setStatusDesc(this.messageSource.
					getMessage("exception.dao.user.group.update", null, null));
			LOGGER.error(e.getLocalizedMessage());
		}
		LOGGER.info("<<<<<< FacilitatorAccountController.updateUserNameFacilitator");
		return myAccountVO;
	}
	/**
	 * Method to update the email address
	 * @param body
	 * @param request
	 * @return return status code only.
	 */
	@RequestMapping(value = ACTION_UPDATE_EMAIL_ID_FACILITATOR, method = RequestMethod.POST)
	@ResponseBody()
	public MyAccountVO updateEmailIdFacilitator(@RequestBody String body, HttpServletRequest request) 
			throws JCTException {
		LOGGER.info(">>>>>> FacilitatorAccountController.updateEmailIdFacilitator");
		MyAccountVO myAccountVO = new MyAccountVO();
		JsonNode node = CommonUtility.getNode(body);
		String valueDesc = node.get("valueDesc").toString().replaceAll("\"" , "");
		Integer tablePkId = Integer.parseInt(node.get("tablePkId").toString().replaceAll("\"" , ""));
		String facilitatorEmail = node.get("facilitatorEmail").asText();
		String updationSuccess = service.searchUserIdInDB(valueDesc);
		if(updationSuccess.equals("exist")){
			myAccountVO = service.populateExistingFacilitatorData(facilitatorEmail);
			myAccountVO.setStatusCode(StatusConstants.ALREADY_EXISTS);
			myAccountVO.setStatusDesc(this.messageSource.
					getMessage("success.msg.email.already.exist", null, null));					
		} else {
			try {
				myAccountVO = service.updateEmailIdFacilitator(valueDesc, tablePkId, facilitatorEmail);	
				if(myAccountVO.getStatusCode() == StatusConstants.STATUS_SUCCESS_WITH_RESULT) {					
					myAccountVO = service.populateExistingFacilitatorData(valueDesc);
					myAccountVO.setStatusDesc(this.messageSource.
							getMessage("success.msg.email.id.update", null, null));
				} else {
					myAccountVO = service.populateExistingFacilitatorData(facilitatorEmail);
				}					
			} catch (JCTException e) {
				myAccountVO = new MyAccountVO();
				myAccountVO.setStatusCode(StatusConstants.STATUS_FAILURE);
				myAccountVO.setStatusDesc(this.messageSource.
						getMessage("exception.dao.user.group.update", null, null));
				LOGGER.error(e.getLocalizedMessage());
			}
		}
		
		LOGGER.info("<<<<<< FacilitatorAccountController.updateEmailIdFacilitator");
		return myAccountVO;
	}
	/**
	 * Method update the password 
	 * @param body
	 * @param request
	 * @return return status code only.
	 */
	@RequestMapping(value = ACTION_UPDATE_PASSWORD_FACILITATOR, method = RequestMethod.POST)
	@ResponseBody()
	public MyAccountVO updatePasswordFacilitator(@RequestBody String body, HttpServletRequest request) {
		LOGGER.info(">>>>>> FacilitatorAccountController.updatePasswordFacilitator");
		MyAccountVO myAccountVO = new MyAccountVO();
		JsonNode node = CommonUtility.getNode(body);
		String valueDesc = node.get("valueDesc").toString().replaceAll("\"" , "");
		Integer tablePkId = Integer.parseInt(node.get("tablePkId").toString().replaceAll("\"" , ""));
		String facilitatorEmail = node.get("facilitatorEmail").asText();
		try {
			myAccountVO = service.updatePasswordFacilitator(valueDesc, tablePkId, facilitatorEmail);		
			myAccountVO = service.populateExistingFacilitatorData(facilitatorEmail);
			myAccountVO.setStatusDesc(this.messageSource.
					getMessage("success.msg.password.update", null, null));
		} catch (JCTException e) {
			myAccountVO = new MyAccountVO();
			myAccountVO.setStatusCode(StatusConstants.STATUS_FAILURE);
			myAccountVO.setStatusDesc(this.messageSource.
					getMessage("exception.dao.user.update", null, null));
			LOGGER.error(e.getLocalizedMessage());
		}
		LOGGER.info("<<<<<< FacilitatorAccountController.updatePasswordFacilitator");
		return myAccountVO;
	}
}
