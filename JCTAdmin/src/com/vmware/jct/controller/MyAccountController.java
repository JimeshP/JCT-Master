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
import com.vmware.jct.service.IMyAccountService;
import com.vmware.jct.service.vo.MyAccountVO;

/**
 * 
 * <p><b>Class name:</b> CententConfigController.java</p>
 * <p><b>Author:</b> InterraIT</p>
 * <p><b>Purpose:</b> This class acts as a controller for all User Content Configuration on Admin Side..
 * CententConfigController extends BasicController  and has following  methods.
 * -populateExistingUserProfile()
 * <p><b>Description:</b> This class acts as a controller for all Content Configuration on Admin Side. </p>
 * <p><b>Copyrights:</b>All rights reserved by Interra IT and should only be used for its internal 
 * application development.</p>
 * <p><b>Created Date:</b> 28/Mar/2014</p>
 * <p><b>Revision History:</b>
 * 	<li></li>
 * </p>
 */
@Controller
@RequestMapping(value = "/myAccount")
public class MyAccountController extends BasicController {
	
	@Autowired
	private IMyAccountService service;
	
	@Autowired  
	private MessageSource messageSource;
	
	private static final Logger LOGGER = LoggerFactory.getLogger(MyAccountController.class);
	/**
	 * Method populates the Existing User Data.
	 * @param body
	 * @param request
	 * @return ContentConfigVO containing Questions.
	 */
	@RequestMapping(value = ACTION_POPULATE_EXTNG_USER_DATA, method = RequestMethod.POST)
	@ResponseBody()
	public MyAccountVO populateExistingUserData(@RequestBody String body, HttpServletRequest request) {
		LOGGER.info(">>>>>> MyAccountController.populateExistingUserData");
		MyAccountVO myAccountVO = new MyAccountVO();	
		try {						
			myAccountVO = service.populateExistingUserData();
			/*MyAccountVO unused = this.service.populateExistingUserData(userId);
			myAccountVO.setExistingUserDataList(unused.getExistingUserDataList());
			unused = null;		*/	
		} catch (JCTException e) {
			myAccountVO = new MyAccountVO();
			myAccountVO.setStatusCode(StatusConstants.STATUS_FAILURE);
			myAccountVO.setStatusDesc(this.messageSource.getMessage("exception.dao.error", null, null));
			LOGGER.error(e.getLocalizedMessage());
		}
		LOGGER.info("<<<<<< MyAccountController.populateExistingUserData");
		return myAccountVO;
	}
	/**
	 * Method update user first name
	 * @param body
	 * @param request
	 * @return return status code only.
	 */
	@RequestMapping(value = ACTION_UPDATE_FIRST_NAME, method = RequestMethod.POST)
	@ResponseBody()
	public MyAccountVO updateFirstName(@RequestBody String body, HttpServletRequest request) {
		LOGGER.info(">>>>>> MyAccountController.updateFirstName");
		MyAccountVO myAccountVO = new MyAccountVO();
		JsonNode node = CommonUtility.getNode(body);
		String valueDesc = node.get("valueDesc").toString().replaceAll("\"" , "");
		Integer tablePkId = Integer.parseInt(node.get("tablePkId").toString().replaceAll("\"" , ""));
		try {
			myAccountVO = service.updateFirstName(valueDesc, tablePkId);
			myAccountVO = service.populateExistingUserData();
			/*MyAccountVO unused = this.service.populateExistingUserData(userId);
			myAccountVO.setExistingUserDataList(unused.getExistingUserDataList());
			unused = null;	*/
		} catch (JCTException e) {
			myAccountVO = new MyAccountVO();
			myAccountVO.setStatusCode(StatusConstants.STATUS_FAILURE);
			myAccountVO.setStatusDesc(this.messageSource.
					getMessage("exception.dao.user.group.update", null, null));
			LOGGER.error(e.getLocalizedMessage());
		}
		LOGGER.info("<<<<<< MyAccountController.updateFirstName");
		return myAccountVO;
	}
	/**
	 * Method update user last name
	 * @param body
	 * @param request
	 * @return return status code only.
	 */
	@RequestMapping(value = ACTION_UPDATE_LAST_NAME, method = RequestMethod.POST)
	@ResponseBody()
	public MyAccountVO updateLastName(@RequestBody String body, HttpServletRequest request) {
		LOGGER.info(">>>>>> MyAccountController.updateLastName");
		MyAccountVO myAccountVO = new MyAccountVO();
		JsonNode node = CommonUtility.getNode(body);
		String valueDesc = node.get("valueDesc").toString().replaceAll("\"" , "");
		Integer tablePkId = Integer.parseInt(node.get("tablePkId").toString().replaceAll("\"" , ""));
		try {
			myAccountVO = service.updateLastName(valueDesc, tablePkId);			
			myAccountVO = service.populateExistingUserData();
			/*MyAccountVO unused = this.service.populateExistingUserData(userId);
			myAccountVO.setExistingUserDataList(unused.getExistingUserDataList());
			unused = null;	*/
		} catch (JCTException e) {
			myAccountVO = new MyAccountVO();
			myAccountVO.setStatusCode(StatusConstants.STATUS_FAILURE);
			myAccountVO.setStatusDesc(this.messageSource.
					getMessage("exception.dao.user.group.update", null, null));
			LOGGER.error(e.getLocalizedMessage());
		}
		LOGGER.info("<<<<<< MyAccountController.updateLastName");
		return myAccountVO;
	}
	/**
	 * Method update user email id
	 * @param body
	 * @param request
	 * @return return status code only.
	 */
	@RequestMapping(value = ACTION_UPDATE_EMAIL_ID, method = RequestMethod.POST)
	@ResponseBody()
	public MyAccountVO updateEmailId(@RequestBody String body, HttpServletRequest request) {
		LOGGER.info(">>>>>> MyAccountController.updateEmailId");
		MyAccountVO myAccountVO = new MyAccountVO();
		JsonNode node = CommonUtility.getNode(body);
		String valueDesc = node.get("valueDesc").toString().replaceAll("\"" , "");
		Integer tablePkId = Integer.parseInt(node.get("tablePkId").toString().replaceAll("\"" , ""));
		try {
			myAccountVO = service.updateEmailId(valueDesc, tablePkId);	
			myAccountVO = service.populateExistingUserData();
			/*MyAccountVO unused = this.service.populateExistingUserData(userId);
			myAccountVO.setExistingUserDataList(unused.getExistingUserDataList());
			unused = null;*/	
		} catch (JCTException e) {
			myAccountVO = new MyAccountVO();
			myAccountVO.setStatusCode(StatusConstants.STATUS_FAILURE);
			myAccountVO.setStatusDesc(this.messageSource.
					getMessage("exception.dao.user.group.update", null, null));
			LOGGER.error(e.getLocalizedMessage());
		}
		LOGGER.info("<<<<<< MyAccountController.updateEmailId");
		return myAccountVO;
	}
	/**
	 * Method update user password
	 * @param body
	 * @param request
	 * @return return status code only.
	 */
	@RequestMapping(value = ACTION_UPDATE_PASSWORD, method = RequestMethod.POST)
	@ResponseBody()
	public MyAccountVO updatePassword(@RequestBody String body, HttpServletRequest request) {
		LOGGER.info(">>>>>> MyAccountController.updatePassword");
		MyAccountVO myAccountVO = new MyAccountVO();
		JsonNode node = CommonUtility.getNode(body);
		String valueDesc = node.get("valueDesc").toString().replaceAll("\"" , "");
		Integer tablePkId = Integer.parseInt(node.get("tablePkId").toString().replaceAll("\"" , ""));
		try {
			myAccountVO = service.updatePassword(valueDesc, tablePkId);			
		} catch (JCTException e) {
			myAccountVO = new MyAccountVO();
			myAccountVO.setStatusCode(StatusConstants.STATUS_FAILURE);
			myAccountVO.setStatusDesc(this.messageSource.
					getMessage("exception.dao.user.update", null, null));
			LOGGER.error(e.getLocalizedMessage());
		}
		LOGGER.info("<<<<<< MyAccountController.updatePassword");
		return myAccountVO;
	}
}
