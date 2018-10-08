package com.vmware.jct.controller;

import java.io.IOException;

import javax.servlet.http.HttpServletRequest;

import org.codehaus.jackson.JsonNode;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.context.MessageSource;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.ModelAttribute;
import org.springframework.web.bind.annotation.RequestBody;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestMethod;
import org.springframework.web.bind.annotation.ResponseBody;
import org.springframework.web.servlet.ModelAndView;
import com.vmware.jct.common.utility.CommonUtility;
import com.vmware.jct.common.utility.StatusConstants;
import com.vmware.jct.exception.JCTException;
import com.vmware.jct.service.IContentConfigService;
import com.vmware.jct.service.IManageUserService;
import com.vmware.jct.service.IVideoUploadService;
import com.vmware.jct.service.vo.ContentConfigVO;
import com.vmware.jct.service.vo.InstructionVO;
import com.vmware.jct.service.vo.PopupInstructionVO;
/**
 * 
 * <p><b>Class name:</b> VideoUploadController.java</p>
 * <p><b>Author:</b> InterraIT</p>
 * <p><b>Purpose:</b> This class acts as a controller Upload video file for popup..
 * FileUploadController extends BasicController  and has following  methods.
 * -uploadFileHandler()
 * <p><b>Description:</b> This class acts as a controller for  Upload excel file for ONet Data on Admin Side. </p>
 * <p><b>Copyrights:</b> All rights reserved by InterraIT and should only be used for its internal application development.</p>
 * <p><b>Created Date:</b> 14/Oct/2014</p>
 * <p><b>Revision History:</b>
 * 	<li></li>
 * </p>
 */

@Controller
public class VideoUploadController extends BasicController {
	
	@Autowired
	private IVideoUploadService service;
	
	@Autowired
	private IManageUserService userService;
	
	@Autowired
	private IContentConfigService instructionService;
	
	@Autowired  
	private MessageSource messageSource;
	
	private static final Logger LOGGER = LoggerFactory.getLogger(ContentConfigController.class);
	
	/**
	* Upload single file using Spring Controller
	* @param body
	* @param request
	* @return ModelAndView containing message
	*/
	@RequestMapping(value = ACTION_SAVE_VIDEO_DATA, method = RequestMethod.POST)
	public @ResponseBody
	ModelAndView uploadFileHandler(@ModelAttribute("uploadForm") FileUploadModel uploadForm) {
		LOGGER.info(">>>>>> FileUploadController.uploadFileHandler");
		String message = "";
		try {
			byte[] bytes = uploadForm.getFile().getBytes();
            String status = service.saveVideo(Integer.parseInt(uploadForm.getInputUserProfileInpt()), 
					bytes, uploadForm.getHiddenFileName(), uploadForm.getFileType(), uploadForm.getInputTextBeforeVideo(), uploadForm.getInputTextAfterVideo());
			if (status.equals("success")) {
				if(null == uploadForm.getHiddenFileName() || uploadForm.getHiddenFileName().equals("") || uploadForm.getHiddenFileName() == "") {
					message = this.messageSource.getMessage("upload.only.text.success.msg", null, null);
				} else if ((null == uploadForm.getInputTextBeforeVideo() || uploadForm.getInputTextBeforeVideo().equals("") || uploadForm.getInputTextBeforeVideo() == "")
						&& (null == uploadForm.getInputTextAfterVideo() || uploadForm.getInputTextAfterVideo().equals("") || uploadForm.getInputTextAfterVideo() == "")){
					message = this.messageSource.getMessage("upload.only.video.success.msg", null, null);
				} else {
					message = this.messageSource.getMessage("upload.text.video.success.msg", null, null);
				}				
			} else {
				this.messageSource.getMessage("video.upload.failure.msg", null, null);
				message = this.messageSource.getMessage("video.upload.failure.msg", null, null);
			}
		} catch (JCTException ex) {
			LOGGER.error("UNABLE TO SAVE VIDEO FILE... REASON: "+ex.getLocalizedMessage());
			message = this.messageSource.getMessage("video.upload.failure.msg", null, null);
		} catch (IOException e) {
			LOGGER.error("UNABLE TO SAVE VIDEO FILE... REASON: "+e.getLocalizedMessage());
			message = this.messageSource.getMessage("video.upload.failure.msg", null, null);
		}
		LOGGER.info("<<<<<<< FileUploadController.uploadFileHandler");
        return new ModelAndView("popupInfoInstruction", "message", message);       
	}
	/**
	 * Method populates the User Group List as well as User profile list.
	 * @param body
	 * @param request
	 * @return PopupInstructionVO containing user profiles and other data.
	 */
	@RequestMapping(value = ACTION_POPULATE_EXTNG_VIDEOS, method = RequestMethod.POST)
	@ResponseBody()
	public PopupInstructionVO populateExistingVideos(@RequestBody String body, HttpServletRequest request) {
		LOGGER.info(">>>>>> ManageUserController.populateExistingVideos");
		PopupInstructionVO popupVO = new PopupInstructionVO();
		try {
			popupVO.setUserProfileMap(this.userService.populateUserProfile());			
			popupVO.setStatusCode(StatusConstants.STATUS_SUCCESS_WITH_RESULT);
		} catch (JCTException e) {
			popupVO.setStatusCode(StatusConstants.STATUS_FAILURE);
			popupVO.setStatusDesc(this.messageSource.getMessage("exception.dao.error",null, null));
			LOGGER.error(e.getLocalizedMessage());
		}
		LOGGER.info("<<<<<< ManageUserController.populateExistingVideos");
		return popupVO;
	}
	/**
	 * Method populates the videos based on profile Id and page.
	 * @param body
	 * @param request
	 * @return PopupInstructionVO containing user profiles and other data.
	 */
	@RequestMapping(value = ACTION_POPULATE_VIDEOS_SELECTN, method = RequestMethod.POST)
	@ResponseBody()
	public PopupInstructionVO populateVideosBySelection(@RequestBody String body, HttpServletRequest request) {
		LOGGER.info(">>>>>> ManageUserController.populateVideosBySelection");
		PopupInstructionVO popupVO = new PopupInstructionVO();
		JsonNode node = CommonUtility.getNode(body);
		int profileId = Integer.parseInt(node.get("profileId").asText());
		String page = node.get("page").asText();
		//String selectedRadioButton = node.get("selectedRadioButton").asText();
		try {
			PopupInstructionVO unused = service.getSelectedVideo(profileId, page);
			popupVO.setVideoLink(unused.getVideoLink());
			
			popupVO.setInstructionBeforeVideo(unused.getInstructionBeforeVideo());
			popupVO.setInstructionAfterVideo(unused.getInstructionAfterVideo());
			
			//popupVO.setInstruction(unused.getInstruction());
			//popupVO.setInstructionHeader(unused.getInstructionHeader());
			popupVO.setStatusCode(StatusConstants.STATUS_SUCCESS_WITH_RESULT);
			if (null == unused.getVideoLink() || unused.getVideoLink().equals("")) {
				popupVO.setStatusDesc(this.messageSource.getMessage("exception.video.not.found",null, null));
			} else {
				popupVO.setStatusDesc("");
			}
		} catch (JCTException e) {
			popupVO = new PopupInstructionVO();
			popupVO.setStatusCode(StatusConstants.STATUS_FAILURE);
			popupVO.setStatusDesc(this.messageSource.getMessage("exception.dao.error",null, null));
			LOGGER.error(e.getLocalizedMessage());
		}
		LOGGER.info("<<<<<< ManageUserController.populateVideosBySelection");
		return popupVO;
	}
	/**
	 * Method saves the popup instruction.
	 * @param body
	 * @param request
	 * @return PopupInstructionVO containing user profiles and other data.
	 */
	@RequestMapping(value = ACTION_SAVE_POPUP_INSTRUCTION, method = RequestMethod.POST)
	@ResponseBody()
	public PopupInstructionVO savePopupInstruction(@RequestBody String body, HttpServletRequest request) {
		LOGGER.info(">>>>>>>> ManageUserController.savePopupInstruction");
		PopupInstructionVO vo = new PopupInstructionVO();
		JsonNode node = CommonUtility.getNode(body);
		String textVal = node.get("insAreaTextVal").asText();
		int userProfileId = Integer.parseInt(node.get("userProfileId").asText());
		String relatedPage = node.get("relatedPage").asText();
		try {
			String status = service.saveTextInstruction(textVal, userProfileId, relatedPage);
			if (status.equals("success")) {
				vo.setStatusDesc(this.messageSource.getMessage("text.ins.success.msg", null, null));
				vo.setStatusCode(StatusConstants.STATUS_SUCCESS_WITH_RESULT);
			} else {
				vo.setStatusDesc(this.messageSource.getMessage("text.ins.failure.msg", null, null));
				vo.setStatusCode(StatusConstants.STATUS_FAILURE);
			}
		} catch (JCTException ex) {
			LOGGER.error("UNABLE TO SAVE TEXT INSTRUCTION: "+ex.getLocalizedMessage());
			vo.setStatusDesc(this.messageSource.getMessage("text.ins.failure.msg", null, null));
			vo.setStatusCode(StatusConstants.STATUS_FAILURE);
		} 
		LOGGER.info("<<<<<<<< ManageUserController.savePopupInstruction");
		return vo;
	}
	/**
	 * Method populates the popuptext  instruction based on profile Id and page.
	 * @param body
	 * @param request
	 * @return PopupInstructionVO containing user profiles and other data.
	 */
	@RequestMapping(value = ACTION_POPULATE_TEXT_INSTRUCTN, method = RequestMethod.POST)
	@ResponseBody()
	public PopupInstructionVO populateTextInsBySelection(@RequestBody String body, HttpServletRequest request) {
		LOGGER.info(">>>>>> ManageUserController.populateTextInsBySelection");
		PopupInstructionVO popupVO = new PopupInstructionVO();
		JsonNode node = CommonUtility.getNode(body);
		int profileId = Integer.parseInt(node.get("profileId").asText());
		String page = node.get("page").asText();
		try {
			popupVO.setInstructionBeforeVideo(this.service.getSelectedInstruction(profileId, page));		
			popupVO.setStatusCode(StatusConstants.STATUS_SUCCESS_WITH_RESULT);
			if (popupVO.getInstructionBeforeVideo().equals("")) {
				popupVO.setStatusDesc(this.messageSource.getMessage("exception.ins.not.found",null, null));
			} else {
				popupVO.setStatusDesc("");
			}
		} catch (JCTException e) {
			popupVO.setStatusCode(StatusConstants.STATUS_FAILURE);
			popupVO.setStatusDesc(this.messageSource.getMessage("exception.dao.error",null, null));
			LOGGER.error(e.getLocalizedMessage());
		}
		LOGGER.info("<<<<<< ManageUserController.populateTextInsBySelection");
		return popupVO;
	}
	
	/**
	* Upload single file using Spring Controller
	* @param body
	* @param request
	* @return ModelAndView containing message
	*/
	@RequestMapping(value = "/uploadInstructionVideo", method = RequestMethod.POST)
	public @ResponseBody
	ModelAndView uploadInstructionVideo(@ModelAttribute("uploadInstructionForm") InstuctionUploadForm uploadInstructionForm) {
		LOGGER.info(">>>>>> FileUploadController.uploadFileHandler");
		String message = "";
		try {
			byte[] bytes = uploadInstructionForm.getFile().getBytes();
			InstructionVO instructionVO = null;
			String insAreaText = uploadInstructionForm.getEditor1();
			String userProfileValue = uploadInstructionForm.getHiddenProfileValInput();
			Integer userProfileId = Integer.parseInt(uploadInstructionForm.getHiddenProfileInput());
			String relatedPage = uploadInstructionForm.getHiddenPageInput();
			String createdBy = uploadInstructionForm.getCreatedBy();
			//TODO
			instructionVO = new InstructionVO();
			instructionVO.setVideoFile(bytes);
			instructionVO.setJctInstructionDesc(insAreaText);
			instructionVO.setJctRelatedPageDesc(relatedPage);
			instructionVO.setUserProfileId(userProfileId);
			instructionVO.setUserProfileName(userProfileValue);
			instructionVO.setCreatedBy(createdBy);				
			try {
				instructionService.saveVideoInstruction(instructionVO);
				message = this.messageSource.getMessage("upload.ins.success.msg", null, null);
			} catch (JCTException e) {
				LOGGER.error(e.getLocalizedMessage());
			}
			} catch (IOException e) {
			LOGGER.error("UNABLE TO SAVE VIDEO FILE... REASON: "+e.getLocalizedMessage());
			message = "Unable to save instruction";
		}
		LOGGER.info("<<<<<<< FileUploadController.uploadFileHandler");
        return new ModelAndView("instruction", "message", message);
	}
}