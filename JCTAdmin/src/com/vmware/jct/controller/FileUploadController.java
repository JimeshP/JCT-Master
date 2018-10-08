package com.vmware.jct.controller;

import java.io.InputStream;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.context.MessageSource;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.ModelAttribute;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestMethod;
import org.springframework.web.bind.annotation.ResponseBody;
import org.springframework.web.multipart.MultipartFile;
import org.springframework.web.servlet.ModelAndView;

import com.vmware.jct.service.IFileUploadService;
import com.vmware.jct.service.vo.FileUploadVO;
/**
 * 
 * <p><b>Class name:</b> FileUploadController.java</p>
 * <p><b>Author:</b> InterraIT</p>
 * <p><b>Purpose:</b> This class acts as a controller Upload excel file for ONet Data..
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
public class FileUploadController extends BasicController {
	
	@Autowired
	private IFileUploadService service;
	
	@Autowired  
	private MessageSource messageSource;
	
	private static final Logger LOGGER = LoggerFactory.getLogger(ContentConfigController.class);
	
	/**
	* Upload single file using Spring Controller
	* @param body
	* @param request
	* @return ModelAndView containing message
	*/
	@RequestMapping(value = ACTION_SAVE_OCCUPATION_DATA, method = RequestMethod.POST)
	public @ResponseBody
	ModelAndView uploadFileHandler(@ModelAttribute("uploadForm") FileUploadModel uploadForm) {
		LOGGER.info(">>>>>> FileUploadController.uploadFileHandler");
		String message = "";
		MultipartFile multipartFile = uploadForm.getFile();
		String fileName = multipartFile.getOriginalFilename();
		FileUploadVO fileUploadVO = new FileUploadVO();
		if (!multipartFile.isEmpty()) {
            try {
            	InputStream inputStream = multipartFile.getInputStream();
            	fileUploadVO = this.service.saveOnetOccupationData(inputStream, fileName);  
            } catch (Exception e) {
            	LOGGER.error(e.getLocalizedMessage());
            	fileUploadVO.setStatusDesc(this.messageSource.
						getMessage("file.upload.error.msg", null, null));
            }
        } 
		/**
		 * JIRA ID - JCTP-3 
		 * Confirmation message after successfully upload a file
		 */
		if(null == fileUploadVO.getStatusDesc()){
			message = this.messageSource.
					getMessage("file.upload.success.msg", null, null);
		} else {
			message = fileUploadVO.getStatusDesc();
		}		
		LOGGER.info("<<<<<<< FileUploadController.uploadFileHandler");
        return new ModelAndView("onetIntegration", "message", message);       
	    }	   
}
