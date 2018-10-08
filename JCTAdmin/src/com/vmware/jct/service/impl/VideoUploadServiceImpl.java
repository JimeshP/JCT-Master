package com.vmware.jct.service.impl;

import java.io.BufferedOutputStream;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Date;
import java.util.Iterator;
import java.util.List;

import org.apache.activemq.util.ByteArrayInputStream;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.context.MessageSource;
import org.springframework.stereotype.Service;
import org.springframework.transaction.annotation.Propagation;
import org.springframework.transaction.annotation.Transactional;

import com.vmware.jct.common.utility.CommonConstants;
import com.vmware.jct.dao.ICommonDao;
import com.vmware.jct.dao.IVideoUploadDAO;
import com.vmware.jct.exception.DAOException;
import com.vmware.jct.exception.JCTException;
import com.vmware.jct.model.JctPopupInstruction;
import com.vmware.jct.service.IStorageService;
import com.vmware.jct.service.IVideoUploadService;
import com.vmware.jct.service.vo.PopupInstructionVO;

@Service
public class VideoUploadServiceImpl implements IVideoUploadService {

	private static final Logger LOGGER = LoggerFactory.getLogger(VideoUploadServiceImpl.class);
	
	@Autowired
	private ICommonDao commonDao;
	
	@Autowired
	private IVideoUploadDAO videoUploadDAO;
	
	@Autowired  
	private MessageSource messageSource;
	
	@Autowired
	private IStorageService storageService;
	
	@Transactional(propagation = Propagation.REQUIRED)
	public String saveVideo(int profileId, byte[] bytes, String fileName,
			String fileType, String instructionTextBeforeVideo, String instructionTextAfterVideo) throws JCTException {
		String status = "failure";
		boolean isSuccess = false;
		LOGGER.info(">>>>>>>> VideoUploadServiceImpl.saveVideo");
		String temp = "JCT_VIDEO";
		try {
			Integer videoFileId = commonDao.generateKey("jct_ins_video_file_id");
			if(null == fileName || fileName.equals("") || fileName == "") {
				isSuccess = true;
			} else if(fileName.contains(temp)) {				
				isSuccess = true;
			} else {
				//
				//isSuccess = saveFileOnDisk(bytes, fileType, fileName, videoFileId);
				storageService.store(getVideoFilePath(fileType, fileName, videoFileId), new ByteArrayInputStream(bytes), "video/mp4");
				isSuccess = true;
				
			}			
			if (isSuccess) {
				// Soft delete all the entries matching the criteria
				List<JctPopupInstruction> list = videoUploadDAO.getElements(profileId, fileType);
				Iterator<JctPopupInstruction> itr = list.iterator();
				while (itr.hasNext()) {
					JctPopupInstruction obj = (JctPopupInstruction) itr.next();
					obj.setJctPopupInstructionSoftDelete(CommonConstants.ENTRY_SOFT_DELETED);
					// Merge it back to the db
					videoUploadDAO.updateEntity(obj);
				}
				JctPopupInstruction object = createEntity(videoFileId, fileType, profileId, fileName, instructionTextBeforeVideo, instructionTextAfterVideo);
				status = videoUploadDAO.saveEntity(object);
			}
        } catch (Exception e) {
           	LOGGER.error("UNABLE TO SAVE VIDEO FILE... REASON: "+e.getLocalizedMessage());
        }
		LOGGER.info("<<<<<<<< VideoUploadServiceImpl.saveVideo");
		return status;
	}
	
	private String getVideoFilePath(String fileType, String fileName, int videoFileId) {
		String baseVideoPath = this.messageSource.getMessage("video.path",null, null);
		String videoFilePath = baseVideoPath + "/VIDEO-" + fileType + "-" + videoFileId + fileName.substring(fileName.indexOf("."), fileName.length());
		return videoFilePath;
	}
	
	@Deprecated
	private boolean saveFileOnDisk (byte[] bytes, String fileType, String fileName, int videoFileId) throws IOException {
		boolean isSuccess = false;
		String videoPath = this.messageSource.getMessage("video.path",null, null);
		File videoFile = new File(videoPath + "/JCT_VIDEO/VIDEO-" + fileType +
   				"-" + videoFileId + fileName.substring(fileName.indexOf("."), fileName.length()));
//		File videoFile = new File("../webapps/JCT_VIDEO/VIDEO-" + fileType +
//   				"-" + videoFileId + fileName.substring(fileName.indexOf("."), fileName.length()));
		BufferedOutputStream stream = new BufferedOutputStream(new FileOutputStream(videoFile));
		LOGGER.info("VIDEO FILE --> "+videoFile.getAbsolutePath());
		stream.write(bytes);
		stream.close();
		isSuccess = true;
		return isSuccess;
	}

	private JctPopupInstruction createEntity(int videoFileId, String fileType, int profileId, 
				String fileName, String instructionTextBeforeVideo, String instructionTextAfterVideo) throws DAOException {
		LOGGER.info("FILE NAME IN SERVICE - CREATE ENTITY: "+fileName);
		String temp = "JCT_VIDEO";
		JctPopupInstruction obj = new JctPopupInstruction();
		obj.setJctPopupInstructionCreatedBy("ADMIN");
		obj.setJctPopupInstructionCreationDate(new Date());
		obj.setJctPopupInstructionLastUpdationDate(new Date());
		obj.setJctPopupInstructionPage(fileType);
		obj.setJctPopupInstructionProfileId(profileId);
		obj.setJctPopupInstructionSoftDelete(CommonConstants.ENTRY_NOT_DELETED);
		obj.setJctPopupInstructionType("TEXTANDVIDEO");
		//obj.setJctPopupInstructionHeaderName(instructionHeader);
		if(null == instructionTextBeforeVideo || instructionTextBeforeVideo == "" || instructionTextBeforeVideo.equals("")) {
			obj.setJctPopupInstructionTextBeforeVideo("");
		} else {
			obj.setJctPopupInstructionTextBeforeVideo(instructionTextBeforeVideo);
		}
		
		if(null == instructionTextAfterVideo || instructionTextAfterVideo == "" || instructionTextAfterVideo.equals("")) {
			obj.setJctPopupInstructionTextAfterVideo("");
		} else {
			obj.setJctPopupInstructionTextAfterVideo(instructionTextAfterVideo);
		}

		if(null == fileName || fileName.equals("") || fileName == "") {
			obj.setJctPopupInstructionVideoName("");
		} else if (fileName.contains(temp)) {
			obj.setJctPopupInstructionVideoName(fileName);
		} else {
//			String videoPath = this.messageSource.getMessage("video.path",null, null);
//			obj.setJctPopupInstructionVideoName(videoPath + "JCT_VIDEO/VIDEO-" + fileType + "-" + videoFileId + fileName.substring(fileName.indexOf("."), fileName.length()));
			obj.setJctPopupInstructionVideoName(getVideoFilePath(fileType, fileName, videoFileId));
		}
		
		/*if(null == instructionText || instructionText == "") {
			obj.setJctPopupInstructionType("VIDEO");
		} else {
			obj.setJctPopupInstructionType("TEXTANDVIDEO");
		}*/		
		//obj.setJctPopupInstructionVideoName("../../JCT_VIDEO/VIDEO-" + fileType + "-" + videoFileId + fileName.substring(fileName.indexOf("."), fileName.length()));
		return obj;
	}

	@Transactional(propagation = Propagation.REQUIRED)
	public PopupInstructionVO getSelectedVideo(int profileId, String page)
			throws JCTException {
		LOGGER.info(">>>>>>>> VideoUploadServiceImpl.getSelectedVideo");
		String videoLink = "";
		PopupInstructionVO popupVO = new PopupInstructionVO();
		try {
			List<JctPopupInstruction> list = videoUploadDAO.getSelectedElements(profileId, page);
			if (list.size() > 0) {
				JctPopupInstruction obj = (JctPopupInstruction) list.get(list.size()-1);
				//videoLink = obj.getJctPopupInstructionVideoName();
				popupVO.setVideoLink(obj.getJctPopupInstructionVideoName());
				popupVO.setInstructionBeforeVideo(obj.getJctPopupInstructionTextBeforeVideo());
				popupVO.setInstructionAfterVideo(obj.getJctPopupInstructionTextAfterVideo());
				LOGGER.info("VIDEO "+videoLink+" FOUND FOR PROFILE-ID: "+profileId+" AND PAGE TYPE: "+page);
			} else {
				LOGGER.info("NO VIDEO FOUND FOR PROFILE-ID: "+profileId+" AND PAGE TYPE: "+page);
			}
		} catch (DAOException ex) {
			LOGGER.error("UNABLE TO FETCH VIDEO FILE... REASON: "+ex.getLocalizedMessage());
		}
		LOGGER.info("<<<<<<<< VideoUploadServiceImpl.getSelectedVideo");
		return popupVO;
	}

	@Transactional(propagation = Propagation.REQUIRED)
	public String saveTextInstruction(String textVal, int userProfileId,
			String relatedPage) throws JCTException {
		String status = "failure";
		try {
			// Soft delete all the entries matching the criteria
			List<JctPopupInstruction> list = videoUploadDAO.getElements(userProfileId, relatedPage);
			Iterator<JctPopupInstruction> itr = list.iterator();
			while (itr.hasNext()) {
				JctPopupInstruction obj = (JctPopupInstruction) itr.next();
				obj.setJctPopupInstructionSoftDelete(CommonConstants.ENTRY_SOFT_DELETED);
				// Merge it back to the db
				videoUploadDAO.updateEntity(obj);
			}
			JctPopupInstruction object = createTextEntity(textVal, relatedPage, userProfileId);
			status = videoUploadDAO.saveEntity(object);
		} catch (DAOException ex) {
			LOGGER.error("UNABLE TO SAVE INSTRUCTION TEXT... REASON: "+ex.getLocalizedMessage());
		}
		return status;
	}
	
	private JctPopupInstruction createTextEntity(String textVal, String fileType, int profileId) throws DAOException {
		JctPopupInstruction obj = new JctPopupInstruction();
		obj.setJctPopupInstructionCreatedBy("ADMIN");
		obj.setJctPopupInstructionCreationDate(new Date());
		obj.setJctPopupInstructionLastUpdationDate(new Date());
		obj.setJctPopupInstructionPage(fileType);
		obj.setJctPopupInstructionProfileId(profileId);
		obj.setJctPopupInstructionSoftDelete(CommonConstants.ENTRY_NOT_DELETED);
		obj.setJctPopupInstructionTextBeforeVideo(textVal);
		obj.setJctPopupInstructionType("TEXT");
		obj.setJctPopupInstructionVideoName("");
		return obj;
	}

	@Transactional(propagation = Propagation.REQUIRED)
	public String getSelectedInstruction(int profileId, String page)
			throws JCTException {
		LOGGER.info(">>>>>>>> VideoUploadServiceImpl.getSelectedInstruction");
		String popupTextIns = "";
		try {
			List<JctPopupInstruction> list = videoUploadDAO.getSelectedInstruction(profileId, page);
			if (list.size() > 0) {
				JctPopupInstruction obj = (JctPopupInstruction) list.get(list.size()-1);
				popupTextIns = obj.getJctPopupInstructionTextBeforeVideo();
			} else {
				LOGGER.info("NO TEXT INSTRUCTION FOUND FOR PROFILE-ID: "+profileId+" AND PAGE TYPE: "+page);
			}
		} catch (DAOException ex) {
			LOGGER.error("UNABLE TO FETCH POPUP INSTRUCTION... REASON: "+ex.getLocalizedMessage());
		}
		LOGGER.info("<<<<<<<< VideoUploadServiceImpl.getSelectedInstruction");
		return popupTextIns;
	}
}