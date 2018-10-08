package com.vmware.jct.service;

import com.vmware.jct.exception.JCTException;
import com.vmware.jct.service.vo.PopupInstructionVO;

public interface IVideoUploadService {
	/**
	 * Method saves the video file on the disk and stores the corresponding entry in 
	 * database.
	 * @param profileId
	 * @param fileName
	 * @param bytes
	 * @param fileType
	 * @return
	 * @throws JCTException
	 */
	public String saveVideo(int profileId, byte[] bytes, String fileName, String fileType, String instructionText, String instructionHeader) throws JCTException;
	/**
	 * Method populates the video
	 * @param profileId
	 * @param page
	 * @return
	 * @throws JCTException
	 */
	public PopupInstructionVO getSelectedVideo(int profileId, String page) throws JCTException;
	/**
	 * Method gets selected instruction
	 * @param profileId
	 * @param page
	 * @return
	 * @throws JCTException
	 */
	public String getSelectedInstruction(int profileId, String page) throws JCTException;
	/**
	 * mathod saves instruction
	 * @param textVal
	 * @param userProfileId
	 * @param relatedPage
	 * @return
	 * @throws JCTException
	 */
	public String saveTextInstruction (String textVal, int userProfileId, String relatedPage) throws JCTException;
}
