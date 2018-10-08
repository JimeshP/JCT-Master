package com.vmware.jct.service;

import java.io.InputStream;

import com.vmware.jct.exception.JCTException;
import com.vmware.jct.service.vo.FileUploadVO;
/**
 * 
 * <p><b>Class name:</b>IFileUploadService.java</p>
 * <p><b>@author:</b> InterraIT </p>
 * <p><b>Purpose:</b> 		This Service interface represents a IFileUploadService interface
 * <p><b>Description:</b> 	This will have all methods where we implement all business logic and all service layer </p>
 * <p><b>Copyrights:</b> 	All rights reserved by InterraIT and should only be used for its internal application development.</p>
 *  <pre>
 * <b>Revision History:</b>
 * InterraIT - 10/Oct/2014 - Implement exception through out the application
 * </pre>
 */
public interface IFileUploadService {

	/**
	 * Method to save the ONet data 
	 * @param file
	 * @param fileName
	 * @return fileUploadVO
	 * @throws JCTException
	 */
	public FileUploadVO saveOnetOccupationData(InputStream file, String fileName) throws JCTException ;
}
