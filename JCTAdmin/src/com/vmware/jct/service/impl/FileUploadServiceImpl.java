package com.vmware.jct.service.impl;

import java.io.InputStream;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.Map;
import java.util.Set;
import java.util.TreeMap;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.context.MessageSource;
import org.springframework.stereotype.Service;
import org.springframework.transaction.annotation.Propagation;
import org.springframework.transaction.annotation.Transactional;

import com.vmware.jct.common.utility.FileUploadUtility;
import com.vmware.jct.common.utility.StatusConstants;
import com.vmware.jct.dao.ICommonDao;
import com.vmware.jct.dao.IFileUploadDAO;
import com.vmware.jct.exception.DAOException;
import com.vmware.jct.exception.JCTException;
import com.vmware.jct.model.JctOnetOccupationList;
import com.vmware.jct.service.IFileUploadService;
import com.vmware.jct.service.vo.FileUploadVO;
/**
 * 
 * <p><b>Class name:</b> FileUploadServiceImpl.java</p>
 * <p><b>Author:</b> InterraIT</p>
 * <p><b>Purpose:</b> This is a FileUploadServiceImpl class. This artifact is Business layer artifact.
 * FileUploadServiceImpl implement IFileUploadService interface and override the following  methods.
 * -saveOnetOccupationData()
 * </p>
 * <p><b>Description:</b> This class used to perform insert, modify ,fetch data from db and delete operation. Save, update, delete operations are performed in DataAccessObject class </p>
 * <p><b>Revision History:</b>
 * 	<li>InterraIT - 01/Oct/2014 - Implement Exception </li>
 * <li>InterraIT - 31/Oct/2014 - Implement comments, introduce Named query, Transactional attribute </li>
 * </p>
 */
@Service
public class FileUploadServiceImpl implements IFileUploadService {

	@Autowired
	private ICommonDao commonDaoImpl;
	
	@Autowired  
	private MessageSource messageSource;
	
	@Autowired
	private IFileUploadDAO serviceDAO;
	
	private static final Logger LOGGER = LoggerFactory.getLogger(ContentConfigServiceImpl.class);
	
	/**
	 * Method to fetch the ONet data from excel 
	 * and save in database
	 * @param file
	 * @param fileName
	 * @return fileUploadVO
	 */
	@Transactional(propagation = Propagation.REQUIRED)
	public FileUploadVO saveOnetOccupationData(InputStream file, String fileName)
			throws JCTException {
		LOGGER.info(">>>>>> FileUploadServiceImpl.saveOnetOccupationData");
		FileUploadVO fileUploadVO = new FileUploadVO();
		
		FileUploadUtility FileUpload = new FileUploadUtility();
		Map<String, ArrayList<String>> OnetOccupationMap = new TreeMap<String, ArrayList<String>>();
		OnetOccupationMap = FileUpload.readOnetOccupationExcel(file, fileName);
		
		if(OnetOccupationMap.isEmpty()) {
			fileUploadVO.setStatusCode(StatusConstants.STATUS_SUCCESS_WITH_RESULT);
			fileUploadVO.setStatusDesc(this.messageSource.
					getMessage("file.upload.wrong.format", null, null));
		} else {
			 ArrayList<String> arraylist = new ArrayList<String>();
			// Get a set of the entries
			 Set allMapData = OnetOccupationMap.entrySet();
			// Get an iterator
		      Iterator i = allMapData.iterator();
		      while(i.hasNext()) {
		          Map.Entry me = (Map.Entry)i.next();	          	         
		          try {
					String occupationTitle = serviceDAO.getOccupationTitle((String) me.getKey());								
					if(null == occupationTitle || occupationTitle.equalsIgnoreCase("null")) { // New entry
						JctOnetOccupationList onetOccupationData = new JctOnetOccupationList();
						 onetOccupationData.setJctOnetOccupationCode((String) me.getKey());						 
						 arraylist = OnetOccupationMap.get((String) me.getKey());
						 onetOccupationData.setJctOnetOccupationTitle(arraylist.get(0));
						 onetOccupationData.setJctOnetOccupationDesc(arraylist.get(1));						 
						// onetOccupationData.setJctOnetOccupationTitle((String) me.getValue());
						 if( onetOccupationData.getJctOnetOccupationCode().equalsIgnoreCase("1") &&
								 onetOccupationData.getJctOnetOccupationTitle().equalsIgnoreCase("1")) {
							 fileUploadVO.setStatusCode(StatusConstants.STATUS_SUCCESS_WITH_RESULT);
								fileUploadVO.setStatusDesc(this.messageSource.
										getMessage("file.upload.wrong.format", null, null));
						 } else {
							 String retString = serviceDAO.saveONetData(onetOccupationData);
							 if (!retString.equals("success")) {
									fileUploadVO.setStatusCode(StatusConstants.STATUS_FAILURE);
									fileUploadVO.setStatusDesc(this.messageSource.
											getMessage("file.upload.error.msg", null, null));
								} else {
									fileUploadVO.setStatusCode(StatusConstants.STATUS_SUCCESS_WITH_RESULT);
									fileUploadVO.setStatusDesc(this.messageSource.
											getMessage("file.upload.success.msg", null, null));
								} 
						 }						
					} else { // update
						JctOnetOccupationList onetOccupationData = serviceDAO.getOccupationObj((String) me.getKey());
						arraylist = OnetOccupationMap.get((String) me.getKey());
						if(!onetOccupationData.getJctOnetOccupationTitle().equalsIgnoreCase(arraylist.get(0))) {
							 onetOccupationData.setJctOnetOccupationTitle(arraylist.get(0));
							 onetOccupationData.setJctOnetOccupationDesc(arraylist.get(1));							 
							//onetOccupationData.setJctOnetOccupationTitle((String) me.getValue());
							 String retString = serviceDAO.updateONetData(onetOccupationData);
							 if( onetOccupationData.getJctOnetOccupationCode().equalsIgnoreCase("1") &&
									 onetOccupationData.getJctOnetOccupationTitle().equalsIgnoreCase("1")) {
								 fileUploadVO.setStatusCode(StatusConstants.STATUS_SUCCESS_WITH_RESULT);
									fileUploadVO.setStatusDesc(this.messageSource.
											getMessage("file.upload.wrong.format", null, null));
							 }
							 else {
								 if (!retString.equals("success")) {
										fileUploadVO.setStatusCode(StatusConstants.STATUS_FAILURE);
										fileUploadVO.setStatusDesc(this.messageSource.
												getMessage("file.upload.error.msg", null, null));
									} else {
										fileUploadVO.setStatusCode(StatusConstants.STATUS_SUCCESS_WITH_RESULT);
										fileUploadVO.setStatusDesc(this.messageSource.
												getMessage("file.upload.success.msg", null, null));
									}
							 }						
						}					 
					}
		          } catch (DAOException e) {
		        	  LOGGER.error("----"+e.getLocalizedMessage()+" ----");
		        	  fileUploadVO.setStatusCode(StatusConstants.STATUS_FAILURE);
		  			  throw new JCTException(e.getLocalizedMessage());
		          }
		         	          
		       }
		}		
	     LOGGER.info("<<<<<<<<< FileUploadServiceImpl.saveOnetOccupationData");
		return fileUploadVO;
	}
}
