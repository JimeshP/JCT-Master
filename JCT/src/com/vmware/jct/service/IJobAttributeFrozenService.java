package com.vmware.jct.service;

import java.util.List;

import com.vmware.jct.dao.dto.JobAttributeDTO;
import com.vmware.jct.exception.JCTException;

/**
 * 
 * <p><b>Class name:</b>IJobAttributeFrozenService.java</p>
 * <p><b>@author:</b> InterraIt </p>
 * <p><b>Purpose:</b> 		This Service interface represents a IJobAttributeFrozenService interface
 * <p><b>Description:</b> 	This will have all methods where we implement all business logic and all dao layer </p>
 * <p><b>Copyrights:</b> 	All rights reserved by Interra IT and should only be used for its internal application development.</p>
 *  <pre>
 * <b>Revision History:</b>
 * InterraIt - 13/Oct/2014 - Implement exception through out the appliaction
 * </pre>
 */
public interface IJobAttributeFrozenService {
	/**
	 * 
	 * @param jrNo
	 * @param createdFor
	 * @param profileId
	 * @param userId
	 * @return
	 * @throws JCTException
	 */
	public String freezeContent (String jrNo, String createdFor, int profileId, int userId) throws JCTException;
	/**
	 * 
	 * @param profileId
	 * @param userName
	 * @return
	 * @throws JCTException
	 */
	public List<JobAttributeDTO> fetchJobAttribute (int profileId, String userName, int userId) throws JCTException;
}
