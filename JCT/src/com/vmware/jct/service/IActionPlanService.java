package com.vmware.jct.service;

import java.util.List;

import com.vmware.jct.exception.JCTException;
import com.vmware.jct.service.vo.ActionPlanVO;
/**
 * 
 * <p><b>Class name:</b>IActionPlanService.java</p>
 * <p><b>@author:</b> InterreIt </p>
 * <p><b>Purpose:</b> 		This Service interface represents a IActionPlanService interface
 * <p><b>Description:</b> 	This will have all methods where we implement all business logic and all dao layer </p>
 * <p><b>Copyrights:</b> 	All rights reserved by Interra IT and should only be used for its internal application development.</p>
 *  <pre>
 * <b>Revision History:</b>
 * </pre>
 */
public interface IActionPlanService {
	/**
	 * 
	 * @param child
	 * @param jobRefNo
	 * @return
	 * @throws JCTException
	 */
	public String saveActinoPlans(List<ActionPlanVO> child, String jobRefNo, String linkClicked) throws JCTException;
	/**
	 * 
	 * @param jobReferenceNos
	 * @return
	 * @throws JCTException
	 */
	public int getNumberOfCompletion (String jobReferenceNos) throws JCTException;
	/**
	 * 
	 * @param jrNo
	 * @param user
	 * @param userId
	 * @param profileId
	 * @return
	 * @throws JCTException
	 */
	public String freezeActionPlanContent (String jrNo, String user, int userId, int profileId) throws JCTException;
	/**
	 * 
	 * @param jrNo
	 * @param user
	 * @param userId
	 * @param profileId
	 * @return
	 * @throws JCTException
	 */
	public int getActionPlanCount (String jrNo, String user, int userId, int profileId) throws JCTException;
}
