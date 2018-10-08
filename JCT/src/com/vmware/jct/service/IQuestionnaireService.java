package com.vmware.jct.service;

import java.util.List;

import com.vmware.jct.exception.JCTException;
import com.vmware.jct.dao.dto.JobAttributeDTO;
import com.vmware.jct.service.vo.QuestionnaireVO;
/**
 * 
 * <p><b>Class name:</b>IQuestionnaireService.java</p>
 * <p><b>@author:</b> InterraIt </p>
 * <p><b>Purpose:</b> 		This Service interface represents a IAuthenticatorService interface
 * <p><b>Description:</b> 	This will have all methods where we implement all business logic and all dao layer </p>
 * <p><b>Copyrights:</b> 	All rights reserved by Interra IT and should only be used for its internal application development.</p>
 *  <pre>
 * <b>Revision History:</b>
 * InterraIt - 31/Jan/2014 - Implement exception through out the appliaction
 * </pre>
 */
public interface IQuestionnaireService {
	/**
	 * @param It will take QuestionnaireVO and  jobRefNumber as a input parameter
	 * @return It returns string object
	 * @throws it throws JCTException
	 */
	public String saveOrUpdateAnswers(List<QuestionnaireVO> vo, String jobRefNumber, String isCompleted) throws JCTException;
	/**
	 * @param int profileId
	 * @return It returns JobAttributeDTO object
	 * @throws it throws JCTException
	 */
	public List<JobAttributeDTO> fetchJobAttribute(int profileId) throws JCTException;
	/**
	 * @param It will take Job reference number as a input parameter
	 * @return It returns integer
	 * @throws it throws JCTException
	 */
	public List<List> fetQuestionnaireDynamic (String jobRefNo, int profileId) throws JCTException;
	/**
	 * 
	 * @param jrNo
	 * @param userName
	 * @param userId
	 * @param profileId
	 * @return
	 * @throws JCTException
	 */
	public String freezeQuestionnaireContent (String jrNo, String user, int userId, int profileId) throws JCTException;
	/**
	 * 
	 * @param jrNo
	 * @param user
	 * @param userId
	 * @param profileId
	 * @return
	 * @throws JCTException
	 */
	public int getQuestCount (String jrNo, String user, int userId, int profileId) throws JCTException;
}
