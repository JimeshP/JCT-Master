package com.vmware.jct.service;

import com.vmware.jct.exception.JCTException;
/**
 * 
 * <p><b>Class name:</b>ISurveyQuestionService.java</p>
 * <p><b>@author:</b> InterreaIt </p>
 * <p><b>Purpose:</b> 		This Service interface represents a ISurveyQuestionService interface
 * <p><b>Description:</b> 	This will have all methods where we implement all business logic and all dao layer </p>
 * <p><b>Copyrights:</b> 	All rights reserved by Interra IT and should only be used for its internal application development.</p>
 *  <pre>
 * <b>Revision History:</b>
 * InterreaIt - 14/Nov/2014 - Implement exception through out the appliaction
 * </pre>
 */
public interface ISurveyQuestionService {
	/**
	 * Method saves all the survey questions.
	 * @param text
	 * @param radio
	 * @param drop
	 * @param check
	 * @param emailId
	 * @param userId
	 * @return
	 * @throws JCTException
	 */
	public String saveSurveyQuestions(String text, String radio, String drop, String check, String emailId, int userId) throws JCTException;
}
