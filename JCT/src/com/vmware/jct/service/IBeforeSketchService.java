package com.vmware.jct.service;

import java.util.List;

import com.vmware.jct.exception.JCTException;
import com.vmware.jct.model.JctQuestion;
import com.vmware.jct.service.vo.JctBeforeSketchDetailsVO;
import com.vmware.jct.service.vo.JctBeforeSketchHeaderVO;
/**
 * 
 * <p><b>Class name:</b>IBeforeSketchService.java</p>
 * <p><b>@author:</b> InterreIt </p>
 * <p><b>Purpose:</b> 		This Service interface represents a IBeforeSketchService interface
 * <p><b>Description:</b> 	This will have all methods where we implement all business logic and all dao layer </p>
 * <p><b>Copyrights:</b> 	All rights reserved by Interra IT and should only be used for its internal application development.</p>
 *  <pre>
 * <b>Revision History:</b>
 * InterreIt - 31/Jan/2014 - Implement exception through out the appliaction
 * </pre>
 */
public interface IBeforeSketchService {
	/**
	 * @param It will take JctBeforeSketchHeaderVO, JctBeforeSketchDetailsVO List  object as a input parameter
	 * @return It returns String Object
	 * @throws it throws JCTException
	 */
	public String saveOrUpdateBeforeSketch(JctBeforeSketchHeaderVO header, List<JctBeforeSketchDetailsVO> childList, String isEdit) throws JCTException;
	/**
	 * @param It will take jobReferenceNo as a input parameter
	 * @return It will returns string object
	 * @throws it throws JCTException
	 */
	public String getJson(String jobReferenceNo, String linkClicked) throws JCTException;
	/**
	 * @param It will take jobReferenceNo as a input parameter
	 * @return It will returns string object
	 * @throws it throws JCTException
	 */
	public String getJsonFromTemp(String jobReferenceNo, String linkClicked) throws JCTException;
	/**
	 * @param It will take jobReferenceNo as a input parameter
	 * @return It will returns string object
	 * @throws it throws JCTException
	 */
	public String getJsonForEdit(String jobReferenceNo, String linkClicked) throws JCTException;
	/**
	 * @param It will take jobReferenceNo as a input parameter
	 * @return It will returns string object
	 * @throws it throws JCTException
	 */
	public String getInitialJson(String jobReferenceNo, String linkClicked) throws JCTException;
	/**
	 * @param It will take jobReferenceNo as a input parameter
	 * @return It will returns string object
	 * @throws it throws JCTException
	 */
	public String getInitialJsonFromTemp(String jobReferenceNo, String linkClicked) throws JCTException;
	/**
	 * @param It will take jobReferenceNo as a input parameter
	 * @return It will returns string object
	 * @throws it throws JCTException
	 */
	public String getInitialJsonForEdit(String jobReferenceNo, String linkClicked) throws JCTException;
	/**
	 * @param It will take profileId as a input parameter
	 * @return It will returns Question list object
	 * @throws it throws JCTException
	 */
	public List<JctQuestion> fetchQuestion(int profileId) throws JCTException;
	/**
	 * @param It will take jobReferenceNo as a input parameter
	 * @return It will returns string object
	 * @throws it throws JCTException
	 */
	public List<String> fetchQuestionAnswer(String jobRefNo, int profileId) throws JCTException;
	/**
	 * @param It will take jobReferenceNo as a input parameter
	 * @return It will returns string object
	 * @throws it throws JCTException
	 */
	public String fetchBase64String(String jobReferenceNo) throws JCTException;
}
