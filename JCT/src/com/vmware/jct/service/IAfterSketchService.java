package com.vmware.jct.service;

import java.util.Date;
import java.util.List;

import com.vmware.jct.exception.DAOException;
import com.vmware.jct.exception.JCTException;
import com.vmware.jct.model.JctGlobalProfileHistory;
import com.vmware.jct.service.vo.JctAfterSketchFinalPageVO;
import com.vmware.jct.service.vo.JctAfterSketchHeaderVO;
import com.vmware.jct.service.vo.JctAfterSketchPageOneVO;

/**
 * 
 * <p><b>Class name:</b>IAfterSketchService.java</p>
 * <p><b>@author:</b> InterreIt </p>
 * <p><b>Purpose:</b> 		This Service interface represents a IAuthenticatorService interface
 * <p><b>Description:</b> 	This will have all methods where we implement all business logic and all dao layer </p>
 * <p><b>Copyrights:</b> 	All rights reserved by Interra IT and should only be used for its internal application development.</p>
 *  <pre>
 * <b>Revision History:</b>
 * InterreIt - 31/Jan/2014 - Implement exception through out the appliaction
 * </pre>
 */
public interface IAfterSketchService {
	/**
	 * @param It will take JctAfterSketchHeader VO object as a input parameter
	 * @return It returns string object
	 * @throws it throws JCTException
	 */
	public String saveOrUpdatePageOnde(JctAfterSketchHeaderVO header, List<JctAfterSketchPageOneVO> childList, String linkClicked) throws JCTException;
	/**
	 * @param It will take jobRefNo as a input parameter
	 * @return It returns string object
	 * @throws it throws JCTException
	 */
	public String fetchCheckedElements(String jobRefNo, String linkClicked) throws JCTException;
	/**
	 * @param It will take jobRefNo as a input parameter
	 * @return It returns string object
	 * @throws it throws JCTException
	 */
	public String fetchASOneJson(String jobRefNo, String linkClicked) throws JCTException;
	/**
	 * @param It will take jobRefNo as a input parameter
	 * @return It returns string object
	 * @throws it throws JCTException
	 */
	public String fetchASOneJsonForEdit(String jobRefNo, String linkClicked) throws JCTException;
	/**
	 * @param It will take jobRefNo as a input parameter
	 * @return It returns string object
	 * @throws it throws JCTException
	 */
	public String fetchASFinalPageJson(String jobRefNo, String linkClicked) throws JCTException;
	/**
	 * @param It will take jobRefNo as a input parameter
	 * @return It returns string object
	 * @throws it throws JCTException
	 */
	public String fetchASFinalPageJsonForEdit(String jobRefNo, String linkClicked) throws JCTException;
	/**
	 * @param It will take jobRefNo as a input parameter
	 * @return It returns List
	 * @throws it throws JCTException
	 */
	public int getPassionCount(String jobRefNo) throws JCTException;
	/**
	 * @param It will take jobRefNo as a input parameter
	 * @return It returns List
	 * @throws it throws JCTException
	 */
	public int getValueCount(String jobRefNo) throws JCTException;
	/**
	 * @param It will take jobRefNo as a input parameter
	 * @return It returns List
	 * @throws it throws JCTException
	 */
	public int getStrengthCount(String jobRefNo) throws JCTException;
	/**
	 * @param It will take JctAfterSketchHeader VO object as a input parameter
	 * @return It returns string object
	 * @throws it throws JCTException
	 */
	public String saveAfterSketchFinalPage(JctAfterSketchHeaderVO header, List<JctAfterSketchFinalPageVO> childList, String isCompleted, String isEdit, String linkClicked) throws JCTException;
	/**
	 * @param It will take Job reference number as a input parameter
	 * @return It returns list of string
	 * @throws it throws JCTException
	 */
	public List<String> fetchActionPlans(String jobRefNo, int profileId) throws JCTException;
	/**
	 * @param It will take Job reference number as a input parameter
	 * @return It returns list of string
	 * @throws it throws JCTException
	 */
	public String getDivHeightAndWidth(String refNo, String linkClicked) throws JCTException;
	/**
	 * @param It will take Job reference number as a input parameter
	 * @return It returns list of string
	 * @throws it throws JCTException
	 */
	public String getDivHeightAndWidthForEdit(String refNo, String linkClicked) throws JCTException;
	/**
	 * @param It will take Job reference number as a input parameter
	 * @return It returns list of string
	 * @throws it throws JCTException
	 */
	public String getCheckedItems(String refNo, String linkClicked) throws JCTException;
	/**
	 * @param It will take Job reference number as a input parameter
	 * @return It returns after sketch diagram base 64 string
	 * @throws it throws JCTException
	 */
	public String getFinalPageSnapShot(String refNo, String linkClicked) throws JCTException;
	/**
	 * @param It will take Job reference number as a input parameter
	 * @return It returns integer
	 * @throws it throws JCTException
	 */
	public Integer getAsTotalCount(String jobRefNo, String linkClicked) throws JCTException;
	/**
	 * @param It will take Job reference number as a input parameter
	 * @return It returns integer
	 * @throws it throws JCTException
	 */
	public Integer getAsTotalCountForEdit(String jobRefNo, String linkClicked) throws JCTException;
	/**
	 * @param It will take Job reference number as a input parameter
	 * @return It returns integer
	 * @throws it throws JCTException
	 */
	public String getFinalPageInitialJson(String jobRefNo,String linkClicked) throws JCTException;
	/**
	 * @param It will take Job reference number as a input parameter
	 * @return It returns integer
	 * @throws it throws JCTException
	 */
	public String getFinalPageInitialJsonForEdit(String jobRefNo,String linkClicked) throws JCTException;
	/**
	 * @param It will take Job reference number as a input parameter
	 * @return It returns integer
	 * @throws it throws JCTException
	 */
	public List<List> fetActionPlanDynamic (String jobRefNo, int profileId) throws JCTException;
	/**
	 * 
	 * @param attributeId
	 * @return
	 * @throws JCTException
	 */
	public String getFrozenAttributeById (int attributeId) throws JCTException;
}
