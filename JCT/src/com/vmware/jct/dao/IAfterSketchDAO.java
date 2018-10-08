package com.vmware.jct.dao;

import java.util.Date;
import java.util.List;

import com.vmware.jct.dao.dto.QuestionearDTO;
import com.vmware.jct.exception.DAOException;
import com.vmware.jct.model.JctAfterSketchFinalpageDetail;
import com.vmware.jct.model.JctAfterSketchFinalpageDetailTemp;
import com.vmware.jct.model.JctAfterSketchHeader;
import com.vmware.jct.model.JctAfterSketchHeaderTemp;
import com.vmware.jct.model.JctAfterSketchPageoneDetail;
import com.vmware.jct.model.JctGlobalProfileHistory;
import com.vmware.jct.model.JctUserLoginInfo;
import com.vmware.jct.service.vo.UserVO;

/**
 * 
 * <p><b>Class name:</b>IAfterSketchDAO.java</p>
 * <p><b>@author:</b> InterraIT </p>
 * <p><b>Purpose:</b> 		This DAO interface represents a IAfterSketchDAO interface
 * <p><b>Description:</b> 	This will have all methods mapped to the database and fetching data from database to be used across the application </p>
 * <p><b>Copyrights:</b> 	All rights reserved by Interra IT and should only be used for its internal application development.</p>
 *  <pre>
 * <b>Revision History:</b>
 * InterraIt - 31/Jan/2014 - Implement comments through out the application
 * </pre>
 */


public interface IAfterSketchDAO {
	
	/**
	 * @param It will take JctAfterSketchHeader object as a input parameter
	 * @return It will returns string object
	 	 * @throws DAOException -it throws DAOException
	 */
	public String saveOrUpdatePageOne(JctAfterSketchHeader entity) throws DAOException;
	/**
	 * @param It will take jobRefNo as a String input parameter
	 * @return It returns list of  JctAfterSketchHeader data from database
	 * @throws DAOException -it throws DAOException
	 */
	public List<JctAfterSketchHeader> getList(String jobRefNo, String linkClicked) throws DAOException;
	/**
	 * @param It will take jobRefNo as a String input parameter
	 * @return It returns list of  JctAfterSketchHeader data from database
	 * @throws DAOException -it throws DAOException
	 */
	public List<JctAfterSketchHeader> getListForEdit(String jobRefNo, String linkClicked) throws DAOException;
	/**
	 * @param It will take jobRefNo object as a input parameter
	 * @return It will returns JctAfterSketchHeader list object
	 * @throws DAOException -it throws DAOException
	 */
	public List<JctAfterSketchHeader> getListForKeepingPrevious(String jobRefNo) throws DAOException;
	/**
	 * @param It will take jobRefNo as a String input parameter
	 * @return It returns list of  JctAfterSketchHeader data from database
	 * @throws DAOException -it throws DAOException
	 */
	public List<JctAfterSketchHeaderTemp> getListTemp(String jobRefNo) throws DAOException;
	/**
	 * @param It will take jobRefNo as a String input parameter
	 * @return It returns string object
	 * @throws DAOException -it throws DAOException
	 */
	public String fetchCheckedElements(String jobRefNo, String linkClicked) throws DAOException;
	/**
	 * <p><b>Description:</b>  Based on job reference number this method fetch page one Json value from jct after sketch header table</p>
	 * @param It will take jobRefNo as a input parameter
	 * @return It returns string object
	 * @throws DAOException -it throws DAOException
	 * 
	 */
	public String fetchASOneJson(String jobRefNo, String linkClicked) throws DAOException;
	/**
	 * <p><b>Description:</b>  Based on job reference number this method fetch page one Json value from jct after sketch header table</p>
	 * @param It will take jobRefNo as a input parameter
	 * @return It returns string object
	 * @throws DAOException -it throws DAOException
	 * 
	 */
	public String fetchASOneJsonForEdit(String jobRefNo, String linkClicked) throws DAOException;
	/**
	 * <p><b>Description:</b>  Based on job reference number this method fetch final page Json value from jct after sketch header table</p>
	 * @param It will take jobRefNo as a input parameter
	 * @return It returns string object
	 * @throws DAOException -it throws DAOException
	 * 
	 */
	public String fetchASFinalPageJson(String jobRefNo, String linkClicked) throws DAOException;
	/**
	 * <p><b>Description:</b>  Based on job reference number this method fetch final page Json value from jct after sketch header table</p>
	 * @param It will take jobRefNo as a input parameter
	 * @return It returns string object
	 * @throws DAOException -it throws DAOException
	 * 
	 */
	public String fetchASFinalPageJsonForEdit(String jobRefNo, String linkClicked) throws DAOException;
	/**
	 * @param It will take JctAfterSketchHeader object as a input parameter
	 * @return It will returns string object
	 * @throws DAOException -it throws DAOException
	 */
	public String remove(JctAfterSketchHeader object) throws DAOException;
	/**
	 * @param It will take JctAfterSketchHeader object as a input parameter
	 * @return It will returns string object
	 * @throws DAOException -it throws DAOException
	 */
	public String removeTemp(JctAfterSketchHeaderTemp object) throws DAOException;
	/**
	 * @param It will take JctAfterSketchFinalpageDetail object as a input parameter
	 * @return It will returns string object
	 * @throws DAOException -it throws DAOException
	 */
	public String removeDetails(JctAfterSketchFinalpageDetail object) throws DAOException;
	/**
	 * @param It will take JctAfterSketchFinalpageDetail object as a input parameter
	 * @return It will returns string object
	 * @throws DAOException -it throws DAOException
	 */
	public String removeDetailsTemp(JctAfterSketchFinalpageDetailTemp object) throws DAOException;
	/**
	 * <p><b>Description:</b>  Based on job reference number this method fetches fetches List of passion count</p>
	 * @param It will take jobRefNo as a input parameter
	 * @return It returns Integer object
	 * @throws DAOException -it throws DAOException
	 * 
	 */
	public List<Integer> getPassionCount(String jobRefNo) throws DAOException;
	/**
	 * <p><b>Description:</b>  Based on job reference number this method fetches List of value count</p>
	 * @param It will take jobRefNo as a input parameter
	 * @return It returns Integer object
	 * @throws DAOException -it throws DAOException
	 * 
	 */
	public List<Integer> getValueCount(String jobRefNo) throws DAOException;
	/**
	 * <p><b>Description:</b>  Based on job reference number this method fetches List of strength count</p>
	 * @param It will take jobRefNo as a input parameter
	 * @return It returns Integer object
	 * @throws DAOException -it throws DAOException
	 * 
	 */
	public List<Integer> getStrengthCount(String jobRefNo) throws DAOException;
	/**
	 * <p><b>Description:</b> This method is used for update existing After Sketch header data and insert new 
	 *final page rows</p>
	 * @param It will take JctAfterSketchHeader as a input parameter
	 * @return It returns Integer object
	 * @throws DAOException -it throws DAOException
	 * 
	 */
	public String mergeAfterSketchDiag(JctAfterSketchHeader entity) throws DAOException;
	
	/**
	 * @param It will take jobTitle object as a input parameter
	 * @return It will returns QuestionearDTO object
	 * @throws DAOException -it throws DAOException
	 */
	public List<QuestionearDTO> getFetchQuestion(int profileId) throws DAOException;
	/**
	 * @param It will take jobRefNo as a input parameter
	 * @return It will returns List of AS final page
	 * @throws DAOException -it throws DAOException
	 */
	public List<JctAfterSketchFinalpageDetail> getASFinalPageList(String jobRefNo) throws DAOException;
	/**
	 * @param It will take jobRefNo as a input parameter
	 * @return It will returns List of AS final page
	 * @throws DAOException -it throws DAOException
	 */
	public List<JctAfterSketchFinalpageDetailTemp> getASFinalPageListTemp(String jobRefNo) throws DAOException;
	/**
	 * @param It will take jobRefNo as a input parameter
	 * @return It will returns List of AS page one
	 * @throws DAOException -it throws DAOException
	 */
	public List<JctAfterSketchPageoneDetail> getASPageOneList(String jobRefNo) throws DAOException;
	/**
	 * 
	 * @param It will take jobRefNo as a input parameter
	 * @return List<JctAfterSketchPageoneDetail>
	 * @throws DAOException
	 */
	public List<JctAfterSketchPageoneDetail> getASFirstPageList(String jobRefNo) throws DAOException;
	/**
	 * 
	 * @param It will take jobRefNo as a input parameter
	 * @return Integer
	 * @throws DAOException
	 */
	public Integer calculateDifference(String jobReferenceNos) throws DAOException;
	/**
	 * 
	 * @param It will take jobRefNo as a input parameter
	 * @return String
	 * @throws DAOException
	 */
	public String changeStatus(String jobReferenceNos) throws DAOException;
	/**
	 * 
	 * @param It will take jobRefNo as a input parameter
	 * @return Integer
	 * @throws DAOException
	 */
	public Integer getAsTotalCount(String jobRefNo, String linkClicked) throws DAOException;
	/**
	 * 
	 * @param It will take jobRefNo as a input parameter
	 * @return Integer
	 * @throws DAOException
	 */
	public Integer getAsTotalCountForEdit(String jobRefNo, String linkClicked) throws DAOException;
	/**
	 * 
	 * @param It will take jobRefNo as a input parameter
	 * @return string
	 * @throws DAOException
	 */
	public String removePageOneDetails(JctAfterSketchPageoneDetail object) throws DAOException;
	/**
	 * <p><b>Description:</b>  Based on job ref no this method fetch after sketch final page's initial jon</p>
	 * @param It will take jobReferenceNos as a input parameter
	 * @return It returns string object
	 * @throws DAOException -it throws DAOException
	 * 
	 */
	public String getFinalPageInitialJson(String jobRefNo, String linkClicked) throws DAOException;
	/**
	 * <p><b>Description:</b>  Based on job ref no this method fetch after sketch final page's initial jon</p>
	 * @param It will take jobReferenceNos as a input parameter
	 * @return It returns string object
	 * @throws DAOException -it throws DAOException
	 * 
	 */
	public String getFinalPageInitialJsonForEdit(String jobRefNo, String linkClicked) throws DAOException;
	/**
	 * 
	 * @param It will take jobRefNo  and Task id as a input parameter
	 * @return Object list
	 * @throws DAOException
	 */
	public List<Object> getBsToAsDifference(String jobRefNo, int taskId) throws DAOException;
	/**
	 * 
	 * @param taskId
	 * @param jobReferenceNumber
	 * @return
	 * @throws DAOException
	 */
	public List<String> getRoleFrames(int taskId, String jobReferenceNumber) throws DAOException;
	
	/**
	 * 
	 * @param It will take jobRefNo   as a input parameter
	 * @return JctAfterSketchFinalpageDetail list
	 * @throws DAOException
	 */
	public List<JctAfterSketchFinalpageDetail> getASFinalPageListDescOrder(String jobRefNo) throws DAOException;
	/**
	 * 
	 * @param jobRefNo
	 * @return
	 * @throws DAOException
	 */
	public List<String> getDistinctRoleFrames(String jobRefNo) throws DAOException;
	
	/**
	 * 
	 * @param It will take profileId   as a input parameter
	 * @return string list
	 * @throws DAOException
	 */
	public List<String> getActionPlanMainQtnList (int profileId) throws DAOException;
	/**
	 * 
	 * @param It will take mainQtn and profileId   as a input parameter
	 * @return string list
	 * @throws DAOException
	 */
	public List<String> getActionPlanSubQtnList (String mainQtn, int profileId) throws DAOException;
	/**
	 * 
	 * @param It will take JctAfterSketchFinalpageDetail as a input parameter
	 * @return string list
	 * @throws DAOException
	 */
	public String updateFinalPageDetails(JctAfterSketchFinalpageDetail entity) throws DAOException;
	/**
	 * @param It will take JctAfterSketchHeaderTemp object as a input parameter
	 * @return It will returns string object
	 * @throws DAOException -it throws DAOException
	 */
	public String saveOrUpdateAsHeaderTemp(JctAfterSketchHeaderTemp entity) throws DAOException;
	/**
	 * @param It will take JctAfterSketchHeader object as a input parameter
	 * @return It will returns string object
	 * @throws DAOException -it throws DAOException
	 */
	public String updateHeaderPage(JctAfterSketchHeader entity) throws DAOException;
	/**
	 * @param It will take jobRefNo object as a input parameter
	 * @param It will take seqNos as input parameter
	 * @return It will returns JctBeforeSketchHeader list object
	 * @throws DAOException -it throws DAOException
	 */
	public List<JctAfterSketchHeader> getASMaxPageSeqHeader(String jobRefNo, int seqNos) throws DAOException;
	/**
	 * 
	 * @param userVo
	 * @return
	 * @throws DAOException
	 */
	public List<JctAfterSketchHeader> getAfterSketchActiveObject(UserVO userVo) throws DAOException;
	/**
	 * 
	 * @param jobRefNo
	 * @return
	 * @throws DAOException
	 */
	public List<Object> getBsToAsDeletedTask(String jobRefNo) throws DAOException;
	/**
	 * 
	 * @param jobRefNo
	 * @param taskDesc
	 * @return List<String>
	 * @throws DAOException
	 */
	public List<String> getDistinctRoleName (String jobRefNo, String taskDesc) throws DAOException;
	/**
	 * 
	 * @param attributeId
	 * @return
	 * @throws DAOException
	 */
	public String getFrozenAttributeById(int attributeId) throws DAOException;
	/**
	 * 
	 * @param jobRefNo
	 * @return
	 * @throws DAOException
	 */
	public List<JctUserLoginInfo> getLoginInfoList (String jobRefNo) throws DAOException;
	/**
	 * 
	 * @param headerDate
	 * @param subQtn
	 * @param mainQtn
	 * @return
	 * @throws DAOException
	 */
	public List<JctGlobalProfileHistory> getGlobalProfileChangedData(Date headerDate) throws DAOException;
}
