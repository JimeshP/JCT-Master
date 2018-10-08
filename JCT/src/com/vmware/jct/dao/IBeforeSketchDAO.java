package com.vmware.jct.dao;

import java.util.List;

import com.vmware.jct.exception.DAOException;
import com.vmware.jct.model.JctBeforeSketchHeader;
import com.vmware.jct.model.JctBeforeSketchHeaderTemp;
import com.vmware.jct.model.JctBeforeSketchQuestion;
import com.vmware.jct.model.JctQuestion;
import com.vmware.jct.service.vo.UserVO;

/**
 * 
 * <p><b>Class name:</b>IBeforeSketchDAO.java</p>
 * <p><b>@author:</b> InterraIt </p>
 * <p><b>Purpose:</b> 		This DAO interface represents a IQuestionnaireDAO interface
 * <p><b>Description:</b> 	This will have all methods mapped to the database and fetching data from database to be used across the application </p>
 * <p><b>Copyrights:</b> 	All rights reserved by Interra IT and should only be used for its internal application development.</p>
 *  <pre>
 * <b>Revision History:</b>
 * InterraIt - 31/Jan/2014 - Implement comments through out the application
 * </pre>
 */
public interface IBeforeSketchDAO {
	/**
	 * @param It will take JctBeforeSketchHeader object as a input parameter
	 * @return It will returns string object
	 	 * @throws DAOException -it throws DAOException
	 */
	public String saveOrUpdateBeforeSketch(JctBeforeSketchHeader entity) throws DAOException;
	/**
	 * @param It will take JctBeforeSketchHeaderTemp object as a input parameter
	 * @return It will returns string object
	 	 * @throws DAOException -it throws DAOException
	 */
	public String saveOrUpdateBeforeSketchTemp(JctBeforeSketchHeaderTemp entity) throws DAOException;
	/**
	 * @param It will take jobRefNo object as a input parameter
	 * @return It will returns JctBeforeSketchHeader list object
	 * @throws DAOException -it throws DAOException
	 */
	public List<JctBeforeSketchHeader> getList(String jobRefNo) throws DAOException;
	/**
	 * @param It will take jobRefNo object as a input parameter
	 * @return It will returns JctBeforeSketchHeader list object
	 * @throws DAOException -it throws DAOException
	 */
	public List<JctBeforeSketchHeader> getListForKeepingPrevious(String jobRefNo) throws DAOException;
	/**
	 * @param It will take jobRefNo object as a input parameter
	 * @param It will take seqNos as input parameter
	 * @return It will returns JctBeforeSketchHeader list object
	 * @throws DAOException -it throws DAOException
	 */
	public List<JctBeforeSketchHeader> getBSMaxPageSeqHeader(String jobRefNo, int seqNos) throws DAOException;
	/**
	 * @param It will take jobRefNo object as a input parameter
	 * @return It will returns JctBeforeSketchHeaderTemp list object
	 * @throws DAOException -it throws DAOException
	 */
	public List<JctBeforeSketchHeaderTemp> getListTemp(String jobRefNo) throws DAOException;
	/**
	 * @param It will take jobRefNo object as a input parameter
	 * @return It will returns String object
	 * @throws DAOException -it throws DAOException
	 */
	public String getJson(String jobReferenceNo, String linkClicked) throws DAOException;
	/**
	 * @param It will take jobRefNo object as a input parameter
	 * @return It will returns String object
	 * @throws DAOException -it throws DAOException
	 */
	public String getJsonForEdit(String jobReferenceNo, String linkClicked) throws DAOException;
	/**
	 * @param It will take jobRefNo object as a input parameter
	 * @return It will returns String object
	 * @throws DAOException -it throws DAOException
	 */
	public String getJsonForEditFromTemp(String jobReferenceNo, String linkClicked) throws DAOException;
	/**
	 * @param It will take jobRefNo object as a input parameter
	 * @return It will returns String object
	 * @throws DAOException -it throws DAOException
	 */
	public String getInitialJson(String jobReferenceNo, String linkClicked) throws DAOException;
	/**
	 * @param It will take jobRefNo object as a input parameter
	 * @return It will returns String object
	 * @throws DAOException -it throws DAOException
	 */
	public String getInitialJsonFromTemp(String jobReferenceNo, String linkClicked) throws DAOException;
	/**
	 * @param It will take jobRefNo object as a input parameter
	 * @return It will returns String object
	 * @throws DAOException -it throws DAOException
	 */
	public String getInitialJsonForEdit(String jobReferenceNo, String linkClicked) throws DAOException;
	/**
	 * @param It will take profileId object as a input parameter
	 * @return It will returns JctQuestion object
	 * @throws DAOException -it throws DAOException
	 */
	public List<JctQuestion> getFetchQuestion(int profileId) throws DAOException;
	/**
	 * @param It will take jobTitle object as a input parameter
	 * @return It will returns JctBeforeSketchQuestion object
	 * @throws DAOException -it throws DAOException
	 */
	public List<JctBeforeSketchQuestion> fetchQuestionAnswer(String jobRefNo) throws DAOException;
	/**
	 * Method removes the parent object and changes cascades
	 * @param object
	 * @return
	 * @throws DAOException
	 */
	public String remove(JctBeforeSketchHeader object) throws DAOException;
	/**
	 * Method removes the parent object and changes cascades
	 * @param object
	 * @return
	 * @throws DAOException
	 */
	public String removeTemp(JctBeforeSketchHeaderTemp object) throws DAOException;
	/**
	 * 
	 * @param userVo
	 * @return
	 * @throws DAOException
	 */
	public List<JctBeforeSketchHeader> getBeforeSketchActiveObject(UserVO userVo) throws DAOException;
	
	public List<JctBeforeSketchHeader> getListEdtd(String jobRefNo) throws DAOException;
}
