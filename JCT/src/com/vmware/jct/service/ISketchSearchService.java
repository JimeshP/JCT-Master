package com.vmware.jct.service;

import java.util.List;

import com.vmware.jct.dao.dto.SearchPdfDTO;
import com.vmware.jct.exception.JCTException;
import com.vmware.jct.service.vo.MyAccountVO;
import com.vmware.jct.service.vo.SearchResultVO;


/**
 * 
 * <p><b>Class name:</b>ISketchSearchService.java</p>
 * <p><b>@author:</b> InterraIt </p>
 * <p><b>Purpose:</b> 		This Service interface represents a ISketchSearchService interface
 * <p><b>Description:</b> 	This will have all methods where we implement all search funcionalities </p>
 * <p><b>Copyrights:</b> 	All rights reserved by Interra IT and should only be used for its internal application development.</p>
 * <p><b>Created Date:</b> 05/Feb/2014</p>
 * <b>Revision History:</b>
 * <li></li>
 */
public interface ISketchSearchService {
	/**
	 * 
	 * @param jobTitle
	 * @param firstResult
	 * @param maxResults
	 * @return SearchResultVO object
	 * @throws JCTException
	 */
	public SearchResultVO getBeforeSketchList (String jobTitle, int firstResult, int maxResults) throws JCTException;
	/**
	 * 
	 * @param jobTitle
	 * @param firstResult
	 * @param maxResults
	 * @return SearchResultVO
	 * @throws JCTException
	 */
	public SearchResultVO getAfterSketchList (String jobTitle, int firstResult, int maxResults) throws JCTException;
	/**
	 * 
	 * @param emailId
	 * @return MyAccountVO
	 * @throws JCTException
	 */
	
	public MyAccountVO getAccountDetails (String emailId) throws JCTException;
	/**
	 * 
	 * @param jobReferenceNo
	 * @return SearchResultVO
	 * @throws JCTException
	 */
	public SearchResultVO getSearchTabsTobeViewed (String jobReferenceNo) throws JCTException;
	/**
	 * 
	 * @param jobReferenceNo,diagDistinction
	 * @return SearchResultVO
	 * @throws JCTException
	 */
	public SearchResultVO getSelectedDiagrams (String jobReferenceNo, String diagDistinction) throws JCTException;
	/**
	 * 
	 * @param jobReferenceNo,diagDistinction
	 * @return SearchResultVO
	 * @throws JCTException
	 */
	public SearchResultVO getPrevDiag (String jobReferenceNo, String distinction) throws JCTException;
	/**
	 * 
	 * @param distinction
	 * @param firstResult
	 * @param maxResults
	 * @return
	 * @throws JCTException
	 */
	public SearchResultVO getAllDiagram (String distinction, int firstResult, int maxResults) throws JCTException;
	
	public SearchResultVO getAllDiagram (String distinction, int firstResult, int maxResults, String Code) throws JCTException;
	
	public Long fetchAllDiagramCount(String distinction) throws JCTException;
	
	public Long fetchAllDiagramCount(String distinction, String code) throws JCTException;
	
	public List<SearchPdfDTO> fetchOldPdf(int userId) throws JCTException;
}
