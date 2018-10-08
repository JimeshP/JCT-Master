package com.vmware.jct.controller;


import java.io.File;
import java.io.FileInputStream;
import java.io.InputStream;

import javax.servlet.ServletOutputStream;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;

import org.codehaus.jackson.JsonNode;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.context.MessageSource;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.PathVariable;
import org.springframework.web.bind.annotation.RequestBody;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestMethod;
import org.springframework.web.bind.annotation.ResponseBody;

import com.vmware.jct.common.utility.CommonConstants;
import com.vmware.jct.common.utility.CommonUtility;
import com.vmware.jct.exception.JCTException;
import com.vmware.jct.service.ISketchSearchService;
import com.vmware.jct.service.vo.MyAccountVO;
import com.vmware.jct.service.vo.PdfSearchVO;
import com.vmware.jct.service.vo.ReturnVO;
import com.vmware.jct.service.vo.SearchResultVO;
/**
 * 
 * <p><b>Class name:</b> SketchSearchController.java</p>
 * <p><b>Author:</b> InterraIT</p>
 * <p><b>Purpose:</b> This class acts as a controller Before and After Sketch diagram search.
 * SketchSearchController extends BasicController  and has following  methods.
 * -searchBeforeSketch()
 * -searchAfterSketch()
 * <p><b>Description:</b> This class acts as a controller Before and After Sketch diagram search. </p>
 * <p><b>Copyrights:</b> 	All rights reserved by Interra IT and should only be used for its internal application development.</p>
 * <p><b>Created Date:</b> 05/Feb/2014</p>
 * <p><b>Revision History:</b>
 * 	<li></li>
 * </p>
 */
@Controller
@RequestMapping("/search")
public class SketchSearchController extends BasicController {
	
	@Autowired
	private ISketchSearchService searchService;
	
	@Autowired  
	private MessageSource messageSource;
	
	private static final Logger logger = LoggerFactory.getLogger(SketchSearchController.class);
	
	/**
	 * Method searches the before Sketch diagrams which are submitted for search.
	 * @param body
	 * @param request
	 * @return List of base64 string
	 */
	@RequestMapping(value = ACTION_SEARCH_BEFORE_SKETCH, method = RequestMethod.POST)
	@ResponseBody()
	public SearchResultVO searchBeforeSketch(@RequestBody String body, HttpServletRequest request) {
		logger.info(">>>> SketchSearchController.storeBeforeSketch");
		JsonNode node = CommonUtility.getNode(body);
		String jobTitle = node.get("jobTitle").toString().replaceAll("\"", "");
		int firstResult = Integer.parseInt(node.get("firstResult").toString().replaceAll("\"", ""));
		int maxResults = Integer.parseInt(node.get("maxResults").toString().replaceAll("\"", ""));
		SearchResultVO returnVO = null;
		try{
			returnVO = searchService.getBeforeSketchList(jobTitle, firstResult, maxResults);
			//returnVO.setStatusCode("200");
			returnVO.setStatusCode(CommonConstants.SUCCESSFUL);
			returnVO.setStatusMsg("Success");
		}catch(JCTException jctExp){
			returnVO = new SearchResultVO();
			returnVO.setStatusCode(jctExp.getMessage());
			returnVO.setStatusMsg(this.messageSource.getMessage("error.bssearchproblem",null, null));
			logger.error(jctExp.getLocalizedMessage());
		}
		logger.info("<<<< SketchSearchController.storeBeforeSketch");
		return returnVO;
	}
	/**
	 * Method searches the before After diagrams which are submitted for search.
	 * @param body
	 * @param request
	 * @return List of base64 string
	 */
	@RequestMapping(value = ACTION_SEARCH_AFTER_SKETCH, method = RequestMethod.POST)
	@ResponseBody()
	public SearchResultVO searchAfterSketch(@RequestBody String body, HttpServletRequest request) {
		logger.info(">>>> SketchSearchController.searchAfterSketch");
		
		JsonNode node = CommonUtility.getNode(body);
		String jobTitle = node.get("jobTitle").toString().replaceAll("\"", "");
		int firstResult = Integer.parseInt(node.get("firstResult").toString().replaceAll("\"", ""));
		int maxResults = Integer.parseInt(node.get("maxResults").toString().replaceAll("\"", ""));
		SearchResultVO returnVO = null;
		try{
			returnVO = searchService.getAfterSketchList(jobTitle, firstResult, maxResults);
			//returnVO.setStatusCode("200");
			returnVO.setStatusCode(CommonConstants.SUCCESSFUL);
			returnVO.setStatusMsg("Success");
		}catch(JCTException jctExp){
			returnVO = new SearchResultVO();
			returnVO.setStatusCode(jctExp.getMessage());
			returnVO.setStatusMsg(this.messageSource.getMessage("error.assearchproblem",null, null));
			logger.error(jctExp.getLocalizedMessage());
		}
		logger.info("<<<< SketchSearchController.searchAfterSketch");
		return returnVO;
	}
	
	/**
	 * Method searches the before After diagrams which are submitted for search.
	 * @param body
	 * @param request
	 * @return List of base64 string
	 */
	@RequestMapping(value = ACTION_POPULATE_MY_ACCOUNT, method = RequestMethod.POST)
	@ResponseBody()
	public MyAccountVO populateMyAccount(@RequestBody String body, HttpServletRequest request) {
		logger.info(">>>> SketchSearchController.populateMyAccount");
		
		MyAccountVO vo = new MyAccountVO();
		JsonNode node = CommonUtility.getNode(body);
		String email = node.get("emailId").toString().replaceAll("\"", "");
		try {
			vo = searchService.getAccountDetails(email);
			vo.setStatusCode(Integer.parseInt(CommonConstants.SUCCESSFUL));
			vo.setStatus("Success");
		} catch (JCTException e) {
			vo.setStatusCode(Integer.parseInt(CommonConstants.FAILURE));
			vo.setStatus("Failure");
			logger.error(e.getLocalizedMessage());
		}
		logger.info("<<<< SketchSearchController.populateMyAccount");
		return vo;
	}
	/**
	 * Method determines which tab search to be shown.
	 * @param body
	 * @param request
	 * @return SearchResultVO
	 */
	@RequestMapping(value = ACTION_IDENTIFY_SEARCH_TABS, method = RequestMethod.POST)
	@ResponseBody()
	public SearchResultVO getSearchTabsTobeViewed(@RequestBody String body, HttpServletRequest request) {
		logger.info(">>>> SketchSearchController.getSearchTabsTobeViewed");
		
		JsonNode node = CommonUtility.getNode(body);
		SearchResultVO resultVO = new SearchResultVO();
		String jobRefNo = node.get("jobReferenceNos").asText();
		try {
			resultVO = searchService.getSearchTabsTobeViewed(jobRefNo);
		} catch (JCTException e) {
			resultVO.setTabToBeShown("none");
			logger.error(e.getLocalizedMessage());
		}
		logger.info("<<<< SketchSearchController.getSearchTabsTobeViewed");
		return resultVO;
	}
	/**
	 * Method fetched the selected diagrams.
	 * @param body
	 * @param request
	 * @return SearchResultVO
	 */
	@RequestMapping(value = ACTION_GET_SELECTED_DIAGRAMS, method = RequestMethod.POST)
	@ResponseBody()
	public SearchResultVO getSelectedDiagrams(@RequestBody String body, HttpServletRequest request) {
		logger.info(">>>> SketchSearchController.getSelectedDiagrams");
		
		JsonNode node = CommonUtility.getNode(body);
		SearchResultVO resultVO = new SearchResultVO();
		String jobRefNo = node.get("jobReferenceNos").asText();
		String distinction = node.get("distinction").asText();
		try {
			resultVO = searchService.getSelectedDiagrams(jobRefNo, distinction);
		} catch (JCTException e) {
			resultVO.setTabToBeShown("none");
			logger.error(e.getLocalizedMessage());
		}
		logger.info("<<<< SketchSearchController.getSelectedDiagrams");
		return resultVO;
	}
	/**
	 * Method fetch before sketch last edited.
	 * @param body
	 * @param request
	 * @return SearchResultVO
	 */
	@RequestMapping(value = ACTION_GET_PREV_DIAGRAM, method = RequestMethod.POST)
	@ResponseBody()
	public SearchResultVO getPrevDiag(@RequestBody String body, HttpServletRequest request) {
		logger.info(">>>> SketchSearchController.getPrevDiag");
		
		JsonNode node = CommonUtility.getNode(body);
		SearchResultVO resultVO = new SearchResultVO();
		String jobRefNo = node.get("jobReferenceNos").asText();
		String distinction = node.get("distinction").asText();
		try {
			resultVO = searchService.getPrevDiag(jobRefNo, distinction);
		} catch (JCTException e) {
			resultVO.setTabToBeShown("none");
			logger.error(e.getLocalizedMessage());
		}
		logger.info("<<<< SketchSearchController.getPrevDiag");
		return resultVO;
	}
	/**
	 * Method searches the before Sketch diagrams which are submitted for search.
	 * @param body
	 * @param request
	 * @return List of base64 string
	 */
	@RequestMapping(value = ACTION_SEARCH_ALL_DIAGRAMS, method = RequestMethod.POST)
	@ResponseBody()
	public SearchResultVO searchAllDiagrams(@RequestBody String body, HttpServletRequest request) {
		logger.info(">>>> SketchSearchController.searchAllDiagrams");
		JsonNode node = CommonUtility.getNode(body);
		String distinction = node.get("distinction").asText();
		int firstResult = Integer.parseInt(node.get("firstResult").toString().replaceAll("\"", ""));
		int maxResults = Integer.parseInt(node.get("maxResults").toString().replaceAll("\"", ""));
		SearchResultVO returnVO = null;
		long count = 0;
		try{
			returnVO = searchService.getAllDiagram(distinction, firstResult, maxResults);
			count = searchService.fetchAllDiagramCount(distinction);
			returnVO.setTotalDiagramCount(count);
			returnVO.setStatusCode(CommonConstants.SUCCESSFUL);
			returnVO.setStatusMsg("Success");
		}catch(JCTException jctExp){
			returnVO = new SearchResultVO();
			returnVO.setStatusCode(jctExp.getMessage());
			returnVO.setStatusMsg(this.messageSource.getMessage("error.bssearchproblem",null, null));
			logger.error(jctExp.getLocalizedMessage());
		}
		logger.info("<<<< SketchSearchController.searchAllDiagrams");
		return returnVO;
	}
	/**
	 * Method searches the before Sketch diagrams which are submitted for search.
	 * @param body
	 * @param request
	 * @return List of base64 string
	 */
	@RequestMapping(value = "/searchAllDiagramsWithCode", method = RequestMethod.POST)
	@ResponseBody()
	public SearchResultVO searchAllDiagramsWithCode(@RequestBody String body, HttpServletRequest request) {
		logger.info(">>>> SketchSearchController.searchAllDiagrams");
		JsonNode node = CommonUtility.getNode(body);
		String distinction = node.get("distinction").asText();
		int firstResult = Integer.parseInt(node.get("firstResult").toString().replaceAll("\"", ""));
		int maxResults = Integer.parseInt(node.get("maxResults").toString().replaceAll("\"", ""));
		String code = node.get("code").asText();
		SearchResultVO returnVO = null;
		long count = 0;
		try{
			returnVO = searchService.getAllDiagram(distinction, firstResult, maxResults, code);
			count = searchService.fetchAllDiagramCount(distinction, code);
			returnVO.setTotalDiagramCount(count);
			returnVO.setStatusCode(CommonConstants.SUCCESSFUL);
			returnVO.setStatusMsg("Success");
		}catch(JCTException jctExp){
			returnVO = new SearchResultVO();
			returnVO.setStatusCode(jctExp.getMessage());
			returnVO.setStatusMsg(this.messageSource.getMessage("error.bssearchproblem",null, null));
			logger.error(jctExp.getLocalizedMessage());
		}
		logger.info("<<<< SketchSearchController.searchAllDiagrams");
		return returnVO;
	}
	
	
	/**
	 * Method searches the old PDFs from DB by userID 
	 * @param body
	 * @param request
	 * @return List of base64 string
	 */
	@RequestMapping(value = ACTION_GET_OLD_PDF, method = RequestMethod.POST)
	@ResponseBody()
	public PdfSearchVO fetchOldPdf(@RequestBody String body, HttpServletRequest request) {
		logger.info(">>>> SketchSearchController.searchAllDiagrams");
		JsonNode node = CommonUtility.getNode(body);
		int jctUserId = Integer.parseInt(node.get("jctUserId").toString().replaceAll("\"", ""));
		PdfSearchVO returnVO = new PdfSearchVO();
		try{
			if(searchService.fetchOldPdf(jctUserId).size() > 0){
				returnVO.setStatuscode(CommonConstants.SUCCESSFUL);
				returnVO.setDto(searchService.fetchOldPdf(jctUserId));
				returnVO.setStatusMsg("Success");				
			} else {				
				returnVO.setStatuscode(CommonConstants.FAILURE);
				returnVO.setStatusMsg("Failure");
			}			
		}catch(JCTException jctExp){			
			logger.error(jctExp.getLocalizedMessage());
		}
		logger.info("<<<< SketchSearchController.searchAllDiagrams");
		return returnVO;
	}
	
	
	/**
	 * Method is used to fetch the generated PDF
	 * @param body
	 * @param request
	 * @return
	 */
	@RequestMapping(value = "/showIndividualPdf/{jctFileName}", method = RequestMethod.GET)
	//@ResponseBody()
	public ReturnVO showIndividualPdf(@PathVariable("jctFileName") String file, HttpServletRequest request, HttpServletResponse response) {
		logger.info(">>>> SketchSearchController.downloadFile");
		
		ServletOutputStream print = null;
        try {
        	
        	String path = this.messageSource.getMessage("pdf.path",null, null);
        	String fileName = path+file+".pdf";
        	File temp = new File (fileName);
        	logger.info(fileName+" Exist -->"+temp.exists());
        	print = response.getOutputStream();
    		File f=new File(fileName);
    		byte[] buf = new byte[8192];
            InputStream is = new FileInputStream(f);
            int c = 0;
            while ((c = is.read(buf, 0, buf.length)) > 0) {
            	print.write(buf, 0, c);
            	print.flush();
            }
	        response.setContentType("application/pdf");
            response.setHeader("Content-Disposition",  "attachment;filename=myFile.pdf");
            response.setHeader("Pragma", "cache");
            response.setHeader("Cache-control", "private, max-age=0");
            
            //returnVO.set()
            is.close();
            print.close();
		
        } catch (Exception e) {
			logger.error(e.getLocalizedMessage());
		}
        logger.info("<<<< SketchSearchController.downloadFile");
        return null;
	}
	
}
