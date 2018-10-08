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
import com.vmware.jct.service.IAfterSketchService;
import com.vmware.jct.service.IAuthenticatorService;
import com.vmware.jct.service.IBeforeSketchService;
import com.vmware.jct.service.IConfirmationService;
import com.vmware.jct.service.IGeneratePDFService;
import com.vmware.jct.service.vo.ReturnVO;
import com.vmware.jct.service.vo.UserVO;

/**
 * 
 * <p><b>Class name:</b> ConfirmationController.java</p>
 * <p><b>Author:</b> InterraIt</p>
 * <p><b>Purpose:</b> This class acts as a controller for confirmation/result page .
 * ConfirmationController extends BasicController  and has following  methods.
 * -addDiagWithoutEdits(@RequestBody String body, HttpServletRequest request)
 * -addEdits(@RequestBody String body, HttpServletRequest request)
 * -populateAfterDiagram(UserVO userVO, String jobRefNo)
 * -downloadFile(@PathVariable("jobRefNo") String jobRefNo, HttpServletRequest request, HttpServletResponse response)
 * -disableStatus(@RequestBody String body, HttpServletRequest request)
 * -fetchJobReferenceNumber(String jrnoEncoded)
 * -generateNumber(String jrChar)
 * -reformJrNo(String jobRefDecoded)
 * <p><b>Description:</b>  </p>
 * <p><b>Copyrights:</b> 	All rights reserved by Interra IT and should only be used for its internal application development.</p>
 * <p><b>Created Date:</b> 20/Jan/2014</p>
 * <p><b>Revision History:</b>
 * 	<li>InterraIt - 04/Feb/2014 - Implemented comments</li>
 * </p>
 */

@Controller
@RequestMapping(value="/confirmation")
public class ConfirmationController extends BasicController{
	@Autowired  
	private MessageSource messageSource;
	
	@Autowired
	private IConfirmationService confirmationService;
	
	@Autowired
	private IAfterSketchService afterSketchService;
	
	@Autowired
	private IBeforeSketchService beforeSketchService;
	
	@Autowired
	private IAuthenticatorService authenticatorService;
	
	@Autowired
	private IGeneratePDFService pdfService;
	
	private static final Logger LOGGER = LoggerFactory.getLogger(ConfirmationController.class);
	
	/**
	 * Method saves and populates search Database
	 * @param body
	 * @param request
	 * @return
	 */
	@RequestMapping(value = ACTION_ADD_DIAG_WITHOUT_EDITS, method = RequestMethod.POST)
	@ResponseBody()
	public ReturnVO addDiagWithoutEdits(@RequestBody String body, HttpServletRequest request) {
		LOGGER.info(">>>> ConfirmationController.addDiagWithoutEdits");
		ReturnVO vo = new ReturnVO();
		String status = "";
		JsonNode node = CommonUtility.getNode(body);
		String jobReferenceNo = node.get("jobReferenceNo").toString().replaceAll("\"", "");
		int rowId = Integer.parseInt(node.get("rowId").asText());
		LOGGER.info("ADDING DIAGRAM WITHOUT EDITS FOR JOB_REFERENCE_NO: "+jobReferenceNo);
		try {
			status = confirmationService.addDiagWithoutEdits(jobReferenceNo);
			authenticatorService.updateLoginInfo(jobReferenceNo, rowId, "/user/view/finalPage.jsp");
			vo.setStatusCode("200");
		} catch(JCTException exp){
			vo.setStatusCode(CommonConstants.FAILURE);
			status = "failure";
			LOGGER.error(exp.getLocalizedMessage());
		}
		vo.setStatusDesc(status);
		LOGGER.info("<<<< ConfirmationController.addDiagWithoutEdits");
		return vo;
	}
	/**
	 * Method saves but do not populate search database.
	 * Again After Sketch diagram will be made editable
	 * @param body
	 * @param request
	 * @return
	 */
	@RequestMapping(value = ACTION_ADD_DIAG_WITH_EDITS, method = RequestMethod.POST)
	@ResponseBody()
	public UserVO addEdits(@RequestBody String body, HttpServletRequest request) {
		LOGGER.info(">>>> ConfirmationController.addEdits");
		UserVO vo = new UserVO();
		String status = "";
		JsonNode node = CommonUtility.getNode(body);
		String jobReferenceNo = node.get("jobReferenceNo").toString().replaceAll("\"", "");
		LOGGER.info("ADDING DIAGRAM WITH EDITS FOR JOB_REFERENCE_NO: "+jobReferenceNo);
		try{
			status = confirmationService.addEdits(jobReferenceNo);
			//Fetch Putting it together....
			vo = populateBeforeSketch(vo, jobReferenceNo);
			vo.setStatusCode("200");
		} catch(JCTException exp){
			vo.setStatusCode(CommonConstants.FAILURE);
			status = "failure";
			LOGGER.error(exp.getLocalizedMessage());
		}
		vo.setStatusDesc(status);
		LOGGER.info("<<<< ConfirmationController.addEdits");
		return vo;
	}
	
	/**
	 * Method is used to fetch the generated PDF
	 * @param body
	 * @param request
	 * @return
	 */
	@RequestMapping(value = "/downloadFile/{jobRefNo}", method = RequestMethod.GET)
	//@ResponseBody()
	public ReturnVO downloadFile(@PathVariable("jobRefNo") String jobRefNo, HttpServletRequest request, HttpServletResponse response) {
		LOGGER.info(">>>> ConfirmationController.downloadFile");
		
		ServletOutputStream print = null;
        try {
        	String jobRefDecoded = fetchJobReferenceNumber(jobRefNo);
        	String jobRefNoActual = reformJrNo(jobRefDecoded);
        	//Get pdf file name and path
        	String file = pdfService.getPathAndFile(jobRefNoActual);
        	print = response.getOutputStream();
    		File f=new File(file);
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
			LOGGER.error(e.getLocalizedMessage());
		}
        LOGGER.info("<<<< ConfirmationController.downloadFile");
        return null;
	}
	/**
	 * Method saves but disable all searching.
	 * @param body
	 * @param request
	 * @return
	 */
	@RequestMapping(value = ACTION_DISABLE_STATUS, method = RequestMethod.POST)
	@ResponseBody()
	public ReturnVO disableStatus(@RequestBody String body, HttpServletRequest request) {
		LOGGER.info(">>>> ConfirmationController.disableStatus");
		
		ReturnVO vo = new ReturnVO();
		String status = "";
		int statusInt = 0;
		JsonNode node = CommonUtility.getNode(body);
		String jobReferenceNo = node.get("jobReferenceNo").toString().replaceAll("\"", "");
		LOGGER.info("DISABLING STATUS FOR JOB_REFERENCE_NO: "+jobReferenceNo);
		try {
			statusInt = confirmationService.disableStatus(jobReferenceNo);
			if ( statusInt == 0 ) {
				vo.setStatusCode("201");
			} else {
				vo.setStatusCode("200");
			}
		} catch(JCTException exp){
			//vo.setStatusCode("500");
			vo.setStatusCode(CommonConstants.FAILURE);
			status = "failure";
			LOGGER.error(exp.getLocalizedMessage());
		}
		vo.setStatusDesc(status);
		LOGGER.info("<<<< ConfirmationController.disableStatus");
		return vo;
	}
	
	private String fetchJobReferenceNumber(String jrnoEncoded) {
		LOGGER.info(">>>> ConfirmationController.fetchJobReferenceNumber");
		StringBuilder strBldr = new StringBuilder("");
		for(int index=0; index<jrnoEncoded.length(); index++){
			strBldr.append(generateNumber(jrnoEncoded.substring(index, index+1)));
		}
		LOGGER.info("<<<< ConfirmationController.fetchJobReferenceNumber");
		return strBldr.toString();
	}
	
	private String generateNumber(String jrChar){
		LOGGER.info(">>>> ConfirmationController.generateNumber");
		String nosStr = "";
		if(jrChar.equals("X")){
			nosStr = "0";
		} else if(jrChar.equals("P")){
			nosStr = "1";
		} else if(jrChar.equals("Y")){
			nosStr = "2";
		} else if(jrChar.equals("A")){
			nosStr = "3";
		} else if(jrChar.equals("L")){
			nosStr = "4";
		} else if(jrChar.equals("B")){
			nosStr = "5";
		} else if(jrChar.equals("M")){
			nosStr = "6";
		} else if(jrChar.equals("T")){
			nosStr = "7";
		} else if(jrChar.equals("S")){
			nosStr = "8";
		} else if(jrChar.equals("H")){
			nosStr = "9";
		}
		LOGGER.info("<<<< ConfirmationController.generateNumber");
		return nosStr;
	}
	
	private String reformJrNo(String jobRefDecoded) {
		LOGGER.info("<<<< ConfirmationController.reformJrNo");
		StringBuilder strBldr = new StringBuilder("");
		
		int startIndex = 0;
        int endIndex = 2;
        int count = 0;
        while(jobRefDecoded.length() > count){
        	strBldr.append((char) Integer.parseInt((jobRefDecoded.substring(startIndex, endIndex))));
        	startIndex = startIndex+2;
        	endIndex = endIndex+2;
        	count = count + 2;
        }
        LOGGER.info("<<<< ConfirmationController.reformJrNo");
		return strBldr.toString();
	}
	
	/**
	 * Method populates the before sketch screen (1st screen)
	 * @param userVO
	 * @param jobRefNo
	 * @return
	 * @throws JCTException
	 */
	private UserVO populateBeforeSketch(UserVO userVO, String jobRefNo) throws JCTException {
		LOGGER.info(">>>> AuthenticatingContoller.populateBeforeSketch");
		//userVO.setBsJson(beforeSketchService.getJson(jobRefNo, "Y"));
		//userVO.setInitialJson(beforeSketchService.getInitialJson(jobRefNo, "Y"));
		userVO.setBsJson(beforeSketchService.getJsonForEdit(jobRefNo, "Y"));
		userVO.setInitialJson(beforeSketchService.getInitialJsonForEdit(jobRefNo, "Y"));
		userVO.setIsCompleted(1);
		LOGGER.info("<<<< AuthenticatingContoller.populateBeforeSketch");
		return userVO;
	}
	/**
	 * Method Keeps the previously drawn diagram in searchable database and present drawn diagram is ignored.
	 * @param body
	 * @param request
	 * @return
	 */
	@RequestMapping(value = ACTION_KEEP_PREVIOUS, method = RequestMethod.POST)
	@ResponseBody()
	public ReturnVO keepPrevious(@RequestBody String body, HttpServletRequest request) {
		LOGGER.info(">>>> ConfirmationController.keepPrevious");
		ReturnVO vo = new ReturnVO();
		JsonNode node = CommonUtility.getNode(body);
		String jobReferenceNo = node.get("jobReferenceNos").asText();
		try {
			String updateResult = confirmationService.updatePrevious(jobReferenceNo);
			if (updateResult.equals("success")) {
				vo.setStatusCode(CommonConstants.SUCCESSFUL);
			} else {
				vo.setStatusCode(CommonConstants.FAILURE);
			}
		} catch (JCTException exp) {
			LOGGER.error(exp.getLocalizedMessage());
		}
		LOGGER.info("<<<< ConfirmationController.keepPrevious");
		return vo;
	}
}
