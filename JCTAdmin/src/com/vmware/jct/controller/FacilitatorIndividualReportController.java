package com.vmware.jct.controller;

import java.io.File;
import java.io.FileInputStream;
import java.io.InputStream;
import java.util.Locale;
import java.util.StringTokenizer;

import javax.servlet.ServletOutputStream;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;

import jxl.Workbook;
import jxl.WorkbookSettings;
import jxl.format.Alignment;
import jxl.format.Border;
import jxl.format.BorderLineStyle;
import jxl.format.Colour;
import jxl.format.UnderlineStyle;
import jxl.format.VerticalAlignment;
import jxl.write.Label;
import jxl.write.WritableCellFormat;
import jxl.write.WritableFont;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;
import jxl.write.biff.RowsExceededException;

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

import com.vmware.jct.common.utility.CommonUtility;
import com.vmware.jct.common.utility.StatusConstants;
import com.vmware.jct.exception.JCTException;
import com.vmware.jct.service.IFacilitatorIndividualReportService;
import com.vmware.jct.service.vo.FacilitatorIndividualReportVO;
import com.vmware.jct.service.vo.ReportsVO;


/**
 * 
 * <p><b>Class name:</b> FacilitatorIndividualReportController.java</p>
 * <p><b>Author:</b> InterraIT</p>
 * <p><b>Purpose:</b> This class acts as a controller for merged reports..
 * FacilitatorIndividualReportController has following public methods:-
 * -generateUIReport()
 * -populateNewProgressReport()
 * -generateFacIndUIReport()
 * -generateFaciIndiExcel()
 * <p><b>Description:</b> This class is responsible for creating merged reports. </p>
 * <p><b>Copyrights:</b> 	All rights reserved by Interra IT and should only be used for its internal application development.</p>
 * <p><b>Created Date:</b> 15/Jan/2015</p>
  */

@Controller
@RequestMapping(value = "/facIndRprt")
public class FacilitatorIndividualReportController extends BasicController {
	@Autowired  
	private MessageSource messageSource;
	
	@Autowired
	private IFacilitatorIndividualReportService iFacilitatorIndividualReportService;
	
	private static final Logger LOGGER = LoggerFactory.getLogger(FacilitatorIndividualReportController.class);
	
	/**
	 * Method creates new progress report
	 * @param body
	 * @param request
	 * @return json
	 */

	

	@RequestMapping(value = ACTION_POPULATE_ALL_REPORT, method = RequestMethod.POST)
	@ResponseBody()
	public FacilitatorIndividualReportVO populateAllProgressReports(@RequestBody String body, HttpServletRequest request) {
		LOGGER.info(">>>>>> FacilitatorIndividualReportController.populateNewProgressReport");
		FacilitatorIndividualReportVO reportVO = new FacilitatorIndividualReportVO();
		JsonNode node = CommonUtility.getNode(body);
		String jctEmail = node.get("jctEmail").toString().replaceAll("\"", "");
		
		try{
			String resultStr = iFacilitatorIndividualReportService.populateAllProgressReports(jctEmail);
			reportVO.setReportList(resultStr);
			reportVO.setStatusCode(StatusConstants.STATUS_SUCCESS);
			reportVO.setStatusMessage("SUCCESS");
		}
		catch(JCTException e){
			LOGGER.error(e.getLocalizedMessage());
		}
		LOGGER.info("<<<<<< FacilitatorIndividualReportController.populateNewProgressReport");
		return reportVO;
	}	
	
	/**
	 * Method populates the Facilitator / Individual Report
	 * @param body
	 * @param request
	 * @return Json
	 */
	@RequestMapping(value = ACTION_GENERATE_FAC_IND_UI_RPRT, method = RequestMethod.POST)
	@ResponseBody()
	public ReportsVO generateFacIndUIReport(@RequestBody String body, HttpServletRequest request) {
		LOGGER.info(">>>>>>>> FacilitatorIndividualReportController.generateFacIndUIReport");
		JsonNode node = CommonUtility.getNode(body);
		//String adminEmailId = node.get("adminEmailId").asText();
		int usrType = node.get("usrType").asInt();	// 1 - Individual, 3 - Facilitator, 0 - all users expect ADMIN
		ReportsVO reportVO = new ReportsVO();
		try {
			String viewList = iFacilitatorIndividualReportService.getUserDetails(usrType);
			reportVO.setReportList(viewList);
		} catch (JCTException ex) {
			LOGGER.error(ex.getLocalizedMessage());
		}		
		reportVO.setStatusCode(StatusConstants.STATUS_SUCCESS);
		LOGGER.info("<<<<<<<< FacilitatorIndividualReportController.generateFacIndUIReport");
		return reportVO;
	}
	/**
	 * Method adds contents in the excel
	 * @param sheet
	 * @param column
	 * @param row
	 * @param stringToWrite
	 * @throws WriteException
	 */
	private void addCaption(WritableSheet sheet, int column, int row, String stringToWrite) 
			throws WriteException {
		LOGGER.info(">>>>>> FacilitatorIndividualReportController.addCaption");
		WritableFont headerFont = new WritableFont(WritableFont.ARIAL, 12, WritableFont.BOLD, false, 
				UnderlineStyle.NO_UNDERLINE, Colour.WHITE);
		WritableCellFormat arialHeader = new WritableCellFormat(headerFont);
		arialHeader.setVerticalAlignment(VerticalAlignment.CENTRE);
		arialHeader.setAlignment(Alignment.CENTRE);
		arialHeader.setBackground(Colour.GRAY_50);
		arialHeader.setBorder(Border.ALL, BorderLineStyle.MEDIUM);
		arialHeader.setWrap(true);
		Label label;
	    label = new Label(column, row, stringToWrite, arialHeader);
	    sheet.addCell(label);
	    LOGGER.info("<<<<<< FacilitatorIndividualReportController.addCaption");
	}
	/**
	 * Method adds lebel to excel
	 * @param sheet
	 * @param column
	 * @param row
	 * @param stringToWrite
	 * @throws WriteException
	 * @throws RowsExceededException
	 */
	private void addLabel(WritableSheet sheet, int column, int row, String stringToWrite) 
			throws WriteException {
		LOGGER.info(">>>>>> FacilitatorIndividualReportController.addLabel");
		WritableFont contentFont = new WritableFont(WritableFont.ARIAL, 10);
		WritableCellFormat arialContent = new WritableCellFormat(contentFont);
		arialContent.setVerticalAlignment(VerticalAlignment.CENTRE);
		arialContent.setAlignment(Alignment.CENTRE);
		//arialContent.setBackground(Colour.GRAY_25);
		arialContent.setBorder(Border.ALL, BorderLineStyle.THIN);
	    Label label;
	    label = new Label(column, row, stringToWrite, arialContent);
	    arialContent.setWrap(true);
	    sheet.addCell(label);
	    LOGGER.info("<<<<<< FacilitatorIndividualReportController.addLabel");
	}
	
	/**
	 * Method generates excel report for Facilitator/Individual
	 * @param userType
	 * @param reportName
	 * @param request
	 * @param response
	 */
	@RequestMapping(value = "/generateFaciIndiExcel/{userType}/{reportName}", method = RequestMethod.GET)
	public void generateFaciIndiExcel(@PathVariable("userType") String userType,
			@PathVariable("reportName") String reportName,
			HttpServletRequest request, HttpServletResponse response) {
		LOGGER.info(">>>>>> FacilitatorIndividualReportController.generateFaciIndiExcel");
		ServletOutputStream print = null;
		InputStream is = null;
		try {
			File file = new File("report.xls");
			WorkbookSettings wbSettings = new WorkbookSettings();
			wbSettings.setLocale(new Locale("en", "EN"));
			WritableWorkbook workbook = Workbook.createWorkbook(file, wbSettings);
			workbook.createSheet("Report", 0);
			WritableSheet excelSheet = workbook.getSheet(0);
			excelSheet.setColumnView(0, 26);
			excelSheet.setColumnView(1, 26);
			excelSheet.setColumnView(2, 26);
			excelSheet.setColumnView(3, 26);
			excelSheet.setColumnView(4, 26);
			createLabelFaciIndi(excelSheet,0);
			createReportFaciIndi(excelSheet,userType);
			workbook.write();
			workbook.close();
			print = response.getOutputStream();
			byte[] buf = new byte[8192];
	        is = new FileInputStream(file);
	        int c = 0;
	        while ((c = is.read(buf, 0, buf.length)) > 0) {
	        	print.write(buf, 0, c);
	        	print.flush();
	        }
	        response.setContentType("application/xls");
	        response.setHeader("Content-Disposition",  "attachment;filename=myFile.xls");
	        response.setHeader("Pragma", "cache");
	        response.setHeader("Cache-control", "private, max-age=0");
	        is.close();
	        LOGGER.info("<<<<<< FacilitatorIndividualReportController.generateFaciIndiExcel");
		} catch (Exception ex) {
			LOGGER.error(ex.getLocalizedMessage());
		}	
	}
	
	/**
	 * Method creates Facilitator/Individual report header
	 * @param sheet
	 * @param distinction
	 * @throws WriteException
	 */
	private void createLabelFaciIndi(WritableSheet sheet, int distinction) throws WriteException {
		LOGGER.info(">>>>>> FacilitatorIndividualReportController.createLabelFaciIndi");
		this.addCaption(sheet, 0, 0, "Sl no.");
		this.addCaption(sheet, 1, 0, "USER NAME");
		this.addCaption(sheet, 2, 0, "Profile Name");
		this.addCaption(sheet, 3, 0, "Group Name");
		this.addCaption(sheet, 4, 0, "Created by"); 
		LOGGER.info("<<<<<< FacilitatorIndividualReportController.createLabelFaciIndi");
	}
	
	/**
	 * Method puts the record in the excel sheet
	 * @param excelSheet
	 * @param userType
	 * @throws RowsExceededException
	 * @throws WriteException
	 * @throws JCTException
	 */
	private void createReportFaciIndi(WritableSheet excelSheet,String userType) throws RowsExceededException, WriteException, JCTException {
		LOGGER.info(">>>>>> FacilitatorIndividualReportController.createReportFaciIndi");
		String resultString = iFacilitatorIndividualReportService.createFaciIndiReport(userType);
		StringTokenizer token = new StringTokenizer(resultString, "$$$");
		int rowIndex = 1;
		int srlCounter = 1;
		while (token.hasMoreTokens()) {
			String[] dispObj = token.nextToken().toString().split("#");
			addLabel(excelSheet, 0, rowIndex, srlCounter+"");
			addLabel(excelSheet, 1, rowIndex, dispObj[0]);
			addLabel(excelSheet, 2, rowIndex, dispObj[1]);
			addLabel(excelSheet, 3, rowIndex, dispObj[2]);
			addLabel(excelSheet, 4, rowIndex, dispObj[3]);
			srlCounter = srlCounter + 1;
			rowIndex = rowIndex + 1;
		}
		LOGGER.info("<<<<<< FacilitatorIndividualReportController.createReportFaciIndi");
	}
	

	@RequestMapping(value = ACTION_POPULATE_PROGRESS_REPORT, method = RequestMethod.POST)
	@ResponseBody()
	public FacilitatorIndividualReportVO populateProgressReports(@RequestBody String body, HttpServletRequest request) {
		LOGGER.info(">>>>>> FacilitatorIndividualReportController.populateNewProgressReport");
		FacilitatorIndividualReportVO reportVO = new FacilitatorIndividualReportVO();
		JsonNode node = CommonUtility.getNode(body);
		String jctEmail = node.get("jctEmail").toString().replaceAll("\"", "");
		String userGroupName = node.get("userGroupName").toString().replaceAll("\"", "");
		
		try{
			String resultStr = iFacilitatorIndividualReportService.populateProgressReports(jctEmail,userGroupName);
			reportVO.setReportList(resultStr);
			reportVO.setStatusCode(StatusConstants.STATUS_SUCCESS);
			reportVO.setStatusMessage("SUCCESS");
		}
		catch(JCTException e){
			LOGGER.error(e.getLocalizedMessage());
		}
		LOGGER.info("<<<<<< FacilitatorIndividualReportController.populateNewProgressReport");
		return reportVO;
	
	}
}
