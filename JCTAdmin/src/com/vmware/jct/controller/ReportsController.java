package com.vmware.jct.controller;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.math.BigInteger;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Locale;
import java.util.Map;
import java.util.StringTokenizer;

import javax.servlet.ServletOutputStream;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;

import jxl.CellView;
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
import org.springframework.util.SystemPropertyUtils;
import org.springframework.web.bind.annotation.PathVariable;
import org.springframework.web.bind.annotation.RequestBody;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestMethod;
import org.springframework.web.bind.annotation.ResponseBody;

import com.vmware.jct.common.utility.CommonUtility;
import com.vmware.jct.common.utility.StatusConstants;
import com.vmware.jct.dao.IFileUploadDAO;
import com.vmware.jct.dao.dto.FunctionDTO;
import com.vmware.jct.dao.dto.LevelDTO;
import com.vmware.jct.dao.dto.OnetOccupationDTO;
import com.vmware.jct.dao.dto.SurveyQuestionDTO;
import com.vmware.jct.dao.dto.UserProfileDTO;
import com.vmware.jct.exception.JCTException;
import com.vmware.jct.service.IAuthenticatorService;
import com.vmware.jct.service.IContentConfigService;
import com.vmware.jct.service.IReportService;
import com.vmware.jct.service.vo.AfterSketchReportListVO;
import com.vmware.jct.service.vo.AfterSketchReportVO;
import com.vmware.jct.service.vo.BeforeSketchReportListVO;
import com.vmware.jct.service.vo.BeforeSketchReportVO;
import com.vmware.jct.service.vo.BeforeSketchToAfterSketchReportListVO;
import com.vmware.jct.service.vo.BeforeSketchToAfterSketchReportVO;
import com.vmware.jct.service.vo.ContentConfigVO;
import com.vmware.jct.service.vo.ReportsVO;
import com.vmware.jct.service.vo.surveyQtnReportVO;
/**
 * 
 * <p><b>Class name:</b> ReportsController.java</p>
 * <p><b>Author:</b> InterraIT</p>
 * <p><b>Purpose:</b> This class acts as a controller for all reports except merged reports..
 * AllReportsController has following public methods:-
 * -populateBeforeSketchReport()
 * -populateBeforeSketchSelectedReport()
 * -generateExcel()
 * -populateActionPlanReport()
 * -populateActionPlanSelectedReport()
 * -generateExcelActionPlan()
 * -populateAfterSketchReport()
 * -populateAfterSketchSelectedReport()
 * -generateASExcel()
 * -populateBeforeSketchToAfterSketchReport()
 * -populateBeforeSketchToAfterSketchSelectedReport()
 * -populateMiscReport()
 * -populateMiscDetailedReport()
 * -generateBsToAsExcel()
 * -generateExcelMisc()
 * -populateAfterSketchMappingDetails()
 * -populateProfileDropDown()
 * -generateExcelGroupProfileReport()
 * -generatePaymentExcel()
 * 
 * <p><b>Description:</b> This class is responsible for creating merged reports. </p>
 * <p><b>Copyrights:</b> 	All rights reserved by Interra IT and should only be used for its internal application development.</p>
 * <p><b>Created Date:</b> 01/Mar/2014</p>
 * <p><b>Revision History:</b>
 * 	<li></li>
 * </p>
 */
@Controller
@RequestMapping(value = "/reports")
public class ReportsController extends BasicController {
	
	@Autowired
	private IAuthenticatorService authenticatorService;
	
	@Autowired
	private IReportService iReportService;
	
	@Autowired
	private IFileUploadDAO serviceDAO;
	
	@Autowired  
	private MessageSource messageSource;
	
	@Autowired
	private IContentConfigService service;
	
	private static final Logger LOGGER = LoggerFactory.getLogger(ReportsController.class);
	/**
	 * Method populates the before sketch report
	 * @param body
	 * @param request
	 * @return Json
	 */
	@RequestMapping(value = ACTION_BEFORE_SKETCH_REPORT, method = RequestMethod.POST)
	@ResponseBody()
	public ReportsVO populateBeforeSketchReport(@RequestBody String body, HttpServletRequest request) {
		LOGGER.info(">>>>>> ReportsController.populateBeforeSketchReport");
		ReportsVO reportVO = new ReportsVO();
		try {
			JsonNode node = CommonUtility.getNode(body);
			int recordIndex = Integer.parseInt(node.get("recordIndex").toString().replaceAll("\"" , ""));
			
			reportVO = this.populateFunctionGroup(reportVO);
			reportVO = this.populateJobLevel(reportVO);
			//List<Object> viewList = iReportService.getBeforeSketchList("" , "" , recordIndex);
			List<Object> viewList = iReportService.getBeforeSketchList("" , recordIndex);
			if (viewList.size() < 3) {
				viewList.clear();
			}
			reportVO = this.populateBSSketchReport(viewList, reportVO);
			reportVO.setTotalCount(iReportService.getTotalCount("" , ""));
			reportVO = this.getPaginationAnchor(reportVO, recordIndex);
			reportVO.setTableHeaderString(this.createTableHeader("BS"));
			reportVO.setStatusCode(StatusConstants.STATUS_SUCCESS);
			reportVO.setStatusDesc("SUCCESS");
		} catch (JCTException e) {
			reportVO.setStatusCode(StatusConstants.STATUS_FAILURE);
			reportVO.setStatusDesc(e.getLocalizedMessage());
			LOGGER.error(e.getLocalizedMessage());
		}
		LOGGER.info("<<<<<< ReportsController.populateBeforeSketchReport");
		return reportVO;
	}
	/**
	 * Method populates the before sketch report based on function group and job level
	 * @param body
	 * @param request
	 * @return json
	 */
	@RequestMapping(value = ACTION_BEFORE_SKETCH_SELECTED_REPORT, method = RequestMethod.POST)
	@ResponseBody()
	public ReportsVO populateBeforeSketchSelectedReport(@RequestBody String body, HttpServletRequest request) {
		LOGGER.info(">>>>>> ReportsController.populateBeforeSketchSelectedReport");
		ReportsVO reportVO = new ReportsVO();
		JsonNode node = CommonUtility.getNode(body);
		String occupationCode = node.get("occupationVal").toString().replaceAll("\"", "");
		int recordIndex = Integer.parseInt(node.get("recordIndex").toString().replaceAll("\"" , ""));
		try {
			reportVO = this.populateFunctionGroup(reportVO);
			reportVO = this.populateJobLevel(reportVO);
			List<Object> viewList = iReportService.getBeforeSketchList(occupationCode, recordIndex);
			reportVO = this.populateBSSketchReport(viewList, reportVO);
			
			reportVO.setTotalCount(iReportService.getTotalCount(occupationCode));
			reportVO = this.getPaginationAnchor(reportVO, recordIndex);
			reportVO.setTableHeaderString(this.createTableHeader("BS"));
			reportVO.setStatusCode(StatusConstants.STATUS_SUCCESS);
			reportVO.setStatusDesc("SUCCESS");
		} catch (JCTException e) {
			reportVO.setStatusCode(StatusConstants.STATUS_FAILURE);
			reportVO.setStatusDesc(e.getLocalizedMessage());
			LOGGER.error(e.getLocalizedMessage());
		}
		LOGGER.info("<<<<<< ReportsController.populateBeforeSketchSelectedReport");
		return reportVO;
	}
	/**
	 * Method populates the function group
	 * @param reportVO
	 * @return map in vo
	 * @throws JCTException
	 */
	private ReportsVO populateFunctionGroup(ReportsVO reportVO) throws JCTException {
		LOGGER.info(">>>>>> ReportsController.populateFunctionGroup");
		List<FunctionDTO> jctFunctionList=null;
		Map<String, String> functionGroupMap=new LinkedHashMap<String, String>();
			jctFunctionList=authenticatorService.jctFunctionList();
			int j=0;
			for(int i=1;i<=jctFunctionList.size();i++){
				functionGroupMap.put(jctFunctionList.get(j).getJctFunctionName(), jctFunctionList.get(j).getJctFunctionName());
				j++;
			}
			reportVO.setFunctionGroup(functionGroupMap);
			LOGGER.info("<<<<<< ReportsController.populateFunctionGroup");
		return reportVO;
	}
	/**
	 * Method populates the job level
	 * @param reportVO
	 * @return map in vo
	 * @throws JCTException
	 */
	private ReportsVO populateJobLevel(ReportsVO reportVO) throws JCTException {
		LOGGER.info(">>>>>> ReportsController.populateJobLevel");
		
		List<LevelDTO> jctLevelList=null;
		Map<String, String> jobLevelMap=new LinkedHashMap<String, String>();
			jctLevelList=authenticatorService.jctLevelList();
			int j=0;
			for(int i=1;i<=jctLevelList.size();i++){
				jobLevelMap.put(jctLevelList.get(j).getJctLevelName(), jctLevelList.get(j).getJctLevelName());
				j++;
			}
			reportVO.setJobLevel(jobLevelMap);
			LOGGER.info("<<<<<< ReportsController.populateJobLevel");
		return reportVO;
	}
	/**
	 * Method decides if next link should be disabled or previous or both or none
	 * @param reportVO
	 * @param recordIndex
	 * @return report vo
	 */
	private ReportsVO getPaginationAnchor(ReportsVO reportVO, int recordIndex) {
		LOGGER.info(">>>>>> ReportsController.getPaginationAnchor");
		//If record index is 0 then it means first set. So disabling the previous link
		if(recordIndex == 0){
			reportVO.setPreviousDisabled(1); // 1: Disabled, 0: Enabled
		} else {
			reportVO.setPreviousDisabled(0);
		}
		//If total count and index
		if((reportVO.getTotalCount()-20 <= recordIndex)){
			reportVO.setNextDisabled(1); // 1: Disabled, 0: Enabled
		} else {
			reportVO.setNextDisabled(0);
		}
		LOGGER.info("<<<<<< ReportsController.getPaginationAnchor");
		return reportVO;
	}
	/**
	 * Method generates excel for before sketch report
	 * @param functionGroup
	 * @param jobLevel
	 * @param reportName
	 * @param request
	 * @param response
	 */
	@RequestMapping(value = "/generateExcel/{occupation}/{reportName}", method = RequestMethod.GET)
	public void generateExcel(@PathVariable("occupation") String occupation, 
			@PathVariable("reportName") String reportName, HttpServletRequest request, HttpServletResponse response) {
		LOGGER.info(">>>>>> ReportsController.generateExcel");
		InputStream is = null;
		/*if (functionGroup.equals("JCTDUMMYVARJCTFN")) {
			functionGroup = "";
		}
		if (jobLevel.equals("JCTDUMMYVARJCTJL")) {
			jobLevel = "";
		}*/
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
			//excelSheet.setColumnView(3, 26);
			excelSheet.setColumnView(3, 26);
			excelSheet.setColumnView(4, 26);
			excelSheet.setColumnView(5, 26);
			excelSheet.setColumnView(6, 26);
			excelSheet.setColumnView(7, 26);
			createLabel(excelSheet,0);
			createReport(excelSheet, occupation);
			
			workbook.write();
			workbook.close();
			writeToStream(response.getOutputStream(), response, file, is, generateExcelFileName("before_sketch_report"));
            LOGGER.info("<<<<<< ReportsController.generateExcel");
		}catch(Exception ex){
			LOGGER.error(ex.getLocalizedMessage());
			//ex.printStackTrace();
		}
		
	}
	/**
	 * Method creates report for before sketch
	 * @param excelSheet
	 * @param functionGrp
	 * @param jobLevel
	 * @throws RowsExceededException
	 * @throws WriteException
	 * @throws JCTException
	 */
	private void createReport(WritableSheet excelSheet, String occupation) throws RowsExceededException, WriteException, JCTException {
		LOGGER.info(">>>>>> ReportsController.createReport");
		String occupationVal = "";
		if (!occupation.equals("NA")) {
			occupationVal = occupation;
		}
		List<Object> viewList = iReportService.getBeforeSketchListForExcel(occupationVal);
		ReportsVO reportVO = new ReportsVO();
		reportVO = this.populateBSSketchReport(viewList, reportVO);
		List<BeforeSketchReportVO> list = reportVO.getBsList();
		int rowNumber = 1;
		int mergeIndexEnd = 0;
		int internalRowNumber = 1;
		int mergerStartIndex = 1;
		int rowNos = 1;
		for (int rowCount = 0; rowCount < list.size(); rowCount++) {
			int cellsToBeMerged = 0;
			BeforeSketchReportVO vo = list.get(rowCount);
			List<BeforeSketchReportListVO> innerListVO = vo.getList();
			if (innerListVO.size() > 2) {
				this.addLabel(excelSheet, 0, rowNumber, Integer.toString(rowNos));
				rowNos = rowNos + 1;
				this.addLabel(excelSheet, 1, rowNumber, vo.getName());
				this.addLabel(excelSheet, 2, rowNumber, vo.getEmailId());
				this.addLabel(excelSheet, 3, rowNumber, vo.getFunctionGroup());
				//this.addLabel(excelSheet, 3, rowNumber, vo.getJobLevel());
				mergeIndexEnd = mergeIndexEnd + innerListVO.size();
				
				this.addLabel(excelSheet, 4, internalRowNumber, Integer.toString(innerListVO.size()));
				
				for (int innerRowCount = 0; innerRowCount < innerListVO.size(); innerRowCount++) {
					BeforeSketchReportListVO listVO = innerListVO.get(innerRowCount);
					
					this.addLabel(excelSheet, 5, internalRowNumber, listVO.getTaskDescription());
					
					if (null != listVO.getAllocation() && listVO.getAllocation().equals("0")) {
						this.addLabel(excelSheet, 6, internalRowNumber, " ");
					} else {
						this.addLabel(excelSheet, 6, internalRowNumber, listVO.getAllocation());
					}
					internalRowNumber = internalRowNumber + 1;
					cellsToBeMerged = cellsToBeMerged + 1;
				}
				this.addLabel(excelSheet, 7, rowNumber, formatHHMMSS(Integer.parseInt(vo.getTotalTime())));
				rowNumber = internalRowNumber;
				
				excelSheet.mergeCells(0, mergerStartIndex, 0, mergeIndexEnd);
				excelSheet.mergeCells(2, mergerStartIndex, 2, mergeIndexEnd);
				excelSheet.mergeCells(3, mergerStartIndex, 3, mergeIndexEnd);
				//excelSheet.mergeCells(3, mergerStartIndex, 3, mergeIndexEnd);
				excelSheet.mergeCells(4, mergerStartIndex, 4, mergeIndexEnd);
				//excelSheet.mergeCells(5, mergerStartIndex, 5, mergeIndexEnd);
				excelSheet.mergeCells(7, mergerStartIndex, 7, mergeIndexEnd);
				mergerStartIndex = mergeIndexEnd + 1;
			}
		}

		LOGGER.info("<<<<<< ReportsController.createReport");
		
	}
	
	/**
	 * Method writes header in excel for before sketch report
	 * @param sheet
	 * @param distinction
	 * @throws WriteException
	 */
	private void createLabel(WritableSheet sheet, int distinction) throws WriteException {
		LOGGER.info(">>>>>> ReportsController.createLabel");
		addCaption(sheet, 0, 0, "SL No.");
		addCaption(sheet, 1, 0, "Name (Last, First)");
		addCaption(sheet, 2, 0, "Email Id");
		addCaption(sheet, 3, 0, "Occupation");
		//addCaption(sheet, 3, 0, "Job Level");
		addCaption(sheet, 4, 0, "Total Tasks");
		addCaption(sheet, 5, 0, "Task Description");
		addCaption(sheet, 6, 0, "Time / Energy");
		addCaption(sheet, 7, 0, "Total time for before sketch(HH:MM:SS)");
		LOGGER.info("<<<<<< ReportsController.createLabel");
	}
	/**
	 * Method writes header in excel for after diagram report
	 * @param sheet
	 * @param distinction
	 * @throws WriteException
	 */
	private void createASLabel(WritableSheet sheet, int distinction) throws WriteException {
		LOGGER.info(">>>>>> ReportsController.createASLabel");
		
		this.addCaption(sheet, 0, 0, "Name (Last, First)");
		this.addCaption(sheet, 1, 0, "Email Id");
		this.addCaption(sheet, 2, 0, "Occupation");
		//this.addCaption(sheet, 2, 0, "Job Level");
		this.addCaption(sheet, 3, 0, "Role Description");
		this.addCaption(sheet, 4, 0, "Task Description");
		this.addCaption(sheet, 5, 0, "Task Time / Energy");
		this.addCaption(sheet, 6, 0, "Mappings");
		this.addCaption(sheet, 7, 0, "Total Tasks");
		this.addCaption(sheet, 8, 0, "Total Mappings");
		this.addCaption(sheet, 9, 0, "Total time after diagram(HH:MM:SS)");
		LOGGER.info("<<<<<< ReportsController.createASLabel");
	}
	/**
	 * Method creates report for After Diagram
	 * @param excelSheet
	 * @param functionGrp
	 * @param jobLevel
	 * @throws RowsExceededException
	 * @throws WriteException
	 * @throws JCTException
	 */
	private void createASReport(WritableSheet excelSheet, String occupation) throws RowsExceededException, WriteException, JCTException {
		LOGGER.info(">>>>>> ReportsController.createASReport");
		List<Object> viewList = iReportService.getAfterSketchListForExcel(occupation);
		ReportsVO reportVO = new ReportsVO();
		reportVO = this.populateASSketchReport(viewList, reportVO);
		List<AfterSketchReportVO> list = reportVO.getList();
		int rowNumber = 1;
		int internalRowNumber = 1;
		int mergerStartIndex = 1;
		int mergeIndexEnd = 0;
		int mappingCount = 0;
		for (int rowCount = 0; rowCount < list.size(); rowCount++) {
			int cellsToBeMerged = 0;
			mappingCount = 0;
			AfterSketchReportVO vo = list.get(rowCount);
			this.addLabel(excelSheet, 0, rowNumber, vo.getName());
			this.addLabel(excelSheet, 1, rowNumber, vo.getEmailId());
			this.addLabel(excelSheet, 2, rowNumber, vo.getFunctionGroup());
			//this.addLabel(excelSheet, 2, rowNumber, vo.getJobLevel());
			
			List<AfterSketchReportListVO> innerListVO = vo.getRepeatingList();
			mergeIndexEnd = mergeIndexEnd + innerListVO.size();
			
			for (int innerRowCount = 0; innerRowCount < innerListVO.size(); innerRowCount++) {
				AfterSketchReportListVO listVO = innerListVO.get(innerRowCount);
				this.addLabel(excelSheet, 3, internalRowNumber, listVO.getRole().equals("N/A") ? "" : listVO.getRole());
				this.addLabel(excelSheet, 4, internalRowNumber, listVO.getTaskDesc().equals("") ? "----" : listVO.getTaskDesc());
				if (null != listVO.getTaskAllocation() && listVO.getTaskAllocation().equals("0")) {
					this.addLabel(excelSheet, 5, internalRowNumber, "----");
				} else {
					this.addLabel(excelSheet, 5, internalRowNumber, listVO.getTaskAllocation());
				}
				this.addLabel(excelSheet, 6, internalRowNumber, listVO.getPhaka().equals("") ? "----" : listVO.getPhaka());	//	mapping				
				if(!listVO.getPhaka().equals(""))
					mappingCount++;
				internalRowNumber = internalRowNumber + 1;
				cellsToBeMerged = cellsToBeMerged + 1;
			}
			this.addLabel(excelSheet, 7, rowNumber, vo.getTotalTaskCount());
			this.addLabel(excelSheet, 8, rowNumber, Integer.toString(mappingCount));
			this.addLabel(excelSheet, 9, rowNumber, formatHHMMSS(Integer.parseInt(vo.getTimeSpent())));
			rowNumber = internalRowNumber;
			
			excelSheet.mergeCells(1, mergerStartIndex, 1, mergeIndexEnd);
			excelSheet.mergeCells(2, mergerStartIndex, 2, mergeIndexEnd);
			//excelSheet.mergeCells(2, mergerStartIndex, 2, mergeIndexEnd);
			excelSheet.mergeCells(7, mergerStartIndex, 7, mergeIndexEnd);
			excelSheet.mergeCells(8, mergerStartIndex, 8, mergeIndexEnd);
			excelSheet.mergeCells(9, mergerStartIndex, 9, mergeIndexEnd);
			mergerStartIndex = mergeIndexEnd + 1;
		}
		LOGGER.info("<<<<<< ReportsController.createASReport");
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
		LOGGER.info(">>>>>> ReportsController.addCaption");
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
	    LOGGER.info("<<<<<< ReportsController.addCaption");
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
		LOGGER.info(">>>>>> ReportsController.addLabel");
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
	    LOGGER.info("<<<<<< ReportsController.addLabel");
	}
	/**
	 * Method generates report for Action Plan
	 * @param body
	 * @param request
	 * @return ReportsVO
	 */
	@RequestMapping(value = ACTION_ACTION_PLAN_REPORT, method = RequestMethod.POST)
	@ResponseBody()
	public ReportsVO populateActionPlanReport(@RequestBody String body, HttpServletRequest request) {
		LOGGER.info(">>>>>> ReportsController.populateActionPlanReport");
		LOGGER.info("--------------------- POPULATING ACTION PLAN REPORT --------------------- ");
		ReportsVO reportVO = new ReportsVO();
		try {
			JsonNode node = CommonUtility.getNode(body);
			int recordIndex = Integer.parseInt(node.get("recordIndex").toString().replaceAll("\"", ""));
			
			reportVO = this.populateFunctionGroup(reportVO);
			reportVO = this.populateJobLevel(reportVO);
			List<Object> viewList = iReportService.getActionPlanList("", recordIndex);
			reportVO = this.populateAPSketchReport(viewList, reportVO);
			reportVO.setTotalCount(iReportService.getTotalCountActionPlan(""));
			reportVO = this.getPaginationAnchor(reportVO, recordIndex);
			reportVO.setTableHeaderString(this.createTableHeader("AP"));
			reportVO.setStatusCode(StatusConstants.STATUS_SUCCESS);
			reportVO.setStatusDesc("SUCCESS");
		} catch (JCTException e) {
			reportVO.setStatusCode(StatusConstants.STATUS_FAILURE);
			reportVO.setStatusDesc(e.getLocalizedMessage());
			LOGGER.error(e.getLocalizedMessage());
		}
		LOGGER.info("<<<<<< ReportsController.populateActionPlanReport");
		return reportVO;
	}
	/**
	 * Method generates report for Action Plan with selected parameters
	 * @param body
	 * @param request
	 * @return ReportsVO
	 */
	@RequestMapping(value = ACTION_ACTION_PLAN_SELECTED_REPORT, method = RequestMethod.POST)
	@ResponseBody()
	public ReportsVO populateActionPlanSelectedReport(@RequestBody String body, HttpServletRequest request) {
		LOGGER.info(">>>>>> ReportsController.populateActionPlanSelectedReport");
		LOGGER.info("------------------- POPULATING ACTION PLAN REPORT WITH SELECTION ------------------- ");
		ReportsVO reportVO = new ReportsVO();
		JsonNode node = CommonUtility.getNode(body);
		String occupationVal = node.get("occupationVal").toString().replaceAll("\"", "");
		int recordIndex = Integer.parseInt(node.get("recordIndex").toString().replaceAll("\"", ""));
		try {
			reportVO = this.populateFunctionGroup(reportVO);
			reportVO = this.populateJobLevel(reportVO);
			List<Object> viewList = iReportService.getActionPlanList(occupationVal, recordIndex);
			reportVO = this.populateAPSketchReport(viewList, reportVO);
			reportVO.setTotalCount(iReportService.getTotalCountActionPlan(occupationVal));
			reportVO = this.getPaginationAnchor(reportVO, recordIndex);
			reportVO.setTableHeaderString(this.createTableHeader("AP"));
			reportVO.setStatusCode(StatusConstants.STATUS_SUCCESS);
			reportVO.setStatusDesc("SUCCESS");
		} catch (JCTException e) {
			reportVO.setStatusCode(StatusConstants.STATUS_FAILURE);
			reportVO.setStatusDesc(e.getLocalizedMessage());
			LOGGER.error(e.getLocalizedMessage());
		}
		LOGGER.info("<<<<<< ReportsController.populateActionPlanSelectedReport");
		return reportVO;
	}
	/**
	 * Method populates the action plan report
	 * @param viewList
	 * @param reportVO
	 * @return report vo
	 */
	private ReportsVO populateAPSketchReport(List<Object> viewList, ReportsVO reportVO) {
		LOGGER.info(">>>>>> ReportsController.populateAPSketchReport");
		StringBuilder data = new StringBuilder("");
		ArrayList<String> keyTrackingList = new ArrayList<String>();
		for (int index = 0; index < viewList.size(); index++) {
			Object[] innerObj = (Object[])viewList.get(index);
			if (!keyTrackingList.contains((String)innerObj[0])) {
				keyTrackingList.add((String)innerObj[0]);
			}
		}
		for (int outerIndex=0; outerIndex<keyTrackingList.size(); outerIndex++) {
			String outerJobRefNo = (String)keyTrackingList.get(outerIndex);
			for (int innerIndex = 0; innerIndex<viewList.size(); innerIndex++) {
				Object[] innerObj = (Object[])viewList.get(innerIndex);
				String innerJobRefNo = (String)innerObj[0];
				if (innerJobRefNo.equals(outerJobRefNo)){
					if (!data.toString().contains(innerJobRefNo)) {
						String occupationTitle = "";
						try {
							occupationTitle = serviceDAO.getOccupationTitle((String) innerObj[1]);
						} catch (Exception e) {
							LOGGER.error("Unable to fetch Occupation Title");
						}
						data.append((String)innerObj[0]+"#"+occupationTitle+"#"+(String)innerObj[2]+"#"+(String)innerObj[3]+"#"+(String)innerObj[7]+","+(String)innerObj[6]+"#");
					}
					data.append((String)innerObj[4]+"!"+(String)innerObj[5]+"~");
				}
			}
			data.append("$^$");
		}
		reportVO.setReportList(data.toString());
		LOGGER.info("<<<<<< ReportsController.populateAPSketchReport");
		return reportVO;
	}
	/**
	 * Method generates excel sheet report for action plan
	 * @param functionGroup
	 * @param jobLevel
	 * @param reportName
	 * @param request
	 * @param response
	 */
	@RequestMapping(value = "/generateExcelActionPlan/{occupation}/{reportName}", method = RequestMethod.GET)
	public void generateExcelActionPlan(@PathVariable("occupation") String occupation,  
			@PathVariable("reportName") String reportName, HttpServletRequest request, HttpServletResponse response) {
		LOGGER.info(">>>>>> ReportsController.generateExcelActionPlan");
		InputStream is = null;
		String occupationVal = "";
		if (!occupation.equals("NA")) {
			occupationVal = occupation;
		}
		
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
			excelSheet.setColumnView(5, 26);
			createLabelActionPlan(excelSheet,0);
			createReportActionPlan(excelSheet, occupationVal);
			
			CellView cellView = new CellView();
		    cellView.setHidden(true); //set hidden
		    excelSheet.setColumnView(2, cellView);
			
		    workbook.write();
			workbook.close();
			writeToStream(response.getOutputStream(), response, file, is, generateExcelFileName("action_plan_report"));
	        LOGGER.info("<<<<<< ReportsController.generateExcelActionPlan");
		} catch (Exception ex) {
			LOGGER.error(ex.getLocalizedMessage());
		}	
		
	}
	/**
	 * Method creates the action plan report
	 * @param excelSheet
	 * @param functionGrp
	 * @param jobLevel
	 * @throws RowsExceededException
	 * @throws WriteException
	 * @throws JCTException
	 */
	private void createReportActionPlan(WritableSheet excelSheet, String occupation) 
			throws RowsExceededException, WriteException, JCTException {
		LOGGER.info(">>>>>> ReportsController.createReportActionPlan");
		
		List<Object> viewList = iReportService.getActionPlanListForExcel(occupation);
		ReportsVO reportVO = new ReportsVO();
		reportVO = this.populateAPSketchReport(viewList, reportVO);
		String rptStr = reportVO.getReportList();
		StringTokenizer token = new StringTokenizer(rptStr, "$^$");
		int rowIndex = 1;
		int colIndex = 0;
		int mergingRwCnt = 0;
		int mergeIndexStart = 1;
		int mergeIndexEnd = 0;
		while (token.hasMoreTokens()) {
			String dispObj = token.nextToken().toString();
			String[] initialArr = dispObj.split("#");
			String[] taskTimeStrArr = initialArr[4].split("~"); //Task energy count
			int rowCountForSameJRNo = taskTimeStrArr.length; //This many rows for same job reference number
			mergingRwCnt = mergingRwCnt + rowCountForSameJRNo;
			int loopIndex = 1;
			int rowCount = 0;
			for (int index = 0; index < rowCountForSameJRNo; index++) {
				if (colIndex == 6) {
					colIndex = 0;
					rowIndex = rowIndex + 1;
					rowCount = rowCount + rowCountForSameJRNo;
				}
				if (loopIndex == 1) {
					this.addLabel(excelSheet, colIndex, rowIndex, initialArr[0]);
				} 
				colIndex = colIndex + 1;
				
				if (loopIndex == 1) {
					this.addLabel(excelSheet, colIndex, rowIndex, initialArr[1]);
				} 
				colIndex = colIndex + 1;
				
				if (loopIndex == 1) {
					this.addLabel(excelSheet, colIndex, rowIndex, initialArr[2]);
				} 
				colIndex = colIndex + 1;
				
				if (loopIndex == 1) {
					this.addLabel(excelSheet, colIndex, rowIndex, initialArr[3]);
				} 
				colIndex = colIndex + 1;
				
				this.addLabel(excelSheet, colIndex, rowIndex, taskTimeStrArr[index].split("!")[0]);
				colIndex = colIndex + 1;
				
				this.addLabel(excelSheet, colIndex, rowIndex, taskTimeStrArr[index].split("!")[1]);
				colIndex = colIndex + 1;
				loopIndex = loopIndex+1;
			}
			mergeIndexEnd = mergeIndexEnd + rowCountForSameJRNo;
			excelSheet.mergeCells(0, mergeIndexStart, 0, mergeIndexEnd);
			excelSheet.mergeCells(1, mergeIndexStart, 1, mergeIndexEnd);
			excelSheet.mergeCells(2, mergeIndexStart, 2, mergeIndexEnd);
			excelSheet.mergeCells(3, mergeIndexStart, 3, mergeIndexEnd);
			mergeIndexStart = mergeIndexStart + rowCountForSameJRNo;
		}
		LOGGER.info("<<<<<< ReportsController.createReportActionPlan");
	}
	/**
	 * Method creates action plan headers in excel
	 * @param sheet
	 * @param distinction
	 * @throws WriteException
	 */
	private void createLabelActionPlan(WritableSheet sheet, int distinction) throws WriteException {
		LOGGER.info(">>>>>> ReportsController.createLabelActionPlan");
		
		this.addCaption(sheet, 0, 0, "Email ID");
		this.addCaption(sheet, 1, 0, "Occupation");
		this.addCaption(sheet, 2, 0, "Job Level");
		this.addCaption(sheet, 3, 0, "Main Question");
		this.addCaption(sheet, 4, 0, "Sub Question");
		this.addCaption(sheet, 5, 0, "Answer");
		LOGGER.info("<<<<<< ReportsController.createLabelActionPlan"); 
	}
	/**
	 * Method creates misc report header
	 * @param sheet
	 * @param distinction
	 * @throws WriteException
	 */
	private void createLabelMisc(WritableSheet sheet, int distinction) throws WriteException {
		LOGGER.info(">>>>>> ReportsController.createLabelMisc");
		this.addCaption(sheet, 0, 0, "Sr No.");
		this.addCaption(sheet, 1, 0, "Name (Last, First)");
		this.addCaption(sheet, 2, 0, "Email ID");
		this.addCaption(sheet, 3, 0, "First Start date/time");
		this.addCaption(sheet, 4, 0, "Last End date/time");
		this.addCaption(sheet, 5, 0, "Number of times logged"); 
		LOGGER.info("<<<<<< ReportsController.createLabelMisc");
	}
	/**
	 * Method creates the After Diagram report
	 * @param body
	 * @param request
	 * @return json
	 */
	@RequestMapping(value = ACTION_AFTER_SKETCH_REPORT, method = RequestMethod.POST)
	@ResponseBody()
	public ReportsVO populateAfterSketchReport(@RequestBody String body, HttpServletRequest request) {
		LOGGER.info(">>>>>> ReportsController.populateAfterSketchReport");
		ReportsVO reportVO = new ReportsVO();
		try {
			JsonNode node = CommonUtility.getNode(body);
			int recordIndex = Integer.parseInt(node.get("recordIndex").toString().replaceAll("\"", ""));
			reportVO = this.populateFunctionGroup(reportVO);
			reportVO = this.populateJobLevel(reportVO);
			List<Object> viewList = iReportService.getAfterSketchList("","", recordIndex);
			reportVO = this.populateASSketchReport(viewList, reportVO);
			reportVO.setTotalCount(iReportService.getASTotalCount("", ""));
			reportVO = this.getPaginationAnchor(reportVO, recordIndex);
			reportVO.setTableHeaderString(this.createTableHeader("AS"));
			reportVO.setStatusCode(StatusConstants.STATUS_SUCCESS);
			reportVO.setStatusDesc("SUCCESS");
		} catch (JCTException e) {
			reportVO.setStatusCode(StatusConstants.STATUS_FAILURE);
			reportVO.setStatusDesc(e.getLocalizedMessage());
			LOGGER.error(e.getLocalizedMessage());
		}
		LOGGER.info("<<<<<< ReportsController.populateAfterSketchReport");
		return reportVO;
	}
	/**
	 * Method populates the before sketch values
	 * @param viewList
	 * @param reportVO
	 * @return report vo
	 */
	private ReportsVO populateBSSketchReport(List<Object> viewList, ReportsVO reportVO) {
		LOGGER.info(">>>>>> ReportsController.populateBSSketchReport");
		List<BeforeSketchReportVO> reportContentsVO = new ArrayList<BeforeSketchReportVO>();
		StringBuilder data = new StringBuilder("");
		ArrayList<String> keyTrackingList = new ArrayList<String>();
		
		
		for (int index = 0; index < viewList.size(); index++) {
			Object[] innerObj = (Object[]) viewList.get(index);
			if (!keyTrackingList.contains((String) innerObj[0])) {
				keyTrackingList.add((String) innerObj[0]);
			}
		}
		
		for (int outerIndex = 0; outerIndex < keyTrackingList.size(); outerIndex++) {
			String outerJobRefNo = (String) keyTrackingList.get(outerIndex);
			BeforeSketchReportVO vo = new BeforeSketchReportVO();
			ArrayList<BeforeSketchReportListVO> list = new ArrayList<BeforeSketchReportListVO>();
			
			for (int innerIndex = 0; innerIndex < viewList.size(); innerIndex++) {
				Object[] innerObj = (Object[]) viewList.get(innerIndex);
				String innerJobRefNo = (String) innerObj[0];
				if (innerJobRefNo.equals(outerJobRefNo)) {
					if (!data.toString().contains(innerJobRefNo)) {
						String occupationTitle = "";
						try {
							occupationTitle = serviceDAO.getOccupationTitle((String) innerObj[1]);
						} catch (Exception e) {
							LOGGER.error("Unable to fetch Occupation Title");
						}
						
						data.append((String) innerObj[0] + "#" + 
									occupationTitle + "#" + // Occupation serviceDAO.getOccupationTitle((String) me.getKey());
									(String) innerObj[2] + "#" + 
									(Integer) innerObj[5] + "#" );
						
						vo.setEmailId((String)innerObj[0]);
						vo.setFunctionGroup(occupationTitle);
						vo.setJobLevel((String)innerObj[2]);
						vo.setTotalTime((Integer) innerObj[5]+"");
						vo.setName((String)innerObj[7]+","+(String)innerObj[6]);
					}
					BeforeSketchReportListVO repetingVO = new BeforeSketchReportListVO();
					data.append((String) innerObj[3]+"!"+(Integer) innerObj[4]+"~"+(String)innerObj[7]+","+(String)innerObj[6]);
					repetingVO.setTaskDescription((String)innerObj[3]);
					repetingVO.setAllocation((Integer)innerObj[4]+"");
					list.add(repetingVO);
				}
				vo.setList(list);
			}
			reportContentsVO.add(vo);
			data.append("$^$");
		}
		reportVO.setReportList(data.toString());
		reportVO.setBsList(reportContentsVO);
		LOGGER.info("<<<<<< ReportsController.populateBSSketchReport");
		return reportVO;
	}
	
	/**
	 * Method populates the After Diagram report
	 * @param viewList
	 * @param reportVO
	 * @return return vo
	 */
	private ReportsVO populateASSketchReport(List<Object> viewList, ReportsVO reportVO) {
		LOGGER.info(">>>>>> ReportsController.populateASSketchReport");
		List<AfterSketchReportVO> reportContentsVO = new ArrayList<AfterSketchReportVO>();
		StringBuilder data = new StringBuilder("");
		ArrayList<String> keyTrackingList = new ArrayList<String>();
		for (int index = 0; index < viewList.size(); index++) {
			Object[] innerObj = (Object[])viewList.get(index);
			if (!keyTrackingList.contains((String)innerObj[0])) {
				keyTrackingList.add((String)innerObj[0]);
			}
		}
		for (int outerIndex = 0; outerIndex < keyTrackingList.size(); outerIndex++) {
			String outerJobRefNo = (String)keyTrackingList.get(outerIndex);
			AfterSketchReportVO vo = new AfterSketchReportVO();
			ArrayList<AfterSketchReportListVO> list = new ArrayList<AfterSketchReportListVO>();
			for (int innerIndex = 0; innerIndex < viewList.size(); innerIndex++) {
				Object[] innerObj = (Object[])viewList.get(innerIndex);
				String innerJobRefNo = (String)innerObj[0];
				if (innerJobRefNo.equals(outerJobRefNo)) {
					if (!data.toString().contains(innerJobRefNo)) {
						String occupationTitle = "";
						try {
							occupationTitle = serviceDAO.getOccupationTitle((String) innerObj[6]);
						} catch (Exception e) {
							LOGGER.error("Unable to fetch Occupation Title");
						}
									//job reference no		//role desc				//time spent
						data.append((String)innerObj[0]+"#"+(Integer)innerObj[4]+"#"+(String)innerObj[6]+"#"+(String)innerObj[7]+"#"+iReportService.getTotalTaskCount((String)innerObj[0])+"#"+iReportService.getTotalMappingCount((String)innerObj[0])+"#"+iReportService.getTotalRoleCount((String)innerObj[0])+"#");
						vo.setEmailId((String)innerObj[0]);
						vo.setTimeSpent((Integer)innerObj[4]+"");
						vo.setFunctionGroup(occupationTitle);
						vo.setJobLevel((String)innerObj[7]);
						vo.setName((String)innerObj[9]+","+(String)innerObj[8]);
						vo.setTotalTaskCount(iReportService.getTotalTaskCount((String)innerObj[0])+"");
						vo.setTotalMappingCount("0");	//	will be calculated later in js manually
						vo.setTotalRoleCount(iReportService.getTotalRoleCount((String)innerObj[0])+"");
					}
					
					AfterSketchReportListVO repetingVO = new AfterSketchReportListVO();
					data.append((String)innerObj[2]+"!"+(Integer)innerObj[3]+"!"+(String)innerObj[5]+"~"+(String)innerObj[1]+"@@"+(String)innerObj[9]+","+(String)innerObj[8]);
					repetingVO.setTaskDesc((String)innerObj[2]);
					repetingVO.setTaskAllocation((Integer)innerObj[3]+"");
					repetingVO.setPhaka((String)innerObj[5]);
					repetingVO.setRole((String)innerObj[1]);
					list.add(repetingVO);
				}
				vo.setRepeatingList(list);
			}
			reportContentsVO.add(vo);
			data.append("$$$");
		}
		reportVO.setReportList(data.toString());
		reportVO.setList(reportContentsVO);
		LOGGER.info("<<<<<< ReportsController.populateASSketchReport");
		return reportVO;
	}
	/**
	 * Method populates the After Diagram report based on function group and joblevel
	 * @param body
	 * @param request
	 * @return json
	 */
	@RequestMapping(value = ACTION_AFTER_SKETCH_SELECTED_REPORT, method = RequestMethod.POST)
	@ResponseBody()
	public ReportsVO populateAfterSketchSelectedReport(@RequestBody String body, HttpServletRequest request) {
		LOGGER.info(">>>>>> ReportsController.populateAfterSketchSelectedReport");
		ReportsVO reportVO = new ReportsVO();
		JsonNode node = CommonUtility.getNode(body);
		String occupationCode = node.get("occupationVal").toString().replaceAll("\"", "");
		int recordIndex = Integer.parseInt(node.get("recordIndex").toString().replaceAll("\"", ""));
		try {
			reportVO = this.populateFunctionGroup(reportVO);
			reportVO = this.populateJobLevel(reportVO);
			List<Object> viewList = iReportService.getAfterSketchList(occupationCode, recordIndex);
			reportVO = this.populateASSketchReport(viewList, reportVO);
			reportVO.setTotalCount(iReportService.getASTotalCount(occupationCode));
			reportVO = this.getPaginationAnchor(reportVO, recordIndex);
			reportVO.setTableHeaderString(this.createTableHeader("AS"));
			reportVO.setStatusCode(StatusConstants.STATUS_SUCCESS);
			reportVO.setStatusDesc("SUCCESS");
		} catch (JCTException e) {
			reportVO.setStatusCode(StatusConstants.STATUS_FAILURE);
			reportVO.setStatusDesc(e.getLocalizedMessage());
			LOGGER.error(e.getLocalizedMessage());
		}
		LOGGER.info("<<<<<< ReportsController.populateAfterSketchSelectedReport");
		return reportVO;
	}
	/**
	 * Method generates the excel sheet report for after sketch diagram
	 * @param functionGroup
	 * @param jobLevel
	 * @param reportName
	 * @param request
	 * @param response
	 */
	@RequestMapping(value = "/generateASExcel/{occupation}/{reportName}", method = RequestMethod.GET)
	public void generateASExcel(@PathVariable("occupation") String occupation, 
			@PathVariable("reportName") String reportName, 
			HttpServletRequest request, HttpServletResponse response) {
		LOGGER.info(">>>>>> ReportsController.generateASExcel");
		InputStream is = null;
		String occupationVal = "";
		if(!occupation.equals("NA")){
			occupationVal = occupation;
		}
		try {
			File file = new File("report.xls");
			WorkbookSettings wbSettings = new WorkbookSettings();
			wbSettings.setLocale(new Locale("en", "EN"));
			WritableWorkbook workbook = Workbook.createWorkbook(file, wbSettings);
			workbook.createSheet("After Diagram Report", 0);
			WritableSheet excelSheet = workbook.getSheet(0);
			excelSheet.setColumnView(0, 26);
			excelSheet.setColumnView(1, 26);
			//excelSheet.setColumnView(2, 26);
			excelSheet.setColumnView(2, 26);
			excelSheet.setColumnView(3, 26);
			excelSheet.setColumnView(4, 26);
			excelSheet.setColumnView(5, 26);
			excelSheet.setColumnView(6, 26);
			excelSheet.setColumnView(7, 26);
			excelSheet.setColumnView(8, 26);
			excelSheet.setColumnView(9, 26);
			this.createASLabel(excelSheet,0);
			this.createASReport(excelSheet, occupationVal);
			workbook.write();
			workbook.close();
			writeToStream(response.getOutputStream(), response, file, is, generateExcelFileName("after_sketch_report"));
            LOGGER.info("<<<<<< ReportsController.generateASExcel");
		}catch(Exception ex){
			LOGGER.error(ex.getLocalizedMessage());
		}
		
	}
	/**
	 * Method creates before sketch to after sketch report
	 * @param body
	 * @param request
	 * @return jsonb
	 */
	@RequestMapping(value = ACTION_BEFORE_SKETCH_AFTER_SKETCH_REPORT, method = RequestMethod.POST)
	@ResponseBody()
	public ReportsVO populateBeforeSketchToAfterSketchReport(@RequestBody String body, HttpServletRequest request) {
		LOGGER.info(">>>>>> ReportsController.populateBeforeSketchToAfterSketchReport");
		ReportsVO reportVO = new ReportsVO();
		try {
			JsonNode node = CommonUtility.getNode(body);
			int recordIndex = Integer.parseInt((node.get("recordIndex").toString().replaceAll("\"", "")));
			
			reportVO = this.populateFunctionGroup(reportVO);
			reportVO = this.populateJobLevel(reportVO);
			List<Object> viewList = iReportService.getBsToAsList("", recordIndex);
			reportVO = this.populateBSTOASSketchReport(viewList, reportVO);
			reportVO.setTotalCount(iReportService.getTotalCountBsToAs("", ""));
			reportVO = this.getPaginationAnchor(reportVO, recordIndex);
			reportVO.setTableHeaderString(this.createTableHeader("BSTOAS"));
			reportVO.setStatusCode(StatusConstants.STATUS_SUCCESS);
			reportVO.setStatusDesc("SUCCESS");
		} catch (JCTException e) {
			reportVO.setStatusCode(StatusConstants.STATUS_FAILURE);
			reportVO.setStatusDesc(e.getLocalizedMessage());
			LOGGER.error(e.getLocalizedMessage());
		}
		LOGGER.info("<<<<<< ReportsController.populateBeforeSketchToAfterSketchReport");
		return reportVO;
	}
	/**
	 * Method creates before sketch to After Diagram report based on selected parameters
	 * @param body
	 * @param request
	 * @return json
	 */
	@RequestMapping(value = ACTION_BEFORE_SKETCH_AFTER_SKETCH_SELECTED_REPORT, method = RequestMethod.POST)
	@ResponseBody()
	public ReportsVO populateBeforeSketchToAfterSketchSelectedReport(@RequestBody String body, HttpServletRequest request) {
		LOGGER.info(">>>>>> ReportsController.populateBeforeSketchToAfterSketchSelectedReport");
		ReportsVO reportVO = new ReportsVO();
		JsonNode node = CommonUtility.getNode(body);
		String occupationVal = node.get("occupationVal").toString().replaceAll("\"", "");
		int recordIndex = Integer.parseInt((node.get("recordIndex").toString().replaceAll("\"", "")));
		try {
			reportVO = this.populateFunctionGroup(reportVO);
			reportVO = this.populateJobLevel(reportVO);
			List<Object> viewList = iReportService.getBsToAsList(occupationVal, recordIndex);
			reportVO = this.populateBSTOASSketchReport(viewList, reportVO);
			reportVO.setTotalCount(iReportService.getTotalCountBsToAs(occupationVal, ""));
			reportVO = this.getPaginationAnchor(reportVO, recordIndex);
			reportVO.setTableHeaderString(this.createTableHeader("BSTOAS"));
			reportVO.setStatusCode(StatusConstants.STATUS_SUCCESS);
			reportVO.setStatusDesc("SUCCESS");
		} catch (JCTException e) {
			reportVO.setStatusCode(StatusConstants.STATUS_FAILURE);
			reportVO.setStatusDesc(e.getLocalizedMessage());
			LOGGER.error(e.getLocalizedMessage());
		}
		LOGGER.info("<<<<<< ReportsController.populateBeforeSketchToAfterSketchSelectedReport");
		return reportVO;
	}
	/**
	 * Method populates the before sketch to After Diagram report
	 * @param viewList
	 * @param reportVO
	 * @return report 
	 */
	private ReportsVO populateBSTOASSketchReport(List<Object> viewList, ReportsVO reportVO) {
		LOGGER.info(">>>>>> ReportsController.populateBSTOASSketchReport");
		StringBuilder data = new StringBuilder("");
		ArrayList<String> keyTrackingList = new ArrayList<String>();
		List<BeforeSketchToAfterSketchReportVO> reportContentsVO = new ArrayList<BeforeSketchToAfterSketchReportVO>();
		for (int index = 0; index < viewList.size(); index++) {
			Object[] innerObj = (Object[])viewList.get(index);
			if (!keyTrackingList.contains((String)innerObj[0])) {
				keyTrackingList.add((String)innerObj[0]);
			}
		}
		for (int outerIndex = 0; outerIndex < keyTrackingList.size(); outerIndex++) {
			String outerJobRefNo = (String)keyTrackingList.get(outerIndex);
			
			BeforeSketchToAfterSketchReportVO vo = new BeforeSketchToAfterSketchReportVO();
			ArrayList<BeforeSketchToAfterSketchReportListVO> list = new ArrayList<BeforeSketchToAfterSketchReportListVO>();
			
			for (int innerIndex = 0; innerIndex < viewList.size(); innerIndex++) {
				Object[] innerObj = (Object[])viewList.get(innerIndex);
				String innerJobRefNo = (String)innerObj[0];
				if (innerJobRefNo.equals(outerJobRefNo)) {
					if (!data.toString().contains(innerJobRefNo)) {
						String occupationTitle = "";
						try {
							occupationTitle = serviceDAO.getOccupationTitle((String) innerObj[1]);
						} catch (Exception e) {
							LOGGER.error("Unable to fetch Occupation Title");
						}
						
						data.append((String)innerObj[0]+"#"+(String)innerObj[1]+"#"+(String)innerObj[2]+"#");
						
						vo.setEmailId((String)innerObj[0]);
						vo.setFunctionGroup(occupationTitle);
						vo.setJobLevel((String)innerObj[2]);
						vo.setName((String)innerObj[11]+","+(String)innerObj[10]);
					}
					data.append((String)innerObj[3]+"!"+(String)innerObj[4]+"!"+(String)innerObj[5]+"!"+(String)innerObj[6]+"!"
							+(Integer)innerObj[7]+"!"+(Integer)innerObj[8]+"!"+(String)innerObj[9]+"~"+(String)innerObj[11]+","+(String)innerObj[10]);
					
					BeforeSketchToAfterSketchReportListVO repetingVO = new BeforeSketchToAfterSketchReportListVO();					
					repetingVO.setTaskStatus((String)innerObj[3]);
					repetingVO.setBeforeSketchTaskDesc((String)innerObj[4]);
					repetingVO.setAfterSketchTaskDesc((String)innerObj[5]);
					repetingVO.setTaskDiff((String)innerObj[6]);
					repetingVO.setBsTime((Integer)innerObj[7]+"");
					repetingVO.setAsTime((Integer)innerObj[8]+"");
					repetingVO.setTimeDiff((String)innerObj[9]);
					list.add(repetingVO);				
				}
				vo.setList(list);
			}
			reportContentsVO.add(vo);
			data.append("$^$");
		}
		reportVO.setReportList(data.toString());
		reportVO.setBsToAsList(reportContentsVO);
		
		
		LOGGER.info("<<<<<< ReportsController.populateBSTOASSketchReport");
		return reportVO;
	}
	/**
	 * Method creates miscellaneous report
	 * @param body
	 * @param request
	 * @return json
	 */
	@RequestMapping(value = ACTION_MISC_REPORT, method = RequestMethod.POST)
	@ResponseBody()
	public ReportsVO populateMiscReport(@RequestBody String body, HttpServletRequest request) {
		LOGGER.info(">>>>>> ReportsController.populateMiscReport");
		ReportsVO reportVO = new ReportsVO();
		try {
			String resultString = iReportService.createMiscHeaderReport();
			reportVO.setReportList(resultString);
			reportVO.setTableHeaderString(this.createTableHeader("MISC"));
			reportVO.setStatusCode(StatusConstants.STATUS_SUCCESS);
			reportVO.setStatusDesc("SUCCESS");
		} catch (JCTException e) {
			reportVO.setStatusCode(StatusConstants.STATUS_FAILURE);
			reportVO.setStatusDesc(e.getLocalizedMessage());
			LOGGER.error(e.getLocalizedMessage());
		}
		LOGGER.info("<<<<<< ReportsController.populateMiscReport");
		return reportVO;
	}
	/**
	 * Method creates detailed miscellaneous report based on job reference nos
	 * @param body
	 * @param request
	 * @return
	 */
	@RequestMapping(value = ACTION_MISC_DETAILED_REPORT, method = RequestMethod.POST)
	@ResponseBody()
	public ReportsVO populateMiscDetailedReport(@RequestBody String body, HttpServletRequest request) {
		LOGGER.info(">>>>>> ReportsController.populateMiscDetailedReport");
		ReportsVO reportVO = new ReportsVO();
		JsonNode node = CommonUtility.getNode(body);
		String jobReferenceNo = node.get("jrNos").toString().replaceAll("\"", "");
		try {
			String resultString = iReportService.createMiscDetailesReport(jobReferenceNo);
			reportVO.setReportList(resultString);
			reportVO.setStatusCode(StatusConstants.STATUS_SUCCESS);
			reportVO.setStatusDesc("SUCCESS");
		} catch (JCTException e) {
			reportVO.setStatusCode(StatusConstants.STATUS_FAILURE);
			reportVO.setStatusDesc(e.getLocalizedMessage());
			LOGGER.error(e.getLocalizedMessage());
		}
		LOGGER.info("<<<<<< ReportsController.populateMiscDetailedReport");
		return reportVO;
	}
	/**
	 * Method creates excel report for before sketch to After Diagram
	 * @param functionGroup
	 * @param jobLevel
	 * @param reportName
	 * @param request
	 * @param response
	 */
	@RequestMapping(value = "/generateBsToAsExcel/{occupation}/{reportName}", method = RequestMethod.GET)
	public void generateBsToAsExcel(@PathVariable("occupation") String occupation, 
			@PathVariable("reportName") String reportName,
			HttpServletRequest request, HttpServletResponse response) {
		LOGGER.info(">>>>>> ReportsController.generateBsToAsExcel");
		InputStream is = null;
		String occupationVal = "";
		if (!occupation.equals("NA")) {
			occupationVal = occupation;
		}
		try {
			File file = new File("report.xls");
			WorkbookSettings wbSettings = new WorkbookSettings();
			wbSettings.setLocale(new Locale("en", "EN"));
			WritableWorkbook workbook = Workbook.createWorkbook(file, wbSettings);
			workbook.createSheet("BS To AS Report", 0);
			WritableSheet excelSheet = workbook.getSheet(0);
			excelSheet.setColumnView(0, 26);
			excelSheet.setColumnView(1, 26);
			//excelSheet.setColumnView(2, 26);
			excelSheet.setColumnView(2, 26);
			excelSheet.setColumnView(3, 26);
			excelSheet.setColumnView(4, 26);
			excelSheet.setColumnView(5, 26);
			excelSheet.setColumnView(6, 26);
			excelSheet.setColumnView(7, 26);
			excelSheet.setColumnView(8, 26);
			excelSheet.setColumnView(9, 26);
			this.createLabelBsToAs(excelSheet,0);			
			this.createReportBsToAs(excelSheet, occupationVal);
			workbook.write();
			workbook.close();
			writeToStream(response.getOutputStream(), response, file, is, generateExcelFileName("bs_to_as_report"));
	        LOGGER.info("<<<<<< ReportsController.generateBsToAsExcel");
		} catch (Exception ex) {
			LOGGER.error(ex.getLocalizedMessage());
		}	
	}
	/**
	 * Method populates before sketch to After Diagram report
	 * @param excelSheet
	 * @param functionGrp
	 * @param jobLevel
	 * @throws RowsExceededException
	 * @throws WriteException
	 * @throws JCTException
	 */
	private void createReportBsToAs(WritableSheet excelSheet, String occupation) 
			throws RowsExceededException, WriteException, JCTException {
		LOGGER.info(">>>>>> ReportsController.createReportBsToAs");
		List<Object> viewList = iReportService.getBsToAsListForExcel(occupation);
		ReportsVO reportVO = new ReportsVO();
		reportVO = populateBSTOASSketchReport(viewList, reportVO);
		List<BeforeSketchToAfterSketchReportVO> list = reportVO.getBsToAsList();
		int rowNumber = 1;
		int internalRowNumber = 1;
		int mergerStartIndex = 1;
		int mergeIndexEnd = 0;		
		for (int rowCount = 0; rowCount < list.size(); rowCount++) {
			int cellsToBeMerged = 0;
			BeforeSketchToAfterSketchReportVO vo = list.get(rowCount);	
			this.addLabel(excelSheet, 0, rowNumber, vo.getName());
			this.addLabel(excelSheet, 1, rowNumber, vo.getEmailId());
			this.addLabel(excelSheet, 2, rowNumber, vo.getFunctionGroup());
			//this.addLabel(excelSheet, 2, rowNumber, vo.getJobLevel());
			
			List<BeforeSketchToAfterSketchReportListVO> innerListVO = vo.getList();
			mergeIndexEnd = mergeIndexEnd + innerListVO.size();			
			for (int innerRowCount = 0; innerRowCount < innerListVO.size(); innerRowCount++) {
				BeforeSketchToAfterSketchReportListVO listVO = innerListVO.get(innerRowCount);
				this.addLabel(excelSheet, 3, internalRowNumber, listVO.getTaskStatus());
				this.addLabel(excelSheet, 4, internalRowNumber, listVO.getBeforeSketchTaskDesc());	
				this.addLabel(excelSheet, 5, internalRowNumber, listVO.getAfterSketchTaskDesc());
				this.addLabel(excelSheet, 6, internalRowNumber, listVO.getTaskDiff().equals("") ? "----" : listVO.getTaskDiff());
				this.addLabel(excelSheet, 7, internalRowNumber, listVO.getBsTime());
				this.addLabel(excelSheet, 8, internalRowNumber, listVO.getAsTime());
				this.addLabel(excelSheet, 9, internalRowNumber, listVO.getTimeDiff());				
				internalRowNumber = internalRowNumber + 1;
				cellsToBeMerged = cellsToBeMerged + 1;
			}
			rowNumber = internalRowNumber;			
			excelSheet.mergeCells(1, mergerStartIndex, 1, mergeIndexEnd);
			excelSheet.mergeCells(2, mergerStartIndex, 2, mergeIndexEnd);
			//excelSheet.mergeCells(2, mergerStartIndex, 2, mergeIndexEnd);
			mergerStartIndex = mergeIndexEnd + 1;
		}
		LOGGER.info("<<<<<< ReportsController.createReportBsToAs");
	}
	/**
	 * Method creates before sketch to After Diagram header for excel
	 * @param sheet
	 * @param distinction
	 * @throws WriteException
	 */
	private void createLabelBsToAs(WritableSheet sheet, int distinction) throws WriteException {
		LOGGER.info(">>>>>> ReportsController.createLabelBsToAs");
		this.addCaption(sheet, 0, 0, "Name (Last, First)");
		this.addCaption(sheet, 1, 0, "Email ID");
		this.addCaption(sheet, 2, 0, "Occupation");
		//this.addCaption(sheet, 2, 0, "Job Level");
		this.addCaption(sheet, 3, 0, "Task Status");
		this.addCaption(sheet, 4, 0, "Before Sketch Task Desc");
		this.addCaption(sheet, 5, 0, "After Diagram Task Desc");
		this.addCaption(sheet, 6, 0, "Diff in Task Desc");
		this.addCaption(sheet, 7, 0, "Before Sketch Energy");
		this.addCaption(sheet, 8, 0, "After Diagram Energy");
		this.addCaption(sheet, 9, 0, "Diff in Energy");
		LOGGER.info("<<<<<< ReportsController.createLabelBsToAs");
	}
	/**
	 * Method generates excel report for Misc
	 * @param reportName
	 * @param request
	 * @param response
	 */
	@RequestMapping(value = "/generateExcelMisc/{reportName}", method = RequestMethod.GET)
	public void generateExcelMisc(@PathVariable("reportName") String reportName, 
			HttpServletRequest request, HttpServletResponse response) {
		LOGGER.info(">>>>>> ReportsController.generateExcelMisc");
		InputStream is = null;
		try {
			File file = new File("report.xls");
			WorkbookSettings wbSettings = new WorkbookSettings();
			wbSettings.setLocale(new Locale("en", "EN"));
			WritableWorkbook workbook = Workbook.createWorkbook(file, wbSettings);
			workbook.createSheet("Report", 0);
			WritableSheet excelSheet = workbook.getSheet(0);
			excelSheet.setColumnView(0, 8);
			excelSheet.setColumnView(1, 26);
			excelSheet.setColumnView(2, 26);
			excelSheet.setColumnView(3, 26);
			excelSheet.setColumnView(4, 26);
			excelSheet.setColumnView(5, 26);
			createLabelMisc(excelSheet,0);
			createReportMisc(excelSheet);
			workbook.write();
			workbook.close();
			writeToStream(response.getOutputStream(), response, file, is, generateExcelFileName("misc_report"));
	        LOGGER.info("<<<<<< ReportsController.generateExcelMisc");
		} catch (Exception ex) {
			LOGGER.error(ex.getLocalizedMessage());
		}	
	}
	/**
	 * Method puts the record in the excel sheet
	 * @param excelSheet
	 * @throws RowsExceededException
	 * @throws WriteException
	 * @throws JCTException
	 */
	private void createReportMisc(WritableSheet excelSheet) throws RowsExceededException, WriteException, JCTException {
		LOGGER.info(">>>>>> ReportsController.createReportMisc");
		String resultString = iReportService.createMiscHeaderReport();
		StringTokenizer token = new StringTokenizer(resultString, "$$$");
		int rowIndex = 1;
		int srlCounter = 1;
		while (token.hasMoreTokens()) {
			String[] dispObj = token.nextToken().toString().split("#");
			addLabel(excelSheet, 0, rowIndex, srlCounter+"");
			addLabel(excelSheet, 1, rowIndex, dispObj[5]);
			addLabel(excelSheet, 2, rowIndex, dispObj[0]);
			addLabel(excelSheet, 3, rowIndex, dispObj[2]);
			addLabel(excelSheet, 4, rowIndex, dispObj[3]);
			addLabel(excelSheet, 5, rowIndex, dispObj[4]);
			srlCounter = srlCounter + 1;
			rowIndex = rowIndex + 1;
		}
		LOGGER.info("<<<<<< ReportsController.createReportMisc");
	}
	/**
	 * Method takes report name as the parameter and returns 
	 * header in customized locale.
	 * @param reportName
	 * @return customized header.
	 */	
	private String createTableHeader (String reportName) {
		LOGGER.info(">>>>>> ReportsController.createTableHeader");
		StringBuilder headerBuilder = new StringBuilder("");		
		if ("MISC".equals(reportName)) {
			headerBuilder.append(this.messageSource.getMessage("report.header.srln", null , null));
			headerBuilder.append("&&&");
			headerBuilder.append(this.messageSource.getMessage("report.header.email", null , null));
			headerBuilder.append("&&&");
			headerBuilder.append(this.messageSource.getMessage("report.header.fstdt", null , null));
			headerBuilder.append("&&&");
			headerBuilder.append(this.messageSource.getMessage("report.header.leddt", null , null));
			headerBuilder.append("&&&");
			headerBuilder.append(this.messageSource.getMessage("report.header.count", null , null));
			headerBuilder.append("&&&");
			headerBuilder.append(this.messageSource.getMessage("report.header.searchable", null , null));
			headerBuilder.append("&&&");
			headerBuilder.append(this.messageSource.getMessage("report.header.flname", null , null));
		} else if ("BS".equals(reportName)) {
			headerBuilder.append(this.messageSource.getMessage("report.header.email", null , null));
			headerBuilder.append("&&&");
			headerBuilder.append(this.messageSource.getMessage("label.occupation.search", null , null));
			headerBuilder.append("&&&");
			headerBuilder.append(this.messageSource.getMessage("report.header.jobLevel", null , null));
			headerBuilder.append("&&&");
			headerBuilder.append(this.messageSource.getMessage("report.header.taskDesc", null , null));
			headerBuilder.append("&&&");
			headerBuilder.append(this.messageSource.getMessage("report.header.timeEnrgy", null , null));
			headerBuilder.append("&&&");
			headerBuilder.append(this.messageSource.getMessage("report.header.totalTimeBs", null , null));
			headerBuilder.append("&&&");
			headerBuilder.append(this.messageSource.getMessage("report.header.display", null , null));
			headerBuilder.append("&&&");
			headerBuilder.append(this.messageSource.getMessage("report.header.to", null , null));
			headerBuilder.append("&&&");
			headerBuilder.append(this.messageSource.getMessage("report.header.outOf", null , null));
			headerBuilder.append("&&&");
			headerBuilder.append(this.messageSource.getMessage("report.header.export", null , null));
			headerBuilder.append("&&&");
			headerBuilder.append(this.messageSource.getMessage("report.header.previous", null , null));
			headerBuilder.append("&&&");
			headerBuilder.append(this.messageSource.getMessage("report.header.next", null , null));
			headerBuilder.append("&&&");
			headerBuilder.append(this.messageSource.getMessage("report.header.noRecord", null , null));
			headerBuilder.append("&&&");
			headerBuilder.append(this.messageSource.getMessage("report.header.totalTask", null , null));
			headerBuilder.append("&&&");
			headerBuilder.append(this.messageSource.getMessage("report.header.flname", null , null));
			
		} else if ("AP".equals(reportName)) {
			headerBuilder.append(this.messageSource.getMessage("report.header.email", null , null));
			headerBuilder.append("&&&");
			headerBuilder.append(this.messageSource.getMessage("label.occupation.search", null , null));
			headerBuilder.append("&&&");
			headerBuilder.append(this.messageSource.getMessage("report.header.jobLevel", null , null));
			headerBuilder.append("&&&");
			headerBuilder.append(this.messageSource.getMessage("report.header.mainQtn", null , null));
			headerBuilder.append("&&&");
			headerBuilder.append(this.messageSource.getMessage("report.header.subQtn", null , null));
			headerBuilder.append("&&&");
			headerBuilder.append(this.messageSource.getMessage("report.header.answer", null , null));
			headerBuilder.append("&&&");
			headerBuilder.append(this.messageSource.getMessage("report.header.display", null , null));
			headerBuilder.append("&&&");
			headerBuilder.append(this.messageSource.getMessage("report.header.to", null , null));
			headerBuilder.append("&&&");
			headerBuilder.append(this.messageSource.getMessage("report.header.outOf", null , null));
			headerBuilder.append("&&&");
			headerBuilder.append(this.messageSource.getMessage("report.header.export", null , null));
			headerBuilder.append("&&&");
			headerBuilder.append(this.messageSource.getMessage("report.header.previous", null , null));
			headerBuilder.append("&&&");
			headerBuilder.append(this.messageSource.getMessage("report.header.next", null , null));
			headerBuilder.append("&&&");
			headerBuilder.append(this.messageSource.getMessage("report.header.noRecord", null , null));
			headerBuilder.append("&&&");
			headerBuilder.append(this.messageSource.getMessage("report.header.flname", null , null));
		} else if ("AS".equals(reportName)) {
			headerBuilder.append(this.messageSource.getMessage("report.header.roleDesc", null , null));
			headerBuilder.append("&&&");
			headerBuilder.append(this.messageSource.getMessage("report.header.taskDesc", null , null));
			headerBuilder.append("&&&");
			headerBuilder.append(this.messageSource.getMessage("report.header.taskTimeEnrgy", null , null));
			headerBuilder.append("&&&");
			headerBuilder.append(this.messageSource.getMessage("report.header.mappings", null , null));
			headerBuilder.append("&&&");
			headerBuilder.append(this.messageSource.getMessage("report.header.totalTask", null , null));
			headerBuilder.append("&&&");
			headerBuilder.append(this.messageSource.getMessage("report.header.totalMappings", null , null));
			headerBuilder.append("&&&");
			headerBuilder.append(this.messageSource.getMessage("report.header.display", null , null));
			headerBuilder.append("&&&");
			headerBuilder.append(this.messageSource.getMessage("report.header.to", null , null));
			headerBuilder.append("&&&");
			headerBuilder.append(this.messageSource.getMessage("report.header.outOf", null , null));
			headerBuilder.append("&&&");
			headerBuilder.append(this.messageSource.getMessage("report.header.export", null , null));
			headerBuilder.append("&&&");
			headerBuilder.append(this.messageSource.getMessage("report.header.previous", null , null));
			headerBuilder.append("&&&");
			headerBuilder.append(this.messageSource.getMessage("report.header.next", null , null));
			headerBuilder.append("&&&");
			headerBuilder.append(this.messageSource.getMessage("report.header.noRecord", null , null));
			headerBuilder.append("&&&");
			headerBuilder.append(this.messageSource.getMessage("report.header.email", null , null));
			headerBuilder.append("&&&");
			headerBuilder.append(this.messageSource.getMessage("label.occupation.search", null , null));
			headerBuilder.append("&&&");
			headerBuilder.append(this.messageSource.getMessage("report.header.jobLevel", null , null));
			headerBuilder.append("&&&");
			headerBuilder.append(this.messageSource.getMessage("report.header.totalTimeAs", null , null));
			headerBuilder.append("&&&");
			headerBuilder.append(this.messageSource.getMessage("report.header.role.count", null , null));
			headerBuilder.append("&&&");
			headerBuilder.append(this.messageSource.getMessage("report.header.flname", null , null));
		} else if ("BSTOAS".equals(reportName)) {
			headerBuilder.append(this.messageSource.getMessage("report.header.email", null , null));
			headerBuilder.append("&&&");
			headerBuilder.append(this.messageSource.getMessage("label.occupation.search", null , null));
			headerBuilder.append("&&&");
			headerBuilder.append(this.messageSource.getMessage("report.header.jobLevel", null , null));
			headerBuilder.append("&&&");			
			headerBuilder.append(this.messageSource.getMessage("report.header.taskStatus", null , null));
			headerBuilder.append("&&&");
			headerBuilder.append(this.messageSource.getMessage("report.header.bsTaskDesc", null , null));
			headerBuilder.append("&&&");
			headerBuilder.append(this.messageSource.getMessage("report.header.asTaskDesc", null , null));
			headerBuilder.append("&&&");
			headerBuilder.append(this.messageSource.getMessage("report.header.diffTaskDesc", null , null));
			headerBuilder.append("&&&");			
			headerBuilder.append(this.messageSource.getMessage("report.header.bsTimeEnrgy", null , null));
			headerBuilder.append("&&&");
			headerBuilder.append(this.messageSource.getMessage("report.header.asTimeEnrgy", null , null));
			headerBuilder.append("&&&");
			headerBuilder.append(this.messageSource.getMessage("report.header.diffTimeEnrgy", null , null));
			headerBuilder.append("&&&");
			headerBuilder.append(this.messageSource.getMessage("report.header.display", null , null));
			headerBuilder.append("&&&");
			headerBuilder.append(this.messageSource.getMessage("report.header.to", null , null));
			headerBuilder.append("&&&");
			headerBuilder.append(this.messageSource.getMessage("report.header.outOf", null , null));
			headerBuilder.append("&&&");
			headerBuilder.append(this.messageSource.getMessage("report.header.export", null , null));
			headerBuilder.append("&&&");
			headerBuilder.append(this.messageSource.getMessage("report.header.previous", null , null));
			headerBuilder.append("&&&");
			headerBuilder.append(this.messageSource.getMessage("report.header.next", null , null));
			headerBuilder.append("&&&");
			headerBuilder.append(this.messageSource.getMessage("report.header.noRecord", null , null));
			headerBuilder.append("&&&");
			headerBuilder.append(this.messageSource.getMessage("report.header.flname", null , null));
		}		
		LOGGER.info("<<<<<< ReportsController.createTableHeader");
		return headerBuilder.toString();
	}
	/**
	 * Method creates After Diagram mapping details report based on user email
	 * @param body
	 * @param request
	 * @return json
	 */
	@RequestMapping(value = ACTION_AFTER_SKETCH_MAPPING_DETAILS, method = RequestMethod.POST)
	@ResponseBody()
	public ReportsVO populateAfterSketchMappingDetails(@RequestBody String body, HttpServletRequest request) {
		LOGGER.info(">>>>>> ReportsController.populateAfterSketchMappingDetails");
		ReportsVO reportVO = new ReportsVO();
		JsonNode node = CommonUtility.getNode(body);
		String email =  node.get("useremail").toString().replaceAll("\"", "");
		String mappings = iReportService.getASDetailedViewOfMappings(email);
		StringTokenizer token = new StringTokenizer(mappings, "&&&");
		while (token.hasMoreTokens()) {
			reportVO.setStrengthString((String) token.nextToken());
			reportVO.setValueString((String) token.nextToken());
			reportVO.setPassionString((String) token.nextToken());
		}
		LOGGER.info("<<<<<< ReportsController.populateAfterSketchMappingDetails");
		return reportVO;
	}
	/**
	 * Method formats the integer to hh:mm:ss format
	 * @param secondsCount
	 * @return string 
	 */
	private String formatHHMMSS(int secondsCount){  
		LOGGER.info(">>>>>> ReportsController.formatHHMMSS");
	    //Calculate the seconds to display:  
	    int seconds = secondsCount % 60;  
	    secondsCount -= seconds;  
	    //Calculate the minutes:  
	    long minutesCount = secondsCount / 60;  
	    long minutes = minutesCount % 60;  
	    minutesCount -= minutes;  
	    //Calculate the hours:  
	    long hoursCount = minutesCount / 60;  
	    //Build the String after formatting
	    StringBuilder sb = new StringBuilder("");
	    if (hoursCount < 10) {
	    	sb.append("0"+hoursCount);
	    } else {
	    	sb.append(hoursCount);
	    }
	    sb.append(" : ");
	    if (minutes < 10) {
	    	sb.append("0"+minutes);
	    } else {
	    	sb.append(minutes);
	    }
	    sb.append(" : ");
	    if (seconds < 10) {
	    	sb.append("0"+seconds);
	    } else {
	    	sb.append(seconds);
	    }
	    LOGGER.info("<<<<<< ReportsController.formatHHMMSS");
	    return sb.toString();  
	}
	/**
	 * Method populates user profile list
	 * @param body
	 * @param request
	 * @return report vo
	 */
	@RequestMapping(value = ACTION_USER_PROFILE_DROPDOWN, method = RequestMethod.POST)
	@ResponseBody()
	public ReportsVO populateProfileDropDown(@RequestBody String body, HttpServletRequest request) {
		LOGGER.info(">>>>>> ReportsController.populateProfileDropDown");
		ReportsVO reportVO = new ReportsVO();
		try {	
			reportVO = this.populateUserProfile(reportVO);
			reportVO.setStatusCode(StatusConstants.STATUS_SUCCESS);
			reportVO.setStatusDesc("SUCCESS");
		} catch (JCTException e) {
			reportVO.setStatusCode(StatusConstants.STATUS_FAILURE);
			reportVO.setStatusDesc(e.getLocalizedMessage());
			LOGGER.error(e.getLocalizedMessage());
		}
		LOGGER.info("<<<<<< ReportsController.populateProfileDropDown");
		return reportVO;
	}
	/**
	 * Method populates user Profile map
	 * @param reportVO
	 * @return map in vo
	 * @throws JCTException
	 */
	private ReportsVO populateUserProfile(ReportsVO reportVO) throws JCTException {
		LOGGER.info(">>>>>> ReportsController.populateUserProfile");
		List<UserProfileDTO> jctUserProfileList = null;
		Map<Integer, String> userProfileMap = new LinkedHashMap<Integer, String>();
		jctUserProfileList = authenticatorService.jctUserProfileList();
			int j = 0;
			for (int i = 1; i<=jctUserProfileList.size();i++) {
				userProfileMap.put(jctUserProfileList.get(j).getUserProfileId(), jctUserProfileList.get(j).getUserProfileDesc());
				j++;
			}
			reportVO.setUserProfile(userProfileMap);
			LOGGER.info("<<<<<< ReportsController.populateUserProfile");
		return reportVO;
	}
	/**
	 * Method generates group profile excel report
	 * @param userProfile
	 * @param reportName
	 * @param request
	 * @param response
	 */
	@RequestMapping(value = "/generateExcelGroupProfileReport/{userProfile}/{reportName}", method = RequestMethod.GET)
	public void generateExcelGroupProfileReport(@PathVariable("userProfile") String userProfile,
			@PathVariable("reportName") String reportName, HttpServletRequest request, HttpServletResponse response) {
		LOGGER.info(">>>>>> ReportsController.generateExcelGroupProfileReport");
		InputStream is = null;
		if (userProfile.equals("JCTDUMMYVARJCTFN")) {
			userProfile = "";
		}		
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
			excelSheet.setColumnView(3, 0);
			excelSheet.setColumnView(4, 0);
			excelSheet.setColumnView(5, 0);
			excelSheet.setColumnView(6, 0);
			excelSheet.setColumnView(7, 46);
			excelSheet.setColumnView(8, 46);
			excelSheet.setColumnView(9, 46);
			excelSheet.setColumnView(10, 26);
			excelSheet.setColumnView(11, 26);
			excelSheet.setColumnView(12, 26);
			excelSheet.setColumnView(13, 26);
			excelSheet.setColumnView(14, 26);
			excelSheet.setColumnView(15, 26);
			excelSheet.setColumnView(16, 26);
			excelSheet.setColumnView(17, 26);
			excelSheet.setColumnView(18, 26);
			excelSheet.setColumnView(19, 26);
			excelSheet.setColumnView(20, 26);
			excelSheet.setColumnView(21, 26);
			excelSheet.setColumnView(22, 26);
			excelSheet.setColumnView(23, 26);
			excelSheet.setColumnView(24, 26);
			excelSheet.setColumnView(25, 26);
			excelSheet.setColumnView(26, 26);
			excelSheet.setColumnView(27, 26);
			excelSheet.setColumnView(28, 26);
			excelSheet.setColumnView(29, 26);
			excelSheet.setColumnView(30, 26);
			excelSheet.setColumnView(31, 26);
			excelSheet.setColumnView(32, 26);
			excelSheet.setColumnView(33, 26);
			excelSheet.setColumnView(34, 26);
			excelSheet.setColumnView(35, 26);
			excelSheet.setColumnView(36, 26);
			excelSheet.setColumnView(37, 26);
			excelSheet.setColumnView(38, 26);
			excelSheet.setColumnView(39, 26);
			excelSheet.setColumnView(40, 26);
			excelSheet.setColumnView(41, 26);
			excelSheet.setColumnView(42, 26);
			excelSheet.setColumnView(43, 26);
			excelSheet.setColumnView(44, 26);
			excelSheet.setColumnView(45, 26);
			excelSheet.setColumnView(46, 26);
			excelSheet.setColumnView(47, 26);
			excelSheet.setColumnView(48, 26);
			excelSheet.setColumnView(49, 26);
			excelSheet.setColumnView(40, 26);
			excelSheet.setColumnView(41, 26);
			excelSheet.setColumnView(42, 26);
			excelSheet.setColumnView(43, 26);
			excelSheet.setColumnView(44, 26);
			excelSheet.setColumnView(45, 26);
			excelSheet.setColumnView(46, 26);
			excelSheet.setColumnView(47, 26);
			excelSheet.setColumnView(48, 26);
			excelSheet.setColumnView(49, 26);
			excelSheet.setColumnView(50, 26);
			excelSheet.setColumnView(51, 26);
			excelSheet.setColumnView(52, 26);
			excelSheet.setColumnView(53, 26);
			excelSheet.setColumnView(54, 26);
			excelSheet.setColumnView(55, 26);
			excelSheet.setColumnView(56, 26);
			excelSheet.setColumnView(57, 26);
			excelSheet.setColumnView(58, 26);
			excelSheet.setColumnView(59, 26);
			excelSheet.setColumnView(60, 26);
			excelSheet.setColumnView(61, 26);
			excelSheet.setColumnView(62, 26);
			excelSheet.setColumnView(63, 26);
			excelSheet.setColumnView(64, 26);
			excelSheet.setColumnView(65, 26);
			excelSheet.setColumnView(66, 26);
			excelSheet.setColumnView(67, 26);
			excelSheet.setColumnView(68, 26);
			excelSheet.setColumnView(69, 26);
			excelSheet.setColumnView(70, 26);
			excelSheet.setColumnView(71, 26);
			excelSheet.setColumnView(72, 26);
			excelSheet.setColumnView(73, 26);
			excelSheet.setColumnView(74, 26);
			excelSheet.setColumnView(75, 26);
			excelSheet.setColumnView(76, 26);
			excelSheet.setColumnView(77, 26);
			excelSheet.setColumnView(78, 26);
			excelSheet.setColumnView(79, 26);
			excelSheet.setColumnView(80, 26);
			excelSheet.setColumnView(81, 26);
			excelSheet.setColumnView(82, 26);
			excelSheet.setColumnView(83, 26);
			excelSheet.setColumnView(84, 26);
			excelSheet.setColumnView(85, 26);
			excelSheet.setColumnView(86, 26);
			excelSheet.setColumnView(87, 26);
			excelSheet.setColumnView(88, 26);
			excelSheet.setColumnView(89, 26);
			excelSheet.setColumnView(90, 26);
			excelSheet.setColumnView(91, 26);
			excelSheet.setColumnView(92, 26);
			excelSheet.setColumnView(93, 26);
			excelSheet.setColumnView(94, 26);
			excelSheet.setColumnView(95, 26);
			excelSheet.setColumnView(96, 26);
			excelSheet.setColumnView(97, 26);
			excelSheet.setColumnView(98, 26);
			excelSheet.setColumnView(99, 26);
			excelSheet.setColumnView(100, 26);
			excelSheet.setColumnView(101, 26);
			excelSheet.setColumnView(102, 26);
			excelSheet.setColumnView(103, 26);
			excelSheet.setColumnView(104, 26);
			excelSheet.setColumnView(105, 26);
			excelSheet.setColumnView(106, 26);
			excelSheet.setColumnView(107, 26);
			excelSheet.setColumnView(108, 26);
			excelSheet.setColumnView(109, 26);
			excelSheet.setColumnView(110, 26);
			excelSheet.setColumnView(111, 26);
			excelSheet.setColumnView(112, 26);
			excelSheet.setColumnView(113, 26);
			excelSheet.setColumnView(114, 26);
			excelSheet.setColumnView(115, 26);
			excelSheet.setColumnView(116, 26);
			excelSheet.setColumnView(117, 26);
			excelSheet.setColumnView(118, 26);
			excelSheet.setColumnView(119, 26);
			excelSheet.setColumnView(120, 26);
			excelSheet.setColumnView(121, 26);
			excelSheet.setColumnView(122, 26);
			excelSheet.setColumnView(123, 26);
			excelSheet.setColumnView(124, 26);
			excelSheet.setColumnView(125, 26);
			excelSheet.setColumnView(126, 26);
			excelSheet.setColumnView(127, 26);
			excelSheet.setColumnView(128, 26);
			excelSheet.setColumnView(129, 26);
			excelSheet.setColumnView(130, 26);
			excelSheet.setColumnView(131, 26);
			excelSheet.setColumnView(132, 26);
			excelSheet.setColumnView(133, 26);
			excelSheet.setColumnView(134, 26);
			excelSheet.setColumnView(135, 26);
			excelSheet.setColumnView(136, 26);
			excelSheet.setColumnView(137, 26);
			excelSheet.setColumnView(138, 26);
			excelSheet.setColumnView(139, 26);
			excelSheet.setColumnView(140, 26);
			excelSheet.setColumnView(141, 26);
			excelSheet.setColumnView(142, 26);
			excelSheet.setColumnView(143, 26);
			excelSheet.setColumnView(144, 26);
			excelSheet.setColumnView(145, 26);
			excelSheet.setColumnView(146, 26);
			excelSheet.setColumnView(147, 26);
			excelSheet.setColumnView(148, 26);
			excelSheet.setColumnView(149, 26);
			excelSheet.setColumnView(150, 26);
			excelSheet.setColumnView(151, 26);
			excelSheet.setColumnView(152, 26);
			excelSheet.setColumnView(153, 26);
			excelSheet.setColumnView(154, 26);
			excelSheet.setColumnView(155, 26);
			excelSheet.setColumnView(156, 26);
			excelSheet.setColumnView(157, 26);
			excelSheet.setColumnView(158, 26);
			excelSheet.setColumnView(159, 26);
			excelSheet.setColumnView(160, 26);
			excelSheet.setColumnView(161, 26);
			excelSheet.setColumnView(162, 26);
			excelSheet.setColumnView(163, 26);
			excelSheet.setColumnView(164, 26);
			excelSheet.setColumnView(165, 26);
			excelSheet.setColumnView(166, 26);
			excelSheet.setColumnView(167, 26);
			excelSheet.setColumnView(168, 26);
			excelSheet.setColumnView(169, 26);
			excelSheet.setColumnView(170, 26);
			excelSheet.setColumnView(171, 26);
			excelSheet.setColumnView(172, 26);
			excelSheet.setColumnView(173, 26);
			excelSheet.setColumnView(174, 26);
			excelSheet.setColumnView(175, 26);
			excelSheet.setColumnView(176, 26);
			excelSheet.setColumnView(177, 26);
			excelSheet.setColumnView(178, 26);
			excelSheet.setColumnView(179, 26);
			excelSheet.setColumnView(180, 26);
			excelSheet.setColumnView(181, 26);
			excelSheet.setColumnView(182, 26);
			excelSheet.setColumnView(183, 26);
			excelSheet.setColumnView(184, 26);
			excelSheet.setColumnView(185, 26);
			excelSheet.setColumnView(186, 26);
			excelSheet.setColumnView(187, 26);
			excelSheet.setColumnView(188, 26);
			excelSheet.setColumnView(189, 26);
			excelSheet.setColumnView(190, 26);
			excelSheet.setColumnView(191, 26);
			excelSheet.setColumnView(192, 26);
			excelSheet.setColumnView(193, 26);
			excelSheet.setColumnView(194, 26);
			excelSheet.setColumnView(195, 26);
			excelSheet.setColumnView(196, 26);
			excelSheet.setColumnView(197, 26);
			excelSheet.setColumnView(198, 26);
			excelSheet.setColumnView(199, 26);
			excelSheet.setColumnView(200, 26);
			excelSheet.setColumnView(201, 26);
			excelSheet.setColumnView(202, 26);
			excelSheet.setColumnView(203, 26);
			excelSheet.setColumnView(204, 26);
			excelSheet.setColumnView(205, 26);
			excelSheet.setColumnView(206, 26);
			excelSheet.setColumnView(207, 26);
			excelSheet.setColumnView(208, 26);
			excelSheet.setColumnView(209, 26);
			excelSheet.setColumnView(210, 26);
			excelSheet.setColumnView(211, 26);
			excelSheet.setColumnView(212, 26);
			excelSheet.setColumnView(213, 26);			
			createGPLabel(excelSheet,0);
			createGroupProfileReport(excelSheet, userProfile);
			workbook.write();
			workbook.close();
			writeToStream(response.getOutputStream(), response, file, is, generateExcelFileName("group_profile_report"));
			LOGGER.info("<<<<<< ReportsController.generateExcelGroupProfileReport");
		}catch(Exception ex){
			LOGGER.error(ex.getLocalizedMessage());
		}		
	}
	/**
	 * Method creates group profile report header in excel
	 * @param sheet
	 * @param distinction
	 * @throws WriteException
	 */
	private void createGPLabel(WritableSheet sheet, int distinction) throws WriteException {
		LOGGER.info(">>>>>> ReportsController.createGPLabel");
		addCaption(sheet, 0, 0, "User Profile Name");
		addCaption(sheet, 1, 0, "User Profile Creation Date");
		addCaption(sheet, 2, 0, "Total No of User  Group");
		/*addCaption(sheet, 3, 0, "Main Header");
		addCaption(sheet, 4, 0, "Before Sketch Header");
		addCaption(sheet, 5, 0, "Mapping Yourself Header");
		addCaption(sheet, 6, 0, "After Diagram Header");*/
		addCaption(sheet, 3, 0, "");
		addCaption(sheet, 4, 0, "");
		addCaption(sheet, 5, 0, "");
		addCaption(sheet, 6, 0, "");
		addCaption(sheet, 7, 0, "Before Sketch Instruction");
		addCaption(sheet, 8, 0, "Mapping Yourself Instruction");
		addCaption(sheet, 9, 0, "After Diagram Instruction");
		addCaption(sheet, 10, 0, "Before Reflection - Question 1");
		addCaption(sheet, 11, 0, "Before Reflection - Sub Question 1a");
		addCaption(sheet, 12, 0, "Before Reflection - Sub Question 1b");
		addCaption(sheet, 13, 0, "Before Reflection - Sub Question 1c");
		addCaption(sheet, 14, 0, "Before Reflection - Sub Question 1d");
		
		addCaption(sheet, 15, 0, "Before Reflection - Question 2");
		addCaption(sheet, 16, 0, "Before Reflection - Sub Question 2a");
		addCaption(sheet, 17, 0, "Before Reflection - Sub Question 2b");
		addCaption(sheet, 18, 0, "Before Reflection - Sub Question 2c");
		addCaption(sheet, 19, 0, "Before Reflection - Sub Question 2d");
		
		addCaption(sheet, 20, 0, "Before Reflection - Question 3");
		addCaption(sheet, 21, 0, "Before Reflection - Sub Question 3a");
		addCaption(sheet, 22, 0, "Before Reflection - Sub Question 3b");
		addCaption(sheet, 23, 0, "Before Reflection - Sub Question 3c");
		addCaption(sheet, 24, 0, "Before Reflection - Sub Question 3d");
		
		addCaption(sheet, 25, 0, "Before Reflection - Question 4");
		addCaption(sheet, 26, 0, "Before Reflection - Sub Question 4a");
		addCaption(sheet, 27, 0, "Before Reflection - Sub Question 4b");
		addCaption(sheet, 28, 0, "Before Reflection - Sub Question 4c");
		addCaption(sheet, 29, 0, "Before Reflection - Sub Question 4d");
		
		addCaption(sheet, 30, 0, "Before Reflection - Question 5");
		addCaption(sheet, 31, 0, "Before Reflection - Sub Question 5a");
		addCaption(sheet, 32, 0, "Before Reflection - Sub Question 5b");
		addCaption(sheet, 33, 0, "Before Reflection - Sub Question 5c");
		addCaption(sheet, 34, 0, "Before Reflection - Sub Question 5d");
		
		addCaption(sheet, 35, 0, "Before Reflection - Question 6");
		addCaption(sheet, 36, 0, "Before Reflection - Question 6a");
		addCaption(sheet, 37, 0, "Before Reflection - Question 6b");
		addCaption(sheet, 38, 0, "Before Reflection - Question 6c");
		addCaption(sheet, 39, 0, "Before Reflection - Question 6d");
		
		addCaption(sheet, 40, 0, "After Action - Question 1");
		addCaption(sheet, 41, 0, "After Action - Sub Question 1a");
		addCaption(sheet, 42, 0, "After Action - Sub Question 1b");
		addCaption(sheet, 43, 0, "After Action - Sub Question 1c");
		addCaption(sheet, 44, 0, "After Action - Sub Question 1d");
		
		addCaption(sheet, 45, 0, "After Action - Question 2");
		addCaption(sheet, 46, 0, "After Action - Sub Question 2a");
		addCaption(sheet, 47, 0, "After Action - Sub Question 2b");
		addCaption(sheet, 48, 0, "After Action - Sub Question 2c");
		addCaption(sheet, 49, 0, "After Action - Sub Question 2d");
		
		addCaption(sheet, 50, 0, "After Action - Question 3");
		addCaption(sheet, 51, 0, "After Action - Sub Question 3a");
		addCaption(sheet, 52, 0, "After Action - Sub Question 3b");
		addCaption(sheet, 53, 0, "After Action - Sub Question 3c");
		addCaption(sheet, 54, 0, "After Action - Sub Question 3d");		
		
		addCaption(sheet, 55, 0, "After Action - Question 4");
		addCaption(sheet, 56, 0, "After Action - Sub Question 4a");
		addCaption(sheet, 57, 0, "After Action - Sub Question 4b");
		addCaption(sheet, 58, 0, "After Action - Sub Question 4c");
		addCaption(sheet, 59, 0, "After Action - Sub Question 4d");
		
		addCaption(sheet, 60, 0, "After Action - Question 5");
		addCaption(sheet, 61, 0, "After Action - Sub Question 5a");
		addCaption(sheet, 62, 0, "After Action - Sub Question 5b");
		addCaption(sheet, 63, 0, "After Action - Sub Question 5c");
		addCaption(sheet, 64, 0, "After Action - Sub Question 5d");
		
		addCaption(sheet, 65, 0, "After Action - Question 6");
		addCaption(sheet, 66, 0, "After Action - Sub Question 6a");
		addCaption(sheet, 67, 0, "After Action - Sub Question 6b");
		addCaption(sheet, 68, 0, "After Action - Sub Question 6c");
		addCaption(sheet, 69, 0, "After Action - Sub Question 6d");
		
		addCaption(sheet, 70, 0, "Value 1");
		addCaption(sheet, 71, 0, "Description");
		addCaption(sheet, 72, 0, "Value 2");
		addCaption(sheet, 73, 0, "Description");
		addCaption(sheet, 74, 0, "Value 3");
		addCaption(sheet, 75, 0, "Description");
		addCaption(sheet, 76, 0, "Value 4");
		addCaption(sheet, 77, 0, "Description");
		addCaption(sheet, 78, 0, "Value 5");
		addCaption(sheet, 79, 0, "Description");
		addCaption(sheet, 80, 0, "Value 6");
		addCaption(sheet, 81, 0, "Description");
		addCaption(sheet, 82, 0, "Value 7");
		addCaption(sheet, 83, 0, "Description");
		addCaption(sheet, 84, 0, "Value 8");
		addCaption(sheet, 85, 0, "Description");
		addCaption(sheet, 86, 0, "Value 9");
		addCaption(sheet, 87, 0, "Description");
		addCaption(sheet, 88, 0, "Value 10");
		addCaption(sheet, 89, 0, "Description");
		addCaption(sheet, 90, 0, "Value 11");
		addCaption(sheet, 91, 0, "Description");
		addCaption(sheet, 92, 0, "Value 12");
		addCaption(sheet, 93, 0, "Description");
		addCaption(sheet, 94, 0, "Value 13");
		addCaption(sheet, 95, 0, "Description");
		addCaption(sheet, 96, 0, "Value 14");
		addCaption(sheet, 97, 0, "Description");
		addCaption(sheet, 98, 0, "Value 15");
		addCaption(sheet, 99, 0, "Description");
		addCaption(sheet, 100, 0, "Value 16");
		addCaption(sheet, 101, 0, "Description");
		addCaption(sheet, 102, 0, "Value 17");
		addCaption(sheet, 103, 0, "Description");
		addCaption(sheet, 104, 0, "Value 18");
		addCaption(sheet, 105, 0, "Description");
		addCaption(sheet, 106, 0, "Value 19");
		addCaption(sheet, 107, 0, "Description");
		addCaption(sheet, 108, 0, "Value 20");
		addCaption(sheet, 109, 0, "Description");
		addCaption(sheet, 110, 0, "Value 21");
		addCaption(sheet, 111, 0, "Description");
		addCaption(sheet, 112, 0, "Value 22");
		addCaption(sheet, 113, 0, "Description");
		addCaption(sheet, 114, 0, "Value 23");
		addCaption(sheet, 115, 0, "Description");
		addCaption(sheet, 116, 0, "Value 24");
		addCaption(sheet, 117, 0, "Description");
		
		
		addCaption(sheet, 118, 0, "Strength 1");
		addCaption(sheet, 119, 0, "Description");
		addCaption(sheet, 120, 0, "Strength 2");
		addCaption(sheet, 121, 0, "Description");
		addCaption(sheet, 122, 0, "Strength 3");
		addCaption(sheet, 123, 0, "Description");
		addCaption(sheet, 124, 0, "Strength 4");
		addCaption(sheet, 125, 0, "Description");
		addCaption(sheet, 126, 0, "Strength 5");
		addCaption(sheet, 127, 0, "Description");
		addCaption(sheet, 128, 0, "Strength 6");
		addCaption(sheet, 129, 0, "Description");
		addCaption(sheet, 130, 0, "Strength 7");
		addCaption(sheet, 131, 0, "Description");
		addCaption(sheet, 132, 0, "Strength 8");
		addCaption(sheet, 133, 0, "Description");
		addCaption(sheet, 134, 0, "Strength 9");
		addCaption(sheet, 135, 0, "Description");
		addCaption(sheet, 136, 0, "Strength 10");
		addCaption(sheet, 137, 0, "Description");
		addCaption(sheet, 138, 0, "Strength 11");
		addCaption(sheet, 139, 0, "Description");
		addCaption(sheet, 140, 0, "Strength 12");
		addCaption(sheet, 141, 0, "Description");
		addCaption(sheet, 142, 0, "Strength 13");
		addCaption(sheet, 143, 0, "Description");
		addCaption(sheet, 144, 0, "Strength 14");
		addCaption(sheet, 145, 0, "Description");
		addCaption(sheet, 146, 0, "Strength 15");
		addCaption(sheet, 147, 0, "Description");
		addCaption(sheet, 148, 0, "Strength 16");
		addCaption(sheet, 149, 0, "Description");
		addCaption(sheet, 150, 0, "Strength 17");
		addCaption(sheet, 151, 0, "Description");
		addCaption(sheet, 152, 0, "Strength 18");
		addCaption(sheet, 153, 0, "Description");
		addCaption(sheet, 154, 0, "Strength 19");
		addCaption(sheet, 155, 0, "Description");
		addCaption(sheet, 156, 0, "Strength 20");
		addCaption(sheet, 157, 0, "Description");
		addCaption(sheet, 158, 0, "Strength 21");
		addCaption(sheet, 159, 0, "Description");
		addCaption(sheet, 160, 0, "Strength 22");
		addCaption(sheet, 161, 0, "Description");
		addCaption(sheet, 162, 0, "Strength 23");
		addCaption(sheet, 163, 0, "Description");
		addCaption(sheet, 164, 0, "Strength 24");
		addCaption(sheet, 165, 0, "Description");
		
		addCaption(sheet, 166, 0, "Passion 1");
		addCaption(sheet, 167, 0, "Description");
		addCaption(sheet, 168, 0, "Passion 2");
		addCaption(sheet, 169, 0, "Description");
		addCaption(sheet, 170, 0, "Passion 3");
		addCaption(sheet, 171, 0, "Description");
		addCaption(sheet, 172, 0, "Passion 4");
		addCaption(sheet, 173, 0, "Description");
		addCaption(sheet, 174, 0, "Passion 5");
		addCaption(sheet, 175, 0, "Description");
		addCaption(sheet, 176, 0, "Passion 6");
		addCaption(sheet, 177, 0, "Description");
		addCaption(sheet, 178, 0, "Passion 7");
		addCaption(sheet, 179, 0, "Description");
		addCaption(sheet, 180, 0, "Passion 8");
		addCaption(sheet, 181, 0, "Description");
		addCaption(sheet, 182, 0, "Passion 9");
		addCaption(sheet, 183, 0, "Description");
		addCaption(sheet, 184, 0, "Passion 10");
		addCaption(sheet, 185, 0, "Description");
		addCaption(sheet, 186, 0, "Passion 11");
		addCaption(sheet, 187, 0, "Description");
		addCaption(sheet, 188, 0, "Passion 12");
		addCaption(sheet, 189, 0, "Description");
		addCaption(sheet, 190, 0, "Passion 13");
		addCaption(sheet, 191, 0, "Description");
		addCaption(sheet, 192, 0, "Passion 14");
		addCaption(sheet, 193, 0, "Description");
		addCaption(sheet, 194, 0, "Passion 15");
		addCaption(sheet, 195, 0, "Description");
		addCaption(sheet, 196, 0, "Passion 16");
		addCaption(sheet, 197, 0, "Description");
		addCaption(sheet, 198, 0, "Passion 17");
		addCaption(sheet, 199, 0, "Description");
		addCaption(sheet, 200, 0, "Passion 18");
		addCaption(sheet, 201, 0, "Description");
		addCaption(sheet, 202, 0, "Passion 19");
		addCaption(sheet, 203, 0, "Description");
		addCaption(sheet, 204, 0, "Passion 20");
		addCaption(sheet, 205, 0, "Description");
		addCaption(sheet, 206, 0, "Passion 21");
		addCaption(sheet, 207, 0, "Description");
		addCaption(sheet, 208, 0, "Passion 22");
		addCaption(sheet, 209, 0, "Description");
		addCaption(sheet, 210, 0, "Passion 23");
		addCaption(sheet, 211, 0, "Description");
		addCaption(sheet, 212, 0, "Passion 24");
		addCaption(sheet, 213, 0, "Description");
		
		LOGGER.info("<<<<<< ReportsController.createGPLabel");
	}
	/**
	 * Method populates the group profile report
	 * @param excelSheet
	 * @param userProfile
	 * @throws RowsExceededException
	 * @throws WriteException
	 * @throws JCTException
	 */
	private void createGroupProfileReport(WritableSheet excelSheet, String userProfile) 
			throws RowsExceededException, WriteException, JCTException {
		LOGGER.info(">>>>>> ReportsController.createGroupProfileReport");
		List<Object> viewList1 = null;
		List<Object> viewList = null;
		viewList1 = iReportService.getProfileDescListForExcel(Integer.parseInt(userProfile));
		viewList = iReportService.getGroupProfileListForExcel(Integer.parseInt(userProfile));
			
		ReportsVO reportVO = new ReportsVO();
		reportVO = this.populateOtherDescReport(viewList, reportVO);
		reportVO = this.populateProfileDescReport(viewList1, reportVO);
		
		int rowIndex = 1;
		int colIndex = 0;		
			/**
			 * For profile name, created date  and total number of group
			 */	
			String rptStrProfile = reportVO.getReportProfileDescList();
			StringTokenizer token6 = new StringTokenizer(rptStrProfile, "$^$");				
			try {
				while (token6.hasMoreTokens()) {
					if (colIndex == 3) {
						colIndex = 0;
						rowIndex = rowIndex + 1;				
					} else {
						rowIndex = 1;
					}
					String dispObjProfile = token6.nextToken().toString();
					String[] initialValue = dispObjProfile.split("#");
					for (int index = 0; index < initialValue.length; index++) {				
						if (null != initialValue[index] || initialValue[index] != "") {
							addLabelGroupProfile(excelSheet, colIndex, rowIndex, initialValue[index]);
							colIndex = colIndex + 1;
						} else {
							addLabelGroupProfile(excelSheet, colIndex, rowIndex, "");
							colIndex = colIndex + 1;
						}
					}
				}
			} catch (Exception e){	
				LOGGER.error(e.getLocalizedMessage());
				for (int index = 0; index < 3; index++) {	
					addLabelGroupProfile(excelSheet, colIndex, rowIndex, "");
					colIndex = colIndex + 1;
				}
			}
			
			/**
			 * For header description
			 */		
			for (int index = 0; index < 4; index++) {	
				addLabelGroupProfile(excelSheet, colIndex, rowIndex, " ");
				colIndex = colIndex + 1;				
			}
		
			/**
			 * For instruction bar description
			 */
			String rptStr = reportVO.getReportInstructionList();
			StringTokenizer token = new StringTokenizer(rptStr, "$^$");	
			
			try {
				while (token.hasMoreTokens()) {
				if (colIndex == 10) {
					colIndex = 7;
					rowIndex = rowIndex + 1;				
				} else {
					rowIndex = 1;
				}
				String dispObj = token.nextToken().toString();
				String[] initialArr = dispObj.split("#");
				String instrctionVal = "";
				for (int index = 0; index < initialArr.length; index++) {				
					if (null != initialArr[index] 
							|| !initialArr[index].equals("")
							|| !initialArr[index].isEmpty()) {
						instrctionVal = initialArr[index].replace("[", "");
						instrctionVal = instrctionVal.replace("]", "");
						instrctionVal = instrctionVal.replace("<li>", "");
						instrctionVal = instrctionVal.replace("</li>", "");
						instrctionVal = instrctionVal.replace("</li>", "");
						instrctionVal = instrctionVal.replace("<ul>", "");
						instrctionVal = instrctionVal.replace("</ul>", "");
						instrctionVal = instrctionVal.replace("<ol>", "");
						instrctionVal = instrctionVal.replace("</ol>", "");
						instrctionVal = instrctionVal.replace("<u>", "");
						instrctionVal = instrctionVal.replace("</u>", "");
						instrctionVal = instrctionVal.replace("<b>", "");
						instrctionVal = instrctionVal.replace("</b>", "");
						instrctionVal = instrctionVal.replace("<i>", "");
						instrctionVal = instrctionVal.replace("</i>", "");
						instrctionVal = instrctionVal.replace("<br>", "");
						instrctionVal = instrctionVal.replace("<br />", "");
						instrctionVal = instrctionVal.replace("&nbsp;", "");
						instrctionVal = instrctionVal.replace("<p>", "");
						instrctionVal = instrctionVal.replace("</p>", "");
						instrctionVal = instrctionVal.replace("</span>", "");
						instrctionVal = instrctionVal.replace("<strong>", "");
						instrctionVal = instrctionVal.replace("</strong>", "");
						int startIndex = 0, endIndex = 0;
						while(instrctionVal.contains("<span")){
							startIndex = instrctionVal.indexOf("<span");
							endIndex = instrctionVal.indexOf(">")+1;
							instrctionVal = instrctionVal.replace(instrctionVal.substring(startIndex, endIndex), "");
							
						}
						addLabelGroupProfile(excelSheet, colIndex, rowIndex, instrctionVal);
						colIndex = colIndex + 1;
					} else {
						addLabelGroupProfile(excelSheet, colIndex, rowIndex, "");
						colIndex = colIndex + 1;
					}
				}
				}
			} catch (Exception e){	
				LOGGER.error(e.getLocalizedMessage());
				for (int index = 0; index < 3; index++) {	
					addLabelGroupProfile(excelSheet, colIndex, rowIndex, "");
					colIndex = colIndex + 1;
				}
			}
			
			
			/**
			 * For reflection question
			 */
			String rptStr1 = reportVO.getReportRefQtnList();
			StringTokenizer tokn = new StringTokenizer(rptStr1, "&");			
			while (tokn.hasMoreTokens()) {
				String toknVal = tokn.nextToken().toString();
				StringTokenizer token1 = new StringTokenizer(toknVal, "$^$");
				try {
					while (token1.hasMoreTokens()) {
						if (colIndex == 40) { //16
							colIndex = 10;  //10
							rowIndex = rowIndex + 1;				
						} else {
							rowIndex = 1;
						}
					String dispObj1 = token1.nextToken().toString();
					String[] initialArr1 = dispObj1.split("#");
					int diff1 = 5 - initialArr1.length;
					for (int index = 0; index < initialArr1.length; index++) {	
						if(initialArr1[index].equals("NA")) {
							initialArr1[index] = "";
						}
						addLabelGroupProfile(excelSheet, colIndex, rowIndex, initialArr1[index]);
						colIndex = colIndex + 1;
					} 
					if (diff1 != 0) {
						for (int index = 0; index < diff1; index++) {
							addLabelGroupProfile(excelSheet, colIndex, rowIndex, "");
							colIndex = colIndex + 1;
						}
						}
					}
					if(colIndex < 40) {
						for (int index = colIndex+1; index <= 40; index++) {
							addLabelGroupProfile(excelSheet, colIndex, rowIndex, "");
							colIndex = colIndex + 1;
						}
					}
				} catch (Exception e){
					LOGGER.error(e.getLocalizedMessage());
					for (int index = 0; index < 30; index++) {	
						addLabelGroupProfile(excelSheet, colIndex, rowIndex, "");
						colIndex = colIndex + 1;
					}
				}
			}
												
			/**
			 * For action plan
			 */
			String rptStr2 = reportVO.getReportActionPlanList();
			StringTokenizer tokn1 = new StringTokenizer(rptStr2, "&");	
			while (tokn1.hasMoreTokens()) {
				String toknVal1 = tokn1.nextToken().toString();
				StringTokenizer token2 = new StringTokenizer(toknVal1, "$^$");
				try {
					while (token2.hasMoreTokens()) {
						if (colIndex == 70) { //22
							colIndex = 40;   //16
							rowIndex = rowIndex + 1;				
						} else {						
							rowIndex = 1;
						}
					String dispObj2 = token2.nextToken().toString();
					String[] initialArr2 = dispObj2.split("#");
					int diff2 = 5 - initialArr2.length;
					for (int index = 0; index < initialArr2.length; index++) {	
						if(initialArr2[index].equals("NA")) {
							initialArr2[index] = "";
						}
						addLabelGroupProfile(excelSheet, colIndex, rowIndex, initialArr2[index]);
						colIndex = colIndex + 1;
					} 
					if (diff2 != 0) {
						for (int index = 0; index < diff2; index++) {
							addLabelGroupProfile(excelSheet, colIndex, rowIndex, "");
							colIndex = colIndex + 1;
						}
					}
					}
					if(colIndex < 70) {
						for (int index = colIndex+1; index <= 70; index++) {
							addLabelGroupProfile(excelSheet, colIndex, rowIndex, "");
							colIndex = colIndex + 1;
						}
					}
				} catch (Exception e){
					LOGGER.error(e.getLocalizedMessage());
					for (int index = 0; index < 30; index++) {	
						addLabelGroupProfile(excelSheet, colIndex, rowIndex, "");
						colIndex = colIndex + 1;
					}
				}
			}			
			
			/**
			 * For Value attribute
			 */
			String rptStr3 = reportVO.getReportValueList();
			StringTokenizer token3 = new StringTokenizer(rptStr3, "$^$");
			try {
				while (token3.hasMoreTokens()) {
					if (colIndex == 118) { //54, 102
						colIndex = 70; //22
						rowIndex = rowIndex + 1;				
					} else {
						//colIndex = 70;
						rowIndex = 1;
					}
				String dispObj3 = token3.nextToken().toString();
				String[] initialArr3 = dispObj3.split("#");
				//int diff3 = 32 - initialArr3.length;
				int diff3 = 48 - initialArr3.length;
				for (int index = 0; index < initialArr3.length; index++) {	
					addLabelGroupProfile(excelSheet, colIndex, rowIndex, initialArr3[index]);
					colIndex = colIndex + 1;
				} 
				if (diff3 != 0) {
					for (int index = 0; index < diff3; index++) {
						addLabelGroupProfile(excelSheet, colIndex, rowIndex, "");
						colIndex = colIndex + 1;
					}
				}
				}
			} catch (Exception e){
				LOGGER.error(e.getLocalizedMessage());
				for (int index = 0; index < 48; index++) {	
					addLabelGroupProfile(excelSheet, colIndex, rowIndex, "");
					colIndex = colIndex + 1;
				}
			}			
		
			/**
			 * For Strength attribute
			 */		
			String rptStr4 = reportVO.getReportStrengthList();
			StringTokenizer token4 = new StringTokenizer(rptStr4, "$^$");
			try {
				while (token4.hasMoreTokens()) {
					if (colIndex == 166) { //86, 134
						colIndex = 118; //54
						rowIndex = rowIndex + 1;				
					} else {
						rowIndex = 1;
					}
				String dispObj4 = token4.nextToken().toString();
				String[] initialArr4 = dispObj4.split("#");
				int diff4 = 48 - initialArr4.length;
				for (int index = 0; index < initialArr4.length; index++) {	
					addLabelGroupProfile(excelSheet, colIndex, rowIndex, initialArr4[index]);
					colIndex = colIndex + 1;
				} 
				if (diff4 != 0) {
					for (int index = 0; index < diff4; index++) {
						addLabelGroupProfile(excelSheet, colIndex, rowIndex, "");
						colIndex = colIndex + 1;
					}
				}
				}
			} catch (Exception e) {
				LOGGER.error(e.getLocalizedMessage());
				for (int index = 0; index < 48; index++) {	
					addLabelGroupProfile(excelSheet, colIndex, rowIndex, "");
					colIndex = colIndex + 1;
				}
			}
							
			/**
			 * For Passion attribute
			 */			
			String rptStr5 = reportVO.getReportPassioList();
			StringTokenizer token5 = new StringTokenizer(rptStr5, "$^$");
			try {
				while (token5.hasMoreTokens()) {
					if (colIndex == 214) { //118
						colIndex = 166; //86
						rowIndex = rowIndex + 1;				
					} else {
						rowIndex = 1;
					}
				String dispObj5 = token5.nextToken().toString();
				String[] initialArr5 = dispObj5.split("#");
				int diff5 = 48 - initialArr5.length;
				for (int index = 0; index < initialArr5.length; index++) {	
					addLabelGroupProfile(excelSheet, colIndex, rowIndex, initialArr5[index]);
					colIndex = colIndex + 1;
				} 
				if (diff5 != 0) {
					for (int index = 0; index < diff5; index++) {
						addLabelGroupProfile(excelSheet, colIndex, rowIndex, "");
						colIndex = colIndex + 1;
					}
				}
				}
			} catch (Exception e) {
				LOGGER.error(e.getLocalizedMessage());
				for (int index = 0; index < 48; index++) {	
					addLabelGroupProfile(excelSheet, colIndex, rowIndex, "");
					colIndex = colIndex + 1;
				}
			}	
			LOGGER.info("<<<<<< ReportsController.createGroupProfileReport");	
	}
	/**
	 * Method to populate report
	 * @param viewList
	 * @param reportVO
	 * @return
	 */
	private ReportsVO populateOtherDescReport(List<Object> viewList, ReportsVO reportVO) {
		LOGGER.info(">>>>>> ReportsController.populateOtherDescReport");
		StringBuilder data1 = new StringBuilder("");
		ArrayList<Integer> innerObj1 = new ArrayList<Integer>();
		StringBuilder data2 = new StringBuilder("");
		//ArrayList<Integer> innerObj2 = new ArrayList<Integer>();
		StringBuilder data3 = new StringBuilder("");
		//ArrayList<Integer> innerObj3 = new ArrayList<Integer>();
		StringBuilder data4 = new StringBuilder("");
		//ArrayList<Integer> innerObj4 = new ArrayList<Integer>();
		StringBuilder data5 = new StringBuilder("");
		//ArrayList<Integer> innerObj5 = new ArrayList<Integer>();
		StringBuilder data6 = new StringBuilder("");
		//ArrayList<Integer> innerObj6 = new ArrayList<Integer>();
		int index = 0;
		
		for (int i = 0; i < viewList.size(); i++ ) {
			if(index < viewList.size()) {			
			/***
			 * Get the instruction bar description list
			 */
			innerObj1 = (ArrayList<Integer>) viewList.get(index);
			for(int innerIndex=0; innerIndex < innerObj1.size(); innerIndex++){
				data1.append(innerObj1.get(innerIndex)+ "#");
			}
			index++;
			data1.append("$^$");			
			reportVO.setReportInstructionList(data1.toString());
			
			/***
			 * Get reflection question list
			 */						
			List innerObj2 = (ArrayList) viewList.get(index);
			if(innerObj2.size() == 0){
				data2.append( " " + "#"+" " + "#");
			} else {
				String mainQtn = "";
				for(int innerIndex=0; innerIndex < innerObj2.size(); innerIndex++){
					Object[] innerIndividual1 = (Object[]) innerObj2.get(innerIndex);					
					if(!mainQtn.equalsIgnoreCase((String) innerIndividual1[0])){
						mainQtn = (String) innerIndividual1[0];
						data2.append("$^$"+(String) innerIndividual1[0]+ "#" +(String) innerIndividual1[1]+ "#");
					} else {
						data2.append((String) innerIndividual1[1]+ "#");
					}				
				}
			}		
			index++;
			data2.append("&");	
			reportVO.setReportRefQtnList(data2.toString());
			
			
			/***
			 * Get action plan list
			 */
			
			List innerObj3 = (ArrayList) viewList.get(index);
			if(innerObj3.size() == 0){
				data3.append( " " + "#"+" " + "#");
			} else {
				String mainActionPlan = "";
				for(int innerIndex=0; innerIndex < innerObj3.size(); innerIndex++){
					Object[] innerIndividual1 = (Object[]) innerObj3.get(innerIndex);					
					if(!mainActionPlan.equalsIgnoreCase((String) innerIndividual1[0])){
						mainActionPlan = (String) innerIndividual1[0];
						data3.append("$^$"+(String) innerIndividual1[0]+ "#" +(String) innerIndividual1[1]+ "#");
					} else {
						data3.append((String) innerIndividual1[1]+ "#");
					}				
				}
			}		
			index++;
			data3.append("&");
			reportVO.setReportActionPlanList(data3.toString());
			
			
			/***
			 * Get Value attribute list
			 */
			//innerObj4 = (ArrayList<Integer>) viewList.get(index);
			List innerObj4 = (ArrayList) viewList.get(index);
			if(innerObj4.size() == 0){
				data4.append( " " + "#"+" " + "#");
			} else {
				for(int innerIndex=0; innerIndex < innerObj4.size(); innerIndex++){
					Object[] innerIndividual1 = (Object[]) innerObj4.get(innerIndex);
					data4.append((String) innerIndividual1[0]+ "#" +(String) innerIndividual1[1]+ "#");
					//data4.append(innerObj4.get(innerIndex)+ "#");
				}
			}		
			index++;
			data4.append("$^$");	
			reportVO.setReportValueList(data4.toString());

			/***
			 * Get Strength attribute list
			 */
			//innerObj5 = (ArrayList<Integer>) viewList.get(index);
			List innerObj5 = (ArrayList) viewList.get(index);
			if(innerObj5.size() == 0){
				data5.append( " " + "#"+" " + "#");
			} else {
				for(int innerIndex=0; innerIndex < innerObj5.size(); innerIndex++){
					Object[] innerIndividual2 = (Object[]) innerObj5.get(innerIndex);
					data5.append((String) innerIndividual2[0]+ "#" +(String) innerIndividual2[1]+ "#");
					//data5.append(innerObj5.get(innerIndex)+ "#");
				}
			}			
			index++;
			data5.append("$^$");
			reportVO.setReportStrengthList(data5.toString());
			
			/***
			 * Get Passion attribute list
			 */
			//innerObj6 = (ArrayList<Integer>) viewList.get(index);
			List innerObj6 = (ArrayList) viewList.get(index);
			if(innerObj6.size() == 0){
				data6.append( " " + "#"+" " + "#");
			} else {
				for(int innerIndex=0; innerIndex < innerObj6.size(); innerIndex++){
					Object[] innerIndividual3 = (Object[]) innerObj6.get(innerIndex);
					data6.append((String) innerIndividual3[0]+ "#" +(String) innerIndividual3[1]+ "#");
					//data6.append(innerObj6.get(innerIndex)+ "#");
				}
			}	
			index++;
			data6.append("$^$");	
			reportVO.setReportPassioList(data6.toString());
			}
			
		}
		LOGGER.info("<<<<<< ReportsController.populateOtherDescReport");
		return reportVO;
	}
	/**
	 * Method populates profile description report
	 * @param viewList
	 * @param reportVO
	 * @return
	 */
	private ReportsVO populateProfileDescReport(List<Object> viewList, ReportsVO reportVO) {
		StringBuilder data = new StringBuilder("");
		for (int outerIndex = 0; outerIndex < viewList.size(); outerIndex ++) {
			List outerIndividual = (ArrayList) viewList.get(outerIndex);
			for (int innerIndex = 0; innerIndex < outerIndividual.size(); innerIndex ++) {
				Object[] innerIndividual = (Object[]) outerIndividual.get(innerIndex);
				data.append((String) innerIndividual[0] + "#" + 
						(java.sql.Timestamp) innerIndividual[1] + "#" +
						(BigInteger) innerIndividual[2] + "#");
			}
			data.append("$^$");
		}
		reportVO.setReportProfileDescList(data.toString());
		return reportVO;
	}
	/**
	 * 
	 * @param sheet
	 * @param column
	 * @param row
	 * @param s
	 * @throws WriteException
	 * @throws RowsExceededException
	 */
	private void addLabelGroupProfile(WritableSheet sheet, int column, int row, String s) 
			throws WriteException {
		WritableFont contentFont = new WritableFont(WritableFont.ARIAL, 10);
		WritableCellFormat arialContent = new WritableCellFormat(contentFont);
		arialContent.setVerticalAlignment(VerticalAlignment.TOP);
		arialContent.setAlignment(Alignment.CENTRE);
		Label label;
		arialContent.setBorder(Border.ALL, BorderLineStyle.THIN);
	    label = new Label(column, row, s, arialContent);
	    arialContent.setWrap(true);
	    sheet.addCell(label);
	}
		
	/**
	 * Method populates the Facilitator / Individual Survey Report
	 * @param body
	 * @param request
	 * @return Json
	 */
	/*
	@RequestMapping(value = ACTION_GENERATE_FAC_IND_UI_SURVEY_RPRT, method = RequestMethod.POST)
	@ResponseBody()
	public ReportsVO generateSurveyReport(@RequestBody String body, HttpServletRequest request) {
		LOGGER.info(">>>>>>>>ReportsController.generateFacIndUIReport");
		ReportsVO reportVO = new ReportsVO();
		JsonNode node = CommonUtility.getNode(body);
		//String adminEmailId = node.get("adminEmailId").asText();
		//int usrType = node.get("usrType").asInt();	// 1 - Individual, 3 - Facilitator		
		String usrType = "I";
		try {
			//String resultString = iReportService.createSurveyReport(usrType);
			String resultString = iReportService.getSurveyListForExcel(usrType);
			StringBuilder sBuilder = new StringBuilder("");
			int srlCounter = 1;
			int mergingRwCnt = 0;
			int mergeIndexStart = 1;
			int mergeIndexEnd = 0;
			int rowSpan=0;
			StringTokenizer token = new StringTokenizer(resultString, "!~!^!");
			String firstArr=token.toString();
			while(token.hasMoreTokens()){	
				String dispObj = token.nextToken().toString();
				StringTokenizer innerToken=new StringTokenizer(dispObj,"$$$");
				System.out.println(innerToken.nextToken().toString());
			//	while(innerToken.hasMoreTokens()){
				String[] outerArr=dispObj.split("$$$");	
				System.out.println(outerArr[0].toString());
				rowSpan=outerArr.length;
				mergingRwCnt = mergingRwCnt + rowSpan;
				srlCounter=srlCounter+1;		
				for(int index=0;index<rowSpan;index++){
					String[] innerArr=outerArr[index].split("#");
					sBuilder.append((String)innerArr[0]+" ");
					sBuilder.append((String)innerArr[1]+" ");
					sBuilder.append((String)innerArr[2]+" ");
					sBuilder.append((String)innerArr[3]+" ");
					sBuilder.append((String)innerArr[4]+" ");
					sBuilder.append((String)innerArr[5]+" ");
					sBuilder.append((String)innerArr[6]+" ");
					sBuilder.append((String)innerArr[7]+" ");
					sBuilder.append((String)innerArr[8]+" ");
					sBuilder.append((String)innerArr[9]+" ");
				}
				mergeIndexEnd = mergeIndexEnd + rowSpan;
				mergeIndexStart = mergeIndexStart + rowSpan;
				}
		//	}
			
			reportVO.setReportList(sBuilder.toString());
			//reportVO.setTableHeaderString(this.createTableHeader("SP"));
			reportVO.setStatusCode(StatusConstants.STATUS_SUCCESS);
			reportVO.setStatusDesc("SUCCESS");
		} catch (JCTException e) {
			reportVO.setStatusCode(StatusConstants.STATUS_FAILURE);
			reportVO.setStatusDesc(e.getLocalizedMessage());
			LOGGER.error(e.getLocalizedMessage());
		}
		LOGGER.info("<<<<<<<<ReportController.generateFacIndUIReport");
		return reportVO;
	}*/	
	
	/**
	 * Method generates the excel sheet report for survey reports
	 * @param userType
	 * @param reportName
	 * @param request
	 * @param response
	 */
	@RequestMapping(value = "/generateSurveyExcel/{userType}/{reportName}", method = RequestMethod.GET)
	public void generateSurveyExcel(@PathVariable("userType") String userType,@PathVariable("reportName") String reportName, 
			HttpServletRequest request, HttpServletResponse response) {
		LOGGER.info(">>>>>> ReportsController.generateSurveyExcel");
		InputStream is = null;
		try {
			File file = new File("report.xls");
			WorkbookSettings wbSettings = new WorkbookSettings();
			wbSettings.setLocale(new Locale("en", "EN"));
			WritableWorkbook workbook = Workbook.createWorkbook(file, wbSettings);
			workbook.createSheet("Survey Report", 0);
			WritableSheet excelSheet = workbook.getSheet(0);
			
			int maxColSize = 0;
			List<SurveyQuestionDTO> answeredUserId = iReportService.fetchUserIdFromSurvey(Integer.parseInt(userType));	//	fetching userId's by userType 
			if (answeredUserId.size() != 0 ){				
				for(int j = 0 ; j < answeredUserId.size() ; j++){	//	calculating max col size
					if(maxColSize < answeredUserId.get(j).getSurveyQtnCount()) {
						maxColSize = (int) answeredUserId.get(j).getSurveyQtnCount();					
					}				
				}							
				for(int col = 0; col < 9+(maxColSize*3) ; col++){	//	creating -n- number of fields including sl no. (8 added for user common details fields)
					excelSheet.setColumnView(col, 26);
				}
				this.createSurveyXLReport(excelSheet, userType, answeredUserId, maxColSize);
			}
			this.createSurveyLabel(excelSheet,0,maxColSize);
			workbook.write();
			workbook.close();
			writeToStream(response.getOutputStream(), response, file, is, generateExcelFileName("survey_report"));            	
		}catch(Exception ex){
			LOGGER.error(ex.getLocalizedMessage());
		}
		LOGGER.info("<<<<<< ReportsController.generateSurveyExcel");
	}
	
	
	/**
	 * Method writes header in excel for survey report
	 * @param sheet
	 * @param distinction
	 * @throws WriteException
	 */
	private void createSurveyLabel(WritableSheet sheet, int distinction, int maxColSize) throws WriteException {
		LOGGER.info(">>>>>> ReportsController.createSurveyLabel");
		this.addCaption(sheet, 0, 0, "SL.No.");
		this.addCaption(sheet, 1, 0, "Name (Last, First)");
		this.addCaption(sheet, 2, 0, "User Name");
		this.addCaption(sheet, 3, 0, "User Profile Name");
		this.addCaption(sheet, 4, 0, "User Group Name");
		this.addCaption(sheet, 5, 0, "Customer Id");
		this.addCaption(sheet, 6, 0, "User Type");
		this.addCaption(sheet, 7, 0, "Created By E-Commerce/Admin");
		this.addCaption(sheet, 8, 0, "Expiry Date");
		for(int i = 8; i < (maxColSize*3)+8  ; i++){	// creating only ques|type|ans
			this.addCaption(sheet, i, 0, "Survey Question");
			this.addCaption(sheet, ++i, 0, "Survey Question Type");
			this.addCaption(sheet, ++i, 0, "Answer");		
		}		
		LOGGER.info("<<<<<< ReportsController.createSurveyLabel");
	}	
	/**
	 * Method creates excel report for Survey Report
	 * @param excelSheet
	 * @param userType
	 * @param answeredUserIdList
	 * @param maxColSize
	 * @throws RowsExceededException
	 * @throws WriteException
	 * @throws JCTException
	 */
	private void createSurveyXLReport(WritableSheet excelSheet, String userType,
			List<SurveyQuestionDTO> answeredUserIdList, int maxColSize) throws RowsExceededException, WriteException, JCTException {
		LOGGER.info(">>>>>> ReportsController.createSurveyXLReport");
		int srlCounter = 1;
		int rowIndex = 1;
		int colIndex = 0;
		if(answeredUserIdList.size() != 0){	
			List<surveyQtnReportVO> list = iReportService.getSurveyListForExcel(answeredUserIdList);			
			for(int i = 0 ; i < list.size() ; i++ ){
				colIndex = 0;				
				/********** Common Fields **************/				
				addLabel(excelSheet, colIndex, rowIndex, srlCounter+"");
				addLabel(excelSheet, ++colIndex, rowIndex, list.get(i).getName());	//First and Last Name
				addLabel(excelSheet, ++colIndex, rowIndex, list.get(i).getUserName());	//User Name
				addLabel(excelSheet, ++colIndex, rowIndex, list.get(i).getProfileName());	//Profile Name
				addLabel(excelSheet, ++colIndex, rowIndex, list.get(i).getGroupName());	//Group Name
				addLabel(excelSheet, ++colIndex, rowIndex, list.get(i).getCustId());	//Customer Id
				addLabel(excelSheet, ++colIndex, rowIndex, list.get(i).getUserType());	//User Type
				addLabel(excelSheet, ++colIndex, rowIndex, list.get(i).getCreatedBy());	//Created By
				addLabel(excelSheet, ++colIndex, rowIndex, list.get(i).getExpirationDate());	//Expiration Date
				/***************************************/
				
				/*********** Redundant Fields **********/				
				for(int j = 0 ; j < list.get(i).getQtnList().size() ; j++){
					//Survey Question| Type | Ans
					addLabel(excelSheet, ++colIndex, rowIndex, list.get(i).getQtnList().get(j).getQtn());
					addLabel(excelSheet, ++colIndex, rowIndex, list.get(i).getQtnList().get(j).getQtnType());
					addLabel(excelSheet, ++colIndex, rowIndex, list.get(i).getQtnList().get(j).getAnswer());
				}
				/***************************************/				
				rowIndex++;
				srlCounter++;	
			}			
		}
		LOGGER.info("<<<<<< ReportsController.createSurveyXLReport");
	}
	
	/**
	 * Method generates excel report for Misc
	 * @param reportName
	 * @param userType
	 * @param request
	 * @param response
	 */
	/*@RequestMapping(value = "/generatePaymentExcel/{userType}/{reportName}", method = RequestMethod.GET)
	public void generatePaymentExcel(  @PathVariable("userType") String userType, @PathVariable("reportName") String reportName,
			HttpServletRequest request, HttpServletResponse response) {
		LOGGER.info(">>>>>> ReportsController.generatePaymentExcel");
		InputStream is = null;		
		try {
			File file = new File("report.xls");
			WorkbookSettings wbSettings = new WorkbookSettings();
			wbSettings.setLocale(new Locale("en", "EN"));
			WritableWorkbook workbook = Workbook.createWorkbook(file, wbSettings);
			workbook.createSheet("Report", 0);
			WritableSheet excelSheet = workbook.getSheet(0);
			excelSheet.setColumnView(0, 8);
			excelSheet.setColumnView(1, 26);
			excelSheet.setColumnView(2, 26);
			excelSheet.setColumnView(3, 26);
			excelSheet.setColumnView(4, 26);
			excelSheet.setColumnView(5, 26);
			excelSheet.setColumnView(6, 26);
			excelSheet.setColumnView(7, 26);
			excelSheet.setColumnView(8, 26);
			excelSheet.setColumnView(9, 26);
			excelSheet.setColumnView(10, 26);
			excelSheet.setColumnView(11, 26);
			excelSheet.setColumnView(12, 26);
			//excelSheet.setColumnView(13, 26);
			createLabelForPayment(excelSheet,0);
			createReportForPayment(excelSheet, Integer.parseInt(userType));
			workbook.write();
			workbook.close();
			generateExcelFileName("payment_report");
			writeToStream(response.getOutputStream(), response, file, is, generateExcelFileName("payment_report"));
	        LOGGER.info("<<<<<< ReportsController.generatePaymentExcel");
		} catch (Exception ex) {
			LOGGER.error(ex.getLocalizedMessage());
		}	
	}	*/
	@RequestMapping(value = "/generatePaymentExcel/{userType}/{reportName}", method = RequestMethod.GET)
	 public void generatePaymentExcel(  @PathVariable("userType") String userType, @PathVariable("reportName") String reportName,
	   HttpServletRequest request, HttpServletResponse response) {
	  LOGGER.info(">>>>>> ReportsController.generatePaymentExcel");
	  InputStream is = null;  
	  try {
	   File file = new File("report.xls");
	   WorkbookSettings wbSettings = new WorkbookSettings();
	   wbSettings.setLocale(new Locale("en", "EN"));
	   WritableWorkbook workbook = Workbook.createWorkbook(file, wbSettings);
	   workbook.createSheet("Report", 0);
	   WritableSheet excelSheet = workbook.getSheet(0);
	   excelSheet.setColumnView(0, 8);
	   excelSheet.setColumnView(1, 26);// added for last + first name
	   excelSheet.setColumnView(2, 26);
	   excelSheet.setColumnView(3, 26);
	   excelSheet.setColumnView(4, 26);
	   excelSheet.setColumnView(5, 26);
	   excelSheet.setColumnView(6, 26);
	   excelSheet.setColumnView(7, 26);
	   excelSheet.setColumnView(8, 26);
	   excelSheet.setColumnView(9, 26);
	   excelSheet.setColumnView(10, 26);
	   excelSheet.setColumnView(11, 26);
	   excelSheet.setColumnView(12, 26);
	   excelSheet.setColumnView(13, 26);
	   //excelSheet.setColumnView(13, 26);
	   createLabelForPayment(excelSheet,0);
	   createReportForPayment(excelSheet, Integer.parseInt(userType));
	   workbook.write();
	   workbook.close();
	   generateExcelFileName("payment_report");
	   writeToStream(response.getOutputStream(), response, file, is, generateExcelFileName("payment_report"));
	         LOGGER.info("<<<<<< ReportsController.generatePaymentExcel");
	  } catch (Exception ex) {
	   LOGGER.error(ex.getLocalizedMessage());
	  } 
	 }
	/**
	 * Method creates misc report header
	 * @param sheet
	 * @param distinction
	 * @throws WriteException
	 */
	private void createLabelForPayment(WritableSheet sheet, int distinction) throws WriteException {
		LOGGER.info(">>>>>> ReportsController.createLabelMisc");
		this.addCaption(sheet, 0, 0, "Sl No.");
		this.addCaption(sheet, 1, 0, "Name (Last, First)");
		this.addCaption(sheet, 2, 0, "User Name");
		this.addCaption(sheet, 3, 0, "User Profile Name");
		this.addCaption(sheet, 4, 0, "user Group Name");
		this.addCaption(sheet, 5, 0, "Customer Id");
		this.addCaption(sheet, 6, 0, "user Type");
		this.addCaption(sheet, 7, 0, "Created By E-Commerce/ Admin");
		this.addCaption(sheet, 8, 0, "Payment Amout");
		this.addCaption(sheet, 9, 0, "Payment Mode");
		this.addCaption(sheet, 10, 0, "Payment Date");
		this.addCaption(sheet, 11, 0, "Payment Transaction Number");
		this.addCaption(sheet, 12, 0, "Check Number");
		//this.addCaption(sheet, 12, 0, "Subscription Type");
		this.addCaption(sheet, 13, 0, "Number Of User(s)");
		
		LOGGER.info("<<<<<< ReportsController.createLabelMisc");
	}
	
	/**
	 * Method puts the record in the excel sheet
	 * @param excelSheet
	 * @throws RowsExceededException
	 * @throws WriteException
	 * @throws JCTException
	 */
	/*private void createReportForPayment(WritableSheet excelSheet, int userType) throws RowsExceededException, WriteException, JCTException {
		LOGGER.info(">>>>>> ReportsController.createReportForPayment");
		String resultString = iReportService.createPaymentReport(userType);
		if(resultString.length() > 0){
			String[] split = resultString.split("#");
	        int counter = 1; 
	        int rowIndex = 1;
	        for(String tokens : split) {
	            String[] dispObj = tokens.split("~");
	        	addLabel(excelSheet, 0, rowIndex, counter +"");		//	Sl no.
				addLabel(excelSheet, 2, rowIndex, dispObj[0]);		//	user name
				addLabel(excelSheet, 3, rowIndex, dispObj[1]);		//	user profile name
				addLabel(excelSheet, 4, rowIndex, dispObj[2]);		//	user group name
				addLabel(excelSheet, 5, rowIndex, dispObj[3]);		//	cust id
				addLabel(excelSheet, 6, rowIndex, dispObj[4]);		//	user type	
				addLabel(excelSheet, 7, rowIndex, dispObj[5]);		//	created by
				addLabel(excelSheet, 8, rowIndex, dispObj[6]);		//	payment amount
				addLabel(excelSheet, 9, rowIndex, dispObj[7]);		//	payment type / mode
				addLabel(excelSheet, 10, rowIndex, dispObj[8]);		//	payment date
				addLabel(excelSheet, 11, rowIndex, dispObj[10]);	// transaction no.	*********
				addLabel(excelSheet, 12, rowIndex, dispObj[9]);	//	check number
				
				if (dispObj[4].equals("Individual")){				
					addLabel(excelSheet, 1, rowIndex, dispObj[12]);		// FirstName and LastName
				} else if(dispObj[4].equals("Facilitator")){
					addLabel(excelSheet, 1, rowIndex, dispObj[13]);		// FirstName and LastName			
				}
				
				//if (dispObj[4].equals("Individual")){				
					//addLabel(excelSheet, 12, rowIndex, "Subscribed");	//	subs type	indi
				//} else {
				//	addLabel(excelSheet, 12, rowIndex, dispObj[12]);	//	subs type	faci				
				//}
				addLabel(excelSheet, 13, rowIndex, dispObj[11]);	//	No of users		
				rowIndex++;
	            counter++;
	        }	
		}        
		LOGGER.info("<<<<<< ReportsController.createReportForPayment");
	}*/
	 private void createReportForPayment(WritableSheet excelSheet, int userType) throws RowsExceededException, WriteException, JCTException {
		  LOGGER.info(">>>>>> ReportsController.createReportForPayment################");
		  String resultString = iReportService.createPaymentReport(userType);
		  if(resultString.length() > 0){
		   String[] split = resultString.split("#");
		         int counter = 1; 
		         int rowIndex = 1;
		         for(String tokens : split) {
		             String[] dispObj = tokens.split("~");
		             String type = dispObj[4];
		             String name = dispObj[12];
		             LOGGER.info(">>>>>> ReportsController.createReportForPayment type ################" + type);
		             if (type.equals("Facilitator")){
		              name = dispObj[12];
		              LOGGER.info(">>>>>> ReportsController.createReportForPayment name before ################" + name);
		              if (name.startsWith("Subscribed")){
		               name = name.substring(10);
		              }
		              if(name.startsWith("Renewed")){
		               name = name.substring(7);               
		              }
		              LOGGER.info(">>>>>> ReportsController.createReportForPayment faci name after ################" + name);
		             }
		             
		             if (type.equals("Individual")){
		              name = dispObj[12];             
		              LOGGER.info(">>>>>> ReportsController.createReportForPayment name indi after ################" + name);
		             }
		             
		             
		          addLabel(excelSheet, 0, rowIndex, counter +"");  // Sl no.
		          addLabel(excelSheet, 1, rowIndex, name);     // last name + first name
		    addLabel(excelSheet, 2, rowIndex, dispObj[0]);  // user name
		    addLabel(excelSheet, 3, rowIndex, dispObj[1]);  // user profile name
		    addLabel(excelSheet, 4, rowIndex, dispObj[2]);  // user group name
		    addLabel(excelSheet, 5, rowIndex, dispObj[3]);  // cust id
		    addLabel(excelSheet, 6, rowIndex, dispObj[4]);  // user type 
		    addLabel(excelSheet, 7, rowIndex, dispObj[5]);  // created by
		    addLabel(excelSheet, 8, rowIndex, dispObj[6]);  // payment amount
		    addLabel(excelSheet, 9, rowIndex, dispObj[7]);  // payment type / mode
		    addLabel(excelSheet, 10, rowIndex, dispObj[8]);  // payment date
		    addLabel(excelSheet, 11, rowIndex, dispObj[10]); // transaction no. *********
		    addLabel(excelSheet, 12, rowIndex, dispObj[9]); // check number
		    
		    /*if (dispObj[4].equals("Individual")){    
		     addLabel(excelSheet, 12, rowIndex, "Subscribed"); // subs type indi
		    } else {
		     addLabel(excelSheet, 12, rowIndex, dispObj[12]); // subs type faci    
		    }*/
		    addLabel(excelSheet, 13, rowIndex, dispObj[11]); // No of users  
		    rowIndex++;
		             counter++;
		         } 
		  }        
		  LOGGER.info("<<<<<< ReportsController.createReportForPayment###########################");
		 }
	/**
	 * Method generates excel report compatible to IE-11
	 * @param out
	 * @param response
	 * @param file
	 * @param is
	 * @throws IOException
	 */
	private void writeToStream(ServletOutputStream out, HttpServletResponse response, 
			File file, InputStream is, String fileName) throws IOException {
		out = response.getOutputStream();
		response.setContentType("application/xls");
		response.setHeader("Content-Disposition",  "attachment;filename="
		   		+ fileName);
		response.setHeader("Pragma", "cache");
		response.setHeader("Cache-control", "private, max-age=0");
		byte[] buf = new byte[8192];
		is = new FileInputStream(file);
		//int c = 0;
		while ((is.read(buf, 0, buf.length)) > 0) {
			out.write(buf, 0, 8192);
		}
		is.close();
		out.flush();
		out.close();
	}
	
	private String generateExcelFileName(String reportName) {
		Date today = new Date();
		SimpleDateFormat dateFormat = new SimpleDateFormat("MMM-dd-yyyy");
        String date = dateFormat.format(today);
        return reportName+"_"+date+".xls";
	}
	
	@RequestMapping(value = "/downloadOnetExcel/{reportName}", method = RequestMethod.GET)
	public void downloadOnetExcel(@PathVariable("reportName") String reportName, HttpServletRequest request, HttpServletResponse response) {
		LOGGER.info(">>>>>> ReportsController.downloadOnetExcel");
		InputStream is = null;
		try {
			File file = new File("downloadOnetExcel.xls");
			WorkbookSettings wbSettings = new WorkbookSettings();
			wbSettings.setLocale(new Locale("en", "EN"));
			WritableWorkbook workbook = Workbook.createWorkbook(file, wbSettings);
			workbook.createSheet("Report", 0);
			WritableSheet excelSheet = workbook.getSheet(0);
			excelSheet.setColumnView(0, 16);
			excelSheet.setColumnView(1, 36);
			excelSheet.setColumnView(2, 1000);
			createOnetLabel(excelSheet,0);
			createOnetReport(excelSheet);
			workbook.write();
			workbook.close();
			writeToStream(response.getOutputStream(), response, file, is, generateExcelFileName("existing_onet_data"));
            LOGGER.info("<<<<<< ReportsController.generateExcel");
		}catch(Exception ex){
			LOGGER.error(ex.getLocalizedMessage());
			//ex.printStackTrace();
		}
		
	}
	private void createOnetLabel(WritableSheet sheet, int distinction) throws WriteException {
		LOGGER.info(">>>>>> ReportsController.createOnetLabel");
		addOnetCaption(sheet, 0, 0, "O*NET-SOC Code");
		addOnetCaption(sheet, 1, 0, "Title");
		addOnetCaption(sheet, 2, 0, "Description");
		LOGGER.info("<<<<<< ReportsController.createOnetLabel");
	}
	
	private void createOnetReport(WritableSheet excelSheet) throws RowsExceededException, WriteException, JCTException {
		LOGGER.info(">>>>>> ReportsController.createOnetReport");
		ContentConfigVO contentCongifVo = new ContentConfigVO();
		try {									
			List<OnetOccupationDTO> dataList = service.populateOnetDataListVo();
			int rowNumber = 1;
			for (int rowCount = 0; rowCount < dataList.size(); rowCount++) {
				OnetOccupationDTO vo = dataList.get(rowCount);
				this.addOnetLabel(excelSheet, 0, rowNumber, vo.getJctOnetOccupationCode());
				this.addOnetLabel(excelSheet, 1, rowNumber, vo.getJctOnetOccupationTitle());
				this.addOnetLabel(excelSheet, 2, rowNumber, vo.getJctOnetOccupationDesc());
				rowNumber = rowNumber + 1;
				
			}
		} catch (JCTException e) {
			contentCongifVo = new ContentConfigVO();
			contentCongifVo.setStatusCode(StatusConstants.STATUS_FAILURE);
			contentCongifVo.setStatusDesc(this.messageSource.getMessage("exception.dao.error", null, null));
			LOGGER.error(e.getLocalizedMessage());
		}
		LOGGER.info("<<<<<< ReportsController.createReport");
	}
	
	private void addOnetCaption(WritableSheet sheet, int column, int row, String stringToWrite) 
			throws WriteException {
		LOGGER.info(">>>>>> ReportsController.addCaption");
		WritableFont headerFont = new WritableFont(WritableFont.ARIAL, 10);
		WritableCellFormat arialHeader = new WritableCellFormat(headerFont);
		Label label;
	    label = new Label(column, row, stringToWrite, arialHeader);
	    sheet.addCell(label);
	    LOGGER.info("<<<<<< ReportsController.addCaption");
	}
	
	private void addOnetLabel(WritableSheet sheet, int column, int row, String stringToWrite) 
			throws WriteException {
		LOGGER.info(">>>>>> ReportsController.addOnetLabel");
		WritableFont contentFont = new WritableFont(WritableFont.ARIAL, 10);
		WritableCellFormat arialContent = new WritableCellFormat(contentFont);
		//arialContent.setVerticalAlignment(VerticalAlignment.CENTRE);
		//arialContent.setAlignment(Alignment.CENTRE);
		//arialContent.setBackground(Colour.GRAY_25);
		//arialContent.setBorder(Border.ALL, BorderLineStyle.THIN);
	    Label label;
	    label = new Label(column, row, stringToWrite, arialContent);
	    arialContent.setWrap(true);
	    sheet.addCell(label);
	    LOGGER.info("<<<<<< ReportsController.addOnetLabel");
	}
}