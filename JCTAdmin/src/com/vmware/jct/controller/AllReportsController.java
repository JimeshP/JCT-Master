package com.vmware.jct.controller;

import java.math.BigInteger;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.PathVariable;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestMethod;

import com.vmware.jct.common.utility.CommonUtility;
import com.vmware.jct.exception.JCTException;
import com.vmware.jct.service.IReportService;
/**
 * 
 * <p><b>Class name:</b> AllReportsController.java</p>
 * <p><b>Author:</b> InterraIT</p>
 * <p><b>Purpose:</b> This class acts as a controller for merged reports..
 * AllReportsController has following public methods:-
 * -generateExcelAllReport()
 * <p><b>Description:</b> This class is responsible for creating merged reports. </p>
 * <p><b>Copyrights:</b> 	All rights reserved by Interra IT and should only be used for its internal application development.</p>
 * <p><b>Created Date:</b> 13/May/2014</p>
 * <p><b>Revision History:</b>
 * 	<li></li>
 * </p>
 */
@Controller
@RequestMapping(value = "/allReports")
public class AllReportsController {
	
	@Autowired
	private IReportService iReportService;
	
	private static final Logger LOGGER = LoggerFactory.getLogger(AllReportsController.class);
	
	/**
	 * Method receives the request for generating merged report, generates
	 * and write it to the stream.
	 * @param checkedPreference
	 * @param reportName
	 * @param request
	 * @param response
	 */
	@RequestMapping(value = "/generateExcelAllReport/{checkedPreference}/{reportName}", method = RequestMethod.GET)
	public void generateExcelAllReport(@PathVariable("checkedPreference") String checkedPreference,
			@PathVariable("reportName") String reportName,
			HttpServletRequest request, HttpServletResponse response) {
		LOGGER.info(">>>> AllReportsController.generateExcelAllReport : checkedPreference : " + checkedPreference + ", reportName : " + reportName);
		String[] hiddenTokens = checkedPreference.split("~");
		try {
			Date today = new Date();
			SimpleDateFormat dateFormat = new SimpleDateFormat("MMM-dd-yyyy");
	        String date = dateFormat.format(today);
			//Create Workbook instance for xlsx/xls file input stream
            Workbook workbook = new XSSFWorkbook();
            Sheet sheet = workbook.createSheet("Merged Report - "+date);
            Row row = sheet.createRow(0);
            sheet.createFreezePane(0,1);
            XSSFCellStyle headerStyle = (XSSFCellStyle) workbook.createCellStyle();
      		Font headerFont = workbook.createFont();
      		headerStyle.setFillForegroundColor(IndexedColors.GREEN.getIndex());
      		headerFont.setColor(IndexedColors.WHITE.getIndex());
      		headerStyle.setFillPattern(CellStyle.SOLID_FOREGROUND);
      		headerStyle.setFont(headerFont);
      		Integer surveyCount = iReportService.getSurveyQuestionCount();
            generateHeader (row, headerStyle, surveyCount);
            
            //Individual Cell Style
            XSSFCellStyle bodyStyle = (XSSFCellStyle) workbook.createCellStyle();
      		Font bodyFont = workbook.createFont();
      		bodyFont.setColor(IndexedColors.BLACK.getIndex());
      		bodyStyle.setFont(bodyFont);
      		bodyStyle.setWrapText(true);
            
            //lets write the excel data to file now
            //FileOutputStream fos = new FileOutputStream("Merged_Report - "+date+".xlsx");
            int autoSizeCol = 593 + surveyCount + 1000; // 1000 is for backup
            for (int allColIndex = 0; allColIndex < autoSizeCol; allColIndex++) {
            	workbook.getSheetAt(0).autoSizeColumn(allColIndex);
            	workbook.getSheetAt(0).setHorizontallyCenter(true);
            	workbook.getSheetAt(0).setVerticallyCenter(true);
            }
            
            List<Object> infoTable = iReportService.getLoginInfoDetails(null);
            List<String> emailIdList = new ArrayList<String>();
            List<String> lastNameList = new ArrayList<String>();
            List<String> firstNameList = new ArrayList<String>();
            
            for ( int index = 0; index < infoTable.size(); index ++ ){
            	Object[] obj = (Object[]) infoTable.get(index);
            	emailIdList.add((String) obj[0]);
            	lastNameList.add((String) obj[2]); // added for last name
                firstNameList.add((String) obj[1]); // added for first name
            }
            
            int valueRowCounter = 1;
            for (int emailIdIndex = 0; emailIdIndex<emailIdList.size(); emailIdIndex++) {
            	String emailId = (String) emailIdList.get(emailIdIndex);
            	String lastName = (String) lastNameList.get(emailIdIndex);
                String firstName = (String) firstNameList.get(emailIdIndex);
            	Row valRow = sheet.createRow(valueRowCounter);
            	valRow.createCell(0).setCellValue(lastName+","+firstName);
            	valRow.createCell(1).setCellValue(emailId);
            	//valRow.getCell(0).setCellStyle(bodyStyle);
            	
            	List<Object> infoList = iReportService.getLoginInfoDetails(emailId);
            	
            	int dtls = 2;
            	for ( int index = 0; index < infoList.size(); index ++ ){
                	Object[] obj = (Object[]) infoList.get(index);
                	String functionGrp = (String) obj[0];
                	String jobLevel = (String) obj[1];               	
                	String userGroupName =iReportService.getUserGroupNameById((Integer)obj[2]);
                	
                	valRow.createCell(dtls).setCellValue(functionGrp);
                	dtls = dtls + 1;
                	
                	valRow.createCell(dtls).setCellValue(jobLevel);
                	dtls = dtls + 1;
                	
                	valRow.createCell(dtls).setCellValue(userGroupName);
                	dtls = dtls + 1;
                	
                	//get group creation date
                	java.sql.Timestamp creationDt = iReportService.getGroupCreationDate((Integer)obj[2]);
                	valRow.createCell(dtls).setCellValue(creationDt.toString());
                	dtls = dtls + 1;
                	
                	//Get first Start date and last end date
                	List<Object> dateList =  iReportService.getLoggedIndetails(emailId);
                	for ( int dateIndex = 0; dateIndex < dateList.size(); dateIndex ++ ){
                    	Object[] dateObj = (Object[]) dateList.get(dateIndex);
                    	java.sql.Timestamp minStDt = (java.sql.Timestamp) dateObj[0];
                    	java.sql.Timestamp maxEndDt = (java.sql.Timestamp) dateObj[1];
                    	java.math.BigInteger nosOfTimesLoggedIn = (java.math.BigInteger) dateObj[2];
                    	
                    	valRow.createCell(dtls).setCellValue(minStDt.toString());
                    	dtls = dtls + 1;
                    	
                    	if(maxEndDt == null){// handling null pointer exception
                            maxEndDt = minStDt; // in case user forgot to looged out, endtime = starttime
                           }            
                    	
                    	valRow.createCell(dtls).setCellValue(maxEndDt.toString());
                    	dtls = dtls + 1;
                    	
                    	long totalMilis = maxEndDt.getTime() - minStDt.getTime();
                    	String totalTimeInHHMMSS = CommonUtility.convertMillis(totalMilis);
                    	
                    	valRow.createCell(dtls).setCellValue(totalTimeInHHMMSS);
                    	dtls = dtls + 1;
                    	
                    	valRow.createCell(dtls).setCellValue(nosOfTimesLoggedIn.toString());
                    	dtls = dtls + 1;
                    	
                    	//Before Sketch
                    	Integer totalTimeSpentBS = null;
                    	try {
                    		totalTimeSpentBS = iReportService.getTotalTimeSpentOnBs(emailId);	
                    	} catch (Exception ex) {}
                    	
                    	if (null == totalTimeSpentBS) {
                    		valRow.createCell(dtls).setCellValue("");
                    	} else {
                    		valRow.createCell(dtls).setCellValue(CommonUtility.convertSeconds(totalTimeSpentBS));
                        }
                    	dtls = dtls + 1;
                    	BigInteger bsTasks = null;
                    	try{
                    		bsTasks = iReportService.getTotalBsTasks(emailId);
                    	} catch (Exception ex) {}
                    	 
                    	int bsmapping = dtls + 1;
                    	if(null == bsTasks) {
                    		valRow.createCell(bsmapping).setCellValue("");
                    	} else {
                    		valRow.createCell(bsmapping).setCellValue(bsTasks.toString());
                    	}
                    	
                    	List<Object> bsDescEnergy = iReportService.getBSDescriptionAndTimeEnergy(emailId);
                    	bsmapping = bsmapping + 6;
                    	for ( int bsDescindex = 0; bsDescindex < bsDescEnergy.size(); bsDescindex ++ ){
                        	Object[] bsDescObj = (Object[]) bsDescEnergy.get(bsDescindex);
                        	if(bsDescindex == 0) {
                        		valRow.createCell(17).setCellValue((String) bsDescObj[0]); 			// bs description
	                        	valRow.createCell(18).setCellValue(((Integer) 
	                        			bsDescObj[1]).toString()); 									// energy
                        	} else if (bsDescindex == 1){
                        		bsmapping = bsmapping + 12;
                        		valRow.createCell(bsmapping).setCellValue((String) bsDescObj[0]); 	// bs description
                        		bsmapping = bsmapping + 1;
                        		valRow.createCell(bsmapping).setCellValue(((Integer) 
                        				bsDescObj[1]).toString()); 									// energy
                        	} else {
                        		bsmapping = bsmapping + 11;
                        		valRow.createCell(bsmapping).setCellValue((String) bsDescObj[0]); 	// bs description
                        		bsmapping = bsmapping + 1;
                        		valRow.createCell(bsmapping).setCellValue(((Integer) 
                        				bsDescObj[1]).toString()); //energy
                        	}
                    	}
                    	
                    	//After Sketch
                    	Integer totalTimeSpentAS = null;
                    	try {
                    		totalTimeSpentAS = iReportService.getTotalTimeSpentOnAs(emailId);
                    	} catch (Exception ex) {}
                    	 
                    	if (null == totalTimeSpentAS) {
                    		valRow.createCell(dtls).setCellValue("NA");
                    	} else if (totalTimeSpentAS == 0 || null == totalTimeSpentAS) {
                    		valRow.createCell(dtls).setCellValue("NA");
                    	} else {
                    		valRow.createCell(dtls).setCellValue(CommonUtility.convertSeconds(totalTimeSpentAS));
                    	}
                    	dtls = dtls + 1;
                    	
                    	List<Object> asTasksRoleFrameCount = iReportService.getTotalAsTasks(emailId);
                    	int asmapping = dtls + 1;
                    	for (int count = 0; count < asTasksRoleFrameCount.size(); count ++ ) {
                    		Object[] objArr = (Object[]) asTasksRoleFrameCount.get(count);
                    		int asTaskCount = ((java.math.BigInteger) objArr[0]).intValue();
                    		if (asTaskCount == 0) {
                    			valRow.createCell(asmapping).setCellValue("NA");
                    		} else {
                    			valRow.createCell(asmapping).setCellValue(asTaskCount+"");
                    		}
                    	}
                    	
                    	asmapping = asmapping + 1;
                    	List<Object> roleFrameList = iReportService.getRoleFrame(emailId);
                    	if (roleFrameList.size() == 0) {
                    		valRow.createCell(asmapping).setCellValue("NA");
                    		valRow.getCell(asmapping).setCellStyle(bodyStyle);
                    	} else {
                    		valRow.createCell(asmapping).setCellValue(roleFrameList.size()+"");
                    		valRow.getCell(asmapping).setCellStyle(bodyStyle);
                    	}
                    	int roleFrameId = 0;
                    	Map<String, Integer> roleMapping = new HashMap<String, Integer>();
                    	for (int count = 0; count < roleFrameList.size(); count ++ ) {
                    		asmapping = asmapping + 1;
                    		valRow.createCell(asmapping).setCellValue(((String) roleFrameList.get(count)+""));
                    		valRow.getCell(asmapping).setCellStyle(bodyStyle);
                    		roleFrameId = roleFrameId + 1;
                    		roleMapping.put(((String) roleFrameList.get(count)+""), roleFrameId);
                    	}
                    	
                    	if (hiddenTokens[1].equals("C")) {											// after sketch selected
                    		populateASDetails (valRow, emailId, roleMapping);
                        	populateAsToBs (valRow, emailId);
                        	populateASTaskLocation (valRow, emailId);
                        	populateMappings (valRow, emailId, roleMapping);
                        }
                    	if (hiddenTokens[2].equals("C")) { 											// reflection question not selected
                    		populateReflectionQuestions (valRow, emailId, bodyStyle);
                        }
                    	if (hiddenTokens[3].equals("C")) { 											// action plans not selected
                    		populateActionPlan(valRow, emailId, bodyStyle);
                        }
                    	populateSurveyQuestion(valRow, emailId, bodyStyle, surveyCount);
                    	
                	}
                }
            	//increment the row for next record
            	valueRowCounter = valueRowCounter + 1;
            	workbook.getSheetAt(0).autoSizeColumn(emailIdIndex);
            }
            if (hiddenTokens[0].equals("N")) { 														// before sketch not selected
            	hideBeforeSketch (sheet);
            }
            if (hiddenTokens[1].equals("N")) { 														// after sketch not selected
            	hideAfterSketch (sheet);
            } 
            if (hiddenTokens[2].equals("N")) { 														// reflection question not selected
            	hideReflectionQuestions (sheet);
            }
            if (hiddenTokens[3].equals("N")) { 														// action plans not selected
            	hideActionPlans (sheet);
            }
            if (hiddenTokens[4].equals("N")) { 														// action plans not selected
            	hideSurveyQuestion (sheet, surveyCount);
            }
            
            // Justin Change - No job level / function group
            sheet.setColumnHidden(2, true);
            sheet.setColumnHidden(3, true);
            
            response.setContentType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
            response.setHeader("Content-Disposition", "attachment; filename=Merged_Report - "+date+".xlsx");
            workbook.write(response.getOutputStream());
            response.getOutputStream().close();
		} catch (Exception ex) {
			LOGGER.error(ex.getLocalizedMessage());
			ex.printStackTrace();
		}
		LOGGER.info("<<<< AllReportsController.generateExcelAllReport");
	}
	
	private void hideSurveyQuestion(Sheet sheet, Integer surveyCount) {
		int endIndex = 594 + surveyCount + surveyCount;
		LOGGER.info(">>>> AllReportsController.hideSurveyQuestion");
		for (int index = 594; index < endIndex; index ++ ) {
			sheet.setColumnHidden(index, true);
		}
		LOGGER.info("<<<< AllReportsController.hideSurveyQuestion");
	}
	/**
	 * Method populates the task location of After sketch
	 * @param valRow
	 * @param emailId
	 * @throws JCTException 
	 */
	private void populateASTaskLocation(Row valRow, String emailId) throws JCTException {
		LOGGER.info(">>>> AllReportsController.populateASTaskLocation");
		List<Object> taskLocationList = iReportService.populateASTaskLocation(emailId);
		int taskLocationPlot = 26;
		for (int taskLocationIndex = 0; taskLocationIndex < taskLocationList.size(); taskLocationIndex++) {
			String taskLocation = (String) taskLocationList.get(taskLocationIndex);
			valRow.createCell(taskLocationPlot).setCellValue(taskLocation);
			taskLocationPlot = taskLocationPlot + 12;
		}	
		LOGGER.info("<<<< AllReportsController.populateASTaskLocation");
	}
	/**
	 * Method populates action plan
	 * @param valRow
	 * @param emailId
	 * @param bodyStyle 
	 * @throws JCTException 
	 */
	private void populateActionPlan(Row valRow, String emailId, XSSFCellStyle bodyStyle) throws JCTException {
		LOGGER.info(">>>> AllReportsController.populateActionPlan");
		int ansPlotIndex = 0;
		List<Object> actionPlanMainQtnList = iReportService.populateActionPlan(emailId);
		if (actionPlanMainQtnList.size() > 0) {
			int mainQtnPlotIndex = 492;
			for (int mainQtnIndex = 0; mainQtnIndex < actionPlanMainQtnList.size(); mainQtnIndex ++){
				String mainQuestion = (String) actionPlanMainQtnList.get(mainQtnIndex);
				valRow.createCell(mainQtnPlotIndex).setCellValue(mainQuestion); 					// question
				valRow.getCell(mainQtnPlotIndex).setCellStyle(bodyStyle);
				mainQtnPlotIndex = mainQtnPlotIndex + 17;
				//get the sub question and answers of the main question
				List<Object> subQtnAnsList = iReportService.populateSubQtnActionPlan(mainQuestion, emailId);
				int subQtnPlottingIndex = 493;
				int loopCtr = 0;
				for (int subQtnIndex = 0; subQtnIndex < subQtnAnsList.size(); subQtnIndex++) {
					String subQtns = (String) subQtnAnsList.get(subQtnIndex);
					loopCtr = loopCtr + 1;
					if (loopCtr == 5) {
						subQtnPlottingIndex = 510;
					} else if (loopCtr == 9) {
						subQtnPlottingIndex = 527;
					} else if (loopCtr == 13) {
						subQtnPlottingIndex = 544;
					} else if (loopCtr == 17) {
						subQtnPlottingIndex = 561;
					} else if (loopCtr == 21) {
						subQtnPlottingIndex = 578;
					}
					valRow.createCell(subQtnPlottingIndex).setCellValue(subQtns); 					// sub question
					valRow.getCell(subQtnPlottingIndex).setCellStyle(bodyStyle);
					subQtnPlottingIndex = subQtnPlottingIndex + 4;
					//get the answers of sub questions
					List<Object> ansList = iReportService.getActionPlanAnswers(mainQuestion, subQtns, emailId);
					if (ansPlotIndex == 0) {
						ansPlotIndex = 493;
					} else if (ansPlotIndex == 493) {
						ansPlotIndex = 497;
					} else if (ansPlotIndex == 497) {
						ansPlotIndex = 501;
					} else if (ansPlotIndex == 501) {
						ansPlotIndex = 505;
					} else if (ansPlotIndex == 505) {
						ansPlotIndex = 510;
					} else if (ansPlotIndex == 510) {
						ansPlotIndex = 514;
					} else if (ansPlotIndex == 514) {
						ansPlotIndex = 518;
					} else if (ansPlotIndex == 518) {
						ansPlotIndex = 522;
					} else if (ansPlotIndex == 522) {
						ansPlotIndex = 527;
					} else if (ansPlotIndex == 527) {
						ansPlotIndex = 531;
					} else if (ansPlotIndex == 531) {
						ansPlotIndex = 535;
					} else if (ansPlotIndex == 535) {
						ansPlotIndex = 539;
					} else if (ansPlotIndex == 539) {
						ansPlotIndex = 544;
					} else if (ansPlotIndex == 544) {
						ansPlotIndex = 548;
					} else if (ansPlotIndex == 548) {
						ansPlotIndex = 552;
					} else if (ansPlotIndex == 552) {
						ansPlotIndex = 556;
					} else if (ansPlotIndex == 556) {
						ansPlotIndex = 561;
					} else if (ansPlotIndex == 561) {
						ansPlotIndex = 565;
					} else if (ansPlotIndex == 565) {
						ansPlotIndex = 569;
					} else if (ansPlotIndex == 569) {
						ansPlotIndex = 573;
					} else if (ansPlotIndex == 573) {
						ansPlotIndex = 578;
					} else if (ansPlotIndex == 578) {
						ansPlotIndex = 582;
					} else if (ansPlotIndex == 582) {
						ansPlotIndex = 586;
					} else if (ansPlotIndex == 586) {
						ansPlotIndex = 590;
					}
					if (ansList.size() > 0) {
						valRow.createCell(ansPlotIndex+1).setCellValue((String) ansList.get(0)); 			// answer
						valRow.getCell(ansPlotIndex+1).setCellStyle(bodyStyle);
						valRow.createCell(ansPlotIndex+2).setCellValue((String) ansList.get(1)); 		// answer
						valRow.getCell(ansPlotIndex+2).setCellStyle(bodyStyle);
						valRow.createCell(ansPlotIndex+3).setCellValue((String) ansList.get(2)); 		// answer
						valRow.getCell(ansPlotIndex+3).setCellStyle(bodyStyle);
					}
				}
			}
		}
		LOGGER.info("<<<< AllReportsController.populateActionPlan");
	}
	/**
	 * Method hides the after sketch panel
	 * @param sheet
	 */
	private void hideAfterSketch (Sheet sheet) {
		LOGGER.info(">>>> AllReportsController.hideAfterSketch");
		sheet.setColumnHidden(10, true);
		hideCols(13, 17, sheet);
		hideCols(20, 29, sheet);
		hideCols(32, 41, sheet);
		hideCols(44, 53, sheet);
		hideCols(56, 65, sheet);
		hideCols(68, 77, sheet);
		hideCols(79, 88, sheet);
		hideCols(92, 101, sheet);
		hideCols(104, 113, sheet);
		hideCols(116, 125, sheet);
		hideCols(128, 137, sheet);
		hideCols(140, 149, sheet);
		hideCols(152, 161, sheet);
		hideCols(164, 173, sheet);
		hideCols(176, 185, sheet);
		hideCols(188, 197, sheet);
		hideCols(200, 209, sheet);
		hideCols(212, 221, sheet);
		hideCols(224, 233, sheet);
		hideCols(236, 245, sheet);
		hideCols(248, 437, sheet);
		LOGGER.info("<<<< AllReportsController.hideAfterSketch");
	}
	private void hideCols (int startIndex, int endIndex, Sheet sheet) {
		LOGGER.info(">>>> AllReportsController.hideCols");
		for (int index = startIndex; index <= endIndex; index ++ ) {
			sheet.setColumnHidden(index, true);
		}
		LOGGER.info("<<<< AllReportsController.hideCols");
	}
	/**
	 * Method populates reflection questions
	 * @param valRow
	 * @param emailId
	 * @param bodyStyle 
	 * @throws JCTException 
	 */
	private void populateReflectionQuestions(Row valRow, String emailId, XSSFCellStyle bodyStyle) throws JCTException {
		LOGGER.info(">>>> AllReportsController.populateReflectionQuestions");
		int ansPlotIndex = 440;
		int subQtnPlottingIndex = 439;
		int loopCtr = 0;
		List<Object> questionnaireMainQtnList = iReportService.getReflectionQuestions(emailId);
		if (questionnaireMainQtnList.size() > 0) {
			int mainQtnPlotIndex = 438;
			for (int mainQtnIndex = 0; mainQtnIndex < questionnaireMainQtnList.size(); mainQtnIndex ++){
				String mainQuestion = (String) questionnaireMainQtnList.get(mainQtnIndex);
				valRow.createCell(mainQtnPlotIndex).setCellValue(mainQuestion); 					// question
				valRow.getCell(mainQtnPlotIndex).setCellStyle(bodyStyle);
				mainQtnPlotIndex = mainQtnPlotIndex + 9;
				
				//get the sub question and answers of the main question
				List<Object> subQtnAnsList = iReportService.populateSubQtnQuestionnaire(mainQuestion, emailId);
				for (int subQtnIndex = 0; subQtnIndex < subQtnAnsList.size(); subQtnIndex++) {
					String subQtns = (String) subQtnAnsList.get(subQtnIndex);
					String subQtnVal = subQtns;
					if (subQtns.trim().equals("NA")) {
						subQtnVal = "";
					}
					if (loopCtr == 1) {
						subQtnPlottingIndex = 448;
					} else if (loopCtr == 2) {
						subQtnPlottingIndex = 457;
					} else if (loopCtr == 3) {
						subQtnPlottingIndex = 466;
					} else if (loopCtr == 4) {
						subQtnPlottingIndex = 475;
					} else if (loopCtr == 5) {
						subQtnPlottingIndex = 484;
					}
					valRow.createCell(subQtnPlottingIndex).setCellValue(subQtnVal); 					// sub question
					valRow.getCell(subQtnPlottingIndex).setCellStyle(bodyStyle);
					subQtnPlottingIndex = subQtnPlottingIndex + 2;
					
					//get the answers of sub questions
					List<Object> ansList = iReportService.getQuestionnaireAnswers(mainQuestion, subQtns, emailId);
					if (loopCtr == 1) {
						ansPlotIndex = 449;
					} else if (loopCtr == 2) {
						ansPlotIndex = 458;
					} else if (loopCtr == 3) {
						ansPlotIndex = 467;
					} else if (loopCtr == 4) {
						ansPlotIndex = 476;
					} else if (loopCtr == 5) {
						ansPlotIndex = 485;
					}
					
					valRow.createCell(ansPlotIndex).setCellValue((String) ansList.get(0)); 			// answer
					valRow.getCell(ansPlotIndex).setCellStyle(bodyStyle);
					ansPlotIndex = ansPlotIndex + 2;
				}
				loopCtr = loopCtr + 1;
			}
		}
		LOGGER.info("<<<< AllReportsController.populateReflectionQuestions");
	}
	/**
	 * Method populates the excel header
	 * @param row
	 * @param headerStyle 
	 */
	private void generateHeader(Row row, XSSFCellStyle headerStyle, Integer surveyCount) {
		LOGGER.info(">>>> AllReportsController.generateHeader");
		
		String headerlist[]={"Name (Last,First)","Email Id","Function Group","Job Level","User Group Profile","User Group Profile Creation Date","First Start date/time","Last End date/time",
				"Total Time (HH:MM:SS","Number of times logged","Time Spent on Before Sketch","Time Spent on After Diagram","Before - Total Tasks","After - Total Tasks",
				"Total Role Frames","Role Frame 1                        ","Role Frame 2                        ","Role Frame 3                        ",
				"Task 1 - Before Description","Task 1 - Before Time / Energy","Task 1 - After Description_Main","Task 1 - After Description_Person","Task 1 - After Time / Energy",
				"Task 1 - Percent from Before to After","Task 1 - Status from Before to After","Task 1 - Description Edit from Before to After","Task 1 - After Location",
				"Task 1 - Role Frame 1","Task 1 - Role Frame 2","Task 1 - Role Frame 3","Task 2 - Before Description","Task 2 - Before Time / Energy",
				"Task 2 - After Description_Main","Task 2 - After Description_Person","Task 2 - After Time / Energy","Task 2 - Percent from Before to After",
				"Task 2 - Status from Before to After","Task 2 - Description Edit from Before to After","Task 2 - After Location","Task 2 - Role Frame 1","Task 2 - Role Frame 2",
				"Task 2 - Role Frame 3","Task 3 - Before Description","Task 3 - Before Time / Energy","Task 3 - After Description_Main","Task 3 - After Description_Person",
				"Task 3 - After Time / Energy","Task 3 - Percent from Before to After","Task 3 - Status from Before to After","Task 3 - Description Edit from Before to After",
				"Task 3 - After Location","Task 3 - Role Frame 1","Task 3 - Role Frame 2","Task 3 - Role Frame 3","Task 4 - Before Description","Task 4 - Before Time / Energy",
				"Task 4 - After Description_Main","Task 4 - After Description_Person","Task 4 - After Time / Energy","Task 4 - Percent from Before to After","Task 4 - Status from Before to After",
				"Task 4 - Description Edit from Before to After","Task 4 - After Location","Task 4 - Role Frame 1","Task 4 - Role Frame 2","Task 4 - Role Frame 3",
				"Task 5 - Before Description","Task 5 - Before Time / Energy","Task 5 - After Description_Main","Task 5 - After Description_Person","Task 5 - After Time / Energy",
				"Task 5 - Percent from Before to After","Task 5 - Status from Before to After","Task 5 - Description Edit from Before to After","Task 5 - After Location",
				"Task 5 - Role Frame 1","Task 5 - Role Frame 2","Task 5 - Role Frame 3","Task 6 - Before Description","Task 6 - Before Time / Energy",
				"Task 6 - After Description_Main","Task 6 - After Description_Person","Task 6 - After Time / Energy","Task 6 - Percent from Before to After",
				"Task 6 - Status from Before to After","Task 6 - Description Edit from Before to After","Task 6 - After Location","Task 6 - Role Frame 1","Task 6 - Role Frame 2",
				"Task 6 - Role Frame 3","Task 7 - Before Description","Task 7 - Before Time / Energy","Task 7 - After Description_Main","Task 7 - After Description_Person",
				"Task 7 - After Time / Energy","Task 7 - Percent from Before to After","Task 7 - Status from Before to After","Task 7 - Description Edit from Before to After",
				"Task 7 - After Location","Task 7 - Role Frame 1","Task 7 - Role Frame 2","Task 7 - Role Frame 3","Task 8 - Before Description","Task 8 - Before Time / Energy",
				"Task 8 - After Description_Main","Task 8 - After Description_Person","Task 8 - After Time / Energy","Task 8 - Percent from Before to After",
				"Task 8 - Status from Before to After","Task 8 - Description Edit from Before to After","Task 8 - After Location","Task 8 - Role Frame 1",
				"Task 8 - Role Frame 2","Task 8 - Role Frame 3","Task 9 - Before Description","Task 9 - Before Time / Energy","Task 9 - After Description_Main",
				"Task 9 - After Description_Person","Task 9 - After Time / Energy","Task 9 - Percent from Before to After","Task 9 - Status from Before to After",
				"Task 9 - Description Edit from Before to After","Task 9 - After Location","Task 9 - Role Frame 1","Task 9 - Role Frame 2","Task 9 - Role Frame 3",
				"Task 10 - Before Description","Task 10 - Before Time / Energy","Task 10 - After Description_Main","Task 10 - After Description_Person",
				"Task 10 - After Time / Energy","Task 10 - Percent from Before to After","Task 10 - Status from Before to After","Task 10 - Description Edit from Before to After",
				"Task 10 - After Location","Task 10 - Role Frame 1","Task 10 - Role Frame 2","Task 10 - Role Frame 3","Task 11 - Before Description",
				"Task 11 - Before Time / Energy","Task 11 - After Description_Main","Task 11 - After Description_Person","Task 11 - After Time / Energy",
				"Task 11 - Percent from Before to After","Task 11 - Status from Before to After","Task 11 - Description Edit from Before to After","Task 11 - After Location",
				"Task 11 - Role Frame 1","Task 11 - Role Frame 2","Task 11 - Role Frame 3","Task 12 - Before Description","Task 12 - Before Time / Energy",
				"Task 12 - After Description_Main","Task 12 - After Description_Person","Task 12 - After Time / Energy","Task 12 - Percent from Before to After",
				"Task 12 - Status from Before to After","Task 12 - Description Edit from Before to After","Task 12 - After Location","Task 12 - Role Frame 1",
				"Task 12 - Role Frame 2","Task 12 - Role Frame 3","Task 13 - Before Description","Task 13 - Before Time / Energy","Task 13 - After Description_Main",
				"Task 13 - After Description_Person","Task 13 - After Time / Energy","Task 13 - Percent from Before to After","Task 13 - Status from Before to After",
				"Task 13 - Description Edit from Before to After","Task 13 - After Location","Task 13 - Role Frame 1","Task 13 - Role Frame 2","Task 13 - Role Frame 3",
				"Task 14 - Before Description","Task 14 - Before Time / Energy","Task 14 - After Description_Main","Task 14 - After Description_Person",
				"Task 14 - After Time / Energy","Task 14 - Percent from Before to After","Task 14 - Status from Before to After","Task 14 - Description Edit from Before to After",
				"Task 14 - After Location","Task 14 - Role Frame 1","Task 14 - Role Frame 2","Task 14 - Role Frame 3","Task 15 - Before Description","Task 15 - Before Time / Energy",
				"Task 15 - After Description_Main","Task 15 - After Description_Person","Task 15 - After Time / Energy","Task 15 - Percent from Before to After",
				"Task 15 - Status from Before to After","Task 15 - Description Edit from Before to After","Task 15 - After Location","Task 15 - Role Frame 1",
				"Task 15 - Role Frame 2","Task 15 - Role Frame 3","Task 16 - Before Description","Task 16 - Before Time / Energy","Task 16 - After Description_Main",
				"Task 16 - After Description_Person","Task 16 - After Time / Energy","Task 16 - Percent from Before to After","Task 16 - Status from Before to After",
				"Task 16 - Description Edit from Before to After","Task 16 - After Location","Task 16 - Role Frame 1","Task 16 - Role Frame 2","Task 16 - Role Frame 3",
				"Task 17 - Before Description","Task 17 - Before Time / Energy","Task 17 - After Description_Main","Task 17 - After Description_Person",
				"Task 17 - After Time / Energy","Task 17 - Percent from Before to After","Task 17 - Status from Before to After","Task 17 - Description Edit from Before to After",
				"Task 17 - After Location","Task 17 - Role Frame 1","Task 17 - Role Frame 2","Task 17 - Role Frame 3","Task 18 - Before Description","Task 18 - Before Time / Energy",
				"Task 18 - After Description_Main","Task 18 - After Description_Person","Task 18 - After Time / Energy","Task 18 - Percent from Before to After",
				"Task 18 - Status from Before to After","Task 18 - Description Edit from Before to After","Task 18 - After Location","Task 18 - Role Frame 1",
				"Task 18 - Role Frame 2","Task 18 - Role Frame 3","Task 19 - Before Description","Task 19 - Before Time / Energy","Task 19 - After Description_Main",
				"Task 19 - After Description_Person","Task 19 - After Time / Energy","Task 19 - Percent from Before to After","Task 19 - Status from Before to After",
				"Task 19 - Description Edit from Before to After","Task 19 - After Location","Task 19 - Role Frame 1","Task 19 - Role Frame 2","Task 19 - Role Frame 3",
				"Task 20 - Before Description","Task 20 - Before Time / Energy","Task 20 - After Description_Main","Task 20 - After Description_Person","Task 20 - After Time / Energy",
				"Task 20 - Percent from Before to After","Task 20 - Status from Before to After","Task 20 - Description Edit from Before to After","Task 20 - After Location",
				"Task 20 - Role Frame 1","Task 20 - Role Frame 2","Task 20 - Role Frame 3","Strength 1a - Description","Strength 1a - Location","Strength 1a - Role Frame 1",
				"Strength 1a - Role Frame 2","Strength 1a - Role Frame 3","Strength 1b - Description","Strength 1b - Location","Strength 1b - Role Frame 1","Strength 1b - Role Frame 2",
				"Strength 1b - Role Frame 3","Strength 1c - Description","Strength 1c - Location","Strength 1c - Role Frame 1","Strength 1c - Role Frame 2","Strength 1c - Role Frame 3",
				"Strength 2a - Description","Strength 2a - Location","Strength 2a - Role Frame 1","Strength 2a - Role Frame 2","Strength 2a - Role Frame 3","Strength 2b - Description",
				"Strength 2b - Location","Strength 2b - Role Frame 1","Strength 2b - Role Frame 2","Strength 2b - Role Frame 3","Strength 2c - Description","Strength 2c - Location",
				"Strength 2c - Role Frame 1","Strength 2c - Role Frame 2","Strength 2c - Role Frame 3","Strength 3a - Description","Strength 3a - Location","Strength 3a - Role Frame 1",
				"Strength 3a - Role Frame 2","Strength 3a - Role Frame 3","Strength 3b - Description","Strength 3b - Location","Strength 3b - Role Frame 1","Strength 3b - Role Frame 2",
				"Strength 3b - Role Frame 3","Strength 3c - Description","Strength 3c - Location","Strength 3c - Role Frame 1","Strength 3c - Role Frame 2","Strength 3c - Role Frame 3",
				"Strength 4a - Description","Strength 4a - Location","Strength 4a - Role Frame 1","Strength 4a - Role Frame 2","Strength 4a - Role Frame 3","Strength 4b - Description",
				"Strength 4b - Location","Strength 4b - Role Frame 1","Strength 4b - Role Frame 2","Strength 4b - Role Frame 3","Strength 4c - Description","Strength 4c - Location",
				"Strength 4c - Role Frame 1","Strength 4c - Role Frame 2","Strength 4c - Role Frame 3","Value 1a - Description","Value 1a - Location",
				"Value 1a - Role Frame 1","Value 1a - Role Frame 2","Value 1a - Role Frame 3","Value 1b - Description","Value 1b - Location","Value 1b - Role Frame 1",
				"Value 1b - Role Frame 2","Value 1b - Role Frame 3","Value 1c - Description","Value 1c - Location","Value 1c - Role Frame 1","Value 1c - Role Frame 2",
				"Value 1c - Role Frame 3","Value 2a - Description","Value 2a - Location","Value 2a - Role Frame 1","Value 2a - Role Frame 2","Value 2a - Role Frame 3",
				"Value 2b - Description","Value 2b - Location","Value 2b - Role Frame 1","Value 2b - Role Frame 2","Value 2b - Role Frame 3","Value 2c - Description",
				"Value 2c - Location","Value 2c - Role Frame 1","Value 2c - Role Frame 2","Value 2c - Role Frame 3","Value 3a - Description","Value 3a - Location",
				"Value 3a - Role Frame 1","Value 3a - Role Frame 2","Value 3a - Role Frame 3","Value 3b - Description","Value 3b - Location","Value 3b - Role Frame 1",
				"Value 3b - Role Frame 2","Value 3b - Role Frame 3","Value 3c - Description","Value 3c - Location","Value 3c - Role Frame 1","Value 3c - Role Frame 2",
				"Value 3c - Role Frame 3","Value 4a - Description","Value 4a - Location","Value 4a - Role Frame 1","Value 4a - Role Frame 2","Value 4a - Role Frame 3",
				"Value 4b - Description","Value 4b - Location","Value 4b - Role Frame 1","Value 4b - Role Frame 2","Value 4b - Role Frame 3","Value 4c - Description",
				"Value 4c - Location","Value 4c - Role Frame 1","Value 4c - Role Frame 2","Value 4c - Role Frame 3","Passion 1a - Description","Passion 1a - Location",
				"Passion 1a - Role Frame 1","Passion 1a - Role Frame 2","Passion 1a - Role Frame 3","Passion 1b - Description","Passion 1b - Location","Passion 1b - Role Frame 1",
				"Passion 1b - Role Frame 2","Passion 1b - Role Frame 3","Passion 1c - Description","Passion 1c - Location","Passion 1c - Role Frame 1","Passion 1c - Role Frame 2",
				"Passion 1c - Role Frame 3","Passion 2a - Description","Passion 2a - Location","Passion 2a - Role Frame 1","Passion 2a - Role Frame 2","Passion 2a - Role Frame 3",
				"Passion 2b - Description","Passion 2b - Location","Passion 2b - Role Frame 1","Passion 2b - Role Frame 2","Passion 2b - Role Frame 3","Passion 2c - Description",
				"Passion 2c - Location","Passion 2c - Role Frame 1","Passion 2c - Role Frame 2","Passion 2c - Role Frame 3","Passion 3a - Description","Passion 3a - Location",
				"Passion 3a - Role Frame 1","Passion 3a - Role Frame 2","Passion 3a - Role Frame 3","Passion 3b - Description","Passion 3b - Location","Passion 3b - Role Frame 1",
				"Passion 3b - Role Frame 2","Passion 3b - Role Frame 3","Passion 3c - Description","Passion 3c - Location","Passion 3c - Role Frame 1","Passion 3c - Role Frame 2",
				"Passion 3c - Role Frame 3","Passion 4a - Description","Passion 4a - Location","Passion 4a - Role Frame 1","Passion 4a - Role Frame 2","Passion 4a - Role Frame 3",
				"Passion 4b - Description","Passion 4b - Location","Passion 4b - Role Frame 1","Passion 4b - Role Frame 2","Passion 4b - Role Frame 3","Passion 4c - Description",
				"Passion 4c - Location","Passion 4c - Role Frame 1","Passion 4c - Role Frame 2","Passion 4c - Role Frame 3","Before Reflection - Question 1                                                                                      ",
				"Before Reflection - Sub Question 1a                                                                                      ","Before Reflection - Answer 1a                                                                                      ",
				"Before Reflection - Sub Question 1b                                                                                      ","Before Reflection - Answer 1b                                                                                      ",
				"Before Reflection - Sub Question 1c                                                                                      ","Before Reflection - Answer 1c                                                                                      ",
				"Before Reflection - Sub Question 1d                                                                                      ","Before Reflection - Answer 1d                                                                                      ",
				"Before Reflection - Question 2                                                                                      ","Before Reflection - Sub Question 2a                                                                                      ",
				"Before Reflection - Answer 2a                                                                                      ","Before Reflection - Sub Question 2b                                                                                      ",
				"Before Reflection - Answer 2b                                                                                      ","Before Reflection - Sub Question 2c                                                                                      ",
				"Before Reflection - Answer 2c                                                                                      ","Before Reflection - Sub Question 2d                                                                                      ",
				"Before Reflection - Answer 2d                                                                                      ","Before Reflection - Question 3                                                                                      ",
				"Before Reflection - Sub Question 3a                                                                                      ","Before Reflection - Answer 3a                                                                                      ",
				"Before Reflection - Sub Question 3b                                                                                      ","Before Reflection - Answer 3b                                                                                      ",
				"Before Reflection - Sub Question 3c                                                                                      ","Before Reflection - Answer 3c                                                                                      ",
				"Before Reflection - Sub Question 3d                                                                                      ","Before Reflection - Answer 3d                                                                                      ",
				"Before Reflection - Question 4                                                                                      ","Before Reflection - Sub Question 4a                                                                                      ",
				"Before Reflection - Answer 4a                                                                                      ","Before Reflection - Sub Question 4b                                                                                      ",
				"Before Reflection - Answer 4b                                                                                      ","Before Reflection - Sub Question 4c                                                                                      ",
				"Before Reflection - Answer 4c                                                                                      ","Before Reflection - Sub Question 4d                                                                                      ",
				"Before Reflection - Answer 4d                                                                                      ","Before Reflection - Question 5                                                                                      ",
				"Before Reflection - Sub Question 5a                                                                                      ","Before Reflection - Answer 5a                                                                                      ",
				"Before Reflection - Sub Question 5b                                                                                      ","Before Reflection - Answer 5b                                                                                      ",
				"Before Reflection - Sub Question 5c                                                                                      ","Before Reflection - Answer 5c                                                                                      ",
				"Before Reflection - Sub Question 5d                                                                                      ","Before Reflection - Answer 5d                                                                                      ",
				"Before Reflection - Question 6                                                                                      ","Before Reflection - Sub Question 6a                                                                                      ",
				"Before Reflection - Answer 6a                                                                                      ","Before Reflection - Sub Question 6b                                                                                      ",
				"Before Reflection - Answer 6b                                                                                      ","Before Reflection - Sub Question 6c                                                                                      ",
				"Before Reflection - Answer 6c                                                                                      ","Before Reflection - Sub Question 6d                                                                                      ",
				"Before Reflection - Answer 6d                                                                                      ","After Action - Question 1                                                                                      ",
				"After Action - Sub Question 1a                                                                ","After Action - Answer 1aa                                                                                      ","After Action - Answer 1ab                                                                                      ",
				"After Action - Answer 1ac                                                                                      ","After Action - Sub Question 1b                                                                ","After Action - Answer 1ba                                                                                      ",
				"After Action - Answer 1bb                                                                                      ","After Action - Answer 1bc                                                                                      ","After Action - Sub Question 1c                                                                ",
				"After Action - Answer 1ca                                                                                      ","After Action - Answer 1cb                                                                                      ","After Action - Answer 1cc                                                                                      ",
				"After Action - Sub Question 1d                                                                ","After Action - Answer 1da                                                                                      ","After Action - Answer 1db                                                                                      ",
				"After Action - Answer 1dc                                                                                      ","After Action - Question 2                                                                                      ","After Action - Sub Question 2a                                                                ",
				"After Action - Answer 2aa                                                                                      ","After Action - Answer 2ab                                                                                      ","After Action - Answer 2ac                                                                                      ",
				"After Action - Sub Question 2b                                                                ","After Action - Answer 2ba                                                                                      ","After Action - Answer 2bb                                                                                      ",
				"After Action - Answer 2bc                                                                                      ","After Action - Sub Question 2c                                                                ","After Action - Answer 2ca                                                                                      ",
				"After Action - Answer 2cb                                                                                      ","After Action - Answer 2cc                                                                                      ","After Action - Sub Question 2d                                                                ",
				"After Action - Answer 2da                                                                                      ","After Action - Answer 2db                                                                                      ","After Action - Answer 2dc                                                                                      ",
				"After Action - Question 3                                                                                      ","After Action - Sub Question 3a                                                                ","After Action - Answer 3aa                                                                                      ",
				"After Action - Answer 3ab                                                                                      ","After Action - Answer 3ac                                                                                      ","After Action - Sub Question 3b                                                                ",
				"After Action - Answer 3ba                                                                                      ","After Action - Answer 3bb                                                                                      ","After Action - Answer 3bc                                                                                      ",
				"After Action - Sub Question 3c                                                                ","After Action - Answer 3ca                                                                                      ","After Action - Answer 3cb                                                                                      ",
				"After Action - Answer 3cc                                                                                      ","After Action - Sub Question 3d                                                                ","After Action - Answer 3da                                                                                      ",
				"After Action - Answer 3db                                                                                      ","After Action - Answer 3dc                                                                                      ","After Action - Question 4                                                                                      ",
				"After Action - Sub Question 4a                                                                ","After Action - Answer 4aa                                                                                      ","After Action - Answer 4ab                                                                                      ",
				"After Action - Answer 4ac                                                                                      ","After Action - Sub Question 4b                                                                ","After Action - Answer 4ba                                                                                      ",
				"After Action - Answer 4bb                                                                                      ","After Action - Answer 4bc                                                                                      ","After Action - Sub Question 4c                                                                ",
				"After Action - Answer 4ca                                                                                      ","After Action - Answer 4cb                                                                                      ","After Action - Answer 4cc                                                                                      ",
				"After Action - Sub Question 4d                                                                ","After Action - Answer 4da                                                                                      ","After Action - Answer 4db                                                                                      ",
				"After Action - Answer 4dc                                                                                      ","After Action - Question 5                                                                                      ","After Action - Sub Question 5a                                                                ",
				"After Action - Answer 5aa                                                                                      ","After Action - Answer 5ab                                                                                      ","After Action - Answer 5ac                                                                                      ",
				"After Action - Sub Question 5b                                                                ","After Action - Answer 5ba                                                                                      ","After Action - Answer 5bb                                                                                      ",
				"After Action - Answer 5bc                                                                                      ","After Action - Sub Question 5c                                                                ","After Action - Answer 5ca                                                                                      ",
				"After Action - Answer 5cb                                                                                      ","After Action - Answer 5cc                                                                                      ","After Action - Sub Question 5d                                                                ",
				"After Action - Answer 5da                                                                                      ","After Action - Answer 5db                                                                                      ","After Action - Answer 5dc                                                                                      ",
				"After Action - Question 6                                                                                      ","After Action - Sub Question 6a                                                                ","After Action - Answer 6aa                                                                                      ",
				"After Action - Answer 6ab                                                                                      ","After Action - Answer 6ac                                                                                      ","After Action - Sub Question 6b                                                                ",
				"After Action - Answer 6ba                                                                                      ","After Action - Answer 6bb                                                                                      ","After Action - Answer 6bc                                                                                      ",
				"After Action - Sub Question 6c                                                                ","After Action - Answer 6ca                                                                                      ","After Action - Answer 6cb                                                                                      ",
				"After Action - Answer 6cc                                                                                      ","After Action - Sub Question 6d                                                                ","After Action - Answer 6da                                                                                      ",
				"After Action - Answer 6db                                                                                      ","After Action - Answer 6dc                                                                                      "};
		for(int i=0;i<headerlist.length;i++)
		{
			row.createCell(i).setCellValue(headerlist[i]);
		}
         // Get maximum number of survey questions answered by individual user
        int surveyColNameCounter = 1;
		int totalColumnCount = headerlist.length-1;
		for (int index=0; index<surveyCount; index++) {
			totalColumnCount = totalColumnCount + 1;
			row.createCell(totalColumnCount).setCellValue("Survey Question: "+surveyColNameCounter+"                                                                       ");
			totalColumnCount = totalColumnCount + 1;
			row.createCell(totalColumnCount).setCellValue("Survey Answer: "+surveyColNameCounter+"                                                                       ");
			surveyColNameCounter = surveyColNameCounter + 1;
		}
		//Apply the header style
         for (int headerIndex = 0; headerIndex <= totalColumnCount; headerIndex++) {
        	 row.getCell(headerIndex).setCellStyle(headerStyle);
         }
         /*//Apply the header style
         for (int headerIndex = 0; headerIndex <= 592; headerIndex++) {
        	 row.getCell(headerIndex).setCellStyle(headerStyle);
         }*/
         LOGGER.info("<<<< AllReportsController.generateHeader");
	}
	/**
	 * Method populates the After sketch details.
	 * @param valRow
	 * @param emailId
	 * @param roleMapping 
	 * @throws JCTException
	 */
	private void populateASDetails(Row valRow, String emailId, Map<String, Integer> roleMapping) throws JCTException {
		LOGGER.info(">>>> AllReportsController.populateASDetails");
		List<Object> asDescAddTEList = iReportService.getASDescriptionMainPersonAndTimeEnergy(emailId);
		int check = 44;
    	for (int asCount = 0; asCount < asDescAddTEList.size(); asCount++ ) {
    		Object[] individualObj = (Object[]) asDescAddTEList.get(asCount);
    		if (asCount == 0) {
    			//Get different roles for the task descriptions
    			String taskDesc = (String) individualObj[0];
    			List<String> roleList = iReportService.getRolesForTasks(emailId, taskDesc);
    			
    			valRow.createCell(20).setCellValue((String) individualObj[0]); 						//as main description
    			valRow.createCell(21).setCellValue((String) individualObj[1]); 						//as user description
            	valRow.createCell(22).setCellValue(((Integer) individualObj[2]).toString()); 		//energy
            	
            	if (roleList.size() == 1) {
            		valRow.createCell(27).setCellValue(roleList.get(0).toString()); 
            	} else if (roleList.size() == 2) {
            		valRow.createCell(27).setCellValue(roleList.get(0).toString());
            		valRow.createCell(28).setCellValue(roleList.get(1).toString()); 
            	} else {
            		valRow.createCell(27).setCellValue(roleList.get(0).toString());
            		valRow.createCell(28).setCellValue(roleList.get(1).toString());
            		valRow.createCell(29).setCellValue(roleList.get(2).toString()); 
            	}
            	
            	/*String role = (String) individualObj[3];
            	int roleMapIndx = roleMapping.get(role);
            	if (roleMapIndx == 1) {
            		valRow.createCell(26).setCellValue(role); 
            	} else if (roleMapIndx == 2) {
            		valRow.createCell(27).setCellValue(role); 
            	} else if (roleMapIndx == 3) {
            		valRow.createCell(28).setCellValue(role); 
            	}*/            	
    		} else if (asCount == 1) {
    			//Get different roles for the task descriptions
    			String taskDesc = (String) individualObj[0];
    			List<String> roleList = iReportService.getRolesForTasks(emailId, taskDesc);
    			
    			valRow.createCell(32).setCellValue((String) individualObj[0]); 						//as main description
    			valRow.createCell(33).setCellValue((String) individualObj[1]); 						//as user description
            	valRow.createCell(34).setCellValue(((Integer) individualObj[2]).toString()); 		//energy
            	/*String role = (String) individualObj[3];
            	int roleMapIndx = roleMapping.get(role);
            	if (roleMapIndx == 1) {
            		valRow.createCell(38).setCellValue(role); 
            	} else if (roleMapIndx == 2) {
            		valRow.createCell(39).setCellValue(role); 
            	} else if (roleMapIndx == 3) {
            		valRow.createCell(40).setCellValue(role); 
            	}*/
            	if (roleList.size() == 1) {
            		valRow.createCell(39).setCellValue(roleList.get(0).toString()); 
            	} else if (roleList.size() == 2) {
            		valRow.createCell(39).setCellValue(roleList.get(0).toString());
            		valRow.createCell(40).setCellValue(roleList.get(1).toString()); 
            	} else {
            		valRow.createCell(39).setCellValue(roleList.get(0).toString());
            		valRow.createCell(40).setCellValue(roleList.get(1).toString());
            		valRow.createCell(41).setCellValue(roleList.get(2).toString()); 
            	}
    		} else {
    			//Get different roles for the task descriptions
    			String taskDesc = (String) individualObj[0];
    			List<String> roleList = iReportService.getRolesForTasks(emailId, taskDesc);
    			
    			valRow.createCell(check).setCellValue((String) individualObj[0]); 					//as main description
    			check = check + 1;
    			valRow.createCell(check).setCellValue((String) individualObj[1]); 					//as user description
    			check = check + 1;
            	valRow.createCell(check).setCellValue(((Integer) individualObj[2]).toString()); 	//energy
            	check = check + 10;
            	int rlind = check - 10;
            	/*String role = (String) individualObj[3];
            	int roleMapIndx = roleMapping.get(role);
            	if (roleMapIndx == 1) {
            		valRow.createCell(rlind+5).setCellValue(role); 
            	} else if (roleMapIndx == 2) {
            		valRow.createCell(rlind+6).setCellValue(role); 
            	} else if (roleMapIndx == 3) {
            		valRow.createCell(rlind+7).setCellValue(role); 
            	}*/
            	if (roleList.size() == 1) {
            		valRow.createCell(rlind+5).setCellValue(roleList.get(0).toString()); 
            	} else if (roleList.size() == 2) {
            		valRow.createCell(rlind+5).setCellValue(roleList.get(0).toString());
            		valRow.createCell(rlind+6).setCellValue(roleList.get(1).toString()); 
            	} else {
            		valRow.createCell(rlind+5).setCellValue(roleList.get(0).toString());
            		valRow.createCell(rlind+6).setCellValue(roleList.get(1).toString());
            		valRow.createCell(rlind+7).setCellValue(roleList.get(2).toString()); 
            	}
            	
    		}
    	}
    	LOGGER.info("<<<< AllReportsController.populateASDetails");
	}
	/**
	 * Method populates the before sketch to after sketch changes
	 * @param valRow
	 * @param emailId
	 * @throws JCTException
	 */
	private void populateAsToBs(Row valRow, String emailId) throws JCTException {
		LOGGER.info(">>>> AllReportsController.populateAsToBs");
		List<Object> bsToAsList = iReportService.getBsToAs(emailId);
    	int bsPlot = 47;
    	for (int bsToAsIndex = 0; bsToAsIndex < bsToAsList.size(); bsToAsIndex++ ) {
    		Object[] bstoasObj = (Object[]) bsToAsList.get(bsToAsIndex);
    		if (bsToAsIndex == 0) {
    			valRow.createCell(23).setCellValue((String) bstoasObj[0]); 							// diff in energy
    			valRow.createCell(24).setCellValue((String) bstoasObj[1]); 							// diff status
            	valRow.createCell(25).setCellValue((String) bstoasObj[2]); 							// edited text desc
    		} else if (bsToAsIndex == 1) {
    			valRow.createCell(35).setCellValue((String) bstoasObj[0]); 							// diff in energy
    			valRow.createCell(36).setCellValue((String) bstoasObj[1]); 							// diff status
            	valRow.createCell(37).setCellValue((String) bstoasObj[2]); 							// edited text desc
    		} else {
    			valRow.createCell(bsPlot).setCellValue((String) bstoasObj[0]); 						// diff in energy
    			bsPlot = bsPlot + 1;
    			valRow.createCell(bsPlot).setCellValue((String) bstoasObj[1]); 						// diff status
    			bsPlot = bsPlot + 1;
            	valRow.createCell(bsPlot).setCellValue((String) bstoasObj[2]); 						// edited text desc
            	bsPlot = bsPlot + 10;
    		}
    	}
    	LOGGER.info("<<<< AllReportsController.populateAsToBs");
	}
	/**
	 * Method populate only the strengh value and passion elements.
	 * @param valRow
	 * @param emailId
	 * @param roleMapping
	 * @throws JCTException
	 */
	private void populateMappings(Row valRow, String emailId,
			Map<String, Integer> roleMapping) throws JCTException {
		LOGGER.info(">>>> AllReportsController.populateMappings");
		List<Object> strValPassList = iReportService.getStrValPassItems(emailId);
		int strengthStartIndex = 258;
		int strengthLocationIndex = 259;
		int strengthLastPlottedIndex = 0;
		String strengthMemory = "";

		int valueStartIndex = 318;
		int valueLocationIndex = 319;
		int valueLastPlottedIndex = 0;
		String valueMemory = "";

		int passionStartIndex = 378;
		int passionLocationIndex = 379;
		int passionLastPlottedIndex = 0;
		String passionMemory = "";
		// jct_as_element_code, jct_as_element_desc, jct_as_role_desc, jct_as_position
		for (int elementIndex = 0; elementIndex < strValPassList.size(); elementIndex++) {
			Object[] elementRow = (Object[]) strValPassList.get(elementIndex);
			String elementCode = (String) elementRow[0];											// jct_as_element_code
			if (elementCode.equals("STR")) {
				String strength = (String) elementRow[1];											// jct_as_element_desc
				if (strengthMemory.equals("") || strengthMemory.equals(strength)) {
					String roleDesc = (String) elementRow[2];										// jct_as_role_desc
					String elementLocation = (String) elementRow[3];	
					int roleMapIndex = roleMapping.get(roleDesc);// jct_as_position
					if (strengthMemory.equals("")) {
						valRow.createCell(strengthStartIndex).setCellValue((String) elementRow[1]);
						strengthStartIndex = strengthStartIndex + 2;
						if (roleMapIndex == 1) {
							valRow.createCell(strengthStartIndex).setCellValue(roleDesc); 				// role desc
							valRow.createCell(strengthLocationIndex).setCellValue(elementLocation); 	// strength location
						} else if (roleMapIndex == 2) {
							valRow.createCell(strengthStartIndex + 1).setCellValue(roleDesc); 			// role desc
							valRow.createCell(strengthLocationIndex).setCellValue(elementLocation); 	// strength location
						} else {
							valRow.createCell(strengthStartIndex + 2).setCellValue(roleDesc); 			// role desc
							valRow.createCell(strengthLocationIndex).setCellValue(elementLocation); 	// strength location
						}
						strengthStartIndex = strengthStartIndex + 3;
						strengthLastPlottedIndex = strengthStartIndex - 3;
						strengthMemory = strength;
						strengthLocationIndex = strengthLocationIndex + 5;
					} else {
						strengthStartIndex = strengthStartIndex - 3;
						strengthLastPlottedIndex = strengthStartIndex + 3;
						strengthMemory = strength;
						strengthLocationIndex = strengthLocationIndex - 5;
						if (roleMapIndex == 1) {
							valRow.createCell(strengthStartIndex).setCellValue(roleDesc); 				// role desc
						} else if (roleMapIndex == 2) {
							valRow.createCell(strengthStartIndex + 1).setCellValue(roleDesc); 			// role desc
						} else {
							valRow.createCell(strengthStartIndex + 2).setCellValue(roleDesc); 			// role desc
						}
						strengthStartIndex = strengthStartIndex + 3;
						strengthLastPlottedIndex = strengthStartIndex - 3;
						strengthMemory = strength;
						strengthLocationIndex = strengthLocationIndex + 5;
					}
					
				} else {
					//For location
					if (strengthLastPlottedIndex == 260
							|| strengthLastPlottedIndex == 261
							|| strengthLastPlottedIndex == 262
							|| strengthLastPlottedIndex == 265
							|| strengthLastPlottedIndex == 266
							|| strengthLastPlottedIndex == 267
							|| strengthLastPlottedIndex == 270
							|| strengthLastPlottedIndex == 271
							|| strengthLastPlottedIndex == 272) {
						strengthStartIndex = 273;
					} else if (strengthLastPlottedIndex == 275
							|| strengthLastPlottedIndex == 276
							|| strengthLastPlottedIndex == 277
							|| strengthLastPlottedIndex == 280
							|| strengthLastPlottedIndex == 281
							|| strengthLastPlottedIndex == 282
							|| strengthLastPlottedIndex == 285
							|| strengthLastPlottedIndex == 286
							|| strengthLastPlottedIndex == 287) {
						strengthStartIndex = 288;
					} else if (strengthLastPlottedIndex == 290
							|| strengthLastPlottedIndex == 291
							|| strengthLastPlottedIndex == 292
							|| strengthLastPlottedIndex == 295
							|| strengthLastPlottedIndex == 296
							|| strengthLastPlottedIndex == 297
							|| strengthLastPlottedIndex == 300
							|| strengthLastPlottedIndex == 301
							|| strengthLastPlottedIndex == 302) {
						strengthStartIndex = 303;
					}

					valRow.createCell(strengthStartIndex).setCellValue((String) elementRow[1]); 	// strength desc
					strengthStartIndex = strengthStartIndex + 2;
					String roleDesc = (String) elementRow[2];
					String elementLocation = (String) elementRow[3];
					int roleMapIndex = roleMapping.get(roleDesc);
					if (roleMapIndex == 1) {
						valRow.createCell(strengthStartIndex).setCellValue(roleDesc); 				// role desc
						valRow.createCell(strengthStartIndex - 1).setCellValue(elementLocation); 	// strength location
					} else if (roleMapIndex == 2) {
						valRow.createCell(strengthStartIndex + 1).setCellValue(roleDesc); 			// role desc
						valRow.createCell(strengthStartIndex - 1).setCellValue(elementLocation); 	// strength location
					} else {
						valRow.createCell(strengthStartIndex + 2).setCellValue(roleDesc); 			// role desc
						valRow.createCell(strengthStartIndex - 1).setCellValue(elementLocation); 	// strength location
					}
					
					strengthStartIndex = strengthStartIndex + 3;
					strengthLastPlottedIndex = strengthStartIndex - 3;
					strengthMemory = strength;
					strengthLocationIndex = strengthLocationIndex + 5;
				}

				/*strengthStartIndex = strengthStartIndex + 3;
				strengthLastPlottedIndex = strengthStartIndex - 3;
				strengthMemory = strength;
				strengthLocationIndex = strengthLocationIndex + 5;*/
			} else if (elementCode.equals("VAL")) {
				String value = (String) elementRow[1];
				if (valueMemory.equals("") || valueMemory.equals(value)) {
					String roleDesc = (String) elementRow[2];
					String elementLocation = (String) elementRow[3];
					int roleMapIndex = roleMapping.get(roleDesc);
					if (valueMemory.equals("")) {
						valRow.createCell(valueStartIndex).setCellValue((String) elementRow[1]); 		// value desc
						valueStartIndex = valueStartIndex + 2;
						if (roleMapIndex == 1) {
							valRow.createCell(valueStartIndex).setCellValue(roleDesc); 					// role desc
							valRow.createCell(valueLocationIndex).setCellValue(elementLocation); 		// value location
						} else if (roleMapIndex == 2) {
							valRow.createCell(valueStartIndex + 1).setCellValue(roleDesc); 				// role desc
							valRow.createCell(valueLocationIndex).setCellValue(elementLocation); 		// value location
						} else {
							valRow.createCell(valueStartIndex + 2).setCellValue(roleDesc); 				// role desc
							valRow.createCell(valueLocationIndex).setCellValue(elementLocation); 		// value location
						}
						valueStartIndex = valueStartIndex + 3;
						valueLastPlottedIndex = valueStartIndex - 3;
						valueMemory = value;
						valueLocationIndex = valueLocationIndex + 5;
					} else {
						valueStartIndex = valueStartIndex - 3;
						valueLastPlottedIndex = valueStartIndex + 3;
						valueMemory = value;
						valueLocationIndex = valueLocationIndex - 5;
						if (roleMapIndex == 1) {
							valRow.createCell(valueStartIndex).setCellValue(roleDesc); 				// role desc
						} else if (roleMapIndex == 2) {
							valRow.createCell(valueStartIndex + 1).setCellValue(roleDesc); 			// role desc
						} else {
							valRow.createCell(valueStartIndex + 2).setCellValue(roleDesc); 			// role desc
						}
						valueStartIndex = valueStartIndex + 3;
						valueLastPlottedIndex = valueStartIndex - 3;
						valueMemory = value;
						valueLocationIndex = valueLocationIndex + 5;
					}
				} else {
					if (valueLastPlottedIndex == 320
							|| valueLastPlottedIndex == 321
							|| valueLastPlottedIndex == 322
							|| valueLastPlottedIndex == 325
							|| valueLastPlottedIndex == 326
							|| valueLastPlottedIndex == 327
							|| valueLastPlottedIndex == 330
							|| valueLastPlottedIndex == 331
							|| valueLastPlottedIndex == 332) { // plotted
						valueStartIndex = 333;
					} else if (valueLastPlottedIndex == 335
							|| valueLastPlottedIndex == 336
							|| valueLastPlottedIndex == 337
							|| valueLastPlottedIndex == 340
							|| valueLastPlottedIndex == 341
							|| valueLastPlottedIndex == 342
							|| valueLastPlottedIndex == 345
							|| valueLastPlottedIndex == 346
							|| valueLastPlottedIndex == 347) {
						valueStartIndex = 348;
					} else if (valueLastPlottedIndex == 350
							|| valueLastPlottedIndex == 351
							|| valueLastPlottedIndex == 352
							|| valueLastPlottedIndex == 355
							|| valueLastPlottedIndex == 356
							|| valueLastPlottedIndex == 357
							|| valueLastPlottedIndex == 360
							|| valueLastPlottedIndex == 361
							|| valueLastPlottedIndex == 362) {
						valueStartIndex = 363;
					}

					valRow.createCell(valueStartIndex).setCellValue(
							(String) elementRow[1]); // value desc
					valueStartIndex = valueStartIndex + 2;
					String roleDesc = (String) elementRow[2];
					String elementLocation = (String) elementRow[3];
					int roleMapIndex = roleMapping.get(roleDesc);
					if (roleMapIndex == 1) {
						valRow.createCell(valueStartIndex).setCellValue(roleDesc); 					// role desc
						valRow.createCell(valueStartIndex - 1).setCellValue(elementLocation); 		// value location
					} else if (roleMapIndex == 2) {
						valRow.createCell(valueStartIndex + 1).setCellValue(roleDesc); 				// role desc
						valRow.createCell(valueStartIndex - 1).setCellValue(elementLocation); 		// value location
					} else {
						valRow.createCell(valueStartIndex + 2).setCellValue(roleDesc); 				// role desc
						valRow.createCell(valueStartIndex - 1).setCellValue(elementLocation); 		// value location
					}
					valueStartIndex = valueStartIndex + 3;
					valueLastPlottedIndex = valueStartIndex - 3;
					valueMemory = value;
					valueLocationIndex = valueLocationIndex + 5;
				}
				/*valueStartIndex = valueStartIndex + 3;
				valueLastPlottedIndex = valueStartIndex - 3;
				valueMemory = value;
				valueLocationIndex = valueLocationIndex + 5;*/
			} else {
				String passion = (String) elementRow[1];
				if (passionMemory.equals("") || passionMemory.equals(passion)) {
					String roleDesc = (String) elementRow[2];
					String elementLocation = (String) elementRow[3];
					int roleMapIndex = roleMapping.get(roleDesc);
					if (passionMemory.equals("")) {
						valRow.createCell(passionStartIndex).setCellValue((String) elementRow[1]); 		// passion desc
						passionStartIndex = passionStartIndex + 2;
						if (roleMapIndex == 1) {
							valRow.createCell(passionStartIndex).setCellValue(roleDesc); 				// role desc
							valRow.createCell(passionLocationIndex).setCellValue(elementLocation); 		// strength location
						} else if (roleMapIndex == 2) {
							valRow.createCell(passionStartIndex + 1).setCellValue(roleDesc); 			// role desc
							valRow.createCell(passionLocationIndex).setCellValue(elementLocation); 		// strength location
						} else {
							valRow.createCell(passionStartIndex + 2).setCellValue(roleDesc); 			// role desc
							valRow.createCell(passionLocationIndex).setCellValue(elementLocation); 		// strength location
						}
						passionStartIndex = passionStartIndex + 3;
						passionLastPlottedIndex = passionStartIndex - 3;
						passionMemory = passion;
						passionLocationIndex = passionLocationIndex + 5;
					} else {
						passionStartIndex = passionStartIndex - 3;
						passionLastPlottedIndex = passionStartIndex + 3;
						passionMemory = passion;
						passionLocationIndex = passionLocationIndex - 5;
						if (roleMapIndex == 1) {
							valRow.createCell(passionStartIndex).setCellValue(roleDesc); 				// role desc
						} else if (roleMapIndex == 2) {
							valRow.createCell(passionStartIndex + 1).setCellValue(roleDesc); 			// role desc
						} else {
							valRow.createCell(passionStartIndex + 2).setCellValue(roleDesc); 			// role desc
						}
						passionStartIndex = passionStartIndex + 3;
						passionLastPlottedIndex = passionStartIndex - 3;
						passionMemory = passion;
						passionLocationIndex = passionLocationIndex + 5;
					}
				} else {
					if (passionLastPlottedIndex == 380
							|| passionLastPlottedIndex == 381
							|| passionLastPlottedIndex == 382
							|| passionLastPlottedIndex == 385
							|| passionLastPlottedIndex == 386
							|| passionLastPlottedIndex == 387
							|| passionLastPlottedIndex == 390
							|| passionLastPlottedIndex == 391
							|| passionLastPlottedIndex == 392) {
						passionStartIndex = 393;
					} else if (passionLastPlottedIndex == 395
							|| passionLastPlottedIndex == 396
							|| passionLastPlottedIndex == 397
							|| passionLastPlottedIndex == 400
							|| passionLastPlottedIndex == 401
							|| passionLastPlottedIndex == 402
							|| passionLastPlottedIndex == 405
							|| passionLastPlottedIndex == 406
							|| passionLastPlottedIndex == 407) {
						passionStartIndex = 408;
					} else if (passionLastPlottedIndex == 410
							|| passionLastPlottedIndex == 411
							|| passionLastPlottedIndex == 412
							|| passionLastPlottedIndex == 415
							|| passionLastPlottedIndex == 416
							|| passionLastPlottedIndex == 417
							|| passionLastPlottedIndex == 420
							|| passionLastPlottedIndex == 421
							|| passionLastPlottedIndex == 422) {
						passionStartIndex = 423;
					}
					valRow.createCell(passionStartIndex).setCellValue((String) elementRow[1]); 		// passion desc
					passionStartIndex = passionStartIndex + 2;
					String roleDesc = (String) elementRow[2];
					String elementLocation = (String) elementRow[3];
					int roleMapIndex = roleMapping.get(roleDesc);
					if (roleMapIndex == 1) {
						valRow.createCell(passionStartIndex).setCellValue(roleDesc); 				// role desc
						valRow.createCell(passionStartIndex - 1).setCellValue(elementLocation); 	// value location
					} else if (roleMapIndex == 2) {
						valRow.createCell(passionStartIndex + 1).setCellValue(roleDesc); 			// role desc
						valRow.createCell(passionStartIndex - 1).setCellValue(elementLocation); 	// value location
					} else {
						valRow.createCell(passionStartIndex + 2).setCellValue(roleDesc); 			// role desc
						valRow.createCell(passionStartIndex - 1).setCellValue(elementLocation); 	// value location
					}
					
					passionStartIndex = passionStartIndex + 3;
					passionLastPlottedIndex = passionStartIndex - 3;
					passionMemory = passion;
					passionLocationIndex = passionLocationIndex + 5;
				}
				/*passionStartIndex = passionStartIndex + 3;
				passionLastPlottedIndex = passionStartIndex - 3;
				passionMemory = passion;
				passionLocationIndex = passionLocationIndex + 5;*/
			}
		}
		LOGGER.info("<<<< AllReportsController.populateMappings");
	}
	/**
	 * Method hides Action plan columns in excel.
	 * @param sheet
	 */
	private void hideActionPlans(Sheet sheet) {
		LOGGER.info(">>>> AllReportsController.hideActionPlans");
		for (int index = 450; index < 552; index ++ ) {
			sheet.setColumnHidden(index, true);
		}
		LOGGER.info("<<<< AllReportsController.hideActionPlans");
	}
	/**
	 * Method hides reflection question columns in excel.
	 * @param sheet
	 */
	private void hideReflectionQuestions(Sheet sheet) {
		LOGGER.info(">>>> AllReportsController.hideReflectionQuestions");
		for (int index = 438; index < 492; index ++ ) {
			sheet.setColumnHidden(index, true);
		}
		LOGGER.info("<<<< AllReportsController.hideReflectionQuestions");
	}
	/**
	 * Method hides before sketch columns in excel.
	 * @param sheet
	 */
	private void hideBeforeSketch(Sheet sheet) {
		LOGGER.info(">>>> AllReportsController.hideBeforeSketch");
		sheet.setColumnHidden(10, true);
		sheet.setColumnHidden(12, true);
		sheet.setColumnHidden(18, true);
		sheet.setColumnHidden(19, true);
		sheet.setColumnHidden(23, true);
		sheet.setColumnHidden(24, true);
		sheet.setColumnHidden(25, true);
		sheet.setColumnHidden(30, true);
		sheet.setColumnHidden(31, true);
		sheet.setColumnHidden(35, true);
		sheet.setColumnHidden(36, true);
		sheet.setColumnHidden(37, true);
		sheet.setColumnHidden(42, true);
		sheet.setColumnHidden(43, true);
		sheet.setColumnHidden(47, true);
		sheet.setColumnHidden(48, true);
		sheet.setColumnHidden(49, true);
		sheet.setColumnHidden(54, true);
		sheet.setColumnHidden(55, true);
		sheet.setColumnHidden(59, true);
		sheet.setColumnHidden(60, true);
		sheet.setColumnHidden(61, true);
		sheet.setColumnHidden(66, true);
		sheet.setColumnHidden(67, true);
		sheet.setColumnHidden(71, true);
		sheet.setColumnHidden(72, true);
		sheet.setColumnHidden(73, true);
		sheet.setColumnHidden(78, true);
		sheet.setColumnHidden(79, true);
		sheet.setColumnHidden(83, true);
		sheet.setColumnHidden(84, true);
		sheet.setColumnHidden(85, true);
		sheet.setColumnHidden(90, true);
		sheet.setColumnHidden(91, true);
		sheet.setColumnHidden(95, true);
		sheet.setColumnHidden(96, true);
		sheet.setColumnHidden(97, true);
		sheet.setColumnHidden(102, true);
		sheet.setColumnHidden(103, true);
		sheet.setColumnHidden(107, true);
		sheet.setColumnHidden(108, true);
		sheet.setColumnHidden(109, true);
		sheet.setColumnHidden(114, true);
		sheet.setColumnHidden(115, true);
		sheet.setColumnHidden(119, true);
		sheet.setColumnHidden(120, true);
		sheet.setColumnHidden(121, true);
		sheet.setColumnHidden(126, true);
		sheet.setColumnHidden(127, true);
		sheet.setColumnHidden(131, true);
		sheet.setColumnHidden(132, true);
		sheet.setColumnHidden(133, true);
		sheet.setColumnHidden(138, true);
		sheet.setColumnHidden(139, true);
		sheet.setColumnHidden(143, true);
		sheet.setColumnHidden(144, true);
		sheet.setColumnHidden(145, true);
		sheet.setColumnHidden(150, true);
		sheet.setColumnHidden(151, true);
		sheet.setColumnHidden(155, true);
		sheet.setColumnHidden(156, true);
		sheet.setColumnHidden(157, true);
		sheet.setColumnHidden(162, true);
		sheet.setColumnHidden(163, true);
		sheet.setColumnHidden(167, true);
		sheet.setColumnHidden(168, true);
		sheet.setColumnHidden(169, true);
		sheet.setColumnHidden(174, true);
		sheet.setColumnHidden(175, true);
		sheet.setColumnHidden(179, true);
		sheet.setColumnHidden(180, true);
		sheet.setColumnHidden(181, true);
		sheet.setColumnHidden(186, true);
		sheet.setColumnHidden(187, true);
		sheet.setColumnHidden(191, true);
		sheet.setColumnHidden(192, true);
		sheet.setColumnHidden(193, true);
		sheet.setColumnHidden(198, true);
		sheet.setColumnHidden(199, true);
		sheet.setColumnHidden(203, true);
		sheet.setColumnHidden(204, true);
		sheet.setColumnHidden(205, true);
		sheet.setColumnHidden(210, true);
		sheet.setColumnHidden(211, true);
		sheet.setColumnHidden(215, true);
		sheet.setColumnHidden(216, true);
		sheet.setColumnHidden(217, true);
		sheet.setColumnHidden(222, true);
		sheet.setColumnHidden(223, true);
		sheet.setColumnHidden(227, true);
		sheet.setColumnHidden(228, true);
		sheet.setColumnHidden(229, true);
		sheet.setColumnHidden(234, true);
		sheet.setColumnHidden(235, true);
		sheet.setColumnHidden(239, true);
		sheet.setColumnHidden(240, true);
		sheet.setColumnHidden(241, true);
		sheet.setColumnHidden(246, true);
		sheet.setColumnHidden(247, true);
		sheet.setColumnHidden(251, true);
		sheet.setColumnHidden(252, true);
		sheet.setColumnHidden(253, true);
		LOGGER.info("<<<< AllReportsController.hideBeforeSketch");
	}
	
	private void populateSurveyQuestion(Row valRow, String emailId,
			XSSFCellStyle bodyStyle, Integer surveyCount) throws JCTException {
		LOGGER.info(">>>> AllReportsController.populateSurveyQuestion");
		List<Object> distinctMainQtnList = iReportService.getSurveyMainQtns(emailId);
		if (distinctMainQtnList.size() > 0) {
			int mainQtnPlotIndex = 594;
			int ansPlotIndex = 595;
			for (int mainQtnIndex = 0; mainQtnIndex < distinctMainQtnList.size(); mainQtnIndex ++ ){
            	Object[] obj = (Object[]) distinctMainQtnList.get(mainQtnIndex);
            	
            	String mainQuestion = (String) obj[0];
            	valRow.createCell(mainQtnPlotIndex).setCellValue(mainQuestion); 					// question
				valRow.getCell(mainQtnPlotIndex).setCellStyle(bodyStyle);
				mainQtnPlotIndex = mainQtnPlotIndex + 2;
				
            	Integer ansType = (Integer) obj[1];
            	// Get the answer of each main qtn
				List<String> ansList = iReportService.getSurveyAnswers(emailId, mainQuestion, ansType);
				//Check the size
				if (ansList.size() > 0) {
					StringBuilder sb = new StringBuilder("");
					for (int ansInd=0; ansInd < ansList.size(); ansInd++) {
						sb.append((String) ansList.get(ansInd));
						sb.append("\n");
					}
					valRow.createCell(ansPlotIndex).setCellValue(sb.toString()); 					// answer
					valRow.getCell(ansPlotIndex).setCellStyle(bodyStyle);
					ansPlotIndex = ansPlotIndex + 2;
				}	
        	}
		}
		LOGGER.info("<<<< AllReportsController.populateSurveyQuestion");
	}
}
