package com.vmware.jct.service.impl;

import java.awt.Color;
import java.awt.Graphics2D;
import java.awt.geom.Rectangle2D;
import java.io.FileOutputStream;
import java.io.IOException;
import java.net.MalformedURLException;
import java.text.DecimalFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.HashMap;
import java.util.Iterator;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.regex.Pattern;

import javax.servlet.http.HttpServletRequest;

import org.jfree.chart.ChartColor;
import org.jfree.chart.ChartFactory;
import org.jfree.chart.JFreeChart;
import org.jfree.chart.labels.PieSectionLabelGenerator;
import org.jfree.chart.labels.StandardPieSectionLabelGenerator;
import org.jfree.chart.plot.PieLabelLinkStyle;
import org.jfree.chart.plot.PiePlot;
import org.jfree.data.general.DefaultPieDataset;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;
import org.springframework.transaction.annotation.Propagation;
import org.springframework.transaction.annotation.Transactional;

import com.lowagie.text.Chunk;
import com.lowagie.text.Document;
import com.lowagie.text.DocumentException;
import com.lowagie.text.Element;
import com.lowagie.text.Font;
import com.lowagie.text.FontFactory;
import com.lowagie.text.Image;
import com.lowagie.text.PageSize;
import com.lowagie.text.Paragraph;
import com.lowagie.text.Phrase;
import com.lowagie.text.pdf.DefaultFontMapper;
import com.lowagie.text.pdf.PdfContentByte;
import com.lowagie.text.pdf.PdfPCell;
import com.lowagie.text.pdf.PdfPTable;
import com.lowagie.text.pdf.PdfTemplate;
import com.lowagie.text.pdf.PdfWriter;
import com.vmware.jct.common.utility.CommonConstants;
import com.vmware.jct.dao.IActionPlanDAO;
import com.vmware.jct.dao.IAfterSketchDAO;
import com.vmware.jct.dao.IAuthenticatorDAO;
import com.vmware.jct.dao.IBeforeSketchDAO;
import com.vmware.jct.dao.ICommonDao;
import com.vmware.jct.dao.IPdfRecordsDAO;
import com.vmware.jct.dao.IQuestionnaireDAO;
import com.vmware.jct.exception.DAOException;
import com.vmware.jct.exception.JCTException;
import com.vmware.jct.model.JctActionPlan;
import com.vmware.jct.model.JctAfterSketchFinalpageDetail;
import com.vmware.jct.model.JctAfterSketchHeader;
import com.vmware.jct.model.JctBeforeSketchHeader;
import com.vmware.jct.model.JctPdfRecords;
import com.vmware.jct.model.JctQuestion;
import com.vmware.jct.model.JctUser;
import com.vmware.jct.service.IGeneratePDFService;
import com.vmware.jct.service.vo.TaskDetailsMappingVO;

/**
 * 
 * <p><b>Class name:</b> GeneratePDFServiceImpl.java</p>
 * <p><b>Author:</b> InterraIt</p>
 * <p><b>Purpose:</b> This is a GeneratePDFServiceImpl class. This artifact is Business layer artifact.
 * GeneratePDFServiceImpl implement IGeneratePDFService interface and override the following  methods.
 * -generatePdf(String jobReferenceNumber)
 * </p>
 * <p><b>Description:</b> This class used to perform pdf generation</p>
 * <p><b>Revision History:</b>
 * 	<li></li>
 * </p>
 */
@Service
public class GeneratePDFServiceImpl implements IGeneratePDFService{
	
	@Autowired
	private IAuthenticatorDAO authenticatorDAO;
	
	@Autowired
	private IBeforeSketchDAO beforeSketchDAO;
	
	@Autowired
	private IAfterSketchDAO afterSketchDAO;
	
	@Autowired
	private IActionPlanDAO actionPlanDAO;
	
	@Autowired
	private ICommonDao commonDaoImpl;
	
	@Autowired
	private IQuestionnaireDAO questionnaireDAO;
	@Autowired
	private IPdfRecordsDAO recordsDAO;
	
	private static final Logger LOGGER = LoggerFactory.getLogger(GeneratePDFServiceImpl.class);
	
	//All the colors are shades of Green only in ascending order of their visibility
	public static final Color[] TASK_PIE_COLORS;
	static {
    	TASK_PIE_COLORS = new Color[ 20 ];
    	TASK_PIE_COLORS[0] = new Color(0, 98, 0);
        TASK_PIE_COLORS[1] = new Color(1, 124, 1);
        TASK_PIE_COLORS[2] = new Color(3, 138, 3);
        TASK_PIE_COLORS[3] = new Color(16, 144, 16);
        TASK_PIE_COLORS[4] = new Color(28, 149, 28);
        TASK_PIE_COLORS[5] = new Color(41, 155, 41);
        TASK_PIE_COLORS[6] = new Color(53, 161, 53);
        TASK_PIE_COLORS[7] = new Color(66, 167, 66);
        TASK_PIE_COLORS[8] = new Color(78, 173, 78);
        TASK_PIE_COLORS[9] = new Color(91, 179, 91);
        TASK_PIE_COLORS[10] = new Color(117, 191, 117);
        TASK_PIE_COLORS[11] = new Color(129, 196, 129);
        TASK_PIE_COLORS[12] = new Color(141, 202, 141);
        TASK_PIE_COLORS[13] = new Color(154, 208, 154);
        TASK_PIE_COLORS[14] = new Color(167, 214, 167);
        TASK_PIE_COLORS[15] = new Color(179, 220, 179);
        TASK_PIE_COLORS[16] = new Color(192, 226, 192);
        TASK_PIE_COLORS[17] = new Color(205, 232, 205);
        TASK_PIE_COLORS[18] = new Color(217, 238, 217);
        TASK_PIE_COLORS[19] = new Color(229, 243, 229);
    }
    /**
     * Method generates the pdf byte.
     * @param job reference no
     * @param email Id
     * @param disk path
     * @param HttpServletRequest
     */
	@Transactional(propagation=Propagation.REQUIRED)
	public String generatePdf(String jobReferenceNumber, String emailId, String diskPath, 
			HttpServletRequest request, int jctUserId) throws JCTException {
		LOGGER.info(">>>> GeneratePDFServiceImpl.generatePdf");
		//Fonts
		Color blue = new Color(0, 0, 128);
		Font headerFont = new Font(Font.HELVETICA, 11, Font.BOLD, blue);
		Font tableKeyFont = new Font(Font.HELVETICA, 10, Font.BOLD, blue);
		Font tableContentFont = new Font(Font.HELVETICA, 10, Font.NORMAL, blue);
		Font tableContentBoldFont = new Font(Font.HELVETICA, 10, Font.BOLD, blue);
		Font urlFont = new Font(Font.HELVETICA, 8, Font.NORMAL, blue);
		Font footerFont = new Font(Font.HELVETICA, 8, Font.NORMAL, blue);
		
		PdfWriter writer = null;
		String file = diskPath+""+new Date().getTime()+".pdf";
		try {
			StringBuilder sb = null;
	    	Document document = new Document();
	    	document.setPageSize(PageSize.A4);
		    writer = PdfWriter.getInstance(document, new FileOutputStream(file));
		    document.open();
			
			Paragraph spacer = new Paragraph("\n\n");
			//Fetch the general information which will be the header part
			List<JctUser> userList = authenticatorDAO.getUserList(emailId, 10);
			String funcGrp = "";
			String jobLvl = "";
			JctUser user = null;
			
			if(userList.size() > 0){
				for (int index = 0; index<userList.size(); index++) {
					JctUser userObj = (JctUser)userList.get(index);
					if (userObj.getJctUserId() == jctUserId) {
						funcGrp = userObj.getJctUserDetails().getJctUserDetailsFunctionGroup();
						if ((null == funcGrp) || (funcGrp.equals("null"))) {
							funcGrp = "N/A";
						}
						jobLvl = userObj.getJctUserDetails().getJctUserDetailsLevels();
						if ((null == jobLvl) || (jobLvl.equals("null"))) {
							jobLvl = "N/A";
						}
						user = userObj; 
					}
				}
				//JctUser user = (JctUser)userList.get(0);
				
				PdfPTable headerTable = new PdfPTable(2); //two column
				headerTable.setWidthPercentage(90);
				float[] columnSize = {50f, 12f};
				headerTable.setWidths(columnSize);
				headerTable.getDefaultCell().setBorder(0);
				PdfPCell headerCell;
				
				//Left aligned
				headerCell = new PdfPCell(new Phrase("Job Crafting Report", headerFont));				
				headerCell.setBorder(0);
				headerTable.addCell(headerCell);
				
				String absoluteDiskPath = request.getRealPath("/img/logo_new.png");
				Image image = Image.getInstance(absoluteDiskPath);
				image.scalePercent(90);
				headerTable.addCell(image);
				
				
				headerCell = new PdfPCell(new Phrase(" ", headerFont));				
				headerCell.setBorder(0);
				headerTable.addCell(headerCell);
				
				headerCell = new PdfPCell(new Phrase("www.jobcrafting.com", urlFont));				
				headerCell.setBorder(0);
				headerCell.setHorizontalAlignment(Element.ALIGN_RIGHT);
				headerTable.addCell(headerCell);
				
				document.add(headerTable);
				// link
				/*PdfPTable linkTable = new PdfPTable(1); 
				linkTable.setWidthPercentage(90);
				//float[] columSize = {50f, 12f};
				//linkTable.setWidths(columSize);
				linkTable.getDefaultCell().setBorder(0);
				PdfPCell link =  new PdfPCell(new Phrase("www.jobcrafting.com", tableContentFont));	
				
				linkTable.addCell(link);
				document.add(linkTable);*/
				
				
				
				
				document.add(spacer);
				
				PdfPTable table = new PdfPTable(2); //Two column
				table.setWidthPercentage(98);
				table.getDefaultCell().setBorder(0);
				float[] colSize = {6f, 20f};
				table.setWidths(colSize);
				PdfPCell cell;
					
				//Name
				cell = new PdfPCell(new Phrase("Name", tableKeyFont));
				cell.setBorder(0);
				table.addCell(cell);			
				sb = new StringBuilder("");				
				if(null != user.getJctUserDetails().getJctUserDetailsFirstName() && null != user.getJctUserDetails().getJctUserDetailsLastName()) {
					sb.append(user.getJctUserDetails().getJctUserDetailsFirstName() +" "+ user.getJctUserDetails().getJctUserDetailsLastName());
				} else if(null == user.getJctUserDetails().getJctUserDetailsFirstName() && null != user.getJctUserDetails().getJctUserDetailsLastName()) {
					sb.append(user.getJctUserDetails().getJctUserDetailsLastName());
				} else if(null != user.getJctUserDetails().getJctUserDetailsFirstName() && null == user.getJctUserDetails().getJctUserDetailsLastName()) {
					sb.append(user.getJctUserDetails().getJctUserDetailsFirstName());
				} else {
					sb.append("");
				}								
				cell = new PdfPCell(new Phrase(sb.toString(), tableContentFont));
				cell.setBorder(0);
				table.addCell(cell);
				
				//Email
				cell = new PdfPCell(new Phrase("Email", tableKeyFont));
				cell.setBorder(0);
				table.addCell(cell);
				cell = new PdfPCell(new Phrase(emailId, tableContentFont));
				cell.setBorder(0);
				table.addCell(cell);
				
				// Occupation (Justin Feedback - dated: 12-07-2015: point: 12)
				cell = new PdfPCell(new Phrase("Occupation", tableKeyFont));
				cell.setBorder(0);
				table.addCell(cell);
				sb = new StringBuilder("");
				if ((null != user.getJctUserDetails().getJctUserOnetOccupation()) && 
						(user.getJctUserDetails().getJctUserOnetOccupation().trim().length() > 0)) {
					// Fetch Occupation Title
					sb.append(actionPlanDAO.getOnetTitle(user.getJctUserDetails().getJctUserOnetOccupation()));
					
				} else {
					sb.append("");
				}
				
				cell = new PdfPCell(new Phrase(sb.toString(), tableContentFont));
				cell.setBorder(0);
				table.addCell(cell);
				
				/*//Function Group
				cell = new PdfPCell(new Phrase("Function Group", tableKeyFont));
				cell.setBorder(0);
				table.addCell(cell);
				sb = new StringBuilder("");
				sb.append(funcGrp);
				cell = new PdfPCell(new Phrase(sb.toString(), tableContentFont));
				cell.setBorder(0);
				table.addCell(cell);
				
				//Job Level
				cell = new PdfPCell(new Phrase("Job Level", tableKeyFont));
				cell.setBorder(0);
				table.addCell(cell);
				sb = new StringBuilder("");
				sb.append(jobLvl);
				cell = new PdfPCell(new Phrase(sb.toString(), tableContentFont));
				cell.setBorder(0);
				table.addCell(cell);*/
				
				document.add(table);
				Paragraph paragraph = new Paragraph("\n\n");
				document.add(paragraph);
				
				document.add(spacer);
				
				//Before Sketch Diagram
				PdfPTable beforeSketchtable = new PdfPTable(1); //1 column
				beforeSketchtable.setWidthPercentage(98);
				beforeSketchtable.getDefaultCell().setBorder(0);
				
				cell = new PdfPCell(new Phrase("Before Sketch: How you've previously spent your time and energy in your job" , tableKeyFont));
				cell.setBorderColor(blue);
				beforeSketchtable.addCell(cell);
				//Fetch the byte[] of before Sketch
				List<JctBeforeSketchHeader> hdrList = beforeSketchDAO.getList(jobReferenceNumber);
				if (hdrList.size() == 0) {
					hdrList = beforeSketchDAO.getListEdtd(jobReferenceNumber);
				}
				JctBeforeSketchHeader hdr = (JctBeforeSketchHeader)hdrList.get(0);
				cell.addElement(Image.getInstance(hdr.getJctBsSnapshot()));
				beforeSketchtable.addCell(cell);
				document.add(beforeSketchtable);
				
				document.add(paragraph);
							
				//Page 1 of 5
				PdfPTable footerTable = new PdfPTable(1); //one column
				footerTable.setWidthPercentage(90);
				footerTable.getDefaultCell().setBorder(0);
				
				//Left aligned
				headerCell = new PdfPCell(new Phrase("Page 1 of 5", footerFont));				
				headerCell.setBorder(0);
				headerCell.setHorizontalAlignment(Element.ALIGN_CENTER);
				headerCell.setVerticalAlignment(Element.ALIGN_BOTTOM);				
				footerTable.addCell(headerCell);
				document.add(footerTable);
				
				document.newPage();
				
				// Before sketch questionnaire
				//Action Plan dynamic
				PdfPTable qtntable = new PdfPTable(1); //1 column
				qtntable.setWidthPercentage(98);
				qtntable.getDefaultCell().setBorder(0);
				cell = new PdfPCell(new Phrase("Reflection Questions - Before Sketch", tableKeyFont));
				cell.setBorderColor(blue);
				qtntable.addCell(cell);
				String subQuestion = "";
				//Fetch Action Plan
				List<List> qtnAnsList = fetchQuestionAnswer(jobReferenceNumber, user.getJctUserDetails().getJctUserDetailsProfileId());
				for (int outerIndex = 0; outerIndex < qtnAnsList.size(); outerIndex ++) {
					ArrayList innerList = (ArrayList) qtnAnsList.get(outerIndex);
					int listSize = innerList.size();
					if (listSize == 3) { // 1 record
						cell = new PdfPCell(new Phrase("Question: "+(String) innerList.get(0)+"\n\n", tableContentFont));
						cell.setUseBorderPadding(true);
					    cell.setBorderWidthTop(0f);
					    cell.setBorderWidthBottom(0f);
					    cell.setBorderColor(blue);
					    cell.setBorderColorBottom(Color.WHITE);
					    
					    qtntable.addCell(cell);
					    subQuestion = (String) innerList.get(1);
					    if (!subQuestion.trim().equals("NA")) {
					    	cell = new PdfPCell(new Phrase("Sub Question: "+subQuestion+"\n", tableContentFont));
							cell.setUseBorderPadding(true);
						    cell.setBorderWidthTop(0f);
						    cell.setBorderWidthBottom(0f);
						    cell.setBorderColor(blue);
						    cell.setBorderColorBottom(Color.WHITE);
						    qtntable.addCell(cell);
					    }
					    String answer = (String) innerList.get(2);
					    if (answer.trim().equals("")) {
					    	answer = "Not Answered";
					    }
					    cell = new PdfPCell(new Phrase("Ans : "+answer+"\n", tableContentFont));
						cell.setUseBorderPadding(true);
					    cell.setBorderWidthTop(0f);
					    cell.setBorderColor(blue);
					    qtntable.addCell(cell);
					} else if (listSize == 5) { //2 records
						cell = new PdfPCell(new Phrase("Question: "+(String) innerList.get(0)+"\n\n", tableContentFont));
						cell.setUseBorderPadding(true);
					    cell.setBorderWidthTop(0f);
					    cell.setBorderWidthBottom(0f);
					    cell.setBorderColor(blue);
					    cell.setBorderColorBottom(Color.WHITE);
					    qtntable.addCell(cell);
					    subQuestion = (String) innerList.get(1);
					    if (!subQuestion.trim().equals("NA")) {
					    	cell = new PdfPCell(new Phrase("Sub Question: "+subQuestion+"\n", tableContentFont));
							cell.setUseBorderPadding(true);
						    cell.setBorderWidthTop(0f);
						    cell.setBorderWidthBottom(0f);
						    cell.setBorderColor(blue);
						    cell.setBorderColorBottom(Color.WHITE);
						    qtntable.addCell(cell);
					    }
						String answer = (String) innerList.get(2);
					    if (answer.trim().equals("")) {
					    	answer = "Not Answered";
					    }
						cell = new PdfPCell(new Phrase("Ans : "+answer+"\n", tableContentFont));
						cell.setUseBorderPadding(true);
					    cell.setBorderWidthTop(0f);
					    cell.setBorderWidthBottom(0f);
					    cell.setBorderColor(blue);
					    cell.setBorderColorBottom(Color.WHITE);
					    qtntable.addCell(cell);
					    subQuestion = (String) innerList.get(3);
					    if (!subQuestion.trim().equals("NA")) {
					    	cell = new PdfPCell(new Phrase("Sub Question: "+subQuestion+"\n", tableContentFont));
							cell.setUseBorderPadding(true);
						    cell.setBorderWidthTop(0f);
						    cell.setBorderWidthBottom(0f);
						    cell.setBorderColor(blue);
						    cell.setBorderColorBottom(Color.WHITE);
						    qtntable.addCell(cell);
					    }
					    answer = (String) innerList.get(4);
					    if (answer.trim().equals("")) {
					    	answer = "Not Answered";
					    }
						cell = new PdfPCell(new Phrase("Ans : "+answer+"\n", tableContentFont));
						cell.setUseBorderPadding(true);
					    cell.setBorderWidthTop(0f);
					    cell.setBorderColor(blue);
					    qtntable.addCell(cell);
					} else if (listSize == 7) { //3 records
						cell = new PdfPCell(new Phrase("Question: "+(String) innerList.get(0)+"\n\n", tableContentFont));
						cell.setUseBorderPadding(true);
					    cell.setBorderWidthTop(0f);
					    cell.setBorderWidthBottom(0f);
					    cell.setBorderColor(blue);
					    cell.setBorderColorBottom(Color.WHITE);
					    qtntable.addCell(cell);
					    subQuestion = (String) innerList.get(1);
					    if (!subQuestion.trim().equals("NA")) {
					    	cell = new PdfPCell(new Phrase("Sub Question: "+subQuestion+"\n", tableContentFont));
							cell.setUseBorderPadding(true);
						    cell.setBorderWidthTop(0f);
						    cell.setBorderWidthBottom(0f);
						    cell.setBorderColor(blue);
						    cell.setBorderColorBottom(Color.WHITE);
						    qtntable.addCell(cell);
					    }						
					    String answer = (String) innerList.get(2);
					    if (answer.trim().equals("")) {
					    	answer = "Not Answered";
					    }
						cell = new PdfPCell(new Phrase("Ans : "+answer+"\n", tableContentFont));
						cell.setUseBorderPadding(true);
					    cell.setBorderWidthTop(0f);
					    cell.setBorderWidthBottom(0f);
					    cell.setBorderColor(blue);
					    cell.setBorderColorBottom(Color.WHITE);
					    qtntable.addCell(cell);
					    subQuestion = (String) innerList.get(3);
					    if (!subQuestion.trim().equals("NA")) {
					    	cell = new PdfPCell(new Phrase("Sub Question: "+subQuestion+"\n", tableContentFont));
							cell.setUseBorderPadding(true);
						    cell.setBorderWidthTop(0f);
						    cell.setBorderWidthBottom(0f);
						    cell.setBorderColor(blue);
						    cell.setBorderColorBottom(Color.WHITE);
						    qtntable.addCell(cell);
					    }
					    answer = (String) innerList.get(4);
					    if (answer.trim().equals("")) {
					    	answer = "Not Answered";
					    }
						cell = new PdfPCell(new Phrase("Ans : "+answer+"\n", tableContentFont));
						cell.setUseBorderPadding(true);
					    cell.setBorderWidthTop(0f);
					    cell.setBorderWidthBottom(0f);
					    cell.setBorderColor(blue);
					    cell.setBorderColorBottom(Color.WHITE);
					    qtntable.addCell(cell);
					    subQuestion = (String) innerList.get(5);
					    if (!subQuestion.trim().equals("NA")) {
					    	cell = new PdfPCell(new Phrase("Sub Question: "+subQuestion+"\n", tableContentFont));
							cell.setUseBorderPadding(true);
						    cell.setBorderWidthTop(0f);
						    cell.setBorderWidthBottom(0f);
						    cell.setBorderColor(blue);
						    cell.setBorderColorBottom(Color.WHITE);
						    qtntable.addCell(cell);
					    }
						answer = (String) innerList.get(6);
					    if (answer.trim().equals("")) {
					    	answer = "Not Answered";
					    }
						cell = new PdfPCell(new Phrase("Ans : "+answer+"\n", tableContentFont));
						cell.setUseBorderPadding(true);
						cell.setBorderColor(blue);
					    cell.setBorderWidthTop(0f);
					    qtntable.addCell(cell);
					} else if (listSize == 9) { //4 records
						cell = new PdfPCell(new Phrase("Question: "+(String) innerList.get(0)+"\n\n", tableContentFont));
						cell.setUseBorderPadding(true);
					    cell.setBorderWidthTop(0f);
					    cell.setBorderWidthBottom(0f);
					    cell.setBorderColor(blue);
					    cell.setBorderColorBottom(Color.WHITE);
					    qtntable.addCell(cell);
					    subQuestion = (String) innerList.get(1);
					    if (!subQuestion.trim().equals("NA")) {
					    	cell = new PdfPCell(new Phrase("Sub Question: "+subQuestion+"\n", tableContentFont));
							cell.setUseBorderPadding(true);
						    cell.setBorderWidthTop(0f);
						    cell.setBorderWidthBottom(0f);
						    cell.setBorderColor(blue);
						    cell.setBorderColorBottom(Color.WHITE);
						    qtntable.addCell(cell);
					    }
						String answer = (String) innerList.get(2);
					    if (answer.trim().equals("")) {
					    	answer = "Not Answered";
					    }
						cell = new PdfPCell(new Phrase("Ans : "+answer+"\n", tableContentFont));
						cell.setUseBorderPadding(true);
					    cell.setBorderWidthTop(0f);
					    cell.setBorderWidthBottom(0f);
					    cell.setBorderColor(blue);
					    cell.setBorderColorBottom(Color.WHITE);
					    qtntable.addCell(cell);
					    subQuestion = (String) innerList.get(3);
					    if (!subQuestion.trim().equals("NA")) {
					    	cell = new PdfPCell(new Phrase("Sub Question: "+subQuestion+"\n", tableContentFont));
							cell.setUseBorderPadding(true);
						    cell.setBorderWidthTop(0f);
						    cell.setBorderWidthBottom(0f);
						    cell.setBorderColor(blue);
						    cell.setBorderColorBottom(Color.WHITE);
						    qtntable.addCell(cell);
					    }
						answer = (String) innerList.get(4);
					    if (answer.trim().equals("")) {
					    	answer = "Not Answered";
					    }
						cell = new PdfPCell(new Phrase("Ans : "+answer+"\n", tableContentFont));
						cell.setUseBorderPadding(true);
					    cell.setBorderWidthTop(0f);
					    cell.setBorderWidthBottom(0f);
					    cell.setBorderColor(blue);
					    cell.setBorderColorBottom(Color.WHITE);
					    qtntable.addCell(cell);
					    subQuestion = (String) innerList.get(5);
					    if (!subQuestion.trim().equals("NA")) {
					    	cell = new PdfPCell(new Phrase("Sub Question: "+subQuestion+"\n", tableContentFont));
							cell.setUseBorderPadding(true);
						    cell.setBorderWidthTop(0f);
						    cell.setBorderWidthBottom(0f);
						    cell.setBorderColor(blue);
						    cell.setBorderColorBottom(Color.WHITE);
						    qtntable.addCell(cell);
					    }
					    answer = (String) innerList.get(6);
					    if (answer.trim().equals("")) {
					    	answer = "Not Answered";
					    }
						cell = new PdfPCell(new Phrase("Ans : "+answer+"\n", tableContentFont));
						cell.setBorderColor(blue);
						cell.setUseBorderPadding(true);
					    cell.setBorderWidthTop(0f);
					    cell.setBorderWidthBottom(0f);
					    cell.setBorderColor(blue);
					    cell.setBorderColorBottom(Color.WHITE);
					    qtntable.addCell(cell);
					    subQuestion = (String) innerList.get(7);
					    if (!subQuestion.trim().equals("NA")) {
					    	cell = new PdfPCell(new Phrase("Sub Question: "+subQuestion+"\n", tableContentFont));
							cell.setUseBorderPadding(true);
						    cell.setBorderWidthTop(0f);
						    cell.setBorderWidthBottom(0f);
						    cell.setBorderColor(blue);
						    cell.setBorderColorBottom(Color.WHITE);
						    qtntable.addCell(cell);
					    }
						answer = (String) innerList.get(8);
					    if (answer.trim().equals("")) {
					    	answer = "Not Answered";
					    }
						cell = new PdfPCell(new Phrase("Ans : "+answer+"\n", tableContentFont));
						cell.setUseBorderPadding(true);
					    cell.setBorderWidthTop(0f);
					    cell.setBorderColor(blue);
					    qtntable.addCell(cell);
					}
				}
				document.add(qtntable);
				
				//Page 2 of 5
				PdfPTable footerTable2 = new PdfPTable(1); //one column
				footerTable2.setWidthPercentage(90);
				footerTable2.getDefaultCell().setBorder(0);
				
				//Left aligned
				headerCell = new PdfPCell(new Phrase("Page 2 of 5", footerFont));				
				headerCell.setBorder(0);
				headerCell.setHorizontalAlignment(Element.ALIGN_CENTER);
				headerCell.setVerticalAlignment(Element.ALIGN_BOTTOM);				
				footerTable2.addCell(headerCell);
				document.add(footerTable2);
				
				document.newPage(); //Changing it to new page
				
				//After Sketch Diagram
				PdfPTable afterSketchtable = new PdfPTable(1); //1 column
				afterSketchtable.setWidthPercentage(98);
				afterSketchtable.getDefaultCell().setBorder(0);
				cell = new PdfPCell(new Phrase("After Diagram: How to craft your job going forward to better suit your values, strengths, and passions" , tableKeyFont));
				cell.setBorderColor(blue);
				afterSketchtable.addCell(cell);
				//Fetch the byte[] of after Sketch
				List<JctAfterSketchHeader> asHdrList = afterSketchDAO.getList(jobReferenceNumber, "N");
				JctAfterSketchHeader asHdr = (JctAfterSketchHeader)asHdrList.get(0);
				cell.addElement(Image.getInstance(asHdr.getJctAsFinalpageSnapshot()));
				cell.setFixedHeight(700f);
				afterSketchtable.addCell(cell);
				document.add(afterSketchtable);
				//Page 3 of 5
				PdfPTable footerTable3 = new PdfPTable(1); //one column
				footerTable3.setWidthPercentage(90);
				footerTable3.getDefaultCell().setBorder(0);
				
				//Left aligned
				headerCell = new PdfPCell(new Phrase("Page 3 of 5", footerFont));				
				headerCell.setBorder(0);
				headerCell.setHorizontalAlignment(Element.ALIGN_CENTER);
				headerCell.setVerticalAlignment(Element.ALIGN_BOTTOM);				
				footerTable3.addCell(headerCell);
				document.add(footerTable3);
				
				document.newPage(); //Changing it to new page
				//Adding the pie here..
				//Fetch After sketch Record List
				List<JctAfterSketchHeader> headerList = getJctAfterSketchHeader(jobReferenceNumber);
				if (headerList != null) {
					float width = 510;
			        float height = 380;
					PdfContentByte contentByte = writer.getDirectContent();
			        PdfTemplate template = contentByte.createTemplate(width, height);
					Graphics2D graphics2d = template.createGraphics(width, height, new DefaultFontMapper());
			        Rectangle2D rectangle2d = new Rectangle2D.Double(0, 0, width, height);
					List<JctAfterSketchFinalpageDetail> dtlList = afterSketchDAO.getASFinalPageListDescOrder(jobReferenceNumber); //Fetching child objs in desc order according to the task energy
					
					JFreeChart dtlsChart = generatePieChart(dtlList);
					dtlsChart.setBackgroundPaint(ChartColor.white);
					dtlsChart.setBorderPaint(ChartColor.white);
					dtlsChart.setBorderVisible(false);

					dtlsChart.draw(graphics2d, rectangle2d);
			        graphics2d.dispose();
			        
			        
			        // Fetch distinct role frames
			        List<String> roleFramesList = afterSketchDAO.getDistinctRoleFrames(jobReferenceNumber);
			        Map<Integer, String> roleMap = new HashMap<Integer, String>();
			        int counter = 1;
			        for (int index = 0; index < roleFramesList.size(); index++) {
			        	 if (!roleFramesList.get(index).toString().equals("N/A")) {
			        		 roleMap.put(counter, roleFramesList.get(index).toString());
					        	counter = counter + 1; 
			        	 }			        	
			        }
			        LOGGER.info("FOUND "+roleFramesList.size()+" ROLE FRAMES FOR JOB REFERENCE NUMBER: "+jobReferenceNumber);
			        
			        //After Sketch Diagram
                    PdfPTable pietable = new PdfPTable(1); //1 column
                    pietable.setWidthPercentage(98);
                    pietable.getDefaultCell().setBorder(0);
                    PdfPCell cells = new PdfPCell(new Phrase("Future Allocation of Time/Energy" , tableKeyFont));
                    cells.setBorderColor(blue);
                    pietable.addCell(cells);
                    //Fetch the byte[] of after Sketch
                    Image chartImage = Image.getInstance(template);
                    cells.addElement(Image.getInstance(chartImage));
                    pietable.addCell(cells);
                    document.add(pietable);

			        
			    	// Plot the main roles
					PdfPTable roleTable = new PdfPTable(2);
					roleTable.setWidthPercentage((float) 98.00);
					float[] roleColsSize = {5.1f, 50.3f};
					roleTable.setWidths(roleColsSize);
					Map<String, String> finalRole = new HashMap<String, String>();
					for (Map.Entry<Integer, String> entry : roleMap.entrySet()) {
					    int key = (Integer) entry.getKey();
						String roleHead = "Role A: ";
						String roleDist = "A";
						if (key == 2) {
							roleHead = "Role B: ";
							roleDist = "B";
						} else if (key == 3) {
							roleHead = "Role C: ";
							roleDist = "C";
						} else if (key == 4) {
							roleHead = "Role D: ";
							roleDist = "D";
						}
						finalRole.put(entry.getValue().toString(), roleDist);
						if (!entry.getValue().toString().equals("N/A")) {
							PdfPCell roleHeadCell = new PdfPCell(new Phrase(roleHead+"\n   ", tableContentBoldFont));
							roleHeadCell.setUseBorderPadding(true);
							roleHeadCell.setBorderWidthTop(0f);
							roleHeadCell.setBorderWidthBottom(0f);
							//roleHeadCell.setBorderWidthLeft(0f);
							roleHeadCell.setBorderWidthRight(0f);
							roleHeadCell.setBorderColorRight(Color.WHITE);
							//roleHeadCell.setBorderColorLeft(Color.WHITE);
							roleHeadCell.setBorderColorTop(Color.WHITE);
							roleHeadCell.setBorderColorBottom(Color.WHITE);
							roleHeadCell.setBorderColor(blue);
							
							roleTable.addCell(roleHeadCell);
							
							PdfPCell roleDescCell = new PdfPCell(new Phrase(entry.getValue().toString()+"\n   ", tableContentFont));
							roleDescCell.setUseBorderPadding(true);
							roleDescCell.setBorderWidthTop(0f);
							roleDescCell.setBorderWidthBottom(0f);
							roleDescCell.setBorderWidthLeft(0f);
							//roleDescCell.setBorderWidthRight(0f);
							//roleDescCell.setBorderColorRight(Color.WHITE);
							roleDescCell.setBorderColorLeft(Color.WHITE);
							roleDescCell.setBorderColorTop(Color.WHITE);
							roleDescCell.setBorderColorBottom(Color.WHITE);
							roleDescCell.setBorderColor(blue);
							roleTable.addCell(roleDescCell);
						}
					}
					document.add(roleTable);
					
					//get the relational crafting text and show it in tabular format...
					Map<String, TaskDetailsMappingVO> taskDetailsMap = getTaskDescMapping(dtlList);
					PdfPTable pieInnertable = new PdfPTable(5);
					//pieInnertable.setSpacingBefore(10f);
					pieInnertable.setWidthPercentage(98);
					float[] colsSize = {4f, 9.3f, 22.1f, 22.1f, 4.6f};
					pieInnertable.setWidths(colsSize);
					//Iterate the map
					Iterator iterator = taskDetailsMap.entrySet().iterator();
					PdfPCell taskMappingCell = null;
					//populate the headers
					//populate the cells
					taskMappingCell = new PdfPCell(new Phrase("Task \n #" , tableKeyFont));
					taskMappingCell.setHorizontalAlignment(Element.ALIGN_CENTER);
					taskMappingCell.setBorderColor(blue);
					pieInnertable.addCell(taskMappingCell);
					
					taskMappingCell = new PdfPCell(new Phrase("Time / Energy % \n(Change from Before Sketch)" , tableKeyFont));
					taskMappingCell.setHorizontalAlignment(Element.ALIGN_CENTER);
					taskMappingCell.setBorderColor(blue);
					pieInnertable.addCell(taskMappingCell);
					
					taskMappingCell = new PdfPCell(new Phrase("Task Crafting" , tableKeyFont));
					taskMappingCell.setHorizontalAlignment(Element.ALIGN_CENTER);
					taskMappingCell.setBorderColor(blue);
					pieInnertable.addCell(taskMappingCell);
					
					taskMappingCell = new PdfPCell(new Phrase("Relational Crafting" , tableKeyFont));
					taskMappingCell.setHorizontalAlignment(Element.ALIGN_CENTER);
					taskMappingCell.setBorderColor(blue);
					pieInnertable.addCell(taskMappingCell);
					
					taskMappingCell = new PdfPCell(new Phrase("Role" , tableKeyFont));
					taskMappingCell.setHorizontalAlignment(Element.ALIGN_CENTER);
					taskMappingCell.setBorderColor(blue);
					pieInnertable.addCell(taskMappingCell);
					int count = 1;
					while (iterator.hasNext()) {
						count = count + 1;
						Map.Entry mapEntry = (Map.Entry) iterator.next();
						String taskKey = mapEntry.getKey().toString();
						TaskDetailsMappingVO voObj = (TaskDetailsMappingVO) mapEntry.getValue();
						
						//populate the cells
						taskMappingCell = new PdfPCell(new Phrase(taskKey , tableContentFont));
						taskMappingCell.setBorderColor(blue);
						taskMappingCell.setHorizontalAlignment(Element.ALIGN_CENTER);
						pieInnertable.addCell(taskMappingCell);
						
						taskMappingCell = new PdfPCell(new Phrase(voObj.getTimeEnergyChangeFromBeforeSketch() , tableContentFont));
						taskMappingCell.setBorderColor(blue);
						taskMappingCell.setHorizontalAlignment(Element.ALIGN_CENTER);
						pieInnertable.addCell(taskMappingCell);
						
						taskMappingCell = new PdfPCell(new Phrase(voObj.getTaskCrafting() , tableContentFont));
						taskMappingCell.setBorderColor(blue);
						pieInnertable.addCell(taskMappingCell);
						
						taskMappingCell = new PdfPCell(new Phrase(voObj.getRelationalCrafting() , tableContentFont));
						taskMappingCell.setBorderColor(blue);
						pieInnertable.addCell(taskMappingCell);
						
						//	Fetch the role for the task
						List<String> roleList = afterSketchDAO.getDistinctRoleName(jobReferenceNumber, voObj.getTaskCrafting());
						StringBuffer roleBuff = new StringBuffer("");
						if (roleList.size() == 3) {
							roleBuff.append("A, B, C");
						} else if (roleList.size() == 2) {
							for(int index=0; index<roleList.size(); index++) {
								roleBuff.append(finalRole.get(roleList.get(index).toString()));
								if (index < (roleList.size()-1)){
									roleBuff.append(" , ");
								}
							}
						} else {
							try {
								roleBuff.append(finalRole.get(roleList.get(0)).toString());
							} catch (Exception ex) {
								roleBuff.append("--");
							}
							/*roleBuff.append(finalRole.get(roleList.get(0)).toString());
							if (roleBuff.equals("N/A")) {
								roleBuff.append("--");
							}*/
						}
						
						taskMappingCell = new PdfPCell(new Phrase(roleBuff.toString() , tableContentFont));
						taskMappingCell.setHorizontalAlignment(Element.ALIGN_CENTER);
						taskMappingCell.setBorderColor(blue);
						pieInnertable.addCell(taskMappingCell);
					}
					
					// Add the dropped task
					List<String> deletedList = getDeletedTask(jobReferenceNumber);
					if (deletedList.size() > 0) {
						for (int i=0; i<deletedList.size(); i++) {
							taskMappingCell = new PdfPCell(new Phrase(count+"" , tableContentFont));
							taskMappingCell.setBorderColor(blue);
							taskMappingCell.setHorizontalAlignment(Element.ALIGN_CENTER);
							pieInnertable.addCell(taskMappingCell);
							
							String string = deletedList.get(i).toString();
							String[] split = string.split("``");
							taskMappingCell = new PdfPCell(new Phrase(split[1] , tableContentFont));
							taskMappingCell.setBorderColor(blue);
							taskMappingCell.setHorizontalAlignment(Element.ALIGN_CENTER);
							pieInnertable.addCell(taskMappingCell);
							
							taskMappingCell = new PdfPCell(new Phrase(split[0] , tableContentFont));
							taskMappingCell.setBorderColor(blue);
							//taskMappingCell.setHorizontalAlignment(Element.ALIGN_CENTER);
							pieInnertable.addCell(taskMappingCell);
							
							taskMappingCell = new PdfPCell(new Phrase("--" , tableContentFont));
							taskMappingCell.setBorderColor(blue);
							//taskMappingCell.setHorizontalAlignment(Element.ALIGN_CENTER);
							pieInnertable.addCell(taskMappingCell);
							
							taskMappingCell = new PdfPCell(new Phrase("--" , tableContentFont));
							taskMappingCell.setBorderColor(blue);
							taskMappingCell.setHorizontalAlignment(Element.ALIGN_CENTER);
							pieInnertable.addCell(taskMappingCell);
							count = count + 1;
						}
						
						
						/*String string = "Task 1``(was 15% of time/energy)";
						String[] s = string.split("``");
						
						taskMappingCell = new PdfPCell(new Phrase(voObj.getTimeEnergyChangeFromBeforeSketch() , tableContentFont));
						taskMappingCell.setBorderColor(blue);
						taskMappingCell.setHorizontalAlignment(Element.ALIGN_CENTER);
						pieInnertable.addCell(taskMappingCell);
						
						//Get the deleted tasks.
						PdfPTable droppedTable = new PdfPTable(1);
						droppedTable.setWidthPercentage(98);
						
						PdfPCell droppedCellInfo = new PdfPCell();
						droppedCellInfo.setUseBorderPadding(true);
						droppedCellInfo.setBorderWidthTop(0f);
						droppedCellInfo.setBorderWidthBottom(0f);
						//droppedCellInfo.setBorderColorBottom(Color.WHITE);
						droppedCellInfo.setBorderColor(blue);
						droppedCellInfo.setBorderColor(Color.WHITE);
					    
						Font boldDeleted = new Font(Font.HELVETICA, 11, Font.BOLD, blue);
						Font deletedNormal = new Font(Font.HELVETICA, 11, Font.NORMAL, blue);
						Phrase droppedTaskInfo = new Phrase();
						droppedTaskInfo.add(new Chunk("\n", boldDeleted));
						droppedTaskInfo.add(new Chunk("Dropped Tasks: ", boldDeleted));
						droppedTaskInfo.add(new Chunk("The following are tasks from your Before Sketch that are not in your After Diagram.", deletedNormal));
						droppedCellInfo.setBorderColor(blue);
						droppedCellInfo.addElement(droppedTaskInfo);
						droppedTable.addCell(droppedCellInfo);*/
						
						//Actual details
						/*PdfPCell detailsInfo = new PdfPCell();
						detailsInfo.setUseBorderPadding(true);
						detailsInfo.setBorderWidthTop(0f);
						detailsInfo.setBorderWidthBottom(0f);
						detailsInfo.setBorderColor(blue);
						detailsInfo.setBorderColor(Color.WHITE);
						com.lowagie.text.List unorderedList = new com.lowagie.text.List(com.lowagie.text.List.UNORDERED);
						com.lowagie.text.ListItem listItem;*/
						
						/*Iterator<String> delTaskItr = deletedList.iterator();
						while (delTaskItr.hasNext()) {
							listItem = new com.lowagie.text.ListItem(delTaskItr.next().toString(), FontFactory.getFont(FontFactory.TIMES_ROMAN, 12));
							unorderedList.add(listItem);
						}
						detailsInfo.addElement(unorderedList);
						droppedTable.addCell(detailsInfo);
						document.add(droppedTable);*/
					}
					document.add(pieInnertable);
					
					
					/*List<String> deletedList = getDeletedTask(jobReferenceNumber);
					if (deletedList.size() > 0) {
						//Get the deleted tasks.
						PdfPTable droppedTable = new PdfPTable(1);
						droppedTable.setWidthPercentage(98);
						
						PdfPCell droppedCellInfo = new PdfPCell();
						droppedCellInfo.setUseBorderPadding(true);
						droppedCellInfo.setBorderWidthTop(0f);
						droppedCellInfo.setBorderWidthBottom(0f);
						//droppedCellInfo.setBorderColorBottom(Color.WHITE);
						droppedCellInfo.setBorderColor(blue);
						droppedCellInfo.setBorderColor(Color.WHITE);
					    
						Font boldDeleted = new Font(Font.HELVETICA, 11, Font.BOLD, blue);
						Font deletedNormal = new Font(Font.HELVETICA, 11, Font.NORMAL, blue);
						Phrase droppedTaskInfo = new Phrase();
						droppedTaskInfo.add(new Chunk("\n", boldDeleted));
						droppedTaskInfo.add(new Chunk("Dropped Tasks: ", boldDeleted));
						droppedTaskInfo.add(new Chunk("The following are tasks from your Before Sketch that are not in your After Diagram.", deletedNormal));
						droppedCellInfo.setBorderColor(blue);
						droppedCellInfo.addElement(droppedTaskInfo);
						droppedTable.addCell(droppedCellInfo);
						
						//Actual details
						PdfPCell detailsInfo = new PdfPCell();
						detailsInfo.setUseBorderPadding(true);
						detailsInfo.setBorderWidthTop(0f);
						detailsInfo.setBorderWidthBottom(0f);
						detailsInfo.setBorderColor(blue);
						detailsInfo.setBorderColor(Color.WHITE);
						com.lowagie.text.List unorderedList = new com.lowagie.text.List(com.lowagie.text.List.UNORDERED);
						com.lowagie.text.ListItem listItem;
						
						Iterator<String> delTaskItr = deletedList.iterator();
						while (delTaskItr.hasNext()) {
							listItem = new com.lowagie.text.ListItem(delTaskItr.next().toString(), FontFactory.getFont(FontFactory.TIMES_ROMAN, 12));
							unorderedList.add(listItem);
						}
						detailsInfo.addElement(unorderedList);
						droppedTable.addCell(detailsInfo);
						document.add(droppedTable);
					}*/
				}
				
				//Page 4 of 5
				PdfPTable footerTable4 = new PdfPTable(1); //one column
				footerTable4.setWidthPercentage(90);
				footerTable4.getDefaultCell().setBorder(0);
				
				//Left aligned
				headerCell = new PdfPCell(new Phrase("Page 4 of 5", footerFont));				
				headerCell.setBorder(0);
				headerCell.setHorizontalAlignment(Element.ALIGN_CENTER);
				headerCell.setVerticalAlignment(Element.ALIGN_BOTTOM);				
				footerTable4.addCell(headerCell);
				document.add(footerTable4);
				
				document.newPage(); //Changing it to new page
				
				//Action Plan dynamic
				PdfPTable actionPlantable = new PdfPTable(1); //1 column
				actionPlantable.setWidthPercentage(98);
				actionPlantable.getDefaultCell().setBorder(0);
				cell = new PdfPCell(new Phrase("Action Plan: Making Your After Diagram a Reality", tableKeyFont));
				cell.setBorderColor(blue);
				actionPlantable.addCell(cell);
				
				//Fetch Action Plan
				List<List> actionPlanList = fetchActionPlans(jobReferenceNumber, user.getJctUserDetails().getJctUserDetailsProfileId());
				for (int outerIndex = 0; outerIndex < actionPlanList.size(); outerIndex ++) {
					ArrayList innerList = (ArrayList) actionPlanList.get(outerIndex);
					int listSize = innerList.size();
					if (listSize == 5) { // 1 record
						cell = new PdfPCell(new Phrase("Question: "+(String) innerList.get(0)+"\n\n", tableContentFont));
						cell.setUseBorderPadding(true);
					    cell.setBorderWidthTop(0f);
					    cell.setBorderWidthBottom(0f);
					    cell.setBorderColor(blue);
					    cell.setBorderColorBottom(Color.WHITE);
						actionPlantable.addCell(cell);
						subQuestion = (String) innerList.get(1);
						if (!subQuestion.trim().equals("NA")) {
							cell = new PdfPCell(new Phrase("Sub Question: "+(String) innerList.get(1)+"\n", tableContentFont));
							cell.setUseBorderPadding(true);
						    cell.setBorderWidthTop(0f);
						    cell.setBorderWidthBottom(0f);
						    cell.setBorderColor(blue);
						    cell.setBorderColorBottom(Color.WHITE);
							actionPlantable.addCell(cell);
						}
						cell = new PdfPCell(new Phrase("Ans 1: "+(String) innerList.get(2), tableContentFont));
						cell.setUseBorderPadding(true);
					    cell.setBorderWidthTop(0f);
					    cell.setBorderWidthBottom(0f);
					    cell.setBorderColor(blue);
					    cell.setBorderColorBottom(Color.WHITE);
						actionPlantable.addCell(cell);
						cell = new PdfPCell(new Phrase("Ans 2: "+(String) innerList.get(3), tableContentFont));
						cell.setUseBorderPadding(true);
					    cell.setBorderWidthTop(0f);
					    cell.setBorderWidthBottom(0f);
					    cell.setBorderColor(blue);
					    cell.setBorderColorBottom(Color.WHITE);
					    actionPlantable.addCell(cell);
						cell = new PdfPCell(new Phrase("Ans 3: "+(String) innerList.get(4)+"\n\n", tableContentFont));
						cell.setUseBorderPadding(true);
					    cell.setBorderWidthTop(0f);
					    cell.setBorderColor(blue);
					    actionPlantable.addCell(cell);
					} else if (listSize == 9) {//2 records
						cell = new PdfPCell(new Phrase("Question: "+(String) innerList.get(0)+"\n\n", tableContentFont));
						cell.setUseBorderPadding(true);
					    cell.setBorderWidthTop(0f);
					    cell.setBorderWidthBottom(0f);
					    cell.setBorderColor(blue);
					    cell.setBorderColorBottom(Color.WHITE);
						actionPlantable.addCell(cell);
						subQuestion = (String) innerList.get(1);
						if (!subQuestion.trim().equals("NA")) {
							cell = new PdfPCell(new Phrase("Sub Question: "+(String) innerList.get(1)+"\n", tableContentFont));
							cell.setUseBorderPadding(true);
						    cell.setBorderWidthTop(0f);
						    cell.setBorderWidthBottom(0f);
						    cell.setBorderColor(blue);
						    cell.setBorderColorBottom(Color.WHITE);
							actionPlantable.addCell(cell);
						}
						cell = new PdfPCell(new Phrase("Ans 1: "+(String) innerList.get(2), tableContentFont));
						cell.setUseBorderPadding(true);
					    cell.setBorderWidthTop(0f);
					    cell.setBorderWidthBottom(0f);
					    cell.setBorderColor(blue);
					    cell.setBorderColorBottom(Color.WHITE);
						actionPlantable.addCell(cell);
						cell = new PdfPCell(new Phrase("Ans 2: "+(String) innerList.get(3), tableContentFont));
						cell.setUseBorderPadding(true);
					    cell.setBorderWidthTop(0f);
					    cell.setBorderWidthBottom(0f);
					    cell.setBorderColor(blue);
					    cell.setBorderColorBottom(Color.WHITE);
						actionPlantable.addCell(cell);
						cell = new PdfPCell(new Phrase("Ans 3: "+(String) innerList.get(4)+"\n\n", tableContentFont));
						cell.setUseBorderPadding(true);
					    cell.setBorderWidthTop(0f);
					    cell.setBorderWidthBottom(0f);
					    cell.setBorderColor(blue);
					    cell.setBorderColorBottom(Color.WHITE);
						actionPlantable.addCell(cell);
						subQuestion = (String) innerList.get(5);
						if (!subQuestion.trim().equals("NA")) {
							cell = new PdfPCell(new Phrase("Sub Question: "+(String) innerList.get(5)+"\n", tableContentFont));
							cell.setUseBorderPadding(true);
						    cell.setBorderWidthTop(0f);
						    cell.setBorderWidthBottom(0f);
						    cell.setBorderColor(blue);
						    cell.setBorderColorBottom(Color.WHITE);
							actionPlantable.addCell(cell);
						}
						cell = new PdfPCell(new Phrase("Ans 1: "+(String) innerList.get(6), tableContentFont));
						cell.setUseBorderPadding(true);
					    cell.setBorderWidthTop(0f);
					    cell.setBorderWidthBottom(0f);
					    cell.setBorderColor(blue);
					    cell.setBorderColorBottom(Color.WHITE);
						actionPlantable.addCell(cell);
						cell = new PdfPCell(new Phrase("Ans 2: "+(String) innerList.get(7), tableContentFont));
						cell.setUseBorderPadding(true);
					    cell.setBorderWidthTop(0f);
					    cell.setBorderWidthBottom(0f);
					    cell.setBorderColor(blue);
					    cell.setBorderColorBottom(Color.WHITE);
						actionPlantable.addCell(cell);
						cell = new PdfPCell(new Phrase("Ans 3: "+(String) innerList.get(8)+"\n\n", tableContentFont));
						cell.setUseBorderPadding(true);
					    cell.setBorderWidthTop(0f);
					    cell.setBorderColor(blue);
					    actionPlantable.addCell(cell);
					} else if (listSize == 13) { //3 records
						cell = new PdfPCell(new Phrase("Question: "+(String) innerList.get(0)+"\n\n", tableContentFont));
						cell.setUseBorderPadding(true);
					    cell.setBorderWidthTop(0f);
					    cell.setBorderWidthBottom(0f);
					    cell.setBorderColor(blue);
					    cell.setBorderColorBottom(Color.WHITE);
						actionPlantable.addCell(cell);
						subQuestion = (String) innerList.get(1);
						if (!subQuestion.trim().equals("NA")) {
							cell = new PdfPCell(new Phrase("Sub Question: "+(String) innerList.get(1)+"\n", tableContentFont));
							cell.setUseBorderPadding(true);
						    cell.setBorderWidthTop(0f);
						    cell.setBorderWidthBottom(0f);
						    cell.setBorderColor(blue);
						    cell.setBorderColorBottom(Color.WHITE);
							actionPlantable.addCell(cell);
						}
						cell = new PdfPCell(new Phrase("Ans 1: "+(String) innerList.get(2), tableContentFont));
						cell.setUseBorderPadding(true);
					    cell.setBorderWidthTop(0f);
					    cell.setBorderWidthBottom(0f);
					    cell.setBorderColor(blue);
					    cell.setBorderColorBottom(Color.WHITE);
						actionPlantable.addCell(cell);
						cell = new PdfPCell(new Phrase("Ans 2: "+(String) innerList.get(3), tableContentFont));
						cell.setUseBorderPadding(true);
					    cell.setBorderWidthTop(0f);
					    cell.setBorderWidthBottom(0f);
					    cell.setBorderColor(blue);
					    cell.setBorderColorBottom(Color.WHITE);
						actionPlantable.addCell(cell);
						cell = new PdfPCell(new Phrase("Ans 3: "+(String) innerList.get(4)+"\n\n", tableContentFont));
						cell.setUseBorderPadding(true);
					    cell.setBorderWidthTop(0f);
					    cell.setBorderWidthBottom(0f);
					    cell.setBorderColor(blue);
					    cell.setBorderColorBottom(Color.WHITE);
						actionPlantable.addCell(cell);
						subQuestion = (String) innerList.get(5);
						if (!subQuestion.trim().equals("NA")) {
							cell = new PdfPCell(new Phrase("Sub Question: "+(String) innerList.get(5)+"\n", tableContentFont));
							cell.setUseBorderPadding(true);
						    cell.setBorderWidthTop(0f);
						    cell.setBorderWidthBottom(0f);
						    cell.setBorderColor(blue);
						    cell.setBorderColorBottom(Color.WHITE);
							actionPlantable.addCell(cell);
						}
						cell = new PdfPCell(new Phrase("Ans 1: "+(String) innerList.get(6), tableContentFont));
						cell.setUseBorderPadding(true);
					    cell.setBorderWidthTop(0f);
					    cell.setBorderWidthBottom(0f);
					    cell.setBorderColor(blue);
					    cell.setBorderColorBottom(Color.WHITE);
						actionPlantable.addCell(cell);
						cell = new PdfPCell(new Phrase("Ans 2: "+(String) innerList.get(7), tableContentFont));
						cell.setUseBorderPadding(true);
					    cell.setBorderWidthTop(0f);
					    cell.setBorderWidthBottom(0f);
					    cell.setBorderColor(blue);
					    cell.setBorderColorBottom(Color.WHITE);
						actionPlantable.addCell(cell);
						cell = new PdfPCell(new Phrase("Ans 3: "+(String) innerList.get(8)+"\n\n", tableContentFont));
						cell.setUseBorderPadding(true);
					    cell.setBorderWidthTop(0f);
					    cell.setBorderWidthBottom(0f);
					    cell.setBorderColor(blue);
					    cell.setBorderColorBottom(Color.WHITE);
					    actionPlantable.addCell(cell);
					    subQuestion = (String) innerList.get(9);
					    if (!subQuestion.trim().equals("NA")) {
							cell = new PdfPCell(new Phrase("Sub Question: "+(String) innerList.get(9)+"\n", tableContentFont));
							cell.setUseBorderPadding(true);
						    cell.setBorderWidthTop(0f);
						    cell.setBorderWidthBottom(0f);
						    cell.setBorderColor(blue);
						    cell.setBorderColorBottom(Color.WHITE);
							actionPlantable.addCell(cell);
					    }
						cell = new PdfPCell(new Phrase("Ans 1: "+(String) innerList.get(10), tableContentFont));
						cell.setUseBorderPadding(true);
					    cell.setBorderWidthTop(0f);
					    cell.setBorderWidthBottom(0f);
					    cell.setBorderColor(blue);
					    cell.setBorderColorBottom(Color.WHITE);
						actionPlantable.addCell(cell);
						cell = new PdfPCell(new Phrase("Ans 2: "+(String) innerList.get(11), tableContentFont));
						cell.setUseBorderPadding(true);
					    cell.setBorderWidthTop(0f);
					    cell.setBorderWidthBottom(0f);
					    cell.setBorderColor(blue);
					    cell.setBorderColorBottom(Color.WHITE);
						actionPlantable.addCell(cell);
						cell = new PdfPCell(new Phrase("Ans 3: "+(String) innerList.get(12)+"\n\n", tableContentFont));
						cell.setUseBorderPadding(true);
						cell.setBorderColor(blue);
					    cell.setBorderWidthTop(0f);
					    actionPlantable.addCell(cell);
					} else if (listSize == 17) { //4 records
						cell = new PdfPCell(new Phrase("Question: "+(String) innerList.get(0)+"\n\n", tableContentFont));
						cell.setUseBorderPadding(true);
					    cell.setBorderWidthTop(0f);
					    cell.setBorderWidthBottom(0f);
					    cell.setBorderColor(blue);
					    cell.setBorderColorBottom(Color.WHITE);
						actionPlantable.addCell(cell);
						subQuestion = (String) innerList.get(1);
						if (!subQuestion.trim().equals("NA")) {
							cell = new PdfPCell(new Phrase("Sub Question: "+(String) innerList.get(1)+"\n", tableContentFont));
							cell.setUseBorderPadding(true);
						    cell.setBorderWidthTop(0f);
						    cell.setBorderWidthBottom(0f);
						    cell.setBorderColor(blue);
						    cell.setBorderColorBottom(Color.WHITE);
							actionPlantable.addCell(cell);
						}
						cell = new PdfPCell(new Phrase("Ans 1: "+(String) innerList.get(2), tableContentFont));
						cell.setUseBorderPadding(true);
					    cell.setBorderWidthTop(0f);
					    cell.setBorderWidthBottom(0f);
					    cell.setBorderColor(blue);
					    cell.setBorderColorBottom(Color.WHITE);
						actionPlantable.addCell(cell);
						cell = new PdfPCell(new Phrase("Ans 2: "+(String) innerList.get(3), tableContentFont));
						cell.setUseBorderPadding(true);
					    cell.setBorderWidthTop(0f);
					    cell.setBorderWidthBottom(0f);
					    cell.setBorderColor(blue);
					    cell.setBorderColorBottom(Color.WHITE);
						actionPlantable.addCell(cell);
						cell = new PdfPCell(new Phrase("Ans 3: "+(String) innerList.get(4)+"\n\n", tableContentFont));
						cell.setUseBorderPadding(true);
					    cell.setBorderWidthTop(0f);
					    cell.setBorderWidthBottom(0f);
					    cell.setBorderColor(blue);
					    cell.setBorderColorBottom(Color.WHITE);
						actionPlantable.addCell(cell);
						subQuestion = (String) innerList.get(5);
						if (!subQuestion.trim().equals("NA")) {
							cell = new PdfPCell(new Phrase("Sub Question: "+(String) innerList.get(5)+"\n", tableContentFont));
							cell.setUseBorderPadding(true);
						    cell.setBorderWidthTop(0f);
						    cell.setBorderWidthBottom(0f);
						    cell.setBorderColor(blue);
						    cell.setBorderColorBottom(Color.WHITE);
							actionPlantable.addCell(cell);
						}
						cell = new PdfPCell(new Phrase("Ans 1: "+(String) innerList.get(6), tableContentFont));
						cell.setUseBorderPadding(true);
					    cell.setBorderWidthTop(0f);
					    cell.setBorderWidthBottom(0f);
					    cell.setBorderColor(blue);
					    cell.setBorderColorBottom(Color.WHITE);
						actionPlantable.addCell(cell);
						cell = new PdfPCell(new Phrase("Ans 2: "+(String) innerList.get(7), tableContentFont));
						cell.setUseBorderPadding(true);
					    cell.setBorderWidthTop(0f);
					    cell.setBorderWidthBottom(0f);
					    cell.setBorderColor(blue);
					    cell.setBorderColorBottom(Color.WHITE);
						actionPlantable.addCell(cell);
						cell = new PdfPCell(new Phrase("Ans 3: "+(String) innerList.get(8)+"\n\n", tableContentFont));
						cell.setUseBorderPadding(true);
					    cell.setBorderWidthTop(0f);
					    cell.setBorderWidthBottom(0f);
					    cell.setBorderColor(blue);
					    cell.setBorderColorBottom(Color.WHITE);
					    actionPlantable.addCell(cell);
					    subQuestion = (String) innerList.get(9);
					    if (!subQuestion.trim().equals("NA")) {
							cell = new PdfPCell(new Phrase("Sub Question: "+(String) innerList.get(9)+"\n", tableContentFont));
							cell.setUseBorderPadding(true);
						    cell.setBorderWidthTop(0f);
						    cell.setBorderWidthBottom(0f);
						    cell.setBorderColor(blue);
						    cell.setBorderColorBottom(Color.WHITE);
							actionPlantable.addCell(cell);
					    }
						cell = new PdfPCell(new Phrase("Ans 1: "+(String) innerList.get(10), tableContentFont));
						cell.setUseBorderPadding(true);
					    cell.setBorderWidthTop(0f);
					    cell.setBorderWidthBottom(0f);
					    cell.setBorderColor(blue);
					    cell.setBorderColorBottom(Color.WHITE);
						actionPlantable.addCell(cell);
						cell = new PdfPCell(new Phrase("Ans 2: "+(String) innerList.get(11), tableContentFont));
						cell.setUseBorderPadding(true);
					    cell.setBorderWidthTop(0f);
					    cell.setBorderWidthBottom(0f);
					    cell.setBorderColor(blue);
					    cell.setBorderColorBottom(Color.WHITE);
						actionPlantable.addCell(cell);
						cell = new PdfPCell(new Phrase("Ans 3: "+(String) innerList.get(12)+"\n\n", tableContentFont));
						cell.setUseBorderPadding(true);
					    cell.setBorderWidthTop(0f);
					    cell.setBorderWidthBottom(0f);
					    cell.setBorderColor(blue);
					    cell.setBorderColorBottom(Color.WHITE);
						actionPlantable.addCell(cell);
						subQuestion = (String) innerList.get(13);
						if (!subQuestion.trim().equals("NA")) {
							cell = new PdfPCell(new Phrase("Sub Question: "+(String) innerList.get(13)+"\n", tableContentFont));
							cell.setUseBorderPadding(true);
						    cell.setBorderWidthTop(0f);
						    cell.setBorderWidthBottom(0f);
						    cell.setBorderColor(blue);
						    cell.setBorderColorBottom(Color.WHITE);
							actionPlantable.addCell(cell);
						}
						cell = new PdfPCell(new Phrase("Ans 1: "+(String) innerList.get(14), tableContentFont));
						cell.setUseBorderPadding(true);
					    cell.setBorderWidthTop(0f);
					    cell.setBorderWidthBottom(0f);
					    cell.setBorderColor(blue);
					    cell.setBorderColorBottom(Color.WHITE);
						actionPlantable.addCell(cell);
						cell = new PdfPCell(new Phrase("Ans 2: "+(String) innerList.get(15), tableContentFont));
						cell.setUseBorderPadding(true);
					    cell.setBorderWidthTop(0f);
					    cell.setBorderWidthBottom(0f);
					    cell.setBorderColor(blue);
					    cell.setBorderColorBottom(Color.WHITE);
						actionPlantable.addCell(cell);
						cell = new PdfPCell(new Phrase("Ans 3: "+(String) innerList.get(16)+"\n\n", tableContentFont));
						cell.setUseBorderPadding(true);
					    cell.setBorderWidthTop(0f);
					    actionPlantable.addCell(cell);
					} else if (listSize == 21) {//5 records
						cell = new PdfPCell(new Phrase("Question: "+(String) innerList.get(0)+"\n\n", tableContentFont));
						cell.setUseBorderPadding(true);
					    cell.setBorderWidthTop(0f);
					    cell.setBorderWidthBottom(0f);
					    cell.setBorderColor(blue);
					    cell.setBorderColorBottom(Color.WHITE);
						actionPlantable.addCell(cell);
						subQuestion = (String) innerList.get(1);
						if (!subQuestion.trim().equals("NA")) {
							cell = new PdfPCell(new Phrase("Sub Question: "+(String) innerList.get(1)+"\n", tableContentFont));
							cell.setUseBorderPadding(true);
						    cell.setBorderWidthTop(0f);
						    cell.setBorderWidthBottom(0f);
						    cell.setBorderColor(blue);
						    cell.setBorderColorBottom(Color.WHITE);
							actionPlantable.addCell(cell);
						}
						cell = new PdfPCell(new Phrase("Ans 1: "+(String) innerList.get(2), tableContentFont));
						cell.setUseBorderPadding(true);
					    cell.setBorderWidthTop(0f);
					    cell.setBorderWidthBottom(0f);
					    cell.setBorderColor(blue);
					    cell.setBorderColorBottom(Color.WHITE);
						actionPlantable.addCell(cell);
						cell = new PdfPCell(new Phrase("Ans 2: "+(String) innerList.get(3), tableContentFont));
						cell.setUseBorderPadding(true);
					    cell.setBorderWidthTop(0f);
					    cell.setBorderWidthBottom(0f);
					    cell.setBorderColor(blue);
					    cell.setBorderColorBottom(Color.WHITE);
						actionPlantable.addCell(cell);
						cell = new PdfPCell(new Phrase("Ans 3: "+(String) innerList.get(4)+"\n\n", tableContentFont));
						cell.setUseBorderPadding(true);
					    cell.setBorderWidthTop(0f);
					    cell.setBorderWidthBottom(0f);
					    cell.setBorderColor(blue);
					    cell.setBorderColorBottom(Color.WHITE);
						actionPlantable.addCell(cell);
						subQuestion = (String) innerList.get(5);
						if (!subQuestion.trim().equals("NA")) {
							cell = new PdfPCell(new Phrase("Sub Question: "+(String) innerList.get(5)+"\n", tableContentFont));
							cell.setUseBorderPadding(true);
						    cell.setBorderWidthTop(0f);
						    cell.setBorderWidthBottom(0f);
						    cell.setBorderColor(blue);
						    cell.setBorderColorBottom(Color.WHITE);
							actionPlantable.addCell(cell);
						}
						cell = new PdfPCell(new Phrase("Ans 1: "+(String) innerList.get(6), tableContentFont));
						cell.setUseBorderPadding(true);
					    cell.setBorderWidthTop(0f);
					    cell.setBorderWidthBottom(0f);
					    cell.setBorderColor(blue);
					    cell.setBorderColorBottom(Color.WHITE);
						actionPlantable.addCell(cell);
						cell = new PdfPCell(new Phrase("Ans 2: "+(String) innerList.get(7), tableContentFont));
						cell.setUseBorderPadding(true);
					    cell.setBorderWidthTop(0f);
					    cell.setBorderWidthBottom(0f);
					    cell.setBorderColor(blue);
					    cell.setBorderColorBottom(Color.WHITE);
						actionPlantable.addCell(cell);
						cell = new PdfPCell(new Phrase("Ans 3: "+(String) innerList.get(8)+"\n\n", tableContentFont));
						cell.setUseBorderPadding(true);
					    cell.setBorderWidthTop(0f);
					    cell.setBorderWidthBottom(0f);
					    cell.setBorderColor(blue);
					    cell.setBorderColorBottom(Color.WHITE);
					    actionPlantable.addCell(cell);
					    subQuestion = (String) innerList.get(9);
					    if (!subQuestion.trim().equals("NA")) {
							cell = new PdfPCell(new Phrase("Sub Question: "+(String) innerList.get(9)+"\n", tableContentFont));
							cell.setUseBorderPadding(true);
						    cell.setBorderWidthTop(0f);
						    cell.setBorderWidthBottom(0f);
						    cell.setBorderColor(blue);
						    cell.setBorderColorBottom(Color.WHITE);
							actionPlantable.addCell(cell);
					    }
						cell = new PdfPCell(new Phrase("Ans 1: "+(String) innerList.get(10), tableContentFont));
						cell.setUseBorderPadding(true);
					    cell.setBorderWidthTop(0f);
					    cell.setBorderWidthBottom(0f);
					    cell.setBorderColor(blue);
					    cell.setBorderColorBottom(Color.WHITE);
						actionPlantable.addCell(cell);
						cell = new PdfPCell(new Phrase("Ans 2: "+(String) innerList.get(11), tableContentFont));
						cell.setUseBorderPadding(true);
					    cell.setBorderWidthTop(0f);
					    cell.setBorderWidthBottom(0f);
					    cell.setBorderColor(blue);
					    cell.setBorderColorBottom(Color.WHITE);
						actionPlantable.addCell(cell);
						cell = new PdfPCell(new Phrase("Ans 3: "+(String) innerList.get(12)+"\n\n", tableContentFont));
						cell.setUseBorderPadding(true);
					    cell.setBorderWidthTop(0f);
					    cell.setBorderWidthBottom(0f);
					    cell.setBorderColor(blue);
					    cell.setBorderColorBottom(Color.WHITE);
						actionPlantable.addCell(cell);
						subQuestion = (String) innerList.get(13);
						if (!subQuestion.trim().equals("NA")) {
							cell = new PdfPCell(new Phrase("Sub Question: "+(String) innerList.get(13)+"\n", tableContentFont));
							cell.setUseBorderPadding(true);
						    cell.setBorderWidthTop(0f);
						    cell.setBorderWidthBottom(0f);
						    cell.setBorderColor(blue);
						    cell.setBorderColorBottom(Color.WHITE);
							actionPlantable.addCell(cell);
						}
						cell = new PdfPCell(new Phrase("Ans 1: "+(String) innerList.get(14), tableContentFont));
						cell.setUseBorderPadding(true);
					    cell.setBorderWidthTop(0f);
					    cell.setBorderWidthBottom(0f);
					    cell.setBorderColor(blue);
					    cell.setBorderColorBottom(Color.WHITE);
						actionPlantable.addCell(cell);
						cell = new PdfPCell(new Phrase("Ans 2: "+(String) innerList.get(15), tableContentFont));
						cell.setUseBorderPadding(true);
					    cell.setBorderWidthTop(0f);
					    cell.setBorderWidthBottom(0f);
					    cell.setBorderColor(blue);
					    cell.setBorderColorBottom(Color.WHITE);
						actionPlantable.addCell(cell);
						cell = new PdfPCell(new Phrase("Ans 3: "+(String) innerList.get(16)+"\n\n", tableContentFont));
						cell.setUseBorderPadding(true);
					    cell.setBorderWidthTop(0f);
					    cell.setBorderWidthBottom(0f);
					    cell.setBorderColor(blue);
					    cell.setBorderColorBottom(Color.WHITE);
					    actionPlantable.addCell(cell);
					    subQuestion = (String) innerList.get(17);
					    if (!subQuestion.trim().equals("NA")) {
						    cell = new PdfPCell(new Phrase("Sub Question: "+(String) innerList.get(17), tableContentFont));
							cell.setUseBorderPadding(true);
						    cell.setBorderWidthTop(0f);
						    cell.setBorderWidthBottom(0f);
						    cell.setBorderColor(blue);
						    cell.setBorderColorBottom(Color.WHITE);
							actionPlantable.addCell(cell);
					    }
						cell = new PdfPCell(new Phrase("Ans 1: "+(String) innerList.get(18), tableContentFont));
						cell.setUseBorderPadding(true);
					    cell.setBorderWidthTop(0f);
					    cell.setBorderWidthBottom(0f);
					    cell.setBorderColor(blue);
					    cell.setBorderColorBottom(Color.WHITE);
						actionPlantable.addCell(cell);
						cell = new PdfPCell(new Phrase("Ans 2: "+(String) innerList.get(19), tableContentFont));
						cell.setUseBorderPadding(true);
					    cell.setBorderWidthTop(0f);
					    cell.setBorderWidthBottom(0f);
					    cell.setBorderColor(blue);
					    cell.setBorderColorBottom(Color.WHITE);
						actionPlantable.addCell(cell);
						cell = new PdfPCell(new Phrase("Ans 3: "+(String) innerList.get(20)+"\n\n", tableContentFont));
						cell.setUseBorderPadding(true);
					    cell.setBorderWidthTop(0f);
					    cell.setBorderColor(blue);
					    actionPlantable.addCell(cell);
					} else if (listSize == 25) {//6 records - buffer
						cell = new PdfPCell(new Phrase("Question: "+(String) innerList.get(0), tableContentFont));
						cell.setUseBorderPadding(true);
					    cell.setBorderWidthTop(0f);
					    cell.setBorderWidthBottom(0f);
					    cell.setBorderColor(blue);
					    cell.setBorderColorBottom(Color.WHITE);
						actionPlantable.addCell(cell);
						subQuestion = (String) innerList.get(1);
						if (!subQuestion.trim().equals("NA")) {
							cell = new PdfPCell(new Phrase("Sub Question: "+(String) innerList.get(1), tableContentFont));
							cell.setUseBorderPadding(true);
						    cell.setBorderWidthTop(0f);
						    cell.setBorderWidthBottom(0f);
						    cell.setBorderColor(blue);
						    cell.setBorderColorBottom(Color.WHITE);
							actionPlantable.addCell(cell);
						}
						cell = new PdfPCell(new Phrase("Ans 1: "+(String) innerList.get(2), tableContentFont));
						cell.setUseBorderPadding(true);
					    cell.setBorderWidthTop(0f);
					    cell.setBorderWidthBottom(0f);
					    cell.setBorderColor(blue);
					    cell.setBorderColorBottom(Color.WHITE);
						actionPlantable.addCell(cell);
						cell = new PdfPCell(new Phrase("Ans 2: "+(String) innerList.get(3), tableContentFont));
						cell.setUseBorderPadding(true);
					    cell.setBorderWidthTop(0f);
					    cell.setBorderWidthBottom(0f);
					    cell.setBorderColor(blue);
					    cell.setBorderColorBottom(Color.WHITE);
						actionPlantable.addCell(cell);
						cell = new PdfPCell(new Phrase("Ans 3: "+(String) innerList.get(4)+"\n\n", tableContentFont));
						cell.setUseBorderPadding(true);
					    cell.setBorderWidthTop(0f);
					    cell.setBorderWidthBottom(0f);
					    cell.setBorderColor(blue);
					    cell.setBorderColorBottom(Color.WHITE);
						actionPlantable.addCell(cell);
						subQuestion = (String) innerList.get(5);
						if (!subQuestion.trim().equals("NA")) {
							cell = new PdfPCell(new Phrase("Sub Question: "+(String) innerList.get(5), tableContentFont));
							cell.setUseBorderPadding(true);
						    cell.setBorderWidthTop(0f);
						    cell.setBorderWidthBottom(0f);
						    cell.setBorderColor(blue);
						    cell.setBorderColorBottom(Color.WHITE);
							actionPlantable.addCell(cell);
						}
						cell = new PdfPCell(new Phrase("Ans 1: "+(String) innerList.get(6), tableContentFont));
						cell.setUseBorderPadding(true);
					    cell.setBorderWidthTop(0f);
					    cell.setBorderWidthBottom(0f);
					    cell.setBorderColor(blue);
					    cell.setBorderColorBottom(Color.WHITE);
						actionPlantable.addCell(cell);
						cell = new PdfPCell(new Phrase("Ans 2: "+(String) innerList.get(7), tableContentFont));
						cell.setUseBorderPadding(true);
					    cell.setBorderWidthTop(0f);
					    cell.setBorderWidthBottom(0f);
					    cell.setBorderColor(blue);
					    cell.setBorderColorBottom(Color.WHITE);
						actionPlantable.addCell(cell);
						cell = new PdfPCell(new Phrase("Ans 3: "+(String) innerList.get(8)+"\n\n", tableContentFont));
						cell.setUseBorderPadding(true);
					    cell.setBorderWidthTop(0f);
					    cell.setBorderWidthBottom(0f);
					    cell.setBorderColor(blue);
					    cell.setBorderColorBottom(Color.WHITE);
					    actionPlantable.addCell(cell);
					    subQuestion = (String) innerList.get(9);
					    if (!subQuestion.trim().equals("NA")) {
							cell = new PdfPCell(new Phrase("Sub Question: "+(String) innerList.get(9), tableContentFont));
							cell.setUseBorderPadding(true);
						    cell.setBorderWidthTop(0f);
						    cell.setBorderWidthBottom(0f);
						    cell.setBorderColor(blue);
						    cell.setBorderColorBottom(Color.WHITE);
							actionPlantable.addCell(cell);
					    }
						cell = new PdfPCell(new Phrase("Ans 1: "+(String) innerList.get(10), tableContentFont));
						cell.setUseBorderPadding(true);
					    cell.setBorderWidthTop(0f);
					    cell.setBorderWidthBottom(0f);
					    cell.setBorderColor(blue);
					    cell.setBorderColorBottom(Color.WHITE);
						actionPlantable.addCell(cell);
						cell = new PdfPCell(new Phrase("Ans 2: "+(String) innerList.get(11), tableContentFont));
						cell.setUseBorderPadding(true);
					    cell.setBorderWidthTop(0f);
					    cell.setBorderWidthBottom(0f);
					    cell.setBorderColor(blue);
					    cell.setBorderColorBottom(Color.WHITE);
						actionPlantable.addCell(cell);
						cell = new PdfPCell(new Phrase("Ans 3: "+(String) innerList.get(12)+"\n\n", tableContentFont));
						cell.setUseBorderPadding(true);
					    cell.setBorderWidthTop(0f);
					    cell.setBorderWidthBottom(0f);
					    cell.setBorderColor(blue);
					    cell.setBorderColorBottom(Color.WHITE);
						actionPlantable.addCell(cell);
						subQuestion = (String) innerList.get(13);
						if (!subQuestion.trim().equals("NA")) {
							cell = new PdfPCell(new Phrase("Sub Question: "+(String) innerList.get(13), tableContentFont));
							cell.setUseBorderPadding(true);
						    cell.setBorderWidthTop(0f);
						    cell.setBorderWidthBottom(0f);
						    cell.setBorderColor(blue);
						    cell.setBorderColorBottom(Color.WHITE);
							actionPlantable.addCell(cell);
						}
						cell = new PdfPCell(new Phrase("Ans 1: "+(String) innerList.get(14), tableContentFont));
						cell.setUseBorderPadding(true);
					    cell.setBorderWidthTop(0f);
					    cell.setBorderWidthBottom(0f);
					    cell.setBorderColor(blue);
					    cell.setBorderColorBottom(Color.WHITE);
						actionPlantable.addCell(cell);
						cell = new PdfPCell(new Phrase("Ans 2: "+(String) innerList.get(15), tableContentFont));
						cell.setUseBorderPadding(true);
					    cell.setBorderWidthTop(0f);
					    cell.setBorderWidthBottom(0f);
					    cell.setBorderColor(blue);
					    cell.setBorderColorBottom(Color.WHITE);
						actionPlantable.addCell(cell);
						cell = new PdfPCell(new Phrase("Ans 3: "+(String) innerList.get(16)+"\n\n", tableContentFont));
						cell.setUseBorderPadding(true);
					    cell.setBorderWidthTop(0f);
					    cell.setBorderWidthBottom(0f);
					    cell.setBorderColor(blue);
					    cell.setBorderColorBottom(Color.WHITE);
					    actionPlantable.addCell(cell);
					    subQuestion = (String) innerList.get(17);
					    if (!subQuestion.trim().equals("NA")) {
						    cell = new PdfPCell(new Phrase("Sub Question: "+(String) innerList.get(17), tableContentFont));
							cell.setUseBorderPadding(true);
						    cell.setBorderWidthTop(0f);
						    cell.setBorderWidthBottom(0f);
						    cell.setBorderColor(blue);
						    cell.setBorderColorBottom(Color.WHITE);
							actionPlantable.addCell(cell);
					    }
						cell = new PdfPCell(new Phrase("Ans 1: "+(String) innerList.get(18), tableContentFont));
						cell.setUseBorderPadding(true);
					    cell.setBorderWidthTop(0f);
					    cell.setBorderWidthBottom(0f);
					    cell.setBorderColor(blue);
					    cell.setBorderColorBottom(Color.WHITE);
						actionPlantable.addCell(cell);
						cell = new PdfPCell(new Phrase("Ans 2: "+(String) innerList.get(19), tableContentFont));
						cell.setUseBorderPadding(true);
					    cell.setBorderWidthTop(0f);
					    cell.setBorderWidthBottom(0f);
					    cell.setBorderColor(blue);
					    cell.setBorderColorBottom(Color.WHITE);
						actionPlantable.addCell(cell);
						cell = new PdfPCell(new Phrase("Ans 3: "+(String) innerList.get(20)+"\n\n", tableContentFont));
						cell.setUseBorderPadding(true);
					    cell.setBorderWidthTop(0f);
					    cell.setBorderWidthBottom(0f);
					    cell.setBorderColor(blue);
					    cell.setBorderColorBottom(Color.WHITE);
						actionPlantable.addCell(cell);
						subQuestion = (String) innerList.get(21);
						if (!subQuestion.trim().equals("NA")) {
					    cell = new PdfPCell(new Phrase("Sub Question: "+(String) innerList.get(21), tableContentFont));
							cell.setUseBorderPadding(true);
						    cell.setBorderWidthTop(0f);
						    cell.setBorderWidthBottom(0f);
						    cell.setBorderColor(blue);
						    cell.setBorderColorBottom(Color.WHITE);
							actionPlantable.addCell(cell);
						}
						cell = new PdfPCell(new Phrase("Ans 1: "+(String) innerList.get(22), tableContentFont));
						cell.setUseBorderPadding(true);
					    cell.setBorderWidthTop(0f);
					    cell.setBorderWidthBottom(0f);
					    cell.setBorderColor(blue);
					    cell.setBorderColorBottom(Color.WHITE);
						actionPlantable.addCell(cell);
						cell = new PdfPCell(new Phrase("Ans 2: "+(String) innerList.get(23), tableContentFont));
						cell.setUseBorderPadding(true);
					    cell.setBorderWidthTop(0f);
					    cell.setBorderWidthBottom(0f);
					    cell.setBorderColor(blue);
					    cell.setBorderColorBottom(Color.WHITE);
						actionPlantable.addCell(cell);
						cell = new PdfPCell(new Phrase("Ans 3: "+(String) innerList.get(24)+"\n\n", tableContentFont));
						cell.setUseBorderPadding(true);
						cell.setBorderColor(blue);
					    cell.setBorderWidthTop(0f);
					    actionPlantable.addCell(cell);
					} else if (listSize == 29) {//7 records - buffer
						cell = new PdfPCell(new Phrase("Question: "+(String) innerList.get(0), tableContentFont));
						cell.setUseBorderPadding(true);
					    cell.setBorderWidthTop(0f);
					    cell.setBorderWidthBottom(0f);
					    cell.setBorderColor(blue);
					    cell.setBorderColorBottom(Color.WHITE);
						actionPlantable.addCell(cell);
						subQuestion = (String) innerList.get(1);
						if (!subQuestion.trim().equals("NA")) {
							cell = new PdfPCell(new Phrase("Sub Question: "+(String) innerList.get(1), tableContentFont));
							cell.setUseBorderPadding(true);
						    cell.setBorderWidthTop(0f);
						    cell.setBorderWidthBottom(0f);
						    cell.setBorderColor(blue);
						    cell.setBorderColorBottom(Color.WHITE);
							actionPlantable.addCell(cell);
						}
						cell = new PdfPCell(new Phrase("Ans 1: "+(String) innerList.get(2), tableContentFont));
						cell.setUseBorderPadding(true);
					    cell.setBorderWidthTop(0f);
					    cell.setBorderWidthBottom(0f);
					    cell.setBorderColor(blue);
					    cell.setBorderColorBottom(Color.WHITE);
						actionPlantable.addCell(cell);
						cell = new PdfPCell(new Phrase("Ans 2: "+(String) innerList.get(3), tableContentFont));
						cell.setUseBorderPadding(true);
					    cell.setBorderWidthTop(0f);
					    cell.setBorderWidthBottom(0f);
					    cell.setBorderColor(blue);
					    cell.setBorderColorBottom(Color.WHITE);
						actionPlantable.addCell(cell);
						cell = new PdfPCell(new Phrase("Ans 3: "+(String) innerList.get(4)+"\n\n", tableContentFont));
						cell.setUseBorderPadding(true);
					    cell.setBorderWidthTop(0f);
					    cell.setBorderWidthBottom(0f);
					    cell.setBorderColor(blue);
					    cell.setBorderColorBottom(Color.WHITE);
						actionPlantable.addCell(cell);
						subQuestion = (String) innerList.get(5);
						if (!subQuestion.trim().equals("NA")) {
							cell = new PdfPCell(new Phrase("Sub Question: "+(String) innerList.get(5), tableContentFont));
							cell.setUseBorderPadding(true);
						    cell.setBorderWidthTop(0f);
						    cell.setBorderWidthBottom(0f);
						    cell.setBorderColor(blue);
						    cell.setBorderColorBottom(Color.WHITE);
							actionPlantable.addCell(cell);
						}
						cell = new PdfPCell(new Phrase("Ans 1: "+(String) innerList.get(6), tableContentFont));
						cell.setUseBorderPadding(true);
					    cell.setBorderWidthTop(0f);
					    cell.setBorderWidthBottom(0f);
					    cell.setBorderColor(blue);
					    cell.setBorderColorBottom(Color.WHITE);
						actionPlantable.addCell(cell);
						cell = new PdfPCell(new Phrase("Ans 2: "+(String) innerList.get(7), tableContentFont));
						cell.setUseBorderPadding(true);
					    cell.setBorderWidthTop(0f);
					    cell.setBorderWidthBottom(0f);
					    cell.setBorderColor(blue);
					    cell.setBorderColorBottom(Color.WHITE);
						actionPlantable.addCell(cell);
						cell = new PdfPCell(new Phrase("Ans 3: "+(String) innerList.get(8)+"\n\n", tableContentFont));
						cell.setUseBorderPadding(true);
					    cell.setBorderWidthTop(0f);
					    cell.setBorderWidthBottom(0f);
					    cell.setBorderColor(blue);
					    cell.setBorderColorBottom(Color.WHITE);
					    actionPlantable.addCell(cell);
					    subQuestion = (String) innerList.get(9);
					    if (!subQuestion.trim().equals("NA")) {
							cell = new PdfPCell(new Phrase("Sub Question: "+(String) innerList.get(9), tableContentFont));
							cell.setUseBorderPadding(true);
						    cell.setBorderWidthTop(0f);
						    cell.setBorderWidthBottom(0f);
						    cell.setBorderColor(blue);
						    cell.setBorderColorBottom(Color.WHITE);
							actionPlantable.addCell(cell);
					    }
						cell = new PdfPCell(new Phrase("Ans 1: "+(String) innerList.get(10), tableContentFont));
						cell.setUseBorderPadding(true);
					    cell.setBorderWidthTop(0f);
					    cell.setBorderWidthBottom(0f);
					    cell.setBorderColor(blue);
					    cell.setBorderColorBottom(Color.WHITE);
						actionPlantable.addCell(cell);
						cell = new PdfPCell(new Phrase("Ans 2: "+(String) innerList.get(11), tableContentFont));
						cell.setUseBorderPadding(true);
					    cell.setBorderWidthTop(0f);
					    cell.setBorderWidthBottom(0f);
					    cell.setBorderColor(blue);
					    cell.setBorderColorBottom(Color.WHITE);
						actionPlantable.addCell(cell);
						cell = new PdfPCell(new Phrase("Ans 3: "+(String) innerList.get(12)+"\n\n", tableContentFont));
						cell.setUseBorderPadding(true);
					    cell.setBorderWidthTop(0f);
					    cell.setBorderWidthBottom(0f);
					    cell.setBorderColor(blue);
					    cell.setBorderColorBottom(Color.WHITE);
						actionPlantable.addCell(cell);
						subQuestion = (String) innerList.get(13);
						if (!subQuestion.trim().equals("NA")) {
							cell = new PdfPCell(new Phrase("Sub Question: "+(String) innerList.get(13), tableContentFont));
							cell.setUseBorderPadding(true);
						    cell.setBorderWidthTop(0f);
						    cell.setBorderWidthBottom(0f);
						    cell.setBorderColor(blue);
						    cell.setBorderColorBottom(Color.WHITE);
							actionPlantable.addCell(cell);
						}
						cell = new PdfPCell(new Phrase("Ans 1: "+(String) innerList.get(14), tableContentFont));
						cell.setUseBorderPadding(true);
					    cell.setBorderWidthTop(0f);
					    cell.setBorderWidthBottom(0f);
					    cell.setBorderColor(blue);
					    cell.setBorderColorBottom(Color.WHITE);
						actionPlantable.addCell(cell);
						cell = new PdfPCell(new Phrase("Ans 2: "+(String) innerList.get(15), tableContentFont));
						cell.setUseBorderPadding(true);
					    cell.setBorderWidthTop(0f);
					    cell.setBorderWidthBottom(0f);
					    cell.setBorderColor(blue);
					    cell.setBorderColorBottom(Color.WHITE);
						actionPlantable.addCell(cell);
						cell = new PdfPCell(new Phrase("Ans 3: "+(String) innerList.get(16)+"\n\n", tableContentFont));
						cell.setUseBorderPadding(true);
					    cell.setBorderWidthTop(0f);
					    cell.setBorderWidthBottom(0f);
					    cell.setBorderColor(blue);
					    cell.setBorderColorBottom(Color.WHITE);
					    actionPlantable.addCell(cell);
					    subQuestion = (String) innerList.get(17);
					    if (!subQuestion.trim().equals("NA")) {
						    cell = new PdfPCell(new Phrase("Sub Question: "+(String) innerList.get(17), tableContentFont));
							cell.setUseBorderPadding(true);
						    cell.setBorderWidthTop(0f);
						    cell.setBorderWidthBottom(0f);
						    cell.setBorderColor(blue);
						    cell.setBorderColorBottom(Color.WHITE);
							actionPlantable.addCell(cell);
					    }
						cell = new PdfPCell(new Phrase("Ans 1: "+(String) innerList.get(18), tableContentFont));
						cell.setUseBorderPadding(true);
					    cell.setBorderWidthTop(0f);
					    cell.setBorderWidthBottom(0f);
					    cell.setBorderColor(blue);
					    cell.setBorderColorBottom(Color.WHITE);
						actionPlantable.addCell(cell);
						cell = new PdfPCell(new Phrase("Ans 2: "+(String) innerList.get(19), tableContentFont));
						cell.setUseBorderPadding(true);
					    cell.setBorderWidthTop(0f);
					    cell.setBorderWidthBottom(0f);
					    cell.setBorderColor(blue);
					    cell.setBorderColorBottom(Color.WHITE);
						actionPlantable.addCell(cell);
						cell = new PdfPCell(new Phrase("Ans 3: "+(String) innerList.get(20)+"\n\n", tableContentFont));
						cell.setUseBorderPadding(true);
					    cell.setBorderWidthTop(0f);
					    cell.setBorderWidthBottom(0f);
					    cell.setBorderColor(blue);
					    cell.setBorderColorBottom(Color.WHITE);
						actionPlantable.addCell(cell);
						subQuestion = (String) innerList.get(21);
						if (!subQuestion.trim().equals("NA")) {
						    cell = new PdfPCell(new Phrase("Sub Question: "+(String) innerList.get(21), tableContentFont));
							cell.setUseBorderPadding(true);
						    cell.setBorderWidthTop(0f);
						    cell.setBorderWidthBottom(0f);
						    cell.setBorderColor(blue);
						    cell.setBorderColorBottom(Color.WHITE);
							actionPlantable.addCell(cell);
						}
						cell = new PdfPCell(new Phrase("Ans 1: "+(String) innerList.get(22), tableContentFont));
						cell.setUseBorderPadding(true);
					    cell.setBorderWidthTop(0f);
					    cell.setBorderWidthBottom(0f);
					    cell.setBorderColor(blue);
					    cell.setBorderColorBottom(Color.WHITE);
						actionPlantable.addCell(cell);
						cell = new PdfPCell(new Phrase("Ans 2: "+(String) innerList.get(23), tableContentFont));
						cell.setUseBorderPadding(true);
					    cell.setBorderWidthTop(0f);
					    cell.setBorderWidthBottom(0f);
					    cell.setBorderColor(blue);
					    cell.setBorderColorBottom(Color.WHITE);
						actionPlantable.addCell(cell);
						cell = new PdfPCell(new Phrase("Ans 3: "+(String) innerList.get(24)+"\n\n", tableContentFont));
						cell.setUseBorderPadding(true);
					    cell.setBorderWidthTop(0f);
					    cell.setBorderWidthBottom(0f);
					    cell.setBorderColor(blue);
					    cell.setBorderColorBottom(Color.WHITE);
					    actionPlantable.addCell(cell);
					    subQuestion = (String) innerList.get(25);
					    if (!subQuestion.trim().equals("NA")) {
						    cell = new PdfPCell(new Phrase("Sub Question: "+(String) innerList.get(25), tableContentFont));
							cell.setUseBorderPadding(true);
						    cell.setBorderWidthTop(0f);
						    cell.setBorderWidthBottom(0f);
						    cell.setBorderColor(blue);
						    cell.setBorderColorBottom(Color.WHITE);
							actionPlantable.addCell(cell);
					    }
						cell = new PdfPCell(new Phrase("Ans 1: "+(String) innerList.get(26), tableContentFont));
						cell.setUseBorderPadding(true);
					    cell.setBorderWidthTop(0f);
					    cell.setBorderWidthBottom(0f);
					    cell.setBorderColor(blue);
					    cell.setBorderColorBottom(Color.WHITE);
						actionPlantable.addCell(cell);
						cell = new PdfPCell(new Phrase("Ans 2: "+(String) innerList.get(27), tableContentFont));
						cell.setUseBorderPadding(true);
					    cell.setBorderWidthTop(0f);
					    cell.setBorderWidthBottom(0f);
					    cell.setBorderColor(blue);
					    cell.setBorderColorBottom(Color.WHITE);
						actionPlantable.addCell(cell);
						cell = new PdfPCell(new Phrase("Ans 3: "+(String) innerList.get(28)+"\n\n", tableContentFont));
						cell.setUseBorderPadding(true);
					    cell.setBorderWidthTop(0f);
					    actionPlantable.addCell(cell);
					}
				}
				document.add(actionPlantable);
			}			
			
			//Page 5 of 5
			PdfPTable footerTable5 = new PdfPTable(1); //one column
			footerTable5.setWidthPercentage(90);
			footerTable5.getDefaultCell().setBorder(0);
			
			//Left aligned
			PdfPCell headerCell = new PdfPCell(new Phrase("Page 5 of 5", footerFont));				
			headerCell.setBorder(0);
			headerCell.setHorizontalAlignment(Element.ALIGN_CENTER);
			headerCell.setVerticalAlignment(Element.ALIGN_BOTTOM);				
			footerTable5.addCell(headerCell);
			document.add(footerTable5);
			
			document.close();
			
			//Now save the reference of the pdf in the db
			savePdfReference(jobReferenceNumber, file, diskPath, jctUserId);
		} catch (DocumentException e) {
			LOGGER.error("DOCUMENT EXCEPTION: "+e.getLocalizedMessage());
			throw new JCTException("DOCUMENT EXCEPTION: "+e.getLocalizedMessage());
		} catch (DAOException e) {
			LOGGER.error("DAO EXCEPTION: "+e.getLocalizedMessage());
			throw new JCTException("DAO EXCEPTION: "+e.getLocalizedMessage());
		} catch (MalformedURLException e) {
			LOGGER.error("MALFORMED URL EXCEPTION: "+e.getLocalizedMessage());
			throw new JCTException("MALFORMED URL EXCEPTION: "+e.getLocalizedMessage());
		} catch (IOException e) {
			LOGGER.error("IO EXCEPTION: "+e.getLocalizedMessage());
			throw new JCTException("IO EXCEPTION: "+e.getLocalizedMessage());
		}
		LOGGER.info("<<<< GeneratePDFServiceImpl.generatePdf");
	    return null;
	}
	/**
	 * Method returns the after sketch header row
	 * @param jobReferenceNumber
	 * @return
	 */
	private List<JctAfterSketchHeader> getJctAfterSketchHeader(
			String jobReferenceNumber) {
		LOGGER.info(">>>> GeneratePDFServiceImpl.getJctAfterSketchHeader");
		List<JctAfterSketchHeader> list = null;
		try {
			list = afterSketchDAO.getList(jobReferenceNumber, "N");
		} catch (DAOException e) {
			LOGGER.error(e.getLocalizedMessage());
			//e.printStackTrace();
		}
		LOGGER.info("<<<< GeneratePDFServiceImpl.getJctAfterSketchHeader");
		return list;
	}
	/**
	 * Method saves the pdf path and file info to the table.
	 * This method do not store the pdf bytep[] to the database.
	 * For actual pdf location, consult the properties file.
	 * @param jobReferenceNo
	 * @param fileName
	 * @param fileLocation
	 * @throws JCTException
	 * @throws DAOException
	 */
	private void savePdfReference(String jobReferenceNo, String fileName, 
			String fileLocation, int jctUserId) throws JCTException, DAOException {
		LOGGER.info(">>>> GeneratePDFServiceImpl.savePdfReference");
		JctPdfRecords records = new JctPdfRecords();
		//records.setJctFileId(commonDaoImpl.generateKey("jct_pdf_records"));
		/** THIS IS FOR PUBLIC VERSION **/
		JctUser user = new JctUser();
		user.setJctUserId(jctUserId);
		records.setJctUser(user);
		/********************************/
		records.setJctJobReferenceNo(jobReferenceNo);
		records.setJctFileName(fileName);
		records.setJctCreatedTimestamp(new Date());
		records.setJctCreatedBy("SELF");
		records.setJctFileLocation(fileLocation);
		recordsDAO.savePDFReference(records);
		LOGGER.info("<<<< GeneratePDFServiceImpl.savePdfReference");
	}
	/**
	 * Method fetches question and answer list
	 * @param jobRefNo
	 * @return
	 * @throws JCTException
	 */
	private List<List> fetchQuestionAnswer(String jobRefNo, int profileId) throws JCTException {
		LOGGER.info(">>>> GeneratePDFServiceImpl.fetchQuestionAnswer");
		//This will be returned
		List<List> mainList = null;
		try{
			List<JctBeforeSketchHeader> list = beforeSketchDAO.getList(jobRefNo);
			if (list.size() == 0) {
				list = beforeSketchDAO.getListEdtd(jobRefNo);
			}
			if(list.size()!=0){
					//JctBeforeSketchHeader parent = (JctBeforeSketchHeader)list.get(0);
				//Check if new or existing
					//List<JctBeforeSketchQuestion> actionPlanlist = questionnaireDAO.getList(jobRefNo);
					mainList = new ArrayList<List>();
					// [[Main Qtn1 subQtn Ans subQtn Ans] [Main Qtn2 subQtn Ans subQtn Ans] .....]
					// Select Distinct Main question List
					List<String> questionnaireMainQtnLst = questionnaireDAO.getDistinctQuestionList(jobRefNo);
					Iterator<String> questionnaireMainQtnLstItr = questionnaireMainQtnLst.iterator();
					while (questionnaireMainQtnLstItr.hasNext()) {
						List<String> innerList = new ArrayList<String>();
						String mainQuestion = (String) questionnaireMainQtnLstItr.next();
						innerList.add(mainQuestion);
						// Select the sub question List 
						List<String> subQtnList = questionnaireDAO.getDistinctSubQuestionListByJrno(mainQuestion, jobRefNo);
						Iterator<String> subQtnListItr = subQtnList.iterator();
						while (subQtnListItr.hasNext()) {
							String subQuestion = (String) subQtnListItr.next();
							innerList.add(subQuestion);
							//Fetch answers given
							List<String> ansList = questionnaireDAO.getAnswerListByJrno(mainQuestion, subQuestion, jobRefNo);
							innerList.add((String) ansList.get(0));
						}
						mainList.add(innerList);
					}
			} else {
				LOGGER.error("------> SOMETHING IS MISSING IN DATABASE: JctBeforeSketchHeader. NO RECORD FOUND WITH GIVEN JOB REFERENCE NUMBER: "+jobRefNo);
			}
		} catch(DAOException daoException){
			LOGGER.error("----"+daoException.getLocalizedMessage()+" ----");
		}
		LOGGER.info("<<<< GeneratePDFServiceImpl.fetchQuestionAnswer");
		return mainList;
	}
	/**
	 * Method fetches question listss
	 * @param jobRefNo
	 * @return
	 * @throws JCTException
	 */
	private List<JctQuestion> fetchQuestion(int profileId) throws JCTException {
		LOGGER.info(">>>> GeneratePDFServiceImpl.fetchQuestion");
		List<JctBeforeSketchHeader> list=null;
		JctBeforeSketchHeader parent = null;
		List<JctQuestion> jctQuestionList=null;
		try{
			//list = beforeSketchDAO.getList(jobRefNo);
			//if(list.size()!=0){
			//	parent = (JctBeforeSketchHeader)list.get(0);
				jctQuestionList=beforeSketchDAO.getFetchQuestion(profileId);
			/*}else{
				
			}*/
			
		}catch(DAOException daoException){
			LOGGER.error(daoException.getLocalizedMessage());
			//daoException.printStackTrace();
		}
		LOGGER.info("<<<< GeneratePDFServiceImpl.fetchQuestion");
		return jctQuestionList;
	}
	
	private List<List> fetchActionPlans(String jobRefNo, int profileId) throws JCTException {
		LOGGER.info(">>>> GeneratePDFServiceImpl.fetchActionPlans");
		//This will be returned
		List<List> mainList = null;
		try{
			List<JctBeforeSketchHeader> list = beforeSketchDAO.getList(jobRefNo);
			if (list.size() == 0) {
				list = beforeSketchDAO.getListEdtd(jobRefNo);
			}
			if(list.size()!=0) {
				//Check if new or existing
				List<JctActionPlan> actionPlanlist = actionPlanDAO.getList(jobRefNo);
				if (actionPlanlist.size() == 0) {
					actionPlanlist = actionPlanDAO.getListEdited(jobRefNo);
				}
				if(actionPlanlist.size() > 0){
					mainList = new ArrayList<List>();
					// [[Main Qtn1 subQtn Ans subQtn Ans] [Main Qtn2 subQtn Ans subQtn Ans] .....]
					// Select Distinct Main question List
					List<String> actionPlanMainQtnLst = actionPlanDAO.getDistinctQuestionList(jobRefNo);
					Iterator<String> actionPlanMainQtnLstItr = actionPlanMainQtnLst.iterator();
					while (actionPlanMainQtnLstItr.hasNext()) {
						List<String> innerList = new ArrayList<String>();
						String mainQuestion = (String) actionPlanMainQtnLstItr.next();
						innerList.add(mainQuestion);
						// Select the sub question List 
						List<String> subQtnList = actionPlanDAO.getDistinctSubQuestionListByJrno(mainQuestion, jobRefNo);
						
						Iterator<String> subQtnListItr = subQtnList.iterator();
						while (subQtnListItr.hasNext()) {
							String subQuestion = (String) subQtnListItr.next();
							innerList.add(subQuestion);
							//Fetch answers given
							List<String> ansList = actionPlanDAO.getAnswerListByJrno(mainQuestion, subQuestion, jobRefNo);
							innerList.add((String) ansList.get(0));
							innerList.add((String) ansList.get(1));
							innerList.add((String) ansList.get(2));
						}
						mainList.add(innerList);
					}
				}
			} else {
				LOGGER.error("------> SOMETHING IS MISSING IN DATABASE: JctBeforeSketchHeader. NO RECORD FOUND WITH GIVEN JOB REFERENCE NUMBER: "+jobRefNo);
			}
		} catch(DAOException daoException){
			LOGGER.error("----"+daoException.getLocalizedMessage()+" ----");
		}
		LOGGER.info("<<<< GeneratePDFServiceImpl.fetchActionPlans");
		return mainList;
	}
	
	public String getPathAndFile(String jobReferenceNo) throws JCTException {
		LOGGER.info(">>>> GeneratePDFServiceImpl.getPathAndFile");
		StringBuilder sb = new StringBuilder("");
		try {
			JctPdfRecords record = recordsDAO.getMaxDocRow(jobReferenceNo);
			sb.append(record.getJctFileName());
		} catch (DAOException e) {
			LOGGER.error("------------> UNABLE TO FIND THE PDF..... THE JOB REFERENCE NUMBER SEARCHED WAS :"+jobReferenceNo+". <-------------------");
			//e.printStackTrace();
			throw new JCTException();
		}
		LOGGER.info("<<<< GeneratePDFServiceImpl.getPathAndFile");
		return sb.toString();
	}
	
	private JFreeChart generatePieChart(List<JctAfterSketchFinalpageDetail> dtlsList){
		LOGGER.info(">>>> GeneratePDFServiceImpl.generatePieChart");
		ArrayList<Integer> dtlsIdBackupList = new ArrayList<Integer>();
		DefaultPieDataset dataSet = new DefaultPieDataset();
		Iterator<JctAfterSketchFinalpageDetail> itr = dtlsList.iterator();
		int taskCounter = 1;
		while(itr.hasNext()){
			JctAfterSketchFinalpageDetail dtls = (JctAfterSketchFinalpageDetail)itr.next();
			if(!dtls.getJctAsElementCode().equals(CommonConstants.JOB_ATTRIBUTE_STRENGTH) && 
					!dtls.getJctAsElementCode().equals(CommonConstants.JOB_ATTRIBUTE_PASSION) && 
					!dtls.getJctAsElementCode().equals(CommonConstants.JOB_ATTRIBUTE_VALUE)) {
				if (!dtlsIdBackupList.contains(dtls.getJctAsTaskId())) {
					//dataSet.setValue("TASK "+taskCounter, dtls.getJctAsTaskEnergy());
					String task = dtls.getJctAsTaskDesc();
					if (task.length() > 13) {
						task = task.substring(0, 12) + "...";
					}
					dataSet.setValue("TASK "+taskCounter+":"+task, dtls.getJctAsTaskEnergy());
					taskCounter++;
					dtlsIdBackupList.add(dtls.getJctAsTaskId());
				}
				
			}
		}
		//JFreeChart chart = ChartFactory.createPieChart("", dataSet, true, true, false);
		JFreeChart chart = ChartFactory.createPieChart("", dataSet, false, true, false);
		PiePlot plot = (PiePlot) chart.getPlot();
		plot.setLabelLinkStyle(PieLabelLinkStyle.QUAD_CURVE);
		plot.setLabelFont(new java.awt.Font("HELVETICA", java.awt.Font.BOLD, 9));
		//plot.setSimpleLabels(true);
        
		List <Comparable> keys = dataSet.getKeys(); 
        int aInt; 
        for (int i = 0; i < keys.size(); i++) { 
            aInt = i % this.TASK_PIE_COLORS.length; 
            plot.setSectionPaint(keys.get(i), this.TASK_PIE_COLORS[aInt]); 
        }
        PieSectionLabelGenerator gen = new StandardPieSectionLabelGenerator(
                "{0} ({1}%)", new DecimalFormat("0"), new DecimalFormat("0%"));
            plot.setLabelGenerator(gen);
        chart.setBackgroundPaint(Color.WHITE);
        
        //chart.getLegend().setItemFont(new java.awt.Font("HELVETICA", java.awt.Font.BOLD, 9));
        LOGGER.info("<<<< GeneratePDFServiceImpl.generatePieChart");
        return chart;
	}
		
	private Map<String, TaskDetailsMappingVO> getTaskDescMapping(List<JctAfterSketchFinalpageDetail> dtlsList) throws DAOException {
		LOGGER.info(">>>> GeneratePDFServiceImpl.getTaskDescMapping");
		ArrayList<Integer> dtlsIdBackupList = new ArrayList<Integer>();
		Map<String, TaskDetailsMappingVO> mapping = new LinkedHashMap<String, TaskDetailsMappingVO>();
		Iterator<JctAfterSketchFinalpageDetail> itr = dtlsList.iterator();
		List<Object> objLst = null;
		int taskCounter = 1;
		while (itr.hasNext()) {
			JctAfterSketchFinalpageDetail dtls = (JctAfterSketchFinalpageDetail)itr.next();
			if (!dtls.getJctAsElementCode().equals(CommonConstants.JOB_ATTRIBUTE_STRENGTH) && 
					!dtls.getJctAsElementCode().equals(CommonConstants.JOB_ATTRIBUTE_PASSION) && 
					!dtls.getJctAsElementCode().equals(CommonConstants.JOB_ATTRIBUTE_VALUE)) {
					if (!dtlsIdBackupList.contains(dtls.getJctAsTaskId())) {
						 	
					String jobRefNo = dtls.getJctAsJobrefNo();
					int taskId = dtls.getJctAsTaskId();
					
					List<String> roleFramesList = afterSketchDAO.getDistinctRoleFrames(jobRefNo);
				    Map<Integer, String> roleMap = new HashMap<Integer, String>();
				    int counter = 1;
				    for (int index = 0; index < roleFramesList.size(); index++) {
				      	roleMap.put(counter, roleFramesList.get(index).toString());
				       	counter = counter + 1;
				    }
					
				    Map<String, String> finalRole = new HashMap<String, String>();
					for (Map.Entry<Integer, String> entry : roleMap.entrySet()) {
					    int key = (Integer) entry.getKey();
					    
					    //entry.getValue()
					    
						String roleDist = "";
						if (key == 1) {
							if (entry.getValue().toString().equals("N/A")) {
								roleDist = "--";
							} else {
								roleDist = "A";
							}
						}
						if (key == 2) {
							if (entry.getValue().toString().equals("N/A")) {
								roleDist = "--";
							} else {
								roleDist = "B";
							}
						} else if (key == 3) {
							if (entry.getValue().toString().equals("N/A")) {
								roleDist = "--";
							} else {
								roleDist = "C";
							}
						} else if (key == 4) {
							if (entry.getValue().toString().equals("N/A")) {
								roleDist = "--";
							} else {
								roleDist = "D";
							}
						}
						finalRole.put(entry.getValue().toString(), roleDist);
					}
				    
					//Get the role Frames
					StringBuffer roleFrameBuff = new StringBuffer("");
					List<String> roleFrameList = afterSketchDAO.getRoleFrames(taskId, jobRefNo);
					Iterator<String> roleItr = roleFrameList.iterator();
					//int indexLoop = 1;
					while (roleItr.hasNext()) {
						roleFrameBuff.append((String) roleItr.next()+"```~``~~``````-_-");
					}
					String patternString = "(```~``~~``````-_-)";
					Pattern pattern = Pattern.compile(patternString);
					objLst = afterSketchDAO.getBsToAsDifference(jobRefNo, taskId);
					for (int index = 0; index < objLst.size(); index++ ) {
						Object[] innerObj = (Object[])objLst.get(index);
						int decider = (Integer)innerObj[1]; //flag denoting change
						String relationalVal = dtls.getJctAsAdditionalDesc();
						int asTaskEnergy = dtls.getJctAsTaskEnergy();
						if (decider == 1) { //Append to the desc
							//diffBuilder
							if(Integer.parseInt((String)innerObj[0]) != 0) {
								String value = (String)innerObj[0];
								String displayValue = value.substring(1, value.length());
								if (Integer.parseInt((String)innerObj[0]) < 0) {
									String[] elementToken = pattern.split(roleFrameBuff.toString());
									StringBuffer sb = new StringBuffer("");
									//for ( int innerIndex = 0; innerIndex < elementToken.length; innerIndex++ ) {
										TaskDetailsMappingVO vo = new TaskDetailsMappingVO();
										vo.setTimeEnergyChangeFromBeforeSketch(asTaskEnergy+"% \n("+displayValue+"% Decrease)");
										vo.setRelationalCrafting(relationalVal);
										vo.setTaskCrafting(dtls.getJctAsTaskDesc());
										//vo.setRoleDesc(dtls.getJctAsRoleDesc());
										
										for ( int innerIndex = 0; innerIndex < elementToken.length; innerIndex++ ) {
											sb.append(finalRole.get(elementToken[innerIndex].toString()).toString());
											int lppCount = innerIndex;
											if ((lppCount + 1) < (elementToken.length)) {
												sb.append(" , ");
											}
										}
										vo.setRoleDesc(sb.toString());
										mapping.put(taskCounter+"", vo);
										//taskCounter++;	
									//}
									
									
								} else if (Integer.parseInt((String)innerObj[0]) > 0) {
									String[] elementToken = pattern.split(roleFrameBuff.toString());
									//for ( int innerIndex = 0; innerIndex < elementToken.length; innerIndex++ ) {
										TaskDetailsMappingVO vo = new TaskDetailsMappingVO();
										vo.setTimeEnergyChangeFromBeforeSketch(asTaskEnergy+"% \n("+value+"% Increase)");
										vo.setRelationalCrafting(relationalVal);
										vo.setTaskCrafting(dtls.getJctAsTaskDesc());
										//vo.setRoleDesc(dtls.getJctAsRoleDesc());
										//vo.setRoleDesc(elementToken[innerIndex].toString());
										StringBuffer sb = new StringBuffer("");
										for ( int innerIndex = 0; innerIndex < elementToken.length; innerIndex++ ) {
											sb.append(finalRole.get(elementToken[innerIndex].toString()).toString());
											int lppCount = innerIndex;
											if ((lppCount + 1) < (elementToken.length)) {
												sb.append(" , ");
											}
										}
										vo.setRoleDesc(sb.toString());
										mapping.put(taskCounter+"", vo);
										//taskCounter++;	
									//}
								}
							} else {
								String[] elementToken = pattern.split(roleFrameBuff.toString());
								//for ( int innerIndex = 0; innerIndex < elementToken.length; innerIndex++ ) {
									TaskDetailsMappingVO vo = new TaskDetailsMappingVO();
									vo.setTimeEnergyChangeFromBeforeSketch(asTaskEnergy+"% \n(No Change)");
									vo.setRelationalCrafting(relationalVal);
									vo.setTaskCrafting(dtls.getJctAsTaskDesc());
									//vo.setRoleDesc(dtls.getJctAsRoleDesc());
									//vo.setRoleDesc(elementToken[innerIndex].toString());
									StringBuffer sb = new StringBuffer("");
									for ( int innerIndex = 0; innerIndex < elementToken.length; innerIndex++ ) {
										sb.append(finalRole.get(elementToken[innerIndex].toString()).toString());
										int lppCount = innerIndex;
										if ((lppCount + 1) < (elementToken.length)) {
											sb.append(" , ");
										}
									}
									vo.setRoleDesc(sb.toString());
									mapping.put(taskCounter+"", vo);
									//taskCounter++;	
								//}
							}
						} else {
							String diffStatus = (String)innerObj[2];
							if (diffStatus.equals("Modified Task")) {
								String[] elementToken = pattern.split(roleFrameBuff.toString());
								//for ( int innerIndex = 0; innerIndex < elementToken.length; innerIndex++ ) {
									TaskDetailsMappingVO vo = new TaskDetailsMappingVO();
									String value = (String)innerObj[0];
									String displayValue = value.substring(1, value.length());
									if (Integer.parseInt((String)innerObj[0]) < 0) {
										vo.setTimeEnergyChangeFromBeforeSketch(asTaskEnergy+"% \n("+displayValue+"% Decrease)");
									} else {
										vo.setTimeEnergyChangeFromBeforeSketch(asTaskEnergy+"% \n("+value+"% Increase)");
									}
									//vo.setTimeEnergyChangeFromBeforeSketch(asTaskEnergy+"% \n("+diffStatus+")");
									vo.setRelationalCrafting(relationalVal);
									vo.setTaskCrafting(dtls.getJctAsTaskDesc());
									//vo.setRoleDesc(dtls.getJctAsRoleDesc());
									//vo.setRoleDesc(elementToken[innerIndex].toString());
									StringBuffer sb = new StringBuffer("");
									for ( int innerIndex = 0; innerIndex < elementToken.length; innerIndex++ ) {
										sb.append(finalRole.get(elementToken[innerIndex].toString()).toString());
										int lppCount = innerIndex;
										if ((lppCount + 1) < (elementToken.length)) {
											sb.append(" , ");
										}
									}
									vo.setRoleDesc(sb.toString());
									mapping.put(taskCounter+"", vo);
									//taskCounter++;	
								//}
							} else {
								String[] elementToken = pattern.split(roleFrameBuff.toString());
								//for ( int innerIndex = 0; innerIndex < elementToken.length; innerIndex++ ) {
									TaskDetailsMappingVO vo = new TaskDetailsMappingVO();
									vo.setTimeEnergyChangeFromBeforeSketch(asTaskEnergy+"% \n("+diffStatus+")");
									vo.setRelationalCrafting(relationalVal);
									vo.setTaskCrafting(dtls.getJctAsTaskDesc());
									//vo.setRoleDesc(dtls.getJctAsRoleDesc());
									//vo.setRoleDesc(elementToken[innerIndex].toString());
									StringBuffer sb = new StringBuffer("");
									for ( int innerIndex = 0; innerIndex < elementToken.length; innerIndex++ ) {
										sb.append(finalRole.get(elementToken[innerIndex].toString()).toString());
										int lppCount = innerIndex;
										if ((lppCount + 1) < (elementToken.length)) {
											sb.append(" , ");
										}
									}
									vo.setRoleDesc(sb.toString());
									mapping.put(taskCounter+"", vo);
									//taskCounter++;	
								//}
							}
						}
					}
					taskCounter++;	
					dtlsIdBackupList.add(dtls.getJctAsTaskId());
					}
			}
		}
		LOGGER.info("<<<< GeneratePDFServiceImpl.getTaskDescMapping");
		return mapping;
	}
	
	/*private List<String> getDeletedTask (String jobRefNo) throws DAOException {
		List<Object> objLst = null;
		List<String> returnList = new ArrayList<String>();
		objLst = afterSketchDAO.getBsToAsDeletedTask(jobRefNo);
		for (int index = 0; index < objLst.size(); index++ ) {
			Object[] innerObj = (Object[])objLst.get(index);
			String bsDesc = (String) innerObj[0];
			int bsEnergy = (Integer) innerObj[1];
			returnList.add(bsDesc+ " (was "+bsEnergy+"% of time/energy)");
		}
		return returnList;
	}*/
	private List<String> getDeletedTask (String jobRefNo) throws DAOException {
		List<Object> objLst = null;
		List<String> returnList = new ArrayList<String>();
		objLst = afterSketchDAO.getBsToAsDeletedTask(jobRefNo);
		for (int index = 0; index < objLst.size(); index++ ) {
			Object[] innerObj = (Object[])objLst.get(index);
			String bsDesc = (String) innerObj[0];
			int bsEnergy = (Integer) innerObj[1];
			returnList.add(bsDesc+ "``"+bsEnergy+"%\n(Dropped Task)");
		}
		return returnList;
	}
}