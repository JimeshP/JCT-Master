package com.vmware.jct.controller;

import java.awt.image.BufferedImage;
import java.io.BufferedWriter;
import java.io.ByteArrayInputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.FileWriter;
import java.io.IOException;
import java.io.InputStream;
import java.util.StringTokenizer;

import javax.imageio.ImageIO;
import javax.servlet.http.HttpServletRequest;

import org.codehaus.jackson.JsonNode;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.context.MessageSource;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.RequestBody;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestMethod;
import org.springframework.web.bind.annotation.ResponseBody;

import com.vmware.jct.common.utility.CommonUtility;
import com.vmware.jct.common.utility.EncoderDecoder;
import com.vmware.jct.common.utility.StatusConstants;
import com.vmware.jct.exception.JCTException;
import com.vmware.jct.service.IAppSettingsService;
import com.vmware.jct.service.IStorageService;
import com.vmware.jct.service.vo.AppSettingsVO;

/**
 * 
 * <p><b>Class name:</b> AppSettingsController.java</p>
 * <p><b>Author:</b> InterraIT</p>
 * <p><b>Purpose:</b> This class acts as a controller for white labeling requests..
 * AppSettingsController extends BasicController  and has following  methods.
 * -saveAppSettings()
 * <p><b>Description:</b> This class basically intercept all the authentication related service calls and 
 * then delegated the it to the service layer for further actions. </p>
 * <p><b>Copyrights:</b> 	All rights reserved by Interra IT and should only be used for its internal application development.</p>
 * <p><b>Created Date:</b> 10/Apr/2014</p>
 * <p><b>Revision History:</b>
 * 	<li></li>
 * </p>
 */
@Controller
@RequestMapping(value="/appSettings")
public class AppSettingsController extends BasicController {
	
	@Autowired  
	private MessageSource messageSource;
	
	@Autowired
	private IAppSettingsService appSettingsService;
	
	@Autowired
	private IStorageService storageService;
	
	private static final Logger LOGGER = LoggerFactory.getLogger(AppSettingsController.class);
	/**
	 * Method saves the white lebeling
	 * @param body
	 * @param request
	 */
	@RequestMapping(value = ACTION_SAVE_APP_SETTINGS, method = RequestMethod.POST)
	@ResponseBody()
	public AppSettingsVO saveAppSettings(@RequestBody String body, HttpServletRequest request) {
		LOGGER.info(">>>>>> AppSettingsController.saveAppSettings");
		AppSettingsVO settingsVO = new AppSettingsVO();
		String status = "failed";
		JsonNode node = CommonUtility.getNode(body);
		String headerColor = node.get("headerColor").toString().replaceAll("\"", "");
		String footerColor = node.get("footerColor").toString().replaceAll("\"", "");
		String subHeaderColor = node.get("subHeaderColor").toString().replaceAll("\"", "");
		String instructionPanelColor = node.get("instructionPanelColor").toString().replaceAll("\"", "");
		String instructionPanelTextColor = node.get("instructionPanelTextColor").toString().replaceAll("\"", "");
		String file = node.get("imgByte").toString();
		String jctEmail = node.get("jctEmail").asText();
		//check if image is provided
		byte[] imageInByte = null;
		//LOGGER.info(">>>>>> AppSettingsController.saveAppSettings : file : " + file);
		if (!file.toString().replaceAll("\"", "").equals("NC")) {
			imageInByte = EncoderDecoder.getImageBytes(file.substring(file.indexOf(",") + 1, file.length()));
		}
		//create a css with all colors and images
		try {
			status = whiteLebel(headerColor, footerColor, subHeaderColor, instructionPanelColor, 
					instructionPanelTextColor, imageInByte, jctEmail, file);
			if (!status.equals("failed")) {
				settingsVO.setStatusCode(StatusConstants.STATUS_SUCCESS);
				settingsVO.setMessage(this.messageSource.getMessage("app.settings.success",null, null));
			} else {
				settingsVO.setStatusCode(StatusConstants.STATUS_FAILED);
				settingsVO.setMessage(this.messageSource.getMessage("app.settings.error",null, null));
			}
		} catch (Exception e) {
			LOGGER.error("------------ UNABLE TO WHITE LEVEL ------------- "+e.getLocalizedMessage());
		}
		LOGGER.info("<<<<<< AppSettingsController.saveAppSettings");
		return settingsVO;
	}
	/**
	 * Method actually white levels the cascading style sheet
	 * @param headerColor
	 * @param footerColor
	 * @param subHeaderColor
	 * @param instructionPanelColor
	 * @param instructionPanelTextColor
	 * @param imageInByte
	 * @return
	 * @throws Exception
	 */
	private String whiteLebel(String headerColor, String footerColor, String subHeaderColor, 
			String instructionPanelColor, String instructionPanelTextColor, byte[] imageInByte, String jctEmail, String base64String) throws Exception {
		AppSettingsVO settingsVO = new AppSettingsVO();
		settingsVO = appSettingsService.fetchColor(settingsVO);
		
		LOGGER.info(">>>>>> AppSettingsController.whiteLebel");
		String status = "failed";
		String cssDefaultString = "h5.heading_main{ "+
													"margin: 0; "+
													"font-size: 22px !important; "+
													"line-height: 50px !important; "+
													"font-weight: normal; "+
												"} "+
								   "h5.logo{ "+
								   			"background: url('/user/img/logo.png') no-repeat 0 center; "+
								   			"display: block; "+
								   			"width: 203px; "+
								   			"height:53px; "+
								   			"cursor:pointer; "+
								   		  "} "+
								    ".header{ "+
								    		"position: relative; "+
								    		"color: #fff; "+
								    		"min-height: 54px; "+
								    		"padding-left: 20px; "+
								    		"padding-right: 20px; "+
								    		"background: "+settingsVO.getHeaderColor()+";  "+
								    		"border: none; "+
								    		"}" +
								    ".footer-wrapper{ background:"+settingsVO.getFooterColor()+"; padding: 15px 0px;}" +
								    ".nav-header{ "+
											"position: relative; "+
											"color: #fff; "+
											"min-height: 50px; "+
											"padding-left: 20px; "+
											"padding-right: 20px; "+
											"background: "+settingsVO.getSubHeaderColor()+";  "+
											"border-bottom: 1px solid #BABABA; "+
											"padding-top: 0; "+
											"border-top: #fdfdfd 1px solid; "+
											"box-shadow: 1px 3px 3px #6C6C6C; }" +
									"#panel { "+
										"background: "+settingsVO.getInstructionPanelColor()+"; "+
										"display: none; "+
										"color: "+settingsVO.getInstructionPanelTextColor()+"; "+
										"padding: 1px 0px;}"
									+ ".slide {"+
									"margin: 0;"+
									"padding: 0px;"+
									"border-top: solid 30px #810123;"+
									"display: block;"+
									"}"+
									".bar_loop{"+
									    "background: none repeat scroll 0 0 #810123;"+
									    "display: block;"+
									    "height: 6px;"+
									    "left: -21px;"+
									    "position: absolute;"+
									    "top: -27px;"+
									    "width: 42px;"+	
									"}"+
									"#drp{"+
										"border-left: 21px solid rgba(0, 0, 0, 0);"+
									    "border-right: 21px solid rgba(0, 0, 0, 0);"+
									    "border-top: 22px solid #810123;"+
									    "bottom: -27px;"+
									    "height: 19px;"+
									    "left: 20px;"+
									    "padding: 0;"+
									    "position: absolute;"+
									    "width: 0;"+
									"}";
		String dirLoc = this.messageSource.getMessage("jct.build.path",null, null);
		LOGGER.info(">>>>>> AppSettingsController.whiteLebel : " + dirLoc);
		File saveDir = new File(dirLoc); //webapps
		//Here comes the existence check
		if(!saveDir.exists())
			saveDir.mkdirs();
		
		if (null != imageInByte) {
			LOGGER.info(">>>>>> AppSettingsController.whiteLebel : " + saveDir.getAbsoluteFile());
			try {
				InputStream in = new ByteArrayInputStream(imageInByte);
				BufferedImage bImageFromConvert = ImageIO.read(in);
				ImageIO.write(bImageFromConvert, "png", new File(dirLoc+"/logo.png"));
				storageService.store("images/logo.png", new FileInputStream(dirLoc+"/logo.png"), "image/png");
			} catch (Exception e) {
				LOGGER.info(">>>>>> AppSettingsController.whiteLebel : writing logo file exception : " + e.getMessage());
				e.printStackTrace();
			}
			LOGGER.info(">>>>>> AppSettingsController.whiteLebel : writing logo file done");
		}
		
		StringTokenizer cssToken = new StringTokenizer(cssDefaultString, "}");
		String[] cssArr = new String[cssToken.countTokens()];
		while (cssToken.hasMoreTokens()) {
			cssArr[0] = (String) cssToken.nextToken();
			cssArr[1] = (String) cssToken.nextToken();
			if (headerColor.equals("NC")) {
				cssArr[2] = (String) cssToken.nextToken(); //header color
			} else {
				cssToken.nextToken(); // this is just for the iteration
				cssArr[2] = ".header{ "+
								    "position: relative; "+
								    "color: #fff; "+
								    "min-height: 50px; "+
								    "padding-left: 20px; "+
								    "padding-right: 20px; "+
								    "background: "+headerColor+";  "+
								    "border: none; ";
				
			}
			if (footerColor.equals("NC")) {
				cssArr[3] = (String) cssToken.nextToken(); //header color
			} else {
				cssToken.nextToken(); // this is just for the iteration
				cssArr[3] = ".footer-wrapper{ background:"+footerColor+"; padding: 15px 0px;";
				
			}
			
			if (subHeaderColor.equals("NC")) {
				cssArr[4] = (String) cssToken.nextToken(); //header color
			} else {
				cssToken.nextToken(); // this is just for the iteration
				cssArr[4] = ".nav-header{ "+
											"position: relative; "+
											"color: #fff; "+
											"min-height: 50px; "+
											"padding-left: 20px; "+
											"padding-right: 20px; "+
											"background: "+subHeaderColor+";  "+
											"border-bottom: 1px solid #BABABA; "+
											"padding-top: 0; "+
											"border-top: #fdfdfd 1px solid; "+
											"box-shadow: 1px 3px 3px #6C6C6C; ";
				
			}
			
			
			if ((instructionPanelColor.equals("NC")) && (!instructionPanelTextColor.equals("NC"))) {
				cssToken.nextToken(); // this is just for the iteration
				cssArr[5] = "#panel { "+
										"background: #810123; "+
										"display: none; "+
										"color: "+instructionPanelTextColor+"; "+
										"padding: 1px 0px";
				cssToken.nextToken();
				cssArr[6] = ".slide {"+
									"margin: 0;"+
									"padding: 0px;"+
									"border-top: solid 30px #810123;"+
									"display: block;";
				
				cssToken.nextToken();
				cssArr[7] = ".bar_loop{"+
				    "background: none repeat scroll 0 0 #810123;"+
				    "display: block;"+
				    "height: 6px;"+
				    "left: -21px;"+
				    "position: absolute;"+
				    "top: -27px;"+
				    "width: 42px;";
				
				cssToken.nextToken();
				cssArr[8] = "#drp{"+
					"border-left: 21px solid rgba(0, 0, 0, 0);"+
				    "border-right: 21px solid rgba(0, 0, 0, 0);"+
				    "border-top: 22px solid #810123;"+
				    "bottom: -27px;"+
				    "height: 19px;"+
				    "left: 20px;"+
				    "padding: 0;"+
				    "position: absolute;"+
				    "width: 0;";
			} else if ((!instructionPanelColor.equals("NC")) && (instructionPanelTextColor.equals("NC"))) {
				cssToken.nextToken(); // this is just for the iteration
				cssArr[5] = "#panel { "+
										"background: "+instructionPanelColor+"; "+
										"display: none; "+
										"color: #fff; "+
										"padding: 1px 0px";
				cssToken.nextToken();
				cssArr[6] = ".slide {"+
									"margin: 0;"+
									"padding: 0px;"+
									"border-top: solid 30px "+instructionPanelColor+";"+
									"display: block;";
				
				cssToken.nextToken();
				cssArr[7] = ".bar_loop{"+
				    "background: none repeat scroll 0 0 "+instructionPanelColor+";"+
				    "display: block;"+
				    "height: 6px;"+
				    "left: -21px;"+
				    "position: absolute;"+
				    "top: -27px;"+
				    "width: 42px;";
				
				cssToken.nextToken();
				cssArr[8] = "#drp{"+
					"border-left: 21px solid rgba(0, 0, 0, 0);"+
				    "border-right: 21px solid rgba(0, 0, 0, 0);"+
				    "border-top: 22px solid "+instructionPanelColor+";"+
				    "bottom: -27px;"+
				    "height: 19px;"+
				    "left: 20px;"+
				    "padding: 0;"+
				    "position: absolute;"+
				    "width: 0;";
			} else if ((!instructionPanelColor.equals("NC")) && (!instructionPanelTextColor.equals("NC"))) {
				cssToken.nextToken(); // this is just for the iteration
				cssArr[5] = "#panel { "+
										"background: "+instructionPanelColor+"; "+
										"display: none; "+
										"color: "+instructionPanelTextColor+"; "+
										"padding: 1px 0px";
				cssToken.nextToken();
				cssArr[6] = ".slide {"+
									"margin: 0;"+
									"padding: 0px;"+
									"border-top: solid 30px "+instructionPanelColor+";"+
									"display: block;";
				
				cssToken.nextToken();
				cssArr[7] = ".bar_loop{"+
				    "background: none repeat scroll 0 0 "+instructionPanelColor+";"+
				    "display: block;"+
				    "height: 6px;"+
				    "left: -21px;"+
				    "position: absolute;"+
				    "top: -27px;"+
				    "width: 42px;";
				
				cssToken.nextToken();
				cssArr[8] = "#drp{"+
					"border-left: 21px solid rgba(0, 0, 0, 0);"+
				    "border-right: 21px solid rgba(0, 0, 0, 0);"+
				    "border-top: 22px solid "+instructionPanelColor+";"+
				    "bottom: -27px;"+
				    "height: 19px;"+
				    "left: 20px;"+
				    "padding: 0;"+
				    "position: absolute;"+
				    "width: 0;";
			} else {
				cssArr[5] = (String) cssToken.nextToken();
				cssArr[6] = (String) cssToken.nextToken();
				cssArr[7] = (String) cssToken.nextToken();
				cssArr[8] = (String) cssToken.nextToken();
			}
		}
		// Write the css in the file
		StringBuilder cssBuilder = new StringBuilder("");
		for (int index = 0; index < cssArr.length; index++) {
			cssBuilder.append(cssArr[index]+"} \n");
		}
		FileOutputStream fop = null;
		File file;
		try {
			file = new File(dirLoc+"/commonStyle.css");

			LOGGER.info(">>>>>> AppSettingsController.whiteLebel : commonStyle.css : " + file.getAbsolutePath());
			fop = new FileOutputStream(file);
 			// if file doesnt exists, then create it
			if (!file.exists()) {
				file.createNewFile();
			}
			FileWriter fw = new FileWriter(file.getAbsoluteFile());
			BufferedWriter bw = new BufferedWriter(fw);
			bw.write(cssBuilder.toString());
			bw.close();
			storageService.store("css/commonStyle.css", new FileInputStream(file), "text/css");
			status = "passed";
			LOGGER.info(">>>>>> AppSettingsController.whiteLebel : commonStyle.css : success");
		} catch (IOException e) {
			LOGGER.info(">>>>>> AppSettingsController.whiteLebel : commonStyle.css : fail : " + e.getMessage());
			LOGGER.error(e.getLocalizedMessage());
		} catch (Exception e) {
			LOGGER.info(">>>>>> AppSettingsController.whiteLebel : commonStyle.css : fail : " + e.getMessage());
			LOGGER.error(e.getLocalizedMessage());
		} finally {
			try {
				if (fop != null) {
					fop.close();
				}
			} catch (IOException e) {
				LOGGER.error(e.getLocalizedMessage());
				status = "failed";
			}
		}		
		if (status.equals("passed")) {
			//Store it in database...
			if (!headerColor.equals("NC")) {
				settingsVO.setHeaderColor(headerColor);
			}
			if (!footerColor.equals("NC")) {
				settingsVO.setFooterColor(footerColor);
			}
			if (!subHeaderColor.equals("NC")) {
				settingsVO.setSubHeaderColor(subHeaderColor);
			}
			if (!instructionPanelColor.equals("NC")) {
				settingsVO.setInstructionPanelColor(instructionPanelColor);
			}
			if (!instructionPanelTextColor.equals("NC")) {
				settingsVO.setInstructionPanelTextColor(instructionPanelTextColor);
			}
			if (null != imageInByte) {
				settingsVO.setImage(base64String.toString().replaceAll("\"", ""));
			}
			try {
				status = appSettingsService.saveNewWhiteLebel(settingsVO, jctEmail);
			} catch (JCTException ex) {
				status = "failed";
			}
		}
		LOGGER.info("<<<<<< AppSettingsController.whiteLebel");
		return status;
	}
	
	/**
	 * Method fetches previously stored master color.
	 * @param body
	 * @param request
	 * @return
	 */
	@RequestMapping(value = ACTION_FETCH_COLOR, method = RequestMethod.POST)
	@ResponseBody()
	public AppSettingsVO fetchColor(@RequestBody String body, HttpServletRequest request) {
		LOGGER.info(">>>>>> AppSettingsController.fetchColor");
		AppSettingsVO settingsVO = new AppSettingsVO();
		try {
			settingsVO = appSettingsService.fetchColor(settingsVO);
		} catch (JCTException exception) {
			LOGGER.error("UNEXPECTED TO GET ERROR WHILE FETCHING VALUES..."+exception.getLocalizedMessage());
			LOGGER.error("WILL SET THE SCREEN TO DEFAULT VALUES.");
			settingsVO.setHeaderColor("#006990");
			settingsVO.setFooterColor("#006990");
			settingsVO.setSubHeaderColor("#88cbdf");
			settingsVO.setInstructionPanelColor("#810123");
			settingsVO.setInstructionPanelTextColor("#fff");
		}
		LOGGER.info("<<<<<< AppSettingsController.fetchColor");
		return settingsVO;
	}
}
