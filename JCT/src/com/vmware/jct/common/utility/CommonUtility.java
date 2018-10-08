package com.vmware.jct.common.utility;

import java.io.IOException;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.Date;

import org.codehaus.jackson.JsonNode;
import org.codehaus.jackson.JsonProcessingException;
import org.codehaus.jackson.map.ObjectMapper;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

/**
 * 
 * <p><b>Class name:</b> CommonUtility.java</p>
 * <p><b>Author:</b> InterraIT</p>
 * <p><b>Purpose:</b> This class acts as a utility class
 * All common methods are written inside this class.
 * -saveAfterSketchBasic()
 * <p><b>Description:</b> This is a utility class. All common methods are written inside this class </p>
 * <p><b>Copyrights:</b> 	All rights reserved by Interra IT and should only be used for its internal application development.</p>
 * <p><b>Created Date:</b> 27/Jan/2014</p>
 * <p><b>Revision History:</b>
 * 	<li>InterraIT - 04/Feb/2014 - Implemented comments</li>
 * </p>
 */
public class CommonUtility {
	private static final Logger LOGGER = LoggerFactory.getLogger(CommonUtility.class);
	
	
	public static Date convertToDateObj(String dateString){
		LOGGER.info(">>>> CommonUtility.convertToDateObj");
		SimpleDateFormat format = new SimpleDateFormat("MM/dd/yyyy");
		Date returnDate = null;
		try {
			returnDate = format.parse(dateString);
		} catch (ParseException e) {
			//e.printStackTrace();
			returnDate = new Date();
			LOGGER.error(e.getLocalizedMessage());
		}
		LOGGER.info("<<<< CommonUtility.convertToDateObj");
		return returnDate;
	}
	
	public static String convertMillis(long milliSeconds){
		LOGGER.info(">>>> CommonUtility.convertMillis");
		long totalSeconds = milliSeconds / 1000L;
		int currentSecond = (int)(totalSeconds % 60L);
		long totalMinutes = totalSeconds / 60L;
		int currentMinute = (int)(totalMinutes % 60L);
		long totalHours = totalMinutes / 60L;
		StringBuilder sb = new StringBuilder("");
		String totHrs = ""+totalHours;
		String totMin = ""+currentMinute;
		String totSec = ""+currentSecond;
		if(totHrs.length() == 1){
			sb.append("0"+totHrs);
		} else{
			sb.append(totHrs);
		}
		sb.append(":");
		if(totMin.length() == 1){
			sb.append("0"+totMin);
		} else{
			sb.append(totMin);
		}
		sb.append(":");
		if(totSec.length() == 1){
			sb.append("0"+totSec);
		} else{
			sb.append(totSec);
		}
		LOGGER.info("<<<< CommonUtility.convertMillis");
		return sb.toString();
	}
	/**
	 * Method returns the node
	 * @param body
	 * @return
	 */
	public static JsonNode getNode(String body) {
		LOGGER.info(">>>> CommonUtility.getNode");
		ObjectMapper mapper = new ObjectMapper();
		JsonNode node = null;
		try {
			node = mapper.readTree(body);
		} catch (JsonProcessingException e) {
			LOGGER.error(e.getLocalizedMessage());
		} catch (IOException e) {
			LOGGER.error(e.getLocalizedMessage());
		}
		LOGGER.info("<<<< CommonUtility.getNode");
		return node;
	}
}
