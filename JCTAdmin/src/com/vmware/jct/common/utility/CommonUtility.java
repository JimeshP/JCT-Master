package com.vmware.jct.common.utility;

import java.io.IOException;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.Calendar;
import java.util.Date;
import java.util.GregorianCalendar;
import java.util.Hashtable;

import javax.naming.NamingException;
import javax.naming.directory.Attribute;
import javax.naming.directory.Attributes;
import javax.naming.directory.DirContext;
import javax.naming.directory.InitialDirContext;

import org.codehaus.jackson.JsonNode;
import org.codehaus.jackson.JsonProcessingException;
import org.codehaus.jackson.map.ObjectMapper;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Component;
import org.springframework.stereotype.Service;

import com.vmware.jct.dao.ICommonDao;
import com.vmware.jct.exception.DAOException;
/**
 * 
 * <p><b>Class name:</b> CommonUtility.java</p>
 * <p><b>Author:</b> InterraIT</p>
 * <p><b>Purpose:</b> This class acts as an utility class with functionalities used 
 * across the project.
 * <p><b>Copyrights:</b>All rights reserved by Interra IT and should only be used for its internal application development.</p>
 * <p><b>Created Date:</b> 10/Apr/2014</p>
 * <p><b>Revision History:</b>
 * 	<li></li>
 * </p>
 */
public class CommonUtility {
	
	@Autowired
	private ICommonDao commonDaoImpl;
	
	private static final Logger LOGGER = LoggerFactory.getLogger(CommonUtility.class);
	/**
	 * Method converts String to date object
	 * @param dateString
	 * @return util.Date object
	 */
	public static Date convertToDateObj(String dateString){
		LOGGER.info(">>>> CommonUtility.convertToDateObj");
		SimpleDateFormat format = new SimpleDateFormat("MM/dd/yyyy");
		Date returnDate = null;
		try {
			returnDate = format.parse(dateString);
		} catch (ParseException e) {
			//e.printStackTrace();
			LOGGER.error(e.getLocalizedMessage());
			returnDate = new Date();
		}
		LOGGER.info("<<<< CommonUtility.convertToDateObj");
		return returnDate;
	}
	/**
	 * Method converts milliseconds to HH:MM:SS format
	 * @param milliSeconds
	 * @return String in HH:MM:SS
	 */
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
	 * Method converts seconds to HH:MM:SS format
	 * @param milliSeconds
	 * @return String in HH:MM:SS
	 */
	public static String convertSeconds(int seconds){
		LOGGER.info(">>>> CommonUtility.convertMillis");
		int totalSeconds = seconds;
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
	 * @return node
	 */
	public static JsonNode getNode(String body) {
		LOGGER.info(">>>> CommonUtility.getNode");
		ObjectMapper mapper = new ObjectMapper();
		JsonNode node = null;
		try {
			node = mapper.readTree(body);
		} catch (JsonProcessingException e) {
			LOGGER.error(e.getLocalizedMessage());
			//e.printStackTrace();
		} catch (IOException e) {
			LOGGER.error(e.getLocalizedMessage());
		}
		LOGGER.info("<<<< CommonUtility.getNode");
		return node;
	}
	
	public static Date getExpDate() {
		Calendar cal = GregorianCalendar.getInstance();
		// add 3 days 
	    cal.add((GregorianCalendar.DATE), 3);
	    return cal.getTime();
	}
	
	public static Date getCurrentDate() {
		Calendar cal = GregorianCalendar.getInstance();
		// add 3 days 
	    cal.add((GregorianCalendar.DATE), -1);
	    return cal.getTime();
	}
	
	public static Boolean isEmailValid (String emailId) {
		java.util.regex.Pattern p = java.util.regex.Pattern.compile(CommonConstants.EMAIL_PATTERN);
	    java.util.regex.Matcher m = p.matcher(emailId);
	    Boolean valid = false;
	    if (m.matches()) {
	    	String hostName = emailId.split("@")[1].toString();
	    	//check if domain name is valid [number of mail servers configured]
	    	try {
				int mailServerConfCount = doLookup(hostName);
				if (mailServerConfCount > 0) {
					valid = true;
				}
			} catch (NamingException e) {
				
			}
	    }
	    return valid;
	}
	
	public static int doLookup( String hostName ) throws NamingException {
	    Hashtable env = new Hashtable();
	    env.put("java.naming.factory.initial", "com.sun.jndi.dns.DnsContextFactory");
	    DirContext ictx = new InitialDirContext( env );
	    Attributes attrs = ictx.getAttributes( hostName, new String[] { "MX" });
	    Attribute attr = attrs.get( "MX" );
	    if( attr == null ) return( 0 );
	    return( attr.size() );
	  }
	
	public String generateTransactionCode(ICommonDao commonDao) {
		StringBuffer sb = new StringBuffer("ADM00");
		Integer id = new Integer(0);
		synchronized (this) {
			try {
				id = commonDao.generateKey("jct_payment_transaction_id");
			} catch (DAOException e) {
				LOGGER.error("UNABLE TO GENERATE TRANSACTION REFERENCE NOS: REASON: "+e.getLocalizedMessage());
			}
		}
		sb.append(generateTransRefNo(id, String.valueOf(id).length()));
		return sb.toString();
	}
	private String generateTransRefNo(int id, int length){
		String idString = "";
		switch (length) {
		case 1:
			idString = "0000000"+id;
			break;
		case 2:
			idString = "000000"+id;
			break;
		case 3:
			idString = "00000"+id;
			break;
		case 4:
			idString = "0000"+id;
			break;
		case 5:
			idString = "000"+id;
			break;
		case 6:
			idString = "00"+id;
			break;
		case 7:
			idString = "0"+id;
			break;

		default:
			idString = ""+id;
			break;
		}
		return idString;
	}
}
