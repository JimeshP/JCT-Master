package com.vmware.jct.common.utility;

import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.StringTokenizer;

import javax.mail.internet.AddressException;
import javax.mail.internet.InternetAddress;
import javax.mail.internet.MimeMessage;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.context.MessageSource;
import org.springframework.mail.javamail.JavaMailSender;
import org.springframework.mail.javamail.MimeMessageHelper;
import org.springframework.mail.javamail.MimeMessagePreparator;
import org.springframework.stereotype.Service;

import com.vmware.jct.exception.JCTException;
import com.vmware.jct.exception.MailingException;
import com.vmware.jct.service.IAccountExpiryNotificationService;

/**
 * 
 * <p><b>Class name:</b> AccountExpiryNotificationTask.java</p>
 * <p><b>Author:</b> InterraIT</p>
 * <p><b>Purpose:</b> .
 * <p><b>Copyrights:</b>All rights reserved by Interra IT and should only be used for its internal application development.</p>
 * <p><b>Created Date:</b> 10/Oct/2014</p>
 * <p><b>Revision History:</b>
 * 	<li></li>
 * </p>
 */

@Service("accountExpiryNotificationTask")
public class AccountExpiryNotificationTask {
private static final Logger LOGGER = LoggerFactory.getLogger(AccountExpiryNotificationTask.class);
	
	@Autowired
	private JavaMailSender mailSender;
	
	@Autowired  
	private MessageSource messageSource;
	
	@Autowired
	private IAccountExpiryNotificationService notificationService;
	/**
	 * Method is initiated by the cron scheduler. Method notifies the user and 
	 * corresponding linked facilitators in case of account expiry in 3 days.
	 */
	public void notifyExpirationOfAccount() {
		LOGGER.info(">>>>>>>>> AccountExpiryNotificationTask.notifyExpirationOfAccount");
		try {
			//key = email Id of facilitator, value = list of users who have been registered by facilitator
			Map<String, List<String>> expiryDetails = notificationService.getAcntExpryNtfcatnForFacilitator();
			Iterator iterator = expiryDetails.entrySet().iterator();
			ArrayList<String> all = new ArrayList<String>();
			while (iterator.hasNext()) {
				StringBuilder facilitatorMailDetailsBuilder = new StringBuilder("");
				InternetAddress[] inet = null;
				ArrayList<String> individual = new ArrayList<String>();
				Map.Entry mapEntry = (Map.Entry) iterator.next();
				String facilitatorEmailId = (String) mapEntry.getKey();
				List<String> userDetails =  (List<String>) mapEntry.getValue();
				Iterator<String> userDtlsItr = userDetails.iterator();
				int itrCount = 0;
				int numberOfUsersGettingDeactivated = 0;
				while (userDtlsItr.hasNext()) {
					String userDetailsStr = (String) userDtlsItr.next(); //UserName1`ExpiryDate1!UserName2`ExpiryDate2!
					StringTokenizer token = new StringTokenizer(userDetailsStr, "!");
					numberOfUsersGettingDeactivated = token.countTokens();
					while (token.hasMoreTokens()) {
						String singleElement = (String) token.nextToken();
						String[] splt = singleElement.split("`");
						itrCount = itrCount + 1;
						individual.add((String)splt[0]);
						facilitatorMailDetailsBuilder.append("<tr>");
						facilitatorMailDetailsBuilder.append("<td>"+itrCount+"</td>");
						facilitatorMailDetailsBuilder.append("<td>"+(String)splt[0]+"</td>");
						facilitatorMailDetailsBuilder.append("<td>"+(String)splt[1]+"</td>");
						facilitatorMailDetailsBuilder.append("</tr>");
					}
					inet = new InternetAddress[individual.size()];
					Object[] arr = individual.toArray();
					for (int i = 0; i <arr.length; i++) {
						try {
							inet[i] = new InternetAddress((String) individual.get(i));
						} catch (AddressException e) {
							LOGGER.error("UNABLE TO CONVERT TO THE INTERNET ADDRESS: "+e.getLocalizedMessage());
						}
					}
					try {
						sendMailToUsersInBCC(inet, facilitatorEmailId);
					} catch (MailingException e) {
						LOGGER.error("UNABLE TO SEND MAIL TO THE USERS: "+e.getLocalizedMessage());
					}
				}
				try {
					if (numberOfUsersGettingDeactivated > 0) {
						sendMailToFacilitator(facilitatorEmailId, facilitatorMailDetailsBuilder.toString());
					}
				} catch (MailingException e) {
					LOGGER.error("UNABLE TO SEND MAIL TO THE FACILITATORS: "+e.getLocalizedMessage());
				}
			}
		} catch (JCTException ex) {
			LOGGER.error("UNABLE TO FETCH ACCOUNT EXPIRY DETAILS: "+ex.getLocalizedMessage());
		}
		LOGGER.info("<<<<<<<<< AccountExpiryNotificationTask.notifyExpirationOfAccount");
	}
	/**
	 * Method notifies facilitator in event of account exipry of its users 
	 * @param mailTo
	 * @param userDetails
	 * @throws MailingException
	 */
	private void sendMailToFacilitator (final String mailTo, final String userDetails) throws MailingException {
		LOGGER.info(">>>>>>>>> AccountExpiryNotificationTask.sendMailToFacilitator");
		final String subject = this.messageSource.getMessage("label.facilitator.exp.subject",null, null);
		final String one = this.messageSource.getMessage("label.facilitator.one",null, null);
		final String two = this.messageSource.getMessage("label.facilitator.two",null, null);
		final String slNo = this.messageSource.getMessage("label.facilitator.sl.no",null, null);
		final String emailId = this.messageSource.getMessage("label.facilitator.email.id",null, null);
		final String expDt = this.messageSource.getMessage("label.facilitator.exp.dt",null, null);
		final String emailFrom = this.messageSource.getMessage("email.msg.from",null, null);
		MimeMessagePreparator preparator = new MimeMessagePreparator() {
			public void prepare(MimeMessage mimeMessage) throws Exception {
				MimeMessageHelper message = new MimeMessageHelper(mimeMessage, true);
				message.setTo(mailTo);
				message.setFrom(new InternetAddress(emailFrom));
				message.setSubject(subject);
				StringBuilder mailBuilder = new StringBuilder(one);
				mailBuilder.append("<table border='1'><tr><th>"+slNo+"</th><th>"+emailId+"</th><th>"+expDt+"</th>");
				mailBuilder.append(userDetails);
				mailBuilder.append("</table>");
				mailBuilder.append(two);
				message.setText(mailBuilder.toString(), true);
			}
		};
		LOGGER.info("...SENDING MAIL TO FACILITATOR...");
		this.mailSender.send(preparator);
		LOGGER.info("...MAIL SENT TO FACILITATOR...");
		LOGGER.info("<<<<<<<<< AccountExpiryNotificationTask.sendMailToFacilitator");
	}
	/**
	 * Method sends mail to the users in event of account expiry	
	 * @param inetArray
	 * @param facilitatorEmailId
	 * @throws MailingException
	 */
	private void sendMailToUsersInBCC (final InternetAddress[] inetArray, final String facilitatorEmailId) throws MailingException {
		LOGGER.info(">>>>>>>>> AccountExpiryNotificationTask.sendMailToUsers");
		final String subject = this.messageSource.getMessage("label.user.exp.subject",null, null);
		final String one = this.messageSource.getMessage("label.user.one",null, null);
		final String two = this.messageSource.getMessage("label.user.two",null, null);
		final String three = this.messageSource.getMessage("label.user.three",null, null);
		final String four = this.messageSource.getMessage("label.user.four",null, null);
		final String emailFrom = this.messageSource.getMessage("email.msg.from",null, null);
		MimeMessagePreparator preparator = new MimeMessagePreparator() {
			public void prepare(MimeMessage mimeMessage) throws Exception {
				MimeMessageHelper message = new MimeMessageHelper(mimeMessage, true);
				message.setBcc(inetArray);
				message.setFrom(new InternetAddress(emailFrom));
				message.setSubject(subject);
				StringBuilder mailBuilder = new StringBuilder(one);
				mailBuilder.append(three);
				mailBuilder.append(facilitatorEmailId);
				mailBuilder.append(four);
				mailBuilder.append(two);
				message.setText(mailBuilder.toString(), true);
			}
		};
		LOGGER.info("...SENDING MAIL TO BCC USERS...");
		this.mailSender.send(preparator);
		LOGGER.info("...MAIL SENT TO BCC USERS...");
		LOGGER.info("<<<<<<<<< AccountExpiryNotificationTask.sendMailToUsers");
	}
}