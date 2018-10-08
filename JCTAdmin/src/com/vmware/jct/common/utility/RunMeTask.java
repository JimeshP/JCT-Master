package com.vmware.jct.common.utility;

import java.util.Date;
import java.util.Iterator;
import java.util.List;

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
import com.vmware.jct.model.JctEmailDetails;
import com.vmware.jct.service.IManageUserService;

/**
 * 
 * <p><b>Class name:</b> RunMeTask.java</p>
 * <p><b>Author:</b> InterraIT</p>
 * <p><b>Purpose:</b> .
 * <p><b>Copyrights:</b>All rights reserved by Interra IT and should only be used for its internal application development.</p>
 * <p><b>Created Date:</b> 10/Apr/2014</p>
 * <p><b>Revision History:</b>
 * 	<li></li>
 * </p>
 */

@Service("runMeTask")
public class RunMeTask {
	
	private static final Logger LOGGER = LoggerFactory.getLogger(RunMeTask.class);
	
	@Autowired
	private JavaMailSender mailSender;
	
	@Autowired  
	private MessageSource messageSource;
	
	@Autowired
	private IManageUserService userService;
	
	/**
	 * Method is initiated by the cron scheduler. Method checks if any registered 
	 * users have not been intimated by the tool. If so, it again sends registration 
	 * email.
	 * @throws MailingException
	 */
	public void initiateMailing() throws MailingException {
		try {
			// Fetch the details
			int maxResultsToFetch = Integer.parseInt(this.messageSource.getMessage("cron.fetch.max.results",null, null));
			List<JctEmailDetails> details = userService.getEmailDetails(maxResultsToFetch);
			if (null != details) {
				if (details.size() > 0) {
					// Initiate mail sending
					Iterator<JctEmailDetails> emailItr = details.iterator();
					while (emailItr.hasNext()) {
						JctEmailDetails obj = (JctEmailDetails) emailItr.next();
						try {
							sendMailToUser(obj.getJctRegisteredMail(), obj.getJctRegisteredPassword());
							
							//Merge the details as mail sent
							obj.setJctMailSentDate(new Date());
							obj.setJctMailDispatched(StatusConstants.USER_EMAIL_SENT);
							String updateStatus = userService.updateEmailDetails(obj);
							LOGGER.info("EMAIL DETAILS MODIFICATION STATUS FOR: "+obj.getJctRegisteredMail()+" IS: "+updateStatus);
						} catch (MailingException ex) {
							LOGGER.error("UNABLE TO SEND MAIL TO: "+obj.getJctRegisteredMail()+". REASON: "+ex.getLocalizedMessage());
						}
					}
				}
			}
		} catch (JCTException ex) {
			LOGGER.error("ERROR: "+ex.getLocalizedMessage());
		}
	}
	/**
	 * Method sends mail to the user
	 * @param mailTo
	 * @param password
	 * @throws MailingException
	 */
	private void sendMailToUser (final String mailTo, final String password) throws MailingException {
		LOGGER.info(">>>>>>>>> RunMeTask.sendMailToUser");
		String emailSubject = this.messageSource.getMessage("email.user.subject.msg",null, null);
		final String emailSubjectFinal = emailSubject;
		final String emailFrom = this.messageSource.getMessage("email.msg.from",null, null);
		final String messageHeader = this.messageSource.getMessage("email.msg.header.user",null, null);
		final String messageSubHeaderThanks = this.messageSource.getMessage("email.msg.user.sub.header.thanks",null, null);
		final String messageInfoOne = this.messageSource.getMessage("email.msg.user.info.one",null, null);
		final String messageInfoTwo = this.messageSource.getMessage("email.msg.user.info.two",null, null);
		final String messageInfoThree = this.messageSource.getMessage("email.msg.user.info.three",null, null);
		final String messageInfoFour = this.messageSource.getMessage("email.msg.user.info.four",null, null); //link
		final String messageInfoFive = this.messageSource.getMessage("email.msg.user.info.five",null, null);
		final String messageInfoSix = this.messageSource.getMessage("email.msg.user.info.six",null, null);
		final String messageInfoSeven = this.messageSource.getMessage("email.msg.user.info.seven",null, null);
		final String emailFooter = this.messageSource.getMessage("email.msg.footer.general",null, null);
		MimeMessagePreparator preparator = new MimeMessagePreparator() {
			public void prepare(MimeMessage mimeMessage) throws Exception {
				MimeMessageHelper message = new MimeMessageHelper(mimeMessage, true);
				message.setTo(mailTo);
				message.setFrom(new InternetAddress(emailFrom));
				message.setSubject(emailSubjectFinal);
				StringBuilder mailBuilder = new StringBuilder("");
				mailBuilder.append(messageHeader);
				mailBuilder.append(messageSubHeaderThanks);
				mailBuilder.append(messageInfoOne);
				mailBuilder.append(password);
				mailBuilder.append(messageInfoTwo);
				mailBuilder.append(messageInfoThree);
				mailBuilder.append(messageInfoFour);
				mailBuilder.append(mailTo);
				mailBuilder.append(messageInfoFive);
				mailBuilder.append(messageInfoSix);
				mailBuilder.append(messageInfoSeven);
				mailBuilder.append(mailTo);
				mailBuilder.append(emailFooter);
				message.setText(mailBuilder.toString(), true);
			}
		};
		this.mailSender.send(preparator);
		LOGGER.info("======= MAIL SENT TO: "+mailTo+" =======");
		LOGGER.info("<<<<<<<<< RunMeTask.sendMailToUser");
	}
}
