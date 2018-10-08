package com.vmware.jct.common.utility;

import javax.mail.internet.InternetAddress;
import javax.mail.internet.MimeMessage;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.context.MessageSource;
import org.springframework.mail.javamail.JavaMailSender;
import org.springframework.mail.javamail.MimeMessageHelper;
import org.springframework.mail.javamail.MimeMessagePreparator;
import org.springframework.scheduling.annotation.Async;
import org.springframework.stereotype.Component;

import com.vmware.jct.exception.MailingException;

@Component
public class MailDispatcher {

private static final Logger LOGGER = LoggerFactory.getLogger(MailDispatcher.class);
	
	@Autowired
	private JavaMailSender mailSender;
	
	@Autowired  
	private MessageSource messageSource;
	
	@Async
	public void sendPasswordByMail(final String to, final String password, final String name, final String firstName, final String lastName) throws MailingException {
		LOGGER.info(">>>> MailNotificationHelper.sendPasswordByMail");
		final String messageSubject = this.messageSource.getMessage("label.email.subject.reset",null, null);
		final String emailFrom = this.messageSource.getMessage("email.msg.from.sender",null, null);
		final String headerOne = this.messageSource.getMessage("email.msg.header.one",null, null);
		final String headerTwo = this.messageSource.getMessage("email.msg.header.two",null, null);
		final String headerThree = this.messageSource.getMessage("email.msg.header.three",null, null);
		final String headerFour = this.messageSource.getMessage("email.msg.header.four",null, null);
		final String headerFive = this.messageSource.getMessage("email.msg.header.five",null, null);
		final String headerSix = this.messageSource.getMessage("email.msg.header.six",null, null);
		final String headerSeven = this.messageSource.getMessage("email.msg.header.seven",null, null);
		final String headerEight = this.messageSource.getMessage("email.msg.header.eight",null, null);
		
		try {
			MimeMessagePreparator preparator = new MimeMessagePreparator() {
				public void prepare(MimeMessage mimeMessage) throws Exception {
					// Start Defect id-VMJCT-131
					String lName = firstName +" "+lastName;
					MimeMessageHelper message = new MimeMessageHelper(mimeMessage, true);
					message.setTo(to);
					message.setFrom(new InternetAddress(emailFrom, "Job Crafting LLC"));
					message.setSubject(messageSubject);
					// Start Defect id-VMJCT-131
					/*if (null==lName || lName=="" || lName.trim().length()==0 || lName.equals("null null")){
						lName="User";
					}*/
					// End Defect id-VMJCT-131
					//Create the whole message
					StringBuilder mailBuilder = new StringBuilder("");
					mailBuilder.append(headerOne+lName+headerTwo);
					mailBuilder.append(headerThree+password+headerFour+headerFive+to+headerSix+headerSeven+to);
					mailBuilder.append(headerEight);
					message.setText(mailBuilder.toString(), true);
				}
			};
			this.mailSender.send(preparator);
			LOGGER.info("----> MAIL SENT TO: "+to);
		} catch (Exception ex) {
			LOGGER.error(ex.getLocalizedMessage());
			throw new MailingException("Unable to send mail: ", ex);
		}
		LOGGER.info("<<< MailNotificationHelper.sendPasswordByMail");
	}

}
