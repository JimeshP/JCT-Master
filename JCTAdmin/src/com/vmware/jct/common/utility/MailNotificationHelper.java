package com.vmware.jct.common.utility;


import java.util.ArrayList;
import java.util.Date;
import java.util.List;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.context.MessageSource;
import org.springframework.mail.javamail.JavaMailSender;
import org.springframework.stereotype.Service;

import com.vmware.jct.dao.dto.NewUserDTO;
import com.vmware.jct.exception.MailingException;
/**
 * 
 * <p><b>Class name:</b> MailNotificationHelper.java</p>
 * <p><b>Author:</b> InterraIT</p>
 * <p><b>Purpose:</b> This class sends email notification to users on successful 
 * registration and re-activation
 * <p><b>Copyrights:</b>All rights reserved by Interra IT and should only be used for its internal application development.</p>
 * <p><b>Created Date:</b> 10/Apr/2014</p>
 * <p><b>Revision History:</b>
 * 	<li></li>
 * </p>
 */
@Service("mailService")
public class MailNotificationHelper {

	
	@Autowired
	MailDispatcher helper;
	
	@Autowired
	private JavaMailSender mailSender;
	
	@Autowired  
	private MessageSource messageSource;
	
	private static final Logger LOGGER = LoggerFactory.getLogger(MailNotificationHelper.class);
	/**
	 * This method will intimate the administration via mail about the user registration 
	 * status
	 * @param userDTO
	 * @param adminMailId
	 * @param invalidEmailIds
	 * @throws MailingException
	 */
	public void intimateAdmin(final NewUserDTO userDTO, final String adminMailId, 
			final ArrayList<String> invalidEmailIds) throws MailingException {
		try {
			helper.intimateAdminAsyc(userDTO, adminMailId, invalidEmailIds);
		} catch (MailingException mailingExp) {
			LOGGER.error("UNABLE TO SEND MAIL TO ADMIN: "+adminMailId+". REASON: "+mailingExp.getLocalizedMessage());
		}
	}
	
	/**
	 * This method will intimate the users about the generated password, via mail, on successful registration 
	 * or reset password done by admin
	 * @param to
	 * @param password
	 * @param discVal
	 * @throws MailingException
	 */
	public void intimateUser(final String to, final String password, final String fName,final String discVal, final String paymentType) throws MailingException {
		try {
			helper.intimateUserAsync(to, password, fName, discVal, paymentType);
		} catch (MailingException mailingExp) {
			LOGGER.error("UNABLE TO SEND MAIL TO: "+to+". REASON: "+mailingExp.getLocalizedMessage());
		}
	}
	/**
	 * This method will intimate the Facilitator about the generated password, via mail, on successful check realization 
	 * @param to
	 * @param password
	 * @param discVal
	 * @throws MailingException
	 */
	public void intimateFaci(final String to, final String password, final String fName, final String discVal, final String paymentType) throws MailingException {
		try {
			helper.intimateFaciAsync(to, password, fName, discVal, paymentType);
		} catch (MailingException mailingExp) {
			LOGGER.error("UNABLE TO SEND MAIL TO: "+to+". REASON: "+mailingExp.getLocalizedMessage());
		}
	}
	/**
	 * This method will intimate the new facilitator
	 * @param to
	 * @param password
	 * @param discVal
	 * @throws MailingException
	 */
	public void intimateNewFacilitator(final String facilitatorEmail, final String password, final String fName, final int totalSubscribedUsers, 
			final String opetedAccount,final String facilitatorId,final String paymentType) throws MailingException {
		try {
			helper.intimateNewFacilitatorAsync(facilitatorEmail, password, fName, totalSubscribedUsers, opetedAccount, facilitatorId, paymentType);
		} catch (MailingException mailingExp) {
			LOGGER.error("UNABLE TO SEND MAIL TO: "+facilitatorEmail+". REASON: "+mailingExp.getLocalizedMessage());
		}
	}
	
	/**
	 * Facilitator forgot password
	 * @param to
	 * @param password
	 * @throws MailingException
	 */
	public void intimateFaciPassword(final String facilitatorEmail, final String password, final String fName) throws MailingException {
		helper.intimateFaciPassword(facilitatorEmail, password, fName);
	}
	
	/**
	 * This method will intimate the administration via mail about the user registration 
	 * status
	 * @param userDTO
	 * @param adminMailId
	 * @param invalidEmailIds
	 * @throws MailingException
	 */
	public void intimateAdminGeneralUserRegistration(final List<NewUserDTO> userDTOList, final String adminMailId, 
			final ArrayList<String> invalidEmailIds) throws MailingException {
		try {
			helper.intimateAdminGeneralUserRegistrationAsync(userDTOList, adminMailId, invalidEmailIds);;
		} catch (MailingException mailingExp) {
			LOGGER.error("UNABLE TO SEND MAIL TO: "+adminMailId+". REASON: "+mailingExp.getLocalizedMessage());
		}
	}
	
	/**
	 * This method will send main on user renewal
	 * @param userEmail
	 * @param renewDate
	 * @throws MailingException
	 */
	public void sendMailToRenewedUser(final String userEmail, final Date renewDate, final String paymentType) throws MailingException {
		try {
		helper.sendMailToRenewedUserAsync(userEmail, renewDate, paymentType);
	} catch (MailingException mailingExp) {
		LOGGER.error("UNABLE TO SEND MAIL TO: "+userEmail+". REASON: "+mailingExp.getLocalizedMessage());
	}
		}
	
	/**
	 * This method will intimate the new facilitator
	 * @param userEmail
	 * @param facilitatorEmail
	 * @throws MailingException
	 */
	/*public void intimateUserToFacilitator(final String userEmail,  final String facilitatorEmail
					) throws MailingException {
		LOGGER.info(">>>>>>>>> MailNotificationHelper.intimateNewFacilitator");
		sendMainToFacilitator(facilitatorEmail);
		sendMainToUser(userEmail);
		LOGGER.info("<<<<<<<<< MailNotificationHelper.intimateNewFacilitator");
	}*/
	public void intimateUserToFacilitator(final String userEmail,  final String facilitatorEmail) 
			throws MailingException {
		try {
			helper.intimateUserToFacilitatorAsync(userEmail,facilitatorEmail);
		} catch (MailingException mailingExp) {
			LOGGER.error("UNABLE TO SEND MAIL TO: "+facilitatorEmail+". REASON: "+mailingExp.getLocalizedMessage());
		}
	}
	
/*	*//**
	 * This method used to send main to the user when the user is selected as facilitator 
	 * @param userEmail
	 * @throws MailingException
	 *//*
	private void sendMainToUser(final String userEmail) {
		final String emailSubject = this.messageSource.getMessage("email.user.subject.msg",null, null);
		final String adminEmailId = this.messageSource.getMessage("email.msg.from",null, null);
		final String msg1 = this.messageSource.getMessage("email.new.user.to.fac.msg1",null, null);
		final String msg2 = this.messageSource.getMessage("email.new.user.to.fac.msg2",null, null);
		final String msg3 = this.messageSource.getMessage("email.new.user.to.fac.msg3",null, null);
;
		final String msg5 = this.messageSource.getMessage("email.new.fac.msg5",null, null);
		MimeMessagePreparator preparator = new MimeMessagePreparator() {
			public void prepare(MimeMessage mimeMessage) throws Exception {
				MimeMessageHelper message = new MimeMessageHelper(mimeMessage, true);
				message.setTo(userEmail);
				message.setFrom(new InternetAddress(adminEmailId, "VMware Global Talent Development"));
				message.setSubject(emailSubject);
				StringBuilder mailBuilder = new StringBuilder("");	
				mailBuilder.append(msg1);
				mailBuilder.append(msg2);
				mailBuilder.append(msg3);				
				mailBuilder.append(msg5);
				message.setText(mailBuilder.toString(), true);
			}
		};
		this.mailSender.send(preparator);	
	}
	
	*//**
	 * This method used to send main to the facilitator when his account has been deactivated
	 * @param facilitatorEmail
	 * @throws MailingException
	 *//*
	private void sendMainToFacilitator(final String facilitatorEmail) {
		final String emailSubject = this.messageSource.getMessage("email.user.subject.msg",null, null);
		final String adminEmailId = this.messageSource.getMessage("email.msg.from",null, null);
		final String msg1 = this.messageSource.getMessage("email.facilitator.deactivated.msg1",null, null);
		final String msg2 = this.messageSource.getMessage("email.facilitator.deactivated.msg2",null, null);
		final String msg3 = this.messageSource.getMessage("email.facilitator.deactivated.msg3",null, null);
		final String msg5 = this.messageSource.getMessage("email.new.fac.msg5",null, null);
		
		MimeMessagePreparator preparator = new MimeMessagePreparator() {
			public void prepare(MimeMessage mimeMessage) throws Exception {
				MimeMessageHelper message = new MimeMessageHelper(mimeMessage, true);
				message.setTo(facilitatorEmail);
				message.setFrom(new InternetAddress(adminEmailId, "VMware Global Talent Development"));
				message.setSubject(emailSubject);
				StringBuilder mailBuilder = new StringBuilder("");
				mailBuilder.append(msg1);
				mailBuilder.append(msg2);
				mailBuilder.append(msg3);				
				mailBuilder.append(msg5);
				message.setText(mailBuilder.toString(), true);
			}
		};
		this.mailSender.send(preparator);
		
	}*/
	
	
	/**
	 * This method will intimate the administration via mail about the user registration 
	 * status
	 * @param userDTO
	 * @param adminMailId
	 * @param invalidEmailIds
	 * @throws MailingException
	 */
	public void intimateFacilitator(final NewUserDTO userDTO, final String adminMailId, 
			final ArrayList<String> invalidEmailIds) throws MailingException {
		try {
			helper.intimateFacilitatorAsync(userDTO, adminMailId, invalidEmailIds);
		} catch (MailingException mailingExp) {
			LOGGER.error("UNABLE TO SEND MAIL TO: "+adminMailId+". REASON: "+mailingExp.getLocalizedMessage());
		}
	}
	
	/**
	 * This method will renew/Subscribe existing facilitator
	 * @param facilitator email
	 * @param facilitator ID
	 * @param no of users
	 * @param inputRenewSub
	 * @throws MailingException
	 */
	public void renewSubscribeFacilitator(final String facilitatorEmail, final String facilitatorID, final int numberOfUsers,
			final String inputRenewSub, final String paymentType) throws MailingException {
		try {
			helper.renewSubscribeFacilitatorAsync(facilitatorEmail, facilitatorID, numberOfUsers, inputRenewSub, paymentType);
		} catch (MailingException mailingExp) {
			LOGGER.error("UNABLE TO SEND MAIL TO: "+facilitatorEmail+". REASON: "+mailingExp.getLocalizedMessage());
		}
	}
/**
 * method to notify user/facilitator about check dishonored via mail
 * @param userName
 * @param checkNo
 * @param role
 * @param facType
 * @throws MailingException
 */
public void intimateDishonoredUser(String userName, String fName, String checkNo, int role, String facType, String paymentDate) throws MailingException {
	try {
		helper.dishonoredCheckUser(userName, fName, checkNo, role, facType, paymentDate);
	} catch (MailingException mailingExp) {
		LOGGER.error("UNABLE TO SEND MAIL TO: "+userName+". REASON: "+mailingExp.getLocalizedMessage());
	}
	
}
/**
 * method to notify user/facilitator about check dishonored via mail
 * @param emailId
 * @throws MailingException
 */
public void userDisable(String emailId) throws MailingException {
	try {
		helper.userDisableMailDispatcher(emailId);
	} catch (MailingException mailingExp) {
		LOGGER.error("UNABLE TO SEND MAIL TO: "+emailId+". REASON: "+mailingExp.getLocalizedMessage());
	}
	
}



/**
 * This method will intimate the users about the generated password, via mail, on successful registration 
 * or reset password done by admin
 * @param to
 * @param password
 * @param discVal
 * @throws MailingException
 */
public void mailIndUserRegFromEcomerce(final String to, final String password) throws MailingException {
	try {
		helper.mailIndUserRegFromEcomerce(to, password);
	} catch (MailingException mailingExp) {
		LOGGER.error("UNABLE TO SEND MAIL TO: "+to+". REASON: "+mailingExp.getLocalizedMessage());
	}
}
/**
 * This method will intimate the users about having a duplicate account 
 * @param to
 * @param password
 * @param discVal
 * @throws MailingException
 */
public void mailIndUserDuplicateAccountEcomerce(final String to) throws MailingException {
	try {
		helper.mailIndUserDuplicateAccountEcomerce(to);
	} catch (MailingException mailingExp) {
		LOGGER.error("UNABLE TO SEND MAIL TO: "+to+". REASON: "+mailingExp.getLocalizedMessage());
	}
}
/**
 * This method will send mail at the time of reset password
 * @param to
 * @param password
 * @param discVal
 * @throws MailingException
 */
public void mailResetIndiFaci(final String emailId, final String password, final String fname, final String usrType) throws MailingException {
	try {
		helper.mailResetIndiFaciAsyn(emailId, password, fname, usrType);
	} catch (MailingException mailingExp) {
		LOGGER.error("UNABLE TO SEND MAIL TO: "+emailId+". REASON: "+mailingExp.getLocalizedMessage());
	}
}
/**
 * This method will send mail at the time of account deactivation
 * @param to
 * @param password
 * @param discVal
 * @throws MailingException
 */
public void mailAccDeactivate(final String emailId, final String fName) throws MailingException {
	try {
		helper.mailAccDeactivateAsyn(emailId, fName);
	} catch (MailingException mailingExp) {
		LOGGER.error("UNABLE TO SEND MAIL TO: "+emailId+". REASON: "+mailingExp.getLocalizedMessage());
	}
}
/**
 * This method will send mail when facilitator create an user
 * @param to
 * @param password
 * @param discVal
 * @throws MailingException
 */
public void intimateUserByFaciliator(final String emailId, final String password, final String dicVal, 
		final String fName, final String createdBy, final String faciFullName) throws MailingException {
	try {
		helper.intimateUserByFaciliatorAsyn(emailId, password, dicVal, fName, createdBy, faciFullName);
	} catch (MailingException mailingExp) {
		LOGGER.error("UNABLE TO SEND MAIL TO: "+emailId+". REASON: "+mailingExp.getLocalizedMessage());
	}
}
}