package com.vmware.jct.common.utility;

import java.text.SimpleDateFormat;
import java.util.ArrayList;
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
import org.springframework.scheduling.annotation.Async;
import org.springframework.stereotype.Component;

import com.vmware.jct.dao.dto.NewUserDTO;
import com.vmware.jct.exception.MailingException;

@Component
public class MailDispatcher {
	private static final Logger LOGGER = LoggerFactory.getLogger(MailDispatcher.class);
	
	@Autowired
	private JavaMailSender mailSender;
	
	@Autowired  
	private MessageSource messageSource;
	
	
	/**
	 * function to send mail for registration / check honored related mail for Facilitator
	 * @param to
	 * @param password
	 * @param discVal
	 * @param paymentType
	 * */
	
	@Async
	public void intimateUserAsync (final String to, final String password, final String fName, final String discVal, final String paymentType) throws MailingException {
		LOGGER.info(">>>>>>>>> MailNotificationHelper.intimateUser");
		String emailSubject = null;
		if (discVal.equals("register")) {
			if(null != paymentType) {
				if(paymentType.equals("CHK-CLEARED-INDIVIDUAL")){
					emailSubject = this.messageSource.getMessage("email.user.subject.ChkClered.subject",null, null);	//	 Congratulations! Your check has been cleared		
				} else {
					emailSubject = this.messageSource.getMessage("email.user.subject.register",null, null);				//	Registration - Job Crafting Tool
				}
			} else {
				emailSubject = this.messageSource.getMessage("email.user.subject.msg",null, null);
			}
						
		} else if (discVal.equals("ChkCleared")) {		
			emailSubject = this.messageSource.getMessage("email.user.subject.ChkClered.subject",null, null);
		} else {
			emailSubject = this.messageSource.getMessage("email.user.subject.msg.reset",null, null);
		}
		
		final String emailSubjectFinal = emailSubject;
		final String emailFrom = this.messageSource.getMessage("email.msg.from",null, null);
		final String messageHeaderIndi = this.messageSource.getMessage("email.msg.header.user",null, null);
		final String messageHeaderFaci = this.messageSource.getMessage("email.msg.header.Faci",null, null);
		final String messageSubHeaderThanks = this.messageSource.getMessage("email.msg.user.sub.header.thanks",null, null);
		final String messageSubHeaderThanksChkDone = this.messageSource.getMessage("email.msg.user.sub.header.thanksChkDone",null, null);
		final String messageSubHeaderEdit = this.messageSource.getMessage("email.msg.user.sub.header.edit",null, null);
		final String ChkCleredMsg = this.messageSource.getMessage("email.msg.user.info.ChkCleredMsg",null, null);
		final String messageInfoOne = this.messageSource.getMessage("email.msg.user.info.one",null, null);
		//final String messageInfoTwo = this.messageSource.getMessage("email.msg.user.info.two",null, null);
		//final String messageInfoThree = this.messageSource.getMessage("email.msg.user.info.three",null, null);
		final String messageInfoFour = this.messageSource.getMessage("email.msg.user.info.four",null, null); //link
		final String messageInfoFive = this.messageSource.getMessage("email.msg.user.info.five",null, null);
		final String messageInfoSix = this.messageSource.getMessage("email.msg.user.info.six",null, null);
		final String messageInfoSeven = this.messageSource.getMessage("email.msg.user.info.seven",null, null);
		final String emailFooter = this.messageSource.getMessage("email.msg.footer.general",null, null);		
		final String noPassword = this.messageSource.getMessage("email.msg.user.noPassword",null, null);
		
		MimeMessagePreparator preparator = new MimeMessagePreparator() {
			public void prepare(MimeMessage mimeMessage) throws Exception {
				MimeMessageHelper message = new MimeMessageHelper(mimeMessage, true);
				message.setTo(to);
				message.setFrom(new InternetAddress(emailFrom, "Job Crafting LLC"));
				message.setSubject(emailSubjectFinal);
				StringBuilder mailBuilder = new StringBuilder("");
				if(discVal.equals("resetFaci")){		
					mailBuilder.append(messageHeaderFaci + fName +",</br></br>");
				} else {
					mailBuilder.append(messageHeaderIndi + fName +",</br></br>");	
				}
				if (discVal.equals("register")) {	//	new registration
					if(null != paymentType) {
						if(paymentType.equals("CHK-CLEARED-INDIVIDUAL")){
							mailBuilder.append(messageSubHeaderThanksChkDone);
						} else {
							mailBuilder.append(messageSubHeaderThanks);						
						}
					} else {
						mailBuilder.append(messageSubHeaderThanks);		
					}
				} else if (discVal.equals("ChkCleared")) {		//	registration check cleared
					mailBuilder.append(messageSubHeaderThanksChkDone);	
				} else {		//	reset
					mailBuilder.append(messageSubHeaderEdit);
				}
				if(null !=  paymentType) {
					if(discVal.equals("ChkCleared")) {
						mailBuilder.append(ChkCleredMsg);
						mailBuilder.append(messageInfoOne);
						mailBuilder.append(password + "</i>");
					} else if(paymentType.equals("CHECK")) {	//	Check registration				
						mailBuilder.append(noPassword);					
					} else {		//	Free or Cash registration
						mailBuilder.append(messageInfoOne);
						mailBuilder.append(password+"</i>");			
					}				
				} else {			//		in case of reset password
					mailBuilder.append(messageInfoOne);
					mailBuilder.append(password + "</i>");			
				}					
//				link for log-in will be provided after registration except CHECK payment
				if( null != paymentType) {
					if((!paymentType.equals("CHECK"))) {						
						//	sign in link
						mailBuilder.append(messageInfoFour);
						mailBuilder.append(to);
						mailBuilder.append(messageInfoFive);
						mailBuilder.append(messageInfoSix);
						mailBuilder.append(messageInfoSeven);
						mailBuilder.append(to+"</u></span>");								
					}	
				} else {
					if(!discVal.equals("resetFaci")){	//	faci reset password mail will not contain password reset link
						mailBuilder.append(messageInfoFour);
						mailBuilder.append(to);
						mailBuilder.append(messageInfoFive);
						mailBuilder.append(messageInfoSix);
						mailBuilder.append(messageInfoSeven);
						mailBuilder.append(to);
					}}
								
				mailBuilder.append(emailFooter);
				message.setText(mailBuilder.toString(), true);
			}
		};
		this.mailSender.send(preparator);
		LOGGER.info("======= MAIL SENT TO: "+to+" =======");
		LOGGER.info("<<<<<<<<< MailNotificationHelper.intimateUser");
	}

	/**
	 * function to send mail for check honored related mail for Facilitator
	 * @param to
	 * @param password
	 * @param discVal
	 * @param paymentType
	 * */	
		@Async
		public void intimateFaciAsync(final String to, final String password, final String fName, final String discVal, final String paymentType) throws MailingException {
			LOGGER.info(">>>>>>>>> MailNotificationHelper.intimateFaciAsync");			
			final String emailFrom = this.messageSource.getMessage("email.msg.from",null, null);
			final String ChkCleredMsgRegistration = this.messageSource.getMessage("email.faci.chkCleared.bodyheader.register",null, null);
			final String dear = this.messageSource.getMessage("email.faci.chkCleared.bodyheader.dear",null, null);
			final String ChkCleredMsgRenewsub = this.messageSource.getMessage("email.faci.chkCleared.bodyheader.renewSub",null, null);
			final String emailFooter = this.messageSource.getMessage("email.msg.footer.general",null, null);			
			final String emailSubject = this.messageSource.getMessage("email.fac.ChkCleredMsg.subject", null,null);
			//  link
			final String messageInfoThree = this.messageSource.getMessage("email.msg.user.info.three.faci",null, null);
			final String messageInfoFour = this.messageSource.getMessage("email.msg.user.info.four.faci",null, null); 
			final String messageInfoFive = this.messageSource.getMessage("email.msg.user.info.five",null, null);
			final String messageInfoSix = this.messageSource.getMessage("email.msg.user.info.six",null, null);
			//final String messageInfoSix_link = this.messageSource.getMessage("email.msg.user.info.seven.faci",null, null);			
			final String messageInfoSeven = this.messageSource.getMessage("email.msg.user.info.seven.faci",null, null);
			
			MimeMessagePreparator preparator = new MimeMessagePreparator() {
				public void prepare(MimeMessage mimeMessage) throws Exception {
					MimeMessageHelper message = new MimeMessageHelper(mimeMessage, true);
					message.setTo(to);
					message.setFrom(new InternetAddress(emailFrom, "Job Crafting LLC"));
					message.setSubject(emailSubject);
					StringBuilder mailBuilder = new StringBuilder("");
									
					if(null !=  paymentType) {
						if (discVal.equals("register")) {
							mailBuilder.append(dear);
							mailBuilder.append(fName + ",");
							mailBuilder.append(ChkCleredMsgRegistration);
							mailBuilder.append(password+"</i></b>");
							if(!paymentType.equals("CHECK")){
								//		login link
								mailBuilder.append(messageInfoThree);
								mailBuilder.append(messageInfoFour);
								mailBuilder.append(to);
								mailBuilder.append(messageInfoFive);
								mailBuilder.append(messageInfoSix);
								mailBuilder.append(messageInfoSeven);
								mailBuilder.append(to);	
								mailBuilder.append("</span></u>");
							}
						} else {							
							mailBuilder.append(ChkCleredMsgRenewsub);
						}								
					}
					mailBuilder.append(emailFooter);
					message.setText(mailBuilder.toString(), true);
				}
			};			
			this.mailSender.send(preparator);
			LOGGER.info("======= MAIL SENT TO: "+to+" =======");
			LOGGER.info("<<<<<<<<< MailNotificationHelper.intimateFaciAsync");
			
		}
/**
* Method to send mail to Individual User after cheque clearing
* @param userDTO
* @param adminMailId
* @param invalidEmailIds
*/	
	@Async
	public void intimateAdminAsyc(final NewUserDTO userDTO, final String adminMailId, 
			final ArrayList<String> invalidEmailIds) throws MailingException {
		LOGGER.info(">>>>>>>>> MailNotificationHelper.intimateAdmin");
		final String messageSubject = this.messageSource.getMessage("email.msg.subject.admin",null, null);
		final String messageHeader = this.messageSource.getMessage("email.msg.header.admin",null, null);
		final String messageSubHeader = this.messageSource.getMessage("email.msg.sub.header.admin",null, null);
		final String emailIdPresent = this.messageSource.getMessage("email.msg.email.id.present.admin",null, null);
		final String emailIdInvalid = this.messageSource.getMessage("email.msg.email.id.invalid.admin",null, null);
		final String emailFooter = this.messageSource.getMessage("email.msg.footer.general",null, null);
		final String emailFrom = this.messageSource.getMessage("email.msg.from",null, null);
		MimeMessagePreparator preparator = new MimeMessagePreparator() {
			public void prepare(MimeMessage mimeMessage) throws Exception {
				MimeMessageHelper message = new MimeMessageHelper(mimeMessage, true);
				message.setTo(adminMailId);
				message.setFrom(new InternetAddress(emailFrom, "Job Crafting LLC"));
				message.setSubject(messageSubject);
				List<NewUserDTO> alreadyRegList = userDTO.getAlreadyRegistered();
				StringBuilder tableBuilder = new StringBuilder("");
				if (alreadyRegList.size() > 0) {
					tableBuilder = getAlreadyRegistrationTable(tableBuilder, alreadyRegList);
				}
				StringBuilder mailBuilder = new StringBuilder("");
				mailBuilder.append(messageHeader);
				mailBuilder.append(messageSubHeader);
				if (tableBuilder.toString().trim().length() > 0) {
					mailBuilder.append(emailIdPresent);
					mailBuilder.append(tableBuilder.toString());
				}
				if (invalidEmailIds.size() > 0 ){
					mailBuilder.append(emailIdInvalid);
					for (int index=0; index<invalidEmailIds.size(); index++) {
						if (invalidEmailIds.size()-1 == index) {
							mailBuilder.append(invalidEmailIds.get(index).toString());
						} else {
							mailBuilder.append(invalidEmailIds.get(index).toString()+",");
						}
					}
				}
				mailBuilder.append(emailFooter);
				message.setText(mailBuilder.toString(), true);
			}
		};
		LOGGER.info("SENDING MAIL TO ADMIN...");
		this.mailSender.send(preparator);
		LOGGER.info("MAIL SENT TO ADMIN...");
		LOGGER.info("<<<<<<<<< MailNotificationHelper.intimateAdmin");
	}

	/**
	 * Method returns already registered users
	 * @param builder
	 * @param dtoList
	 * @return String of list of users comma separated
	 */
	private StringBuilder getAlreadyRegistrationTable(StringBuilder builder, List<NewUserDTO> dtoList) {
		LOGGER.info(">>>>>>>>> MailNotificationHelper.getAlreadyRegistrationTable");
		Iterator<NewUserDTO> itr = dtoList.iterator();
		int index = 1;
		int size = dtoList.size();
		int nextLine = 1;
		while (itr.hasNext()) {
			NewUserDTO dto = (NewUserDTO) itr.next();
			if(index != size) {
				builder.append(dto.getEmailId()+" , ");
			} else if(nextLine == 5) {
				builder.append(dto.getEmailId() +", <br />");
				nextLine = 1;
			} else {
				builder.append(dto.getEmailId());
			}
			index = index + 1;
			nextLine = nextLine + 1;
		}
		LOGGER.info("<<<<<<<<< MailNotificationHelper.getAlreadyRegistrationTable");
		return builder;
	}
	/**
	 * This method will intimate the new facilitator
	 * @param to
	 * @param password
	 * @param discVal
	 * @throws MailingException
	 */
	
	@Async
	public void intimateNewFacilitatorAsync(final String facilitatorEmail, final String password, final String fName, final int totalSubscribedUsers, 
			final String opetedAccount,final String facilitatorId,final String paymentType) throws MailingException {
		LOGGER.info(">>>>>>>>> MailNotificationHelper.intimateNewFacilitator");	
		final String emailSubject = this.messageSource.getMessage("email.user.subject.register.facilitator",null, null);
		final String adminEmailId = this.messageSource.getMessage("email.msg.from",null, null);		
		final String body1 = this.messageSource.getMessage("email.reg.faci.body1",null, null);		
		final String body2_1 = this.messageSource.getMessage("email.reg.faci.body2.1",null, null);
		final String body2_2 = this.messageSource.getMessage("email.reg.faci.body2.2",null, null);
		//final String body3 = this.messageSource.getMessage("email.reg.faci.body3",null, null);	//	cust Id	
		final String body4 = this.messageSource.getMessage("email.reg.faci.body4",null, null);
		final String body4_1 = this.messageSource.getMessage("email.reg.faci.body4.1",null, null);
		final String body5 = this.messageSource.getMessage("email.reg.faci.body5",null, null);
		final String footer = this.messageSource.getMessage("email.reg.faci.footer",null, null);
		final String noPassword = this.messageSource.getMessage("email.msg.user.noPassword",null, null);
		
		
		final String genUserExist = this.messageSource.getMessage("msg.registered8",null, null);
		
		MimeMessagePreparator preparator = new MimeMessagePreparator() {
			public void prepare(MimeMessage mimeMessage) throws Exception {
				MimeMessageHelper message = new MimeMessageHelper(mimeMessage, true);
				message.setTo(facilitatorEmail);
				message.setFrom(new InternetAddress(adminEmailId, "Job Crafting LLC"));
				message.setSubject(emailSubject);
				StringBuilder mailBuilder = new StringBuilder("");
				mailBuilder.append(body1);
				mailBuilder.append(fName);
				mailBuilder.append(body2_1);
				if(paymentType.equals("CHECK")){	//	if payment type == CHECK
					mailBuilder.append(noPassword);										
				} else {
					mailBuilder.append(body2_2);
					mailBuilder.append(password + "</b></i>");
					mailBuilder.append(body4);			//	hyperlink
					mailBuilder.append(facilitatorEmail);
					mailBuilder.append(body4_1);
					mailBuilder.append(body5);
					mailBuilder.append(facilitatorEmail);
				}				
				//mailBuilder.append(body3);
				//mailBuilder.append(facilitatorId);				
				mailBuilder.append("</u></span><br>");
								
				/*if (opetedAccount.equals("USER-EXIST")){
					mailBuilder.append(genUserExist);					
				}*/
				mailBuilder.append(footer);
				message.setText(mailBuilder.toString(), true);
			}
		};
		this.mailSender.send(preparator);		
		LOGGER.info("<<<<<<<< MailNotificationHelper.intimateNewFacilitator");	
	}	
	
	/**
	 * This method will intimate the administration via mail about the user registration 
	 * status
	 * @param userDTO
	 * @param adminMailId
	 * @param invalidEmailIds
	 * @throws MailingException
	 */
	@Async
	public void intimateAdminGeneralUserRegistrationAsync(final List<NewUserDTO> userDTOList, final String adminMailId, 
			final ArrayList<String> invalidEmailIds) throws MailingException {
		LOGGER.info(">>>>>>>>> MailNotificationHelper.intimateAdminGeneralUserRegistration");
		final String messageSubject = this.messageSource.getMessage("email.msg.subject.admin",null, null);
		final String messageHeader = this.messageSource.getMessage("email.msg.header.admin",null, null);
		final String messageSubHeader = this.messageSource.getMessage("email.msg.sub.header.admin",null, null);
		final String emailIdPresent = this.messageSource.getMessage("email.msg.email.id.present.admin",null, null);
		final String emailIdInvalid = this.messageSource.getMessage("email.msg.email.id.invalid.admin",null, null);
		final String emailFooter = this.messageSource.getMessage("email.msg.footer.general",null, null);
		final String emailFrom = this.messageSource.getMessage("email.msg.from",null, null);
		MimeMessagePreparator preparator = new MimeMessagePreparator() {
			public void prepare(MimeMessage mimeMessage) throws Exception {
				MimeMessageHelper message = new MimeMessageHelper(mimeMessage, true);
				message.setTo(adminMailId);
				message.setFrom(new InternetAddress(emailFrom, "Job Crafting LLC"));
				message.setSubject(messageSubject);
				StringBuilder tableBuilder = new StringBuilder("");
				if (userDTOList.size() > 0) {
					tableBuilder = getAlreadyRegistrationTable(tableBuilder, userDTOList);
				}
				StringBuilder mailBuilder = new StringBuilder("");
				mailBuilder.append(messageHeader);
				mailBuilder.append(messageSubHeader);
				if (tableBuilder.toString().trim().length() > 0) {
					mailBuilder.append(emailIdPresent);
					mailBuilder.append(tableBuilder.toString());
				}
				if (invalidEmailIds.size() > 0 ){
					mailBuilder.append(emailIdInvalid);
					for (int index=0; index<invalidEmailIds.size(); index++) {
						if (invalidEmailIds.size()-1 == index) {
							mailBuilder.append(invalidEmailIds.get(index).toString());
						} else {
							mailBuilder.append(invalidEmailIds.get(index).toString()+",");
						}
					}
				}
				mailBuilder.append(emailFooter);
				message.setText(mailBuilder.toString(), true);
			}
		};
		LOGGER.info("SENDING MAIL TO ADMIN...");
		this.mailSender.send(preparator);
		LOGGER.info("MAIL SENT TO ADMIN...");
		LOGGER.info("<<<<<<<<< MailNotificationHelper.intimateAdminGeneralUserRegistration");
	}
	/**
	 * This method will send main on user renewal
	 * @param userEmail
	 * @param renewDate
	 * @throws MailingException
	 */
	@Async
	public void sendMailToRenewedUserAsync(final String userEmail, final Date renewDate, final String paymentType) throws MailingException {
		LOGGER.info(">>>>>>>>> MailNotificationHelper.intimateNewFacilitator");
		final SimpleDateFormat obDateFormat = new SimpleDateFormat("MMM d, yyyy");
		final String emailSubject = this.messageSource.getMessage("email.user.subject.msg",null, null);
		final String adminEmailId = this.messageSource.getMessage("email.msg.from",null, null);
		final String msg1 = this.messageSource.getMessage("email.new.user.renew.msg1",null, null);
		final String msg2 = this.messageSource.getMessage("email.new.user.renew.msg2",null, null);
		final String msg3 = this.messageSource.getMessage("email.new.user.renew.msg3",null, null);
		final String msg5 = this.messageSource.getMessage("email.new.fac.msg5",null, null);
		final String msgCheck = this.messageSource.getMessage("email.new.user.renew.check.under.process",null, null);
		MimeMessagePreparator preparator = new MimeMessagePreparator() {
			public void prepare(MimeMessage mimeMessage) throws Exception {
				MimeMessageHelper message = new MimeMessageHelper(mimeMessage, true);
				message.setTo(userEmail);
				message.setFrom(new InternetAddress(adminEmailId, "Job Crafting LLC"));
				message.setSubject(emailSubject);
				StringBuilder mailBuilder = new StringBuilder("");	
				mailBuilder.append(msg1);
				mailBuilder.append(msg2);				
				if(null != paymentType) {
					if(paymentType.equalsIgnoreCase("CHECK")) {
						mailBuilder.append(msgCheck);	
					} else {
						mailBuilder.append(msg3);	
						mailBuilder.append(obDateFormat.format(renewDate));
					}
				} else {
					mailBuilder.append(msg3);	
					mailBuilder.append(obDateFormat.format(renewDate));
				}
				mailBuilder.append(msg5);
				message.setText(mailBuilder.toString(), true);
			}
		};
		this.mailSender.send(preparator);
		LOGGER.info("<<<<<<<<< MailNotificationHelper.intimateNewFacilitator");
	}
	/**
	 * This method will intimate the administration via mail about the user registration 
	 * status
	 * @param userDTO
	 * @param adminMailId
	 * @param invalidEmailIds
	 * @throws MailingException
	 */
	@Async
	public void intimateFacilitatorAsync(final NewUserDTO userDTO, final String adminMailId, 
			final ArrayList<String> invalidEmailIds) throws MailingException {
		LOGGER.info(">>>>>>>>> MailNotificationHelper.intimateAdmin");
		final String messageSubject = this.messageSource.getMessage("email.msg.subject.admin",null, null);
		final String messageHeader = this.messageSource.getMessage("email.msg.header.admin",null, null);
		final String messageSubHeader = this.messageSource.getMessage("email.msg.sub.header.admin",null, null);
		final String emailIdPresent = this.messageSource.getMessage("email.msg.email.id.present.admin",null, null);
		final String emailIdInvalid = this.messageSource.getMessage("email.msg.email.id.invalid.admin",null, null);
		final String emailFooter = this.messageSource.getMessage("email.msg.footer.general",null, null);
		final String emailFrom = this.messageSource.getMessage("email.msg.from",null, null);
		MimeMessagePreparator preparator = new MimeMessagePreparator() {
			public void prepare(MimeMessage mimeMessage) throws Exception {
				MimeMessageHelper message = new MimeMessageHelper(mimeMessage, true);
				message.setTo(adminMailId);
				message.setFrom(new InternetAddress(emailFrom, "Job Crafting LLC"));
				message.setSubject(messageSubject); 
				List<NewUserDTO> alreadyRegList = null;
				if (userDTO.getAlreadyRegistered() != null){
					alreadyRegList = userDTO.getAlreadyRegistered();
				}		
				StringBuilder tableBuilder = new StringBuilder("");
				if (alreadyRegList != null) {
					tableBuilder = getAlreadyRegistrationTable(tableBuilder, alreadyRegList);
				}
				
				StringBuilder mailBuilder = new StringBuilder("");
				mailBuilder.append(messageHeader);
				mailBuilder.append(messageSubHeader);
				if (tableBuilder.toString().trim().length() > 0) {
					mailBuilder.append(emailIdPresent);
					mailBuilder.append(tableBuilder.toString());
				}
				if (invalidEmailIds.size() > 0 ){
					mailBuilder.append(emailIdInvalid);
					for (int index=0; index<invalidEmailIds.size(); index++) {
						if (invalidEmailIds.size()-1 == index) {
							mailBuilder.append(invalidEmailIds.get(index).toString());
						} else {
							mailBuilder.append(invalidEmailIds.get(index).toString()+",");
						}
					}
				}
				mailBuilder.append(emailFooter);
				message.setText(mailBuilder.toString(), true);
			}
		};
		LOGGER.info("SENDING MAIL TO ADMIN...");
		this.mailSender.send(preparator);
		LOGGER.info("MAIL SENT TO ADMIN...");
		LOGGER.info("<<<<<<<<< MailNotificationHelper.intimateAdmin");
	}
	/**
	 * This method will renew/Subscribe existing facilitator
	 * @param facilitator email
	 * @param facilitator ID
	 * @param no of users
	 * @param inputRenewSub
	 * @throws MailingException
	 */
	@Async
	public void renewSubscribeFacilitatorAsync(final String facilitatorEmail, final String facilitatorID, final int numberOfUsers,
			final String inputRenewSub, final String paymentType) throws MailingException {		
		LOGGER.info(">>>>>>>>> MailNotificationHelper.renewSubscribeFacilitator");
		
		final String adminEmailId = this.messageSource.getMessage("email.renewSubscribe.from",null, null);				
		final String body1 = this.messageSource.getMessage("email.renewSub.body1",null, null);
		final String renewMsg = this.messageSource.getMessage("email.renewSub.renewMsg",null, null);
		final String subMsg = this.messageSource.getMessage("email.renewSub.subMsg",null, null);
		final String body2 = this.messageSource.getMessage("email.renewSub.body2",null, null);
		final String noOfUserMsg = this.messageSource.getMessage("email.renewSub.noOfUserMsg",null, null);
		final String noOfRenewedUser = this.messageSource.getMessage("email.renewSub.noOfRenewedUser",null, null);
		final String bodyFooter = this.messageSource.getMessage("email.renewSub.bodyFooter",null, null);
		final String emailSubject = this.messageSource.getMessage("email.renewSubscribe.subject",null, null);
 		final String chkPaymentMsg = this.messageSource.getMessage("error.renewSubs.ChequeProblem",null, null);
		
		final String noOfRenewSubMsg = inputRenewSub.equals("RW") ? noOfRenewedUser : noOfUserMsg;		
		MimeMessagePreparator preparator = new MimeMessagePreparator() {			
			public void prepare(MimeMessage mimeMessage) throws Exception {
				MimeMessageHelper message = new MimeMessageHelper(mimeMessage, true);
				message.setTo(facilitatorEmail);
				message.setFrom(new InternetAddress(adminEmailId, "Job Crafting LLC"));	
				message.setSubject(emailSubject);
				
				StringBuilder mailBuilder = new StringBuilder("");				
				mailBuilder.append(body1);
				if( inputRenewSub.equals("RW")) {
					mailBuilder.append(renewMsg);
				} else {
					mailBuilder.append(subMsg);
				}
				mailBuilder.append(body2);			
				if(numberOfUsers > 0) {
					mailBuilder.append(noOfRenewSubMsg + " " + numberOfUsers);				
				}
				if(paymentType.equals("CHECK")){
					mailBuilder.append(chkPaymentMsg);	
				}
				mailBuilder.append(bodyFooter);
				message.setText(mailBuilder.toString(), true);
			}
		};
		this.mailSender.send(preparator);
		LOGGER.info("<<<<<<<<< MailNotificationHelper.renewSubscribeFacilitator");
	}
	
	@Async
	public void dishonoredCheckUser(final String userName, final String fName, final String checkNo, final int role, final String facType, final String paymentDate) throws MailingException {
		
		LOGGER.info(">>>>>>>>> MailNotificationHelper.dishonoredCheckUser");
		final String adminEmailId = this.messageSource.getMessage("email.renewSubscribe.from",null, null);
		final String emailSubject = this.messageSource.getMessage("label.check.dishonored.subject",null, null);
		final String dear = this.messageSource.getMessage("label.ckeck.dishonored.dear",null, null);
		final String body1 = this.messageSource.getMessage("label.check.dishonored.body1",null, null);
		final String registraionFailed = this.messageSource.getMessage("label.check.dishonored.registrationFailed",null, null);
		final String subRenewFailed= this.messageSource.getMessage("label.check.dishonored.subRenewFailed",null, null);
		final String faciIndiRegistratnFailed = this.messageSource.getMessage("label.check.dishonored.faciIndiRegistratnFailed",null, null);
		final String contactAdmin = this.messageSource.getMessage("label.check.dishonored.contactAdmin",null, null);
		final String paymentDateMsg = this.messageSource.getMessage("label.check.dishonored.paymentDate",null, null);		
		final String bodyFooter = this.messageSource.getMessage("label.check.dishonored.bodyFooter",null, null);
		
		MimeMessagePreparator preparator = new MimeMessagePreparator() {
			public void prepare(MimeMessage mimeMessage) throws Exception {
				MimeMessageHelper message = new MimeMessageHelper(mimeMessage, true);
				message.setTo(userName);
				message.setFrom(new InternetAddress(adminEmailId, "Job Crafting LLC"));
				message.setSubject(emailSubject);
				StringBuilder mailBuilder = new StringBuilder("");			
				//role
				mailBuilder.append(dear);
				mailBuilder.append(" " + fName + ",");
				mailBuilder.append(body1);
				if (facType.equals("NEW_FACILITATOR")) {	//	message for new individual registration
					mailBuilder.append(registraionFailed);					
				} else if (facType.equals("NEW_USER")) {	//	message for new individual registration
					mailBuilder.append(registraionFailed);
				} else {		// message for	subscribe / renew more user
					mailBuilder.append(subRenewFailed);
				}
				mailBuilder.append(" "+checkNo);
				if( role == 5 ){
					mailBuilder.append(faciIndiRegistratnFailed);
				}
				mailBuilder.append(paymentDateMsg);
				mailBuilder.append(" "+paymentDate);				
				mailBuilder.append(contactAdmin);
				mailBuilder.append(bodyFooter);
				message.setText(mailBuilder.toString(), true);
			}
		};
		this.mailSender.send(preparator);
		LOGGER.info("<<<<<<<<< MailNotificationHelper.dishonoredCheckUser");
		
	}
	@Async	
	public void userDisableMailDispatcher(final String emailId) throws MailingException {
		
		LOGGER.info(">>>>>>>>> MailNotificationHelper.userDisableMailDispatcher");
		final String adminEmailId = this.messageSource.getMessage("email.renewSubscribe.from",null, null);
		final String emailSubject = this.messageSource.getMessage("label.refund.subject",null, null);
		final String body1 = this.messageSource.getMessage("email.refund.body1",null, null);
		final String body2 = this.messageSource.getMessage("email.refund.body2",null, null);		
		final String contactAdmin = this.messageSource.getMessage("email.refund.body3",null, null);	
		final String bodyFooter = this.messageSource.getMessage("email.refund.thanksAndFooter",null, null);
		
 		
		MimeMessagePreparator preparator = new MimeMessagePreparator() {
			public void prepare(MimeMessage mimeMessage) throws Exception {
				MimeMessageHelper message = new MimeMessageHelper(mimeMessage, true);
				message.setTo(emailId);
				message.setFrom(new InternetAddress(adminEmailId, "Job Crafting LLC"));
				message.setSubject(emailSubject);
				
				StringBuilder mailBuilder = new StringBuilder("");
				mailBuilder.append(body1);
				mailBuilder.append(body2);				
				mailBuilder.append(contactAdmin);
				mailBuilder.append(bodyFooter);
				message.setText(mailBuilder.toString(), true);
			}
		};
		this.mailSender.send(preparator);
		LOGGER.info("<<<<<<<<< MailNotificationHelper.userDisableMailDispatcher");
		
	}

	/**
	 * This method will intimate the new facilitator
	 * @param userEmail
	 * @param facilitatorEmail
	 * @throws MailingException
	 */
	@Async
	public void intimateUserToFacilitatorAsync(final String userEmail,  final String facilitatorEmail
			) throws MailingException {
	LOGGER.info(">>>>>>>>> MailNotificationHelper.intimateNewFacilitator");
	sendMainToFacilitator(facilitatorEmail);
	sendMainToUser(userEmail);
	LOGGER.info("<<<<<<<<< MailNotificationHelper.intimateNewFacilitator");
	}

/**
 * This method used to send main to the user when the user is selected as facilitator 
 * @param userEmail
 * @throws MailingException
 */
private void sendMainToUser(final String userEmail) {
	final String emailSubject = this.messageSource.getMessage("email.user.subject.msg",null, null);
	final String adminEmailId = this.messageSource.getMessage("email.msg.from",null, null);
	final String msg1 = this.messageSource.getMessage("email.new.user.to.fac.msg1",null, null);
	final String msg2 = this.messageSource.getMessage("email.new.user.to.fac.msg2",null, null);
	final String msg3 = this.messageSource.getMessage("email.new.user.to.fac.msg3",null, null);
	final String msg5 = this.messageSource.getMessage("email.new.fac.msg5",null, null);
	MimeMessagePreparator preparator = new MimeMessagePreparator() {
		public void prepare(MimeMessage mimeMessage) throws Exception {
			MimeMessageHelper message = new MimeMessageHelper(mimeMessage, true);
			message.setTo(userEmail);
			message.setFrom(new InternetAddress(adminEmailId, "Job Crafting LLC"));
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

/**
 * This method used to send mail to the facilitator when his account has been deactivated
 * @param facilitatorEmail
 * @throws MailingException
 */
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
			message.setFrom(new InternetAddress(adminEmailId, "Job Crafting LLC"));
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

/**
 * function to send mail for registration / check honored related mail for Facilitator
 * @param to
 * @param password
 * @param discVal
 * @param paymentType
 * */

@Async
public void mailIndUserRegFromEcomerce (final String to, final String password) throws MailingException {
	LOGGER.info(">>>>>>>>> MailDispatch.mailIndUserRegFromEcomerce");
	String emailSubject = null;
	emailSubject = this.messageSource.getMessage("email.user.subject.msg",null, null);	//	tanks for registering				
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
			message.setTo(to);
			message.setFrom(new InternetAddress(emailFrom, "Job Crafting LLC"));
			message.setSubject(emailSubjectFinal);
			StringBuilder mailBuilder = new StringBuilder("");
			mailBuilder.append(messageHeader);
			mailBuilder.append(messageSubHeaderThanks);		
			mailBuilder.append(messageInfoOne);
			mailBuilder.append(password);			
			mailBuilder.append(messageInfoTwo);
			mailBuilder.append(messageInfoThree);
			mailBuilder.append(messageInfoFour);
			mailBuilder.append(to);
			mailBuilder.append(messageInfoFive);
			mailBuilder.append(messageInfoSix);
			mailBuilder.append(messageInfoSeven);
			mailBuilder.append(to);		
			mailBuilder.append(emailFooter);
			message.setText(mailBuilder.toString(), true);
		}
	};
	this.mailSender.send(preparator);
	LOGGER.info("======= MAIL SENT TO: "+to+" =======");
	LOGGER.info("<<<<<<<<< MailDispatch.mailIndUserRegFromEcomerce");
}

/**
 * function to send mail for registration / check honored related mail for Facilitator
 * @param to
 * @param password
 * @param discVal
 * @param paymentType
 * */
	@Async
	public void mailIndUserDuplicateAccountEcomerce (final String to) throws MailingException {
		LOGGER.info(">>>>>>>>> MailDispatch.mailIndUserDuplicateAccountEcomerce");
		String emailSubject = null;
		emailSubject = this.messageSource.getMessage("email.user.duplicate.subject.msg",null, null);	//	Unable to register				
		final String emailSubjectFinal = emailSubject;
		final String emailFrom = this.messageSource.getMessage("email.msg.from",null, null);
		final String messageHeader = this.messageSource.getMessage("email.msg.header.user",null, null);
		final String messageInfoOne = this.messageSource.getMessage("email.msg.user.duplicate.info.one",null, null);
		final String emailFooter = this.messageSource.getMessage("email.msg.footer.general",null, null);		
		
		MimeMessagePreparator preparator = new MimeMessagePreparator() {
			public void prepare(MimeMessage mimeMessage) throws Exception {
				MimeMessageHelper message = new MimeMessageHelper(mimeMessage, true);
				message.setTo(to);
				message.setFrom(new InternetAddress(emailFrom, "Job Crafting LLC"));
				message.setSubject(emailSubjectFinal);
				StringBuilder mailBuilder = new StringBuilder("");
				mailBuilder.append(messageHeader);
				mailBuilder.append(messageInfoOne);
				mailBuilder.append(emailFooter);
				message.setText(mailBuilder.toString(), true);
			}
		};
		this.mailSender.send(preparator);
		LOGGER.info("======= MAIL SENT TO: "+to+" =======");
		LOGGER.info("<<<<<<<<< MailDispatch.mailIndUserDuplicateAccountEcomerce");
	}
	/**
	 * This method will send mail at the time of reset password
	 * @param to
	 * @param password
	 * @param usrType
	 * */
		@Async
		public void mailResetIndiFaciAsyn (final String emailId, final String password, final String fname,
				final String usrType) throws MailingException {
			LOGGER.info(">>>>>>>>> MailDispatch.mailResetIndiFaciAsyn");
			String emailSubject = null;
			emailSubject = this.messageSource.getMessage("email.user.subject.msg.reset",null, null);	//	Unable to register				
			final String emailSubjectFinal = emailSubject;
			final String emailFrom = this.messageSource.getMessage("email.msg.from",null, null);
			final String messageHeader = this.messageSource.getMessage("email.msg.header.user",null, null);
			final String body2 = this.messageSource.getMessage("email.reset.password",null, null);
			final String body3 = this.messageSource.getMessage("email.msg.user.info.one",null, null);
			final String body4 = this.messageSource.getMessage("email.msg.user.info.five",null, null);
			final String body6 = this.messageSource.getMessage("email.msg.user.info.six",null, null);
			final String footer = this.messageSource.getMessage("email.reg.faci.footer",null, null);
			
			String rawLink = "";
			String link = "";
			if(usrType.equals("reset")){ //	reset individual
				rawLink = this.messageSource.getMessage("email.msg.user.info.four",null, null);		
				link = this.messageSource.getMessage("email.msg.user.info.seven",null, null);		
			} else {	
				rawLink = this.messageSource.getMessage("email.msg.user.info.four.faci",null, null);			
				link = this.messageSource.getMessage("email.msg.user.info.seven.faci",null, null);			
			}		
			final String rawLinkFinal = rawLink;
			final String linkFinal = link;
			
			MimeMessagePreparator preparator = new MimeMessagePreparator() {
				public void prepare(MimeMessage mimeMessage) throws Exception {
					MimeMessageHelper message = new MimeMessageHelper(mimeMessage, true);
					message.setTo(emailId);
					message.setFrom(new InternetAddress(emailFrom, "Job Crafting LLC"));
					message.setSubject(emailSubjectFinal);
					StringBuilder mailBuilder = new StringBuilder("");					
					mailBuilder.append(messageHeader);
					mailBuilder.append(fname + ",");
					mailBuilder.append(body2);					
					mailBuilder.append(body3);					
					mailBuilder.append(password+"<br></i>");					
					mailBuilder.append(rawLinkFinal);	//	link
					mailBuilder.append(emailId);
					mailBuilder.append(body4);
					mailBuilder.append(body6);
					mailBuilder.append(linkFinal);	//	link
					mailBuilder.append(emailId);
					mailBuilder.append(footer);					
					message.setText(mailBuilder.toString(), true);					
				}
			};
			this.mailSender.send(preparator);
			LOGGER.info("======= MAIL SENT TO: "+emailId+" =======");
			LOGGER.info("<<<<<<<<< MailDispatch.mailResetIndiFaciAsyn");
		}
/**
 * This method will send mail at the time of Account deactivation by admin / facilitator
 * @param to
 * @param password
 * @param usrType
 * */
	@Async
	public void mailAccDeactivateAsyn (final String emailId, final String fName) throws MailingException {
		LOGGER.info(">>>>>>>>> MailDispatch.mailResetIndiFaciAsyn");
		String emailSubject = null;
		emailSubject = this.messageSource.getMessage("email.user.subject.msg.reset",null, null);	//	Unable to register				
		final String emailSubjectFinal = emailSubject;
		final String emailFrom = this.messageSource.getMessage("email.msg.from",null, null);
		final String messageHeader = this.messageSource.getMessage("email.msg.header.user",null, null);
		final String body1 = this.messageSource.getMessage("email.deActive.acc.body1",null, null);
		final String body2 = this.messageSource.getMessage("email.deActive.acc.body2",null, null);
		final String footer = this.messageSource.getMessage("email.reg.faci.footer",null, null);
		MimeMessagePreparator preparator = new MimeMessagePreparator() {
			public void prepare(MimeMessage mimeMessage) throws Exception {
				MimeMessageHelper message = new MimeMessageHelper(mimeMessage, true);
				message.setTo(emailId);
				message.setFrom(new InternetAddress(emailFrom, "Job Crafting LLC"));
				message.setSubject(emailSubjectFinal);
				StringBuilder mailBuilder = new StringBuilder("");					
				mailBuilder.append(messageHeader + " " + fName + ",");
				mailBuilder.append(body1);					
				mailBuilder.append(body2);
				mailBuilder.append(footer);					
				message.setText(mailBuilder.toString(), true);					
			}
		};
		this.mailSender.send(preparator);
		LOGGER.info("======= MAIL SENT TO: "+emailId+" =======");
		LOGGER.info("<<<<<<<<< MailDispatch.mailResetIndiFaciAsyn");
	}

	
	/**
	 * This method will send registration mail to user when registered by facilitator
	 * @param to
	 * @param password
	 * @param usrType
	 * */
	@Async	
public void intimateUserByFaciliatorAsyn(final String emailId, final String password,
		final String dicVal, final String fName, final String createdBy, final String faciName) throws MailingException {
	LOGGER.info(">>>>>>>>> MailDispatch.intimateUserByFaciliatorAsyn");
	String emailSubject = null;
	emailSubject = this.messageSource.getMessage("email.user.subject.register",null, null);	
	final String emailSubjectFinal = emailSubject;
	final String emailFrom = this.messageSource.getMessage("email.msg.from",null, null);
	final String messageHeader = this.messageSource.getMessage("email.msg.header.user",null, null);
	final String body1 = this.messageSource.getMessage("email.msg.user.sub.header.thanks",null, null);
	final String body2 = this.messageSource.getMessage("email.msg.reg.by",null, null);
	final String body3 = this.messageSource.getMessage("email.msg.user.info.one",null, null);
	final String body4 = this.messageSource.getMessage("email.msg.user.info.four",null, null);
	final String body5 = this.messageSource.getMessage("email.msg.user.info.five",null, null);	//	raw link
	final String body6 = this.messageSource.getMessage("email.msg.user.info.six",null, null);
	final String body7 = this.messageSource.getMessage("email.msg.user.info.seven",null, null);	
	
	final String footer = this.messageSource.getMessage("email.reg.faci.footer",null, null);
	MimeMessagePreparator preparator = new MimeMessagePreparator() {
		public void prepare(MimeMessage mimeMessage) throws Exception {
			MimeMessageHelper message = new MimeMessageHelper(mimeMessage, true);
			message.setTo(emailId);
			message.setFrom(new InternetAddress(emailFrom, "Job Crafting LLC"));
			message.setSubject(emailSubjectFinal);
			StringBuilder mailBuilder = new StringBuilder("");					
			mailBuilder.append(messageHeader + " " + fName + ",");			
			mailBuilder.append(body1);
			mailBuilder.append(body2);
			mailBuilder.append(faciName);
			mailBuilder.append(" (" + createdBy + ").</br>");
			mailBuilder.append(body3);
			mailBuilder.append(password+"</i>");
			mailBuilder.append(body4);
			mailBuilder.append(emailId);
			mailBuilder.append(body5);	//	raw link
			mailBuilder.append(body6);
			mailBuilder.append(body7);	//	link
			mailBuilder.append(emailId+"</i></u></span>");		
			mailBuilder.append(footer);					
			message.setText(mailBuilder.toString(), true);					
		}
	};
	this.mailSender.send(preparator);
	LOGGER.info("======= MAIL SENT TO: "+emailId+" =======");
	LOGGER.info("<<<<<<<<< MailDispatch.intimateUserByFaciliatorAsyn");	
}
	
	
	@Async
	public void intimateFaciPassword(final String facilitatorEmail, final String password, final String fName) {
		LOGGER.info(">>>>>>>>> MailDispatch.intimateFaciPassword");
		String emailSubject = null;
		emailSubject = this.messageSource.getMessage("email.user.subject.msg.reset",null, null);	//	Unable to register				
		final String emailSubjectFinal = emailSubject;
		final String emailFrom = this.messageSource.getMessage("email.msg.from",null, null);
		final String messageHeader = this.messageSource.getMessage("email.msg.header.user",null, null);
		final String body2 = this.messageSource.getMessage("email.reset.faci.password",null, null);				
		final String body3 = this.messageSource.getMessage("email.msg.user.info.one",null, null);
		
		final String body3_1 = this.messageSource.getMessage("email.reg.faci.body4",null, null);
		final String body3_2 = this.messageSource.getMessage("email.reg.faci.body4.1",null, null);
		final String body4 = this.messageSource.getMessage("email.reg.faci.body5",null, null);
		
		
		//final String body4 = this.messageSource.getMessage("email.faci.resetPwd.loginAgain",null, null);
		final String footer = this.messageSource.getMessage("email.reg.faci.footer",null, null);
				
		MimeMessagePreparator preparator = new MimeMessagePreparator() {
			public void prepare(MimeMessage mimeMessage) throws Exception {
				MimeMessageHelper message = new MimeMessageHelper(mimeMessage, true);
				message.setTo(facilitatorEmail);
				message.setFrom(new InternetAddress(emailFrom, "Job Crafting LLC"));
				message.setSubject(emailSubjectFinal);
				StringBuilder mailBuilder = new StringBuilder("");					
				mailBuilder.append(messageHeader);
				mailBuilder.append(fName + ",");
				mailBuilder.append(body2);					
				mailBuilder.append(body3);
				mailBuilder.append(password+"</b></i>");
				
				mailBuilder.append(body3_1);			//	hyperlink
				mailBuilder.append(facilitatorEmail);
				mailBuilder.append(body3_2);
				mailBuilder.append(body4);
				mailBuilder.append(facilitatorEmail);
				
				
				//mailBuilder.append(body4);
				mailBuilder.append("</u></span><br>");
				mailBuilder.append(footer);					
				message.setText(mailBuilder.toString(), true);					
			}
		};
		this.mailSender.send(preparator);
		LOGGER.info("======= MAIL SENT TO: "+facilitatorEmail+" =======");
		LOGGER.info("<<<<<<<<< MailDispatch.intimateFaciPassword");	
		
	}
	/**
	 * Mail to notify user about their account expiration before 1 month /2 weeks /5 days
	 * @param emailId
	 * @param dicVal
	 * @param fName
	 * @param expDate
	 * */
	 /*
	@Async	
	public void intimateUserExpNotification(final String emailId, 
			final String dicVal, final String fName, final String expDate) throws MailingException {
		LOGGER.info(">>>>>>>>> MailDispatch.intimateUserExpNotification");
		String emailSubject = null;
		emailSubject = this.messageSource.getMessage("email.user.subject.msg.reset",null, null);	
		final String emailSubjectFinal = emailSubject;
		final String emailFrom = this.messageSource.getMessage("email.msg.from",null, null);
		final String messageHeader = this.messageSource.getMessage("email.msg.header.user",null, null);
		final String body1 = this.messageSource.getMessage("lebel.msg.exp.body1",null, null);
		final String body2 = this.messageSource.getMessage("lebel.msg.exp.body2",null, null);
		final String body3 = this.messageSource.getMessage("lebel.msg.exp.clickHere1",null, null);
		final String body4 = this.messageSource.getMessage("lebel.msg.exp.clickHere2",null, null);	//	raw link
		final String body5 = this.messageSource.getMessage("lebel.msg.exp.linkNotWorking",null, null);
		final String footer =this.messageSource.getMessage("email.reg.faci.footer",null, null);
		MimeMessagePreparator preparator = new MimeMessagePreparator() {
			public void prepare(MimeMessage mimeMessage) throws Exception {
				MimeMessageHelper message = new MimeMessageHelper(mimeMessage, true);
				message.setTo(emailId);
				message.setFrom(new InternetAddress(emailFrom, "Job Crafting LLC"));
				message.setSubject(emailSubjectFinal);
				StringBuilder mailBuilder = new StringBuilder("");					
				mailBuilder.append(messageHeader + " " + fName + ",");			
				mailBuilder.append(body1);
				mailBuilder.append(body2);
				mailBuilder.append(expDate);
				mailBuilder.append(body3);
				mailBuilder.append(body4);
				mailBuilder.append(body5);				
				mailBuilder.append(footer);					
				message.setText(mailBuilder.toString(), true);					
			}
		};
		this.mailSender.send(preparator);
		LOGGER.info("======= MAIL SENT TO: "+emailId+" =======");
		LOGGER.info("<<<<<<<<< MailDispatch.intimateUserExpNotification");	
	}*/
}
