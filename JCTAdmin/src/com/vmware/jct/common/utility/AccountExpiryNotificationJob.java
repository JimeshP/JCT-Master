package com.vmware.jct.common.utility;

import org.quartz.JobExecutionContext;
import org.quartz.JobExecutionException;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.scheduling.quartz.QuartzJobBean;
import org.springframework.stereotype.Service;
/**
 * 
 * <p><b>Class name:</b> AccountExpiryNotificationJob.java</p>
 * <p><b>Author:</b> InterraIT</p>
 * <p><b>Purpose:</b> .
 * <p><b>Copyrights:</b>All rights reserved by Interra IT and should only be used for its internal application development.</p>
 * <p><b>Created Date:</b> 10/Oct/2014</p>
 * <p><b>Revision History:</b>
 * 	<li></li>
 * </p>
 */
@Service("accountExpiryNotificationJob")
public class AccountExpiryNotificationJob extends QuartzJobBean {
	private AccountExpiryNotificationTask accountExpiryNotificationTask;
	private static final Logger LOGGER = LoggerFactory.getLogger(AccountExpiryNotificationJob.class);
	public void setAccountExpiryNotificationTask(AccountExpiryNotificationTask accountExpiryNotificationTask) {
		this.accountExpiryNotificationTask = accountExpiryNotificationTask;
	}
 
	protected void executeInternal(JobExecutionContext context)
		throws JobExecutionException {
		LOGGER.info(">>>>>>>>>>>>accountExpiryNotificationJob started...");
		accountExpiryNotificationTask.notifyExpirationOfAccount();
		LOGGER.info("<<<<<<<<<<<<accountExpiryNotificationJob started...");
	}
}
