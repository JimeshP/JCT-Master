/**
 * 
 */
package com.vmware.jct.common.utility;


import org.quartz.JobExecutionContext;
import org.quartz.JobExecutionException;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.scheduling.quartz.QuartzJobBean;
import org.springframework.stereotype.Service;

import com.vmware.jct.exception.MailingException;
 
/**
 * 
 * <p><b>Class name:</b> RunMeJob.java</p>
 * <p><b>Author:</b> InterraIT</p>
 * <p><b>Purpose:</b> .
 * <p><b>Copyrights:</b>All rights reserved by Interra IT and should only be used for its internal application development.</p>
 * <p><b>Created Date:</b> 10/Apr/2014</p>
 * <p><b>Revision History:</b>
 * 	<li></li>
 * </p>
 */

@Service("runMeJob")
public class RunMeJob extends QuartzJobBean {
	private RunMeTask runMeTask;
	
	private static final Logger LOGGER = LoggerFactory.getLogger(RunMeJob.class);
 
	public void setRunMeTask(RunMeTask runMeTask) {
		this.runMeTask = runMeTask;
	}
 
	protected void executeInternal(JobExecutionContext context)
		throws JobExecutionException {
		try {
			runMeTask.initiateMailing();
		} catch (MailingException e) {
			LOGGER.error(e.getLocalizedMessage());
		}
 
	}
}
