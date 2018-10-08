package com.vmware.jct.common.utility;

import org.quartz.JobExecutionContext;
import org.quartz.JobExecutionException;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.scheduling.quartz.QuartzJobBean;
import org.springframework.stereotype.Service;
/**
 * 
 * <p><b>Class name:</b> ShopifyIntegratorJob.java</p>
 * <p><b>Author:</b> InterraIT</p>
 * <p><b>Purpose:</b> .
 * <p><b>Copyrights:</b>All rights reserved by Interra IT and should only be used for its internal application development.</p>
 * <p><b>Created Date:</b> 02/Apr/2015</p>
 * <p><b>Revision History:</b>
 * 	<li></li>
 * </p>
 */
@Service("shopifyIntegratorJob")
public class ShopifyIntegratorJob extends QuartzJobBean {
	private ShopifyIntegratorTask shopifyIntegratorTask;
	private static final Logger LOGGER = LoggerFactory.getLogger(ShopifyIntegratorJob.class);
	public void setShopifyIntegratorTask(ShopifyIntegratorTask shopifyIntegratorTask) {
		this.shopifyIntegratorTask = shopifyIntegratorTask;
	}
 
	protected void executeInternal(JobExecutionContext context)
		throws JobExecutionException {
		LOGGER.info(">>>>>>>>>>>>accountExpiryNotificationJob started...");
		shopifyIntegratorTask.pingShopify();
		LOGGER.info("<<<<<<<<<<<<accountExpiryNotificationJob started...");
	}
}
