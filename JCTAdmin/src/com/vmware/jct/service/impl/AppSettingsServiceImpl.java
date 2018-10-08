package com.vmware.jct.service.impl;

import java.util.Date;
import java.util.Iterator;
import java.util.List;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;
import org.springframework.transaction.annotation.Propagation;
import org.springframework.transaction.annotation.Transactional;

import com.vmware.jct.common.utility.CommonConstants;
import com.vmware.jct.dao.IAppSettingsDAO;
import com.vmware.jct.exception.DAOException;
import com.vmware.jct.exception.JCTException;
import com.vmware.jct.model.JctAppSettingsMaster;
import com.vmware.jct.service.IAppSettingsService;
import com.vmware.jct.service.vo.AppSettingsVO;
/**
 * 
 * <p><b>Class name:</b> AppSettingsServiceImpl.java</p>
 * <p><b>Author:</b> InterraIT</p>
 * <p><b>Purpose:</b> This is a AppSettingsServiceImpl class. This artifact is Business layer artifact.
 * AppSettingsServiceImpl implement IAppSettingsService interface and override the following  methods.
 * -fetchColor(AppSettingsVO settingsVO)
 * -saveNewWhiteLebel(AppSettingsVO settingsVO)
 * </p>
 * <p><b>Description:</b> This class used to perform insert, modify ,fetch data from db and delete operation. Save, update, delete operations are performed in DataAccessObject class </p>
 * <p><b>Revision History:</b>
 * </p>
 */
@Service
public class AppSettingsServiceImpl implements IAppSettingsService {

	private static final Logger LOGGER = LoggerFactory.getLogger(AppSettingsServiceImpl.class);
	
	@Autowired
	private IAppSettingsDAO appSettingsDAO;
	
	@Transactional(propagation=Propagation.REQUIRED)
	public AppSettingsVO fetchColor(AppSettingsVO settingsVO)
			throws JCTException {
		LOGGER.info(">>>>>> AppSettingsServiceImpl.fetchColor");
		try {
			List<JctAppSettingsMaster> list = appSettingsDAO.getColors(0);
			if (list.size() > 0) {
				JctAppSettingsMaster master = (JctAppSettingsMaster) list.get(list.size() - 1); //Last one
				settingsVO.setHeaderColor(master.getAppHeaderColor());
				settingsVO.setFooterColor(master.getAppFooterColor());
				settingsVO.setSubHeaderColor(master.getAppSubHeaderColor());
				settingsVO.setInstructionPanelColor(master.getAppInsPanelColor());
				settingsVO.setInstructionPanelTextColor(master.getAppInsPanelTxtColor());
				settingsVO.setImage(master.getAppImage());
			} else {
				settingsVO.setHeaderColor("#006990");
				settingsVO.setFooterColor("#006990");
				settingsVO.setSubHeaderColor("#88cbdf");
				settingsVO.setInstructionPanelColor("#810123");
				settingsVO.setInstructionPanelTextColor("#fff");
			}
		} catch (DAOException exception) {
			LOGGER.error("-------"+exception.getLocalizedMessage());
			throw new JCTException();
		}
		LOGGER.info("<<<<<< AppSettingsServiceImpl.fetchColor");
		return settingsVO;
	}

	@Transactional(propagation=Propagation.REQUIRED)
	public String saveNewWhiteLebel(AppSettingsVO settingsVO, String jctEmail) throws JCTException {
		String status = "failed";
		try {
		//Fetch Existing
		List<JctAppSettingsMaster> list = appSettingsDAO.getColors(0);
		Iterator<JctAppSettingsMaster> listItr = list.iterator();
		while (listItr.hasNext()) {
			//UPDATE THE SOFT DELETE STATUS
			JctAppSettingsMaster existingMaster = (JctAppSettingsMaster) listItr.next();
			existingMaster.setAppSoftDelete(CommonConstants.ENTRY_SOFT_DELETED);
			appSettingsDAO.updateSoftDelete(existingMaster);
		}
		
		JctAppSettingsMaster master = new JctAppSettingsMaster();
		master.setAppHeaderColor(settingsVO.getHeaderColor());
		master.setAppFooterColor(settingsVO.getFooterColor());
		master.setAppSubHeaderColor(settingsVO.getSubHeaderColor());
		master.setAppInsPanelColor(settingsVO.getInstructionPanelColor());
		master.setAppInsPanelTxtColor(settingsVO.getInstructionPanelTextColor());
		master.setAppLastUpdatedTs(new Date());
		master.setAppLastUpdatedBy(jctEmail);
		master.setAppImage(settingsVO.getImage());
		status = appSettingsDAO.saveWhiteLebel(master);
		} catch (DAOException ex) {
			LOGGER.error("-------"+ex.getLocalizedMessage());
			throw new JCTException();
		}
		return status;
	}

}