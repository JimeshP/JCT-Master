package com.vmware.jct.service.impl;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;
import org.springframework.transaction.annotation.Propagation;
import org.springframework.transaction.annotation.Transactional;

import com.vmware.jct.common.utility.CommonUtility;
import com.vmware.jct.dao.IAccountExpiryNotificationDAO;
import com.vmware.jct.dao.dto.AccountExpiryDTO;
import com.vmware.jct.exception.DAOException;
import com.vmware.jct.exception.JCTException;
import com.vmware.jct.service.IAccountExpiryNotificationService;

/**
 * 
 * <p><b>Class name:</b> AccountExpiryNotificationServiceImpl.java</p>
 * <p><b>Author:</b> InterraIT</p>
 * <p><b>Purpose:</b> This is a AccountExpiryNotificationServiceImpl class. This artifact is Business layer artifact.
 * AccountExpiryNotificationServiceImpl implement IAccountExpiryNotificationService interface and override the following  methods.
 * - getAcntExpryNtfcatnForFacilitator()
 * </p>
 * <p><b>Description:</b> This class used to perform insert, modify ,fetch data from db and delete operation. Save, update, delete operations are performed in DataAccessObject class </p>
 * <p><b>Revision History:</b>
 * <li>20/Oct/2014 - Introduced the class</li>
 * </p>
 */
@Service
public class AccountExpiryNotificationServiceImpl implements IAccountExpiryNotificationService {

	private static final Logger LOGGER = LoggerFactory.getLogger(AccountExpiryNotificationServiceImpl.class);
	
	@Autowired
	private IAccountExpiryNotificationDAO notificationDAO;
	/**
	 * Method fetches all the expiry details..
	 * @return Map<Email, List of all would be expiring account>
	 * @throws JCTException
	 */
	@Transactional(propagation=Propagation.REQUIRED)
	public Map<String, List<String>> getAcntExpryNtfcatnForFacilitator()
			throws JCTException {
		LOGGER.info(">>>>>>> AccountExpiryNotificationServiceImpl.getAcntExpryNtfcatnForFacilitator");
		Map<String, List<String>> detailsMap = new HashMap<String, List<String>>();
		try {
			List<Integer> facilitatorList = notificationDAO.getAllActiveFacilitators();
			Iterator<Integer> facilitatorListItr = facilitatorList.iterator();
			while (facilitatorListItr.hasNext()) {
				Integer facilitatorId = (Integer) facilitatorListItr.next();
				//Get the email Id
				String facEmailId = notificationDAO.getEmailIdByUserId(facilitatorId, 3);
				// Fetch the user details with the facilitator Id
				List<AccountExpiryDTO> userDetailList = notificationDAO.getExpirationDetails(facilitatorId, 
														CommonUtility.getCurrentDate(), CommonUtility.getExpDate());
				Iterator<AccountExpiryDTO> userDtlsItr = userDetailList.iterator();
				List<String> userList = new ArrayList<String>();
				while (userDtlsItr.hasNext()) {
					AccountExpiryDTO dto = (AccountExpiryDTO) userDtlsItr.next();
					userList.add(dto.getUserName()+"`"+dto.getExpiryDate()+"!");
				}
				detailsMap.put(facEmailId, userList);
			}
		} catch (DAOException dooEx) {
			throw new JCTException();
		}
		LOGGER.info("<<<<<<< AccountExpiryNotificationServiceImpl.getAcntExpryNtfcatnForFacilitator");
		return detailsMap;
	}
}