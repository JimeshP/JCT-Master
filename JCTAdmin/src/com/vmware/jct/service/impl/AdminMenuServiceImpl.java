package com.vmware.jct.service.impl;

import java.util.Iterator;
import java.util.List;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;
import org.springframework.transaction.annotation.Propagation;
import org.springframework.transaction.annotation.Transactional;

import com.vmware.jct.dao.IAdminMenuDAO;
import com.vmware.jct.exception.DAOException;
import com.vmware.jct.exception.JCTException;
import com.vmware.jct.model.JctAdminMenu;
import com.vmware.jct.service.IAdminMenuService;

/**
 * 
 * <p><b>Class name:</b> AdminMenuServiceImpl.java</p>
 * <p><b>Author:</b> InterraIT</p>
 * <p><b>Purpose:</b> This is a AdminMenuServiceImpl class. This artifact is Business layer artifact.
 * AdminMenuServiceImpl implement IAdminMenuService interface and override the following  methods.
 * -getMenuString(int userType)
 * </p>
 * <p><b>Description:</b> This class used to perform insert, modify ,fetch data from db and delete operation. Save, update, delete operations are performed in DataAccessObject class </p>
 * <p><b>Revision History:</b>
 * </p>
 */
@Service
public class AdminMenuServiceImpl implements IAdminMenuService {

	@Autowired
	private IAdminMenuDAO adminMenuDAO;
	
	private static final Logger LOGGER = LoggerFactory.getLogger(AdminMenuServiceImpl.class);
	
	/**
	 * Method accepts the roleId of the user and fetches the 
	 * corresponding menu item list, formats it and sends it 
	 * back.
	 */
	@Transactional(propagation=Propagation.REQUIRED)
	public String getMenuString ( int roleId ) throws JCTException {
		LOGGER.info(">>>>> AdminMenuServiceImpl.getMenuString");
		StringBuilder menuBuilder = new StringBuilder("");
		try {
			List<String> distinctMenuList = adminMenuDAO.getDistinctMenu(roleId);
			Iterator<String> menuListItr = distinctMenuList.iterator();
			while (menuListItr.hasNext()) {
				String mainMenu = menuListItr.next().toString(); // Main Menu
				menuBuilder.append(mainMenu);
				menuBuilder.append("`");
				List<JctAdminMenu> details = adminMenuDAO.getMenuDetails(roleId, mainMenu);
				Iterator<JctAdminMenu> detailsItr = details.iterator();
				while (detailsItr.hasNext()) {
					JctAdminMenu menuObject = (JctAdminMenu) detailsItr.next();
					menuBuilder.append(menuObject.getJctSubMenu());
					menuBuilder.append("~");
					menuBuilder.append(menuObject.getJctSubMenuLinkedPage());
					menuBuilder.append("+");
				}
				menuBuilder.append("<!>");
			}
		} catch (DAOException ex) {
			LOGGER.error(ex.getLocalizedMessage());
			throw new JCTException("Unable to fetch any menu items");
		}
		LOGGER.info("<<<<< AdminMenuServiceImpl.getMenuString");
		return menuBuilder.toString();
	}
}