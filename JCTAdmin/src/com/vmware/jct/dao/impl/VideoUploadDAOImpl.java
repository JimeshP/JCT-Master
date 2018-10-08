package com.vmware.jct.dao.impl;

import java.util.List;

import org.hibernate.SessionFactory;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Repository;
import org.springframework.transaction.annotation.Propagation;
import org.springframework.transaction.annotation.Transactional;

import com.vmware.jct.dao.DataAccessObject;
import com.vmware.jct.dao.IVideoUploadDAO;
import com.vmware.jct.exception.DAOException;
import com.vmware.jct.model.JctPopupInstruction;

@Repository
public class VideoUploadDAOImpl extends DataAccessObject implements IVideoUploadDAO {
	
	@Autowired
	private SessionFactory sessionFactory;
		
	@Transactional(propagation=Propagation.REQUIRED)
	public List<JctPopupInstruction> getElements(int profileId, 
			String insPage) throws DAOException {
		return sessionFactory.getCurrentSession().getNamedQuery("getActiveVideos")
				.setParameter("profileId", profileId)
				.setParameter("insPage", insPage)
				.list();
	}

	@Transactional(propagation=Propagation.REQUIRED)
	public String updateEntity(JctPopupInstruction obj) throws DAOException {
		return update(obj);
	}

	@Transactional(propagation=Propagation.REQUIRED)
	public String saveEntity(JctPopupInstruction obj) throws DAOException {
		return save(obj);
	}

	@Transactional(propagation=Propagation.REQUIRED)
	public List<JctPopupInstruction> getSelectedElements(int profileId,
			String page) throws DAOException {
		String type = "";
		return sessionFactory.getCurrentSession().getNamedQuery("getActiveVideosByType")
				.setParameter("profileId", profileId)
				.setParameter("page", page)
				.list();
	}
	
	@Transactional(propagation=Propagation.REQUIRED)
	public List<JctPopupInstruction> getSelectedInstruction(int profileId,
			String page) throws DAOException {
		return sessionFactory.getCurrentSession().getNamedQuery("getActiveInstructionByType")
				.setParameter("profileId", profileId)
				.setParameter("page", page)
				.list();
	}
}
