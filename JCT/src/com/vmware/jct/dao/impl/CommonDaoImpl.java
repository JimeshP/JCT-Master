package com.vmware.jct.dao.impl;


import java.util.List;
import org.hibernate.SessionFactory;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Repository;
import org.springframework.transaction.annotation.Propagation;
import org.springframework.transaction.annotation.Transactional;


import com.vmware.jct.dao.ICommonDao;
import com.vmware.jct.exception.DAOException;
import com.vmware.jct.exception.DAONullException;
import com.vmware.jct.exception.JCTException;
import com.vmware.jct.model.JctIdGenTable;


/**
 * 
 * <p><b>Class name:</b> CommonDaoImpl.java</p>
 * <p><b>Author:</b> InterraIt</p>
 * <p><b>Purpose:</b> This is a CommonDaoImpl class. This artifact is Persistence Manager layer artifact.
 * QuestionnaireDAOImpl implement ICommonDao interface and override the following  methods.
 * -generateKey(String keyName)
 * <p><b>Description:</b>  This class used to perform unique identification number </p>
 * <p><b>Copyrights:</b> 	All rights reserved by Interra IT and should only be used for its internal application development.</p>
 * <p><b>Revision History:</b>
 * 	<li>InterraIt - 31/Jan/2014 - Implement comments, introduce Named query, Transactional attribute </li>
 * </p>
 */
@Repository
public class CommonDaoImpl implements ICommonDao { 
	@Autowired
	private SessionFactory sessionFactory;
	
	private static final Logger logger = LoggerFactory.getLogger(BeforeSketchDAOImpl.class);
	
	/**
	 * 	* <p><b>Description:</b>  This method generate primary key from id gen table </p>
	 * @param It will take string  object as a input parameter
	 * @return unique identification number
	 * @throws DAOException
	 */
	@Transactional(propagation=Propagation.REQUIRED)
	public Integer generateKey(String keyName) throws DAOException {
		logger.info(">>>> CommonDaoImpl.generateKey"); 
		List<JctIdGenTable> worfFlowLst = sessionFactory.getCurrentSession().createQuery("from JctIdGenTable where jctIdName='"+keyName+"'").list();
		if (worfFlowLst != null) {
			if(worfFlowLst.size() == 0)
				throw new DAONullException("No Key entry found..");
			
			}
		JctIdGenTable idGen = worfFlowLst.get(0);
		idGen.setJctIdVal(idGen.getJctIdVal()+1);
		sessionFactory.getCurrentSession().saveOrUpdate(idGen);
		logger.info("<<<< CommonDaoImpl.generateKey"); 
		return idGen.getJctIdVal();
	}
}
