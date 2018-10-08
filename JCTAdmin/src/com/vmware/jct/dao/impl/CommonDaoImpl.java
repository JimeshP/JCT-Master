package com.vmware.jct.dao.impl;


import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.List;
import java.util.Calendar;
import org.hibernate.SessionFactory;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Repository;
import org.springframework.transaction.annotation.Propagation;
import org.springframework.transaction.annotation.Transactional;

import com.vmware.jct.common.utility.CommonConstants;
import com.vmware.jct.dao.ICommonDao;
import com.vmware.jct.exception.DAOException;
import com.vmware.jct.exception.DAONullException;
import com.vmware.jct.model.JctIdGenTable;


/**
 * 
 * <p><b>Class name:</b> CommonDaoImpl.java</p>
 * <p><b>Author:</b> InterraIT</p>
 * <p><b>Purpose:</b> This is a CommonDaoImpl class. This artifact is Persistence Manager layer artifact.
 * QuestionnaireDAOImpl implement ICommonDao interface and override the following  methods.
 * -generateKey(String keyName)
 * -generateUniqueCustomerId(int roleId)
 * <p><b>Description:</b>  This class used to perform unique identification number </p>
 * <p><b>Copyrights:</b> 	All rights reserved by InterraIT and should only be used for its internal application development.</p>
 * <p><b>Revision History:</b>
 * 	<li>InterraIT - 31/Jan/2014 - Implement comments, introduce Named query, Transactional attribute </li>
 *  <li>InterraIT - 27/Oct/2014 - Implemented generateUniqueCustomerId </li>
 * </p>
 */
@Repository
public class CommonDaoImpl implements ICommonDao { 
	
	private static final Logger LOGGER = LoggerFactory.getLogger(CommonDaoImpl.class);
	
	@Autowired
	private SessionFactory sessionFactory;
	/**
	 * 	
	 * @param It will take string  object as a input parameter
	 * @return unique identification number
	 * @throws DAOException
	 */
	@Transactional(propagation=Propagation.REQUIRED)
	public Integer generateKey(String keyName) throws DAOException {
		LOGGER.info(">>>> CommonDaoImpl.generateKey");
		List<JctIdGenTable> worfFlowLst = sessionFactory.getCurrentSession().createQuery("from JctIdGenTable where jctIdName='"+keyName+"'").list();
		if (worfFlowLst != null) {
			if(worfFlowLst.size() == 0)
				throw new DAONullException("No Key entry found..");
			
			}
		JctIdGenTable idGen = worfFlowLst.get(0);
		idGen.setJctIdVal(idGen.getJctIdVal()+1);
		sessionFactory.getCurrentSession().saveOrUpdate(idGen);
		LOGGER.info("<<<< CommonDaoImpl.generateKey");
		return idGen.getJctIdVal();
	}
	
	/**
	 * Method creates unique customer Id
	 * @param roleId
	 * @return
	 */
	@Transactional(propagation=Propagation.REQUIRED)
	public String generateUniqueCustomerId(int roleId) throws DAOException {
		StringBuffer customerIdMaker = new StringBuffer("");
		String id = "";
	/******/	
		int newYear = (Integer) sessionFactory.getCurrentSession().getNamedQuery("fetchMaxDate").uniqueResult();  //get the maximum year from database
	    Calendar cal = Calendar.getInstance();
	    int currentYear = (Integer)cal.get(Calendar.YEAR);														//get current date	
		/*update id to 0 in the database when year changes*/
	    if(newYear!=currentYear){
	    	if(roleId == CommonConstants.FACILITATOR_USER){
			sessionFactory.getCurrentSession().getNamedQuery("updateIdGenerationYear").setParameter("keyName", "jct_user_facilitator_id").executeUpdate();
	    	}
	    	else
	    	{
	    	sessionFactory.getCurrentSession().getNamedQuery("updateIdGenerationYear").setParameter("keyName", "jct_user_customer_id").executeUpdate();
	    	}
		}
	 /******/   
		if (roleId == CommonConstants.FACILITATOR_USER) {
			customerIdMaker.append(CommonConstants.FACILITATOR_USER_CUST_CODE); // 99
			id = generateKey("jct_user_facilitator_id") + "";
		} else {
			customerIdMaker.append(CommonConstants.INDIVIDUAL_USER_CUST_CODE);  // 98
			id = generateKey("jct_user_customer_id") + "";
		}
    	
		SimpleDateFormat sdf = new SimpleDateFormat("ddMMyyyy");
		customerIdMaker.append(sdf.format(new Date()));

		int idLength = id.length();
		
		switch (idLength) {
		case 1:
			customerIdMaker.append("00000000" + id);
			break;
		case 2:
			customerIdMaker.append("0000000" + id);
			break;
		case 3:
			customerIdMaker.append("000000" + id);
			break;
		case 4:
			customerIdMaker.append("00000" + id);
			break;
		case 5:
			customerIdMaker.append("0000" + id);
			break;
		case 6:
			customerIdMaker.append("000" + id);
			break;
		case 7:
			customerIdMaker.append("00" + id);
			break;
		case 8:
			customerIdMaker.append("0" + id);
			break;
		case 9:
			customerIdMaker.append(id);
			break;
		default:
			customerIdMaker.append(id);
			break;
		}
		return customerIdMaker.toString();
	}
}