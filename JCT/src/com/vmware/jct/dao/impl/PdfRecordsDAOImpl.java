package com.vmware.jct.dao.impl;

import org.hibernate.SessionFactory;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Repository;
import org.springframework.transaction.annotation.Propagation;
import org.springframework.transaction.annotation.Transactional;

import com.vmware.jct.dao.DataAccessObject;
import com.vmware.jct.dao.IPdfRecordsDAO;
import com.vmware.jct.exception.DAOException;
import com.vmware.jct.model.JctPdfRecords;

/**
 * 
 * <p><b>Class name:</b> PdfRecordsDAOImpl.java</p>
 * <p><b>Author:</b> InterraIT</p>
 * <p><b>Purpose:</b> This is a PdfRecordsDAOImpl class. This artifact is Persistence Manager layer artifact.
 * PdfRecordsDAOImpl implement IPdfRecordsDAO interface and override the following  methods.
 * -savePDFReference(JctPdfRecords records)
 * -getMaxDocRow(String jobReferenceNo)
 * <p><b>Description:</b> This class used to perform pdf save and retrieve functionality</p>
 * <p><b>Revision History:</b>
 * </p>
 */

@Repository
public class PdfRecordsDAOImpl extends DataAccessObject implements IPdfRecordsDAO{

	@Autowired
	private SessionFactory sessionFactory;
	
	
	/**
	 * <p><b>Description:</b>  Method saves the Pdf file's actual refecence</p>
	 * @param It will take JctPdfRecords as a input parameter
	 * @return  It returns string object
	 * @throws DAOException -it throws DAOException
	 * 
	 */
	@Transactional(propagation=Propagation.REQUIRED)
	public String savePDFReference(JctPdfRecords records) throws DAOException {
		return save(records);
	}
	
	/**
	 * <p><b>Description:</b> Method saves Pdf file's last row matching the job reference number.</p>
	 * @param It will take JctPdfRecords as a input parameter
	 * @return  It returns string object
	 * @throws DAOException -it throws DAOException
	 * 
	 */
	@Transactional(propagation=Propagation.REQUIRED)
	public JctPdfRecords getMaxDocRow(String jobReferenceNo) throws DAOException {
		return (JctPdfRecords)sessionFactory.getCurrentSession().getNamedQuery("fetchLastUpdated").setParameter("refNo", jobReferenceNo).uniqueResult();
	}

}
