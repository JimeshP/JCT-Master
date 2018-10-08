/**
 * 
 */
package com.vmware.jct.dao;

import javax.persistence.EntityExistsException;

import org.hibernate.HibernateException;
import org.hibernate.SessionFactory;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.beans.factory.annotation.Autowired;

import com.vmware.jct.exception.DAOEntityExistsException;
import com.vmware.jct.exception.DAOException;
import com.vmware.jct.exception.DAONullException;

/**
 * 
 * <p><b>Class name:</b>DataAccessObject.java</p>
 * <p><b>Author:</b>  InterraIT</p>
 * <p><b>Purpose:</b>DataAccessObject provides persistence context to all DAO implementations. It also provides basic CRUD methods.</p>
 * <p><b>Description:</b></p>
 * <pre>
 * 		<b>Revision History:</b>
 * 		
 * </pre>
 * <p><b></b></p>
 */
public abstract class DataAccessObject {
	
	
	private static final Logger logger = LoggerFactory.getLogger(DataAccessObject.class);
	@Autowired
	private SessionFactory sessionFactory;
	
	public SessionFactory getSessionFactory() {
		return sessionFactory;
	}

	/**
	 *<p><b>Description:</b>  This method is used for storing relevent object</p>
	 */
	public String save(Object object) throws DAONullException, DAOEntityExistsException {
		logger.info(">>>>   DataAccessObject.save");
		String retStr = "failure";
		if (object == null) {
			logger.warn("DAO object is null.");
			throw new DAONullException("DAO object is null.");
		}
		try {
			logger.debug("Calling save on [" + object.getClass() + "]");
			getSessionFactory().getCurrentSession().persist(object);
			getSessionFactory().getCurrentSession().flush();
			retStr = "success";
		} catch (EntityExistsException exception) {
				logger.error(object.getClass() + " could not be saved as the entity already exists");
			throw new DAOEntityExistsException(exception);
		}
		logger.info("<<<<   DataAccessObject.save");
		return retStr;
	}
	/**
	 *<p><b>Description:</b>  This method is used for updating relevent object</p>
	 */
	public String update(Object object) throws DAONullException,DAOException {
		logger.info(">>>>   DataAccessObject.update");
		String retStr = "failure";
		if (object == null) {
			throw new DAONullException("DAO object is null.");
		}
		logger.debug("Calling update on [" + object.getClass() + "]");
		
		try {
			logger.debug("Calling save on [" + object.getClass() + "]");
			getSessionFactory().getCurrentSession().saveOrUpdate(object);
			getSessionFactory().getCurrentSession().flush();
			retStr = "success";
		} catch (HibernateException exception) {
				logger.error(object.getClass() + " could not be saved as the entity already exists");
			throw new HibernateException(exception);
		}
		
		logger.info("<<<<   DataAccessObject.update");
		return  retStr;
	}
	
	/**
	 *<p><b>Description:</b>  This method is used for updating delete object</p>
	 */
	public void Delete(Object object) throws DAONullException {
		logger.info(">>>>   DataAccessObject.Delete");
		if (object == null) {
			throw new DAONullException("DAO object is null.");
		}
		logger.debug("Calling delete on [" + object.getClass() + "]");
		getSessionFactory().getCurrentSession().delete(object);
		getSessionFactory().getCurrentSession().flush();
		logger.info("<<<<   DataAccessObject.Delete");
	}

	/*public void softDelete(Object object) throws DAONullException {
		if (object == null) {
			throw new DAONullException("DAO object is null.");
		}
		logger.debug("Calling delete on [" + object.getClass() + "]");
		getSessionFactory().getCurrentSession().merge(getSessionFactory().getCurrentSession().getReference(object.getClass(), ((JCTEntity)object).fetchPrimaryKey()));
	}*/


}
