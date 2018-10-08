package com.vmware.jct.dao.impl;

import java.math.BigInteger;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

import org.hibernate.SessionFactory;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Repository;
import org.springframework.transaction.annotation.Propagation;
import org.springframework.transaction.annotation.Transactional;

import com.vmware.jct.dao.IReportDAO;
import com.vmware.jct.dao.dto.SurveyQuestionDTO;
import com.vmware.jct.exception.DAOException;
/**
 * 
 * <p><b>Class name:</b> ReportDAOImpl.java</p>
 * <p><b>Author:</b> InterraIT</p>
 * <p><b>Purpose:</b> This is a ReportDAOImpl class. This artifact is Persistence Manager layer artifact.
 * ReportDAOImpl implement IReportDAO interface.
 * </p>
 * <p><b>Description:</b> This class used to perform insert, modify ,fetch data from db and delete operation. Save, update, delete operations are performed in DataAccessObject class </p>
 * <p><b>Revision History:</b>
 * 	<li>InterraIT - 19/Jan/2014 - Implement Exception </li>
 * <li>InterraIT -  - 31/Jan/2014 - Implement comments, introduce Named query, Transactional attribute </li>
 * <li>InterraIT - 16/May/2014 - Updated comments </li>
 * </p>
 */
@Repository
public class ReportDAOImpl implements IReportDAO {
	
	private static final Logger LOGGER = LoggerFactory.getLogger(ReportDAOImpl.class);
	
	@Autowired
	private SessionFactory sessionFactory;

	@Transactional(propagation=Propagation.REQUIRED)
	public List<Object> getBeforeSketchList(String function, String jobLevel, int recordIndex, int status, int softDelete) throws DAOException {
		LOGGER.info(">>>> ReportDAOImpl.getBeforeSketchList");
		StringBuilder queryBldr = null;
		StringBuilder jrNoBuilder = new StringBuilder("");
		List<Object> returnList = null;
		List<String> jrList = null;
		queryBldr = new StringBuilder("Select distinct(bsview.jct_bs_created_by) from jct_before_sketch_view bsview");
		if ((function.trim().length() > 0) && (jobLevel.trim().length() > 0)) {
			queryBldr.append(" where bsview.jct_user_function_group = :fnGrp " +
					"					and bsview.jct_user_levels = :jLevel and jct_bs_status_id = "+status+" and jct_bs_soft_delete = "+softDelete+"");
			jrList = sessionFactory.getCurrentSession().
					 createSQLQuery(queryBldr.toString()).setParameter("fnGrp", function).
					 setParameter("jLevel", jobLevel).
					 setFirstResult(recordIndex).setMaxResults(20).list();
		} else if ((function.trim().length() > 0) && (jobLevel.trim().length() == 0)) {
			queryBldr.append(" where bsview.jct_user_function_group = :fnGrp and jct_bs_status_id = "+status+" and jct_bs_soft_delete = "+softDelete+"");
			jrList = sessionFactory.getCurrentSession().
					 createSQLQuery(queryBldr.toString()).setParameter("fnGrp", function).
					 setFirstResult(recordIndex).setMaxResults(20).list();
		} else if ((function.trim().length() == 0) && (jobLevel.trim().length() > 0)) {
			queryBldr.append(" where bsview.jct_user_levels = :jLevel and jct_bs_status_id = "+status+" and jct_bs_soft_delete = "+softDelete+"");
			jrList = sessionFactory.getCurrentSession().
					createSQLQuery(queryBldr.toString()).setParameter("jLevel", jobLevel).
					setFirstResult(recordIndex).setMaxResults(20).list();
		} else {
			queryBldr.append(" where jct_bs_status_id = "+status+" and jct_bs_soft_delete = "+softDelete+"");
			jrList = sessionFactory.getCurrentSession().
					 createSQLQuery(queryBldr.toString()).
					 setFirstResult(recordIndex).setMaxResults(20).list();
		}
		
		if (jrList.size() > 0) {
			int size = 0;
			Iterator itr = jrList.iterator();
			while (itr.hasNext()) {
				jrNoBuilder.append("'" + (String) itr.next() + "'");
				if (jrList.size() - 1 > size) {
					jrNoBuilder.append(",");	
				} 
				size = size + 1;
			}
			queryBldr = new StringBuilder ("Select bsview.jct_bs_created_by, " +
					"bsview.jct_user_function_group, " +
					"bsview.jct_user_levels, bsview.jct_bs_task_desc, " +
					"bsview.jct_bs_energy, bsview.jct_bs_time_spent, " +
					"bsview.jct_first_name, bsview.jct_last_name " +
					"from jct_before_sketch_view bsview");
			if ((function.trim().length() > 0) && (jobLevel.trim().length() > 0)) {
				queryBldr.append(" where bsview.jct_user_function_group = :fnGrp " +
						"and bsview.jct_user_levels = :jLevel");
				queryBldr.append(" and bsview.jct_bs_created_by in (" + jrNoBuilder.toString() + ") and jct_bs_status_id = "+status+" and jct_bs_soft_delete = "+softDelete+"");
				queryBldr.append(" and bsview.jct_bs_task_desc != '' ");
			} else if ((function.trim().length() > 0) && (jobLevel.trim().length() == 0)) {
				queryBldr.append(" where bsview.jct_user_function_group = :fnGrp");
				queryBldr.append(" and bsview.jct_bs_created_by in (" + jrNoBuilder.toString() + ") and jct_bs_status_id = "+status+" and jct_bs_soft_delete = "+softDelete+"");
				queryBldr.append(" and bsview.jct_bs_task_desc != '' ");
			} else if ((function.trim().length() == 0) && (jobLevel.trim().length() > 0)) {
				queryBldr.append(" where bsview.jct_user_levels = :jLevel");
				queryBldr.append(" and bsview.jct_bs_created_by in (" + jrNoBuilder.toString() + ") and jct_bs_status_id = "+status+" and jct_bs_soft_delete = "+softDelete+"");
				queryBldr.append(" and bsview.jct_bs_task_desc != '' ");
			} else {
				queryBldr.append(" where bsview.jct_bs_created_by in (" + jrNoBuilder.toString() + ") and jct_bs_status_id = "+status+" and jct_bs_soft_delete = "+softDelete+"");
				queryBldr.append(" and bsview.jct_bs_task_desc != '' ");
			}
			queryBldr.append(" order by bsview.jct_bs_created_by");
			
			if ((function.trim().length() > 0) && (jobLevel.trim().length() > 0)) {
				returnList = sessionFactory.getCurrentSession().
						createSQLQuery(queryBldr.toString()).setParameter("fnGrp", function).
						setParameter("jLevel", jobLevel).list();
			} else if ((function.trim().length() > 0) && (jobLevel.trim().length() == 0)) {
				returnList = sessionFactory.getCurrentSession().
						createSQLQuery(queryBldr.toString()).setParameter("fnGrp", function)
						.list();
			} else if ((function.trim().length() == 0) && (jobLevel.trim().length() > 0)) {
				returnList = sessionFactory.getCurrentSession().
						createSQLQuery(queryBldr.toString()).setParameter("jLevel", jobLevel)
						.list();
			} else {
				returnList = sessionFactory.getCurrentSession().
						createSQLQuery(queryBldr.toString())
						.list();
			}	
		} else {
			returnList = new ArrayList<Object>();
		}
		LOGGER.info("here");
		System.out.println("Query..");
    	for (int i=0;i<returnList.size();i++)
     	{
    		System.out.println(returnList.get(i));
     	}
		LOGGER.info("<<<< ReportDAOImpl.getBeforeSketchList");
    	return returnList;
	}
	
	@Transactional(propagation=Propagation.REQUIRED)
	public List<Object> getBeforeSketchListMax(String function, String jobLevel, int recordIndex) throws DAOException {
		LOGGER.info(">>>> ReportDAOImpl.getBeforeSketchListMax");
		StringBuilder queryBldr = null;
		StringBuilder maxHeaderId = new StringBuilder("");
		List<Object> returnList = null;
		List<String> jrList = null;
		queryBldr = new StringBuilder("select max(jct_bs_header_id),jct_bs_created_by  from jct_before_sketch_view bsview ");
		if ((function.trim().length() > 0) && (jobLevel.trim().length() > 0)) {
			queryBldr.append(" where bsview.jct_user_function_group = :fnGrp " +
					"					and bsview.jct_user_levels = :jLevel and jct_bs_status_id = 5 and jct_bs_soft_delete = 1 group by jct_bs_created_by");
			jrList = sessionFactory.getCurrentSession().
					 createSQLQuery(queryBldr.toString()).setParameter("fnGrp", function).
					 setParameter("jLevel", jobLevel).
					 setFirstResult(recordIndex).setMaxResults(20).list();
		} else if ((function.trim().length() > 0) && (jobLevel.trim().length() == 0)) {
			queryBldr.append(" where bsview.jct_user_function_group = :fnGrp and jct_bs_status_id = 5 and jct_bs_soft_delete = 1 group by jct_bs_created_by");
			jrList = sessionFactory.getCurrentSession().
					 createSQLQuery(queryBldr.toString()).setParameter("fnGrp", function).
					 setFirstResult(recordIndex).setMaxResults(20).list();
		} else if ((function.trim().length() == 0) && (jobLevel.trim().length() > 0)) {
			queryBldr.append(" where bsview.jct_user_levels = :jLevel and jct_bs_status_id = 5 and jct_bs_soft_delete = 1 group by jct_bs_created_by");
			jrList = sessionFactory.getCurrentSession().
					createSQLQuery(queryBldr.toString()).setParameter("jLevel", jobLevel).
					setFirstResult(recordIndex).setMaxResults(20).list();
		} else {
			queryBldr.append(" where jct_bs_status_id = 5 and jct_bs_soft_delete = 1 group by jct_bs_created_by");
			jrList = sessionFactory.getCurrentSession().
					 createSQLQuery(queryBldr.toString()).
					 setFirstResult(recordIndex).setMaxResults(20).list();
		}
		
		if (jrList.size() > 0) {
			int size = 0;
			Iterator itr = jrList.iterator();
			while (itr.hasNext()) {
				if (jrList.size() == 1) {
					try {
					maxHeaderId.append("'"+(BigInteger) itr.next()+"'");
					if(jrList.size()-1 > size){
						maxHeaderId.append(",");	
					} 
					size = size + 1;
					}catch(Exception e) {
						Iterator itrs = jrList.iterator();
						while (itrs.hasNext()) {
							Object[] obj = (Object[]) itrs.next();
							maxHeaderId.append("'" + (Integer) obj[0] + "'");
							if (jrList.size() - 1 > size) {
								maxHeaderId.append(",");	
							} 
							size = size + 1;
						}
					}
				} else {
					Object[] obj = (Object[]) itr.next();
					maxHeaderId.append("'" + (Integer) obj[0] + "'");
					if (jrList.size() - 1 > size) {
						maxHeaderId.append(",");	
					} 
					size = size + 1;
				}}
			queryBldr = new StringBuilder ("Select bsview.jct_bs_created_by, " +
					"bsview.jct_user_function_group, " +
					"bsview.jct_user_levels, bsview.jct_bs_task_desc, " +
					"bsview.jct_bs_energy, bsview.jct_bs_time_spent, " +
					"bsview.jct_first_name, bsview.jct_last_name " +
					"from jct_before_sketch_view bsview");
			if ((function.trim().length() > 0) && (jobLevel.trim().length() > 0)) {
				queryBldr.append(" where bsview.jct_user_function_group = :fnGrp " +
						"and bsview.jct_user_levels = :jLevel");
				queryBldr.append(" and bsview.jct_bs_header_id in (" + maxHeaderId.toString() + ") and jct_bs_status_id = 5 and jct_bs_soft_delete = 1");
				queryBldr.append(" and bsview.jct_bs_task_desc != '' ");
			} else if ((function.trim().length() > 0) && (jobLevel.trim().length() == 0)) {
				queryBldr.append(" where bsview.jct_user_function_group = :fnGrp");
				queryBldr.append(" and bsview.jct_bs_header_id in (" + maxHeaderId.toString() + ") and jct_bs_status_id = 5 and jct_bs_soft_delete = 1");
				queryBldr.append(" and bsview.jct_bs_task_desc != '' ");
			} else if ((function.trim().length() == 0) && (jobLevel.trim().length() > 0)) {
				queryBldr.append(" where bsview.jct_user_levels = :jLevel");
				queryBldr.append(" and bsview.jct_bs_header_id in (" + maxHeaderId.toString() + ") and jct_bs_status_id = 5 and jct_bs_soft_delete = 1");
				queryBldr.append(" and bsview.jct_bs_task_desc != '' ");
			} else {
				queryBldr.append(" where bsview.jct_bs_header_id in (" + maxHeaderId.toString() + ") and jct_bs_status_id = 5 and jct_bs_soft_delete = 1");
				queryBldr.append(" and bsview.jct_bs_task_desc != '' ");
			}
			queryBldr.append(" order by bsview.jct_bs_created_by");
			
			if ((function.trim().length() > 0) && (jobLevel.trim().length() > 0)) {
				returnList = sessionFactory.getCurrentSession().
						createSQLQuery(queryBldr.toString()).setParameter("fnGrp", function).
						setParameter("jLevel", jobLevel).list();
			} else if ((function.trim().length() > 0) && (jobLevel.trim().length() == 0)) {
				returnList = sessionFactory.getCurrentSession().
						createSQLQuery(queryBldr.toString()).setParameter("fnGrp", function)
						.list();
			} else if ((function.trim().length() == 0) && (jobLevel.trim().length() > 0)) {
				returnList = sessionFactory.getCurrentSession().
						createSQLQuery(queryBldr.toString()).setParameter("jLevel", jobLevel)
						.list();
			} else {
				returnList = sessionFactory.getCurrentSession().
						createSQLQuery(queryBldr.toString())
						.list();
			}	
		} else {
			returnList = new ArrayList<Object>();
		}
		LOGGER.info("<<<< ReportDAOImpl.getBeforeSketchListMax");
		return returnList;
	}
	@Transactional(propagation=Propagation.REQUIRED)
	public List<String> getTotalCount(List<String> emailIdList, String function, String jobLevel, int status, int softDel) throws DAOException {
		LOGGER.info(">>>> ReportDAOImpl.getTotalCount");
		StringBuilder jrNoBuilder = new StringBuilder("");
		List<String> jrList = null;
		jrNoBuilder = new StringBuilder("Select distinct(bsview.jct_bs_created_by) " +
				    						"from jct_before_sketch_view bsview");
		if ((function.trim().length() > 0) && (jobLevel.trim().length() > 0)) {
			if (emailIdList == null) {
			jrNoBuilder.append(" where bsview.jct_user_function_group = :fnGrp " +
					"and bsview.jct_user_levels = :jLevel  and jct_bs_status_id = "+status+" and jct_bs_soft_delete = "+softDel+"");
			} else {
				jrNoBuilder.append(" where bsview.jct_user_function_group = :fnGrp " +
						"and bsview.jct_user_levels = :jLevel  and jct_bs_status_id = "+status+" and jct_bs_soft_delete = "+softDel+" and bsview.jct_bs_created_by NOT IN :emailIdList");
			}
			if (emailIdList == null) {
				jrList = sessionFactory.getCurrentSession().
						createSQLQuery(jrNoBuilder.toString()).
						setParameter("fnGrp", function).setParameter("jLevel", jobLevel).list();
			} else {
				jrList = sessionFactory.getCurrentSession().
						createSQLQuery(jrNoBuilder.toString()).
						setParameter("fnGrp", function).setParameter("jLevel", jobLevel).
						setParameterList("emailIdList", emailIdList).list();
			}
		} else if ((function.trim().length() > 0) && (jobLevel.trim().length() == 0)) {
			if (emailIdList == null) {
				jrNoBuilder.append(" where bsview.jct_user_function_group = :fnGrp and jct_bs_status_id = "+status+" and jct_bs_soft_delete = "+softDel+"");
				jrList = sessionFactory.getCurrentSession().
						createSQLQuery(jrNoBuilder.toString()).setParameter("fnGrp", function).list();
			} else {
				jrNoBuilder.append(" where bsview.jct_user_function_group = :fnGrp and jct_bs_status_id = "+status+" and jct_bs_soft_delete = "+softDel+" and bsview.jct_bs_created_by NOT IN :emailIdList");
				jrList = sessionFactory.getCurrentSession().
						createSQLQuery(jrNoBuilder.toString()).setParameter("fnGrp", function).
						setParameterList("emailIdList", emailIdList).list();
			}
			
		} else if ((function.trim().length() == 0) && (jobLevel.trim().length() > 0)) {
			if (emailIdList == null) {
				jrNoBuilder.append(" where bsview.jct_user_levels = :jLevel and jct_bs_status_id = "+status+" and jct_bs_soft_delete = "+softDel+"");
				jrList = sessionFactory.getCurrentSession().
						createSQLQuery(jrNoBuilder.toString()).setParameter("jLevel", jobLevel).list();
			} else {
				jrNoBuilder.append(" where bsview.jct_user_levels = :jLevel and jct_bs_status_id = "+status+" and jct_bs_soft_delete = "+softDel+" and bsview.jct_bs_created_by NOT IN :emailIdList");
				jrList = sessionFactory.getCurrentSession().
						createSQLQuery(jrNoBuilder.toString()).setParameter("jLevel", jobLevel).
						setParameterList("emailIdList", emailIdList).list();
			}
			
		} else {
			if (emailIdList == null) {
				jrNoBuilder.append(" where jct_bs_status_id = "+status+" and jct_bs_soft_delete = "+softDel+"");
				jrList = sessionFactory.getCurrentSession().
						createSQLQuery(jrNoBuilder.toString()).list();
			} else {
				jrNoBuilder.append(" where jct_bs_status_id = "+status+" and jct_bs_soft_delete = "+softDel+" and bsview.jct_bs_created_by NOT IN :emailIdList");
				jrList = sessionFactory.getCurrentSession().
						createSQLQuery(jrNoBuilder.toString()).setParameterList("emailIdList", emailIdList).list();
			}
			
		}
		LOGGER.info("<<<< ReportDAOImpl.getTotalCount");
		return jrList;
	}
	@Transactional(propagation=Propagation.REQUIRED)
	public List<String> getTotalCountMax(String function, String jobLevel) throws DAOException {
		LOGGER.info(">>>> ReportDAOImpl.getTotalCountMax");
		StringBuilder jrNoBuilder = new StringBuilder("");
		List<String> jrList = null;
		jrNoBuilder = new StringBuilder("Select distinct(bsview.jct_bs_header_id) " +
				    						"from jct_before_sketch_view bsview");
		if ((function.trim().length() > 0) && (jobLevel.trim().length() > 0)) {
			jrNoBuilder.append(" where bsview.jct_user_function_group = :fnGrp " +
					"and bsview.jct_user_levels = :jLevel  and jct_bs_status_id = 5 and jct_bs_soft_delete = 1");
			jrList = sessionFactory.getCurrentSession().
					createSQLQuery(jrNoBuilder.toString()).
					setParameter("fnGrp", function).setParameter("jLevel", jobLevel).list();
		} else if ((function.trim().length() > 0) && (jobLevel.trim().length() == 0)) {
			jrNoBuilder.append(" where bsview.jct_user_function_group = :fnGrp and jct_bs_status_id = 5 and jct_bs_soft_delete = 1");
			jrList = sessionFactory.getCurrentSession().
					createSQLQuery(jrNoBuilder.toString()).setParameter("fnGrp", function).list();
		} else if ((function.trim().length() == 0) && (jobLevel.trim().length() > 0)) {
			jrNoBuilder.append(" where bsview.jct_user_levels = :jLevel and jct_bs_status_id = 5 and jct_bs_soft_delete = 1");
			jrList = sessionFactory.getCurrentSession().
					createSQLQuery(jrNoBuilder.toString()).setParameter("jLevel", jobLevel).list();
		} else {
			jrNoBuilder.append(" where jct_bs_status_id = 5 and jct_bs_soft_delete = 1");
			jrList = sessionFactory.getCurrentSession().
					createSQLQuery(jrNoBuilder.toString()).list();
		}
		LOGGER.info("<<<< ReportDAOImpl.getTotalCountMax");
		return jrList;
	}
	
	@Transactional(propagation=Propagation.REQUIRED)
	public List<Object> getBeforeSketchListForExcel(String occupationCode, int status, int softDelete) throws DAOException {
		LOGGER.info(">>>> ReportDAOImpl.getBeforeSketchList");
		StringBuilder queryBldr = null;
		StringBuilder jrNoBuilder = new StringBuilder("");
		List<Object> returnList = null;
		List<String> jrList = null;
		queryBldr = new StringBuilder("Select distinct(bsview.jct_bs_created_by) from jct_before_sketch_view bsview");
		if ((occupationCode.trim().length() > 0)) {
			queryBldr.append(" where bsview.jct_user_occupation = :occupationCode " +
					"					and jct_bs_status_id = "+status+" and jct_bs_soft_delete = "+softDelete+"");
			jrList = sessionFactory.getCurrentSession().
					 createSQLQuery(queryBldr.toString())
					 .setParameter("occupationCode", occupationCode).list();
		}  else {
			queryBldr.append(" where jct_bs_status_id = "+status+" and jct_bs_soft_delete = "+softDelete+"");
			jrList = sessionFactory.getCurrentSession().
					 createSQLQuery(queryBldr.toString()).list();
		}
		if (jrList.size() > 0) {
			int size = 0;
			Iterator itr = jrList.iterator();
			while (itr.hasNext()) {
				jrNoBuilder.append("'" + (String) itr.next() + "'");
				if (jrList.size() - 1 > size) {
					jrNoBuilder.append(",");	
				} 
				size = size + 1;
			}
			queryBldr = new StringBuilder ("Select bsview.jct_bs_created_by, " +
					"bsview.jct_user_occupation, " +
					"bsview.jct_user_levels, bsview.jct_bs_task_desc, " +
					"bsview.jct_bs_energy, bsview.jct_bs_time_spent, " +
					"bsview.jct_first_name, bsview.jct_last_name " +
					"from jct_before_sketch_view bsview");
			if ((occupationCode.trim().length() > 0)) {
				queryBldr.append(" where bsview.jct_user_occupation = :occupationCode ");
				queryBldr.append(" and bsview.jct_bs_created_by in (" + jrNoBuilder.toString() + ") and jct_bs_status_id = "+status+" and jct_bs_soft_delete = "+softDelete+"");
				queryBldr.append(" and bsview.jct_bs_task_desc != '' ");
				queryBldr.append(" order by bsview.jct_bs_created_by");
				returnList = sessionFactory.getCurrentSession().
						createSQLQuery(queryBldr.toString()).setParameter("occupationCode", occupationCode).list();
			} else {
				queryBldr.append(" where bsview.jct_bs_created_by in (" + jrNoBuilder.toString() + ") and jct_bs_status_id = "+status+" and jct_bs_soft_delete = "+softDelete+"");
				queryBldr.append(" and bsview.jct_bs_task_desc != '' ");
				queryBldr.append(" order by bsview.jct_bs_created_by");
				returnList = sessionFactory.getCurrentSession().
						createSQLQuery(queryBldr.toString())
						.list();
			}
		} else {
			returnList = new ArrayList<Object>();
		}
		LOGGER.info("<<<< ReportDAOImpl.getBeforeSketchList");
		return returnList;
		
	
	}
	
	@Transactional(propagation=Propagation.REQUIRED)
	public List<Object> getBeforeSketchListForExcelMax(String function, String jobLevel) throws DAOException {
		LOGGER.info(">>>> ReportDAOImpl.getBeforeSketchListForExcel");
		StringBuilder queryBldr = null;
		StringBuilder maxHeaderId = new StringBuilder("");
		List<Object> returnList = null;
		List<String> jrList = null;
		queryBldr = new StringBuilder("select max(jct_bs_header_id),jct_bs_created_by  from jct_before_sketch_view bsview ");
		if ((function.trim().length() > 0) && (jobLevel.trim().length() > 0)) {
			queryBldr.append(" where bsview.jct_user_function_group = :fnGrp " +
									"and bsview.jct_user_levels = :jLevel  and jct_bs_status_id = 5 and jct_bs_soft_delete = 1 group by jct_bs_created_by");
			jrList = sessionFactory.getCurrentSession().
					createSQLQuery(queryBldr.toString()).setParameter("fnGrp", function).
					setParameter("jLevel", jobLevel).list();
		} else if ((function.trim().length() > 0) && (jobLevel.trim().length() == 0)) {
			queryBldr.append(" where bsview.jct_user_function_group = :fnGrp and jct_bs_status_id = 5 and jct_bs_soft_delete = 1 group by jct_bs_created_by");
			jrList = sessionFactory.getCurrentSession().
					createSQLQuery(queryBldr.toString()).setParameter("fnGrp", function).list();
		} else if ((function.trim().length() == 0) && (jobLevel.trim().length() > 0)) {
			queryBldr.append(" where bsview.jct_user_levels = :jLevel and jct_bs_status_id = 5 and jct_bs_soft_delete = 1 group by jct_bs_created_by");
			jrList = sessionFactory.getCurrentSession().
					createSQLQuery(queryBldr.toString()).setParameter("jLevel", jobLevel).list();
		} else {
			queryBldr.append(" where jct_bs_status_id = 5 and jct_bs_soft_delete = 1 group by jct_bs_created_by");
			jrList = sessionFactory.getCurrentSession().
					createSQLQuery(queryBldr.toString()).list();
		}
		
		if (jrList.size() > 0) {
			int size = 0;
			Iterator itr = jrList.iterator();
			while (itr.hasNext()) {
				if (jrList.size() == 1) {
					try {
					maxHeaderId.append("'"+(BigInteger) itr.next()+"'");
					if(jrList.size()-1 > size){
						maxHeaderId.append(",");	
					} 
					size = size + 1;
					}catch(Exception e) {
						Iterator itrs = jrList.iterator();
						while (itrs.hasNext()) {
							Object[] obj = (Object[]) itrs.next();
							maxHeaderId.append("'" + (Integer) obj[0] + "'");
							if (jrList.size() - 1 > size) {
								maxHeaderId.append(",");	
							} 
							size = size + 1;
						}
					}
				} else {
					Object[] obj = (Object[]) itr.next();
					maxHeaderId.append("'" + (Integer) obj[0] + "'");
					if (jrList.size() - 1 > size) {
						maxHeaderId.append(",");	
					} 
					size = size + 1;
				}}
			queryBldr = new StringBuilder("Select bsview.jct_bs_created_by,bsview." +
					"jct_user_function_group, bsview.jct_user_levels, " +
					"bsview.jct_bs_task_desc, bsview.jct_bs_energy, " +
					"bsview.jct_first_name, bsview.jct_last_name " +
					"bsview.jct_bs_time_spent from jct_before_sketch_view bsview");
			if ((function.trim().length() > 0) && (jobLevel.trim().length() > 0)) {
				queryBldr.append(" where bsview.jct_user_function_group = :fnGrp " +
						"and bsview.jct_user_levels = :jLevel");
				queryBldr.append(" and bsview.jct_bs_header_id in (" + maxHeaderId.toString() + ") and jct_bs_status_id=5 and jct_bs_soft_delete=1");
			} else if ((function.trim().length() > 0) && (jobLevel.trim().length() == 0)) {
				queryBldr.append(" where bsview.jct_user_function_group = :fnGrp");
				queryBldr.append(" and bsview.jct_bs_header_id in (" + maxHeaderId.toString() + ") and jct_bs_status_id=5 and jct_bs_soft_delete=1");
			} else if ((function.trim().length() == 0) && (jobLevel.trim().length() > 0)) {
				queryBldr.append(" where bsview.jct_user_levels = :jLevel");
				queryBldr.append(" and bsview.jct_bs_header_id in (" + maxHeaderId.toString() + ") and jct_bs_status_id=5 and jct_bs_soft_delete=1");
			} else {
				queryBldr.append(" where bsview.jct_bs_header_id in (" + maxHeaderId.toString() + ") and jct_bs_status_id=5 and jct_bs_soft_delete=1");
			}
			queryBldr.append(" order by bsview.jct_bs_created_by");
			
			if ((function.trim().length() > 0) && (jobLevel.trim().length() > 0)) {
				returnList = sessionFactory.getCurrentSession().
						createSQLQuery(queryBldr.toString()).setParameter("fnGrp", function).
						setParameter("jLevel", jobLevel)
						.list();
			} else if ((function.trim().length() > 0) && (jobLevel.trim().length() == 0)) {
				returnList = sessionFactory.getCurrentSession().
						createSQLQuery(queryBldr.toString()).setParameter("fnGrp", function)
						.list();
			} else if ((function.trim().length() == 0) && (jobLevel.trim().length() > 0)) {
				returnList = sessionFactory.getCurrentSession().
						createSQLQuery(queryBldr.toString()).setParameter("jLevel", jobLevel)
						.list();
			} else {
				returnList = sessionFactory.getCurrentSession().
						createSQLQuery(queryBldr.toString())
						.list();
			}	
		} else {
			returnList = new ArrayList<Object>();
		}
		LOGGER.info("<<<< ReportDAOImpl.getBeforeSketchListForExcelMax");
		return returnList;
	}
	
	/***************************** ACTION PLAN******************************************/	
	@Transactional(propagation=Propagation.REQUIRED)
	public List<Object> getActionPlanList(List<String> emailIdList, String occupation, int recordIndex, int statusId, int softDel) throws DAOException {
		LOGGER.info(">>>> ReportDAOImpl.getActionPlanList");
		StringBuilder queryBldr = null;
		StringBuilder jrNoBuilder = new StringBuilder("");
		List<Object> returnList = null;
		List<String> jrList = null;
		queryBldr = new StringBuilder("Select distinct(apview.jct_as_created_by) from jct_action_plan_view apview");
		if ((occupation.trim().length() > 0) ) {
			queryBldr.append(" where apview.jct_user_occupation = :fnGrp " +
					"and jct_status_id="+statusId+" and jct_action_plan_soft_delete="+softDel+"");
			jrList = sessionFactory.getCurrentSession().
					createSQLQuery(queryBldr.toString())
					.setParameter("fnGrp", occupation)
					.setFirstResult(recordIndex)
					.setMaxResults(20).list();
		}  else {
			queryBldr.append(" where jct_status_id="+statusId+" and jct_action_plan_soft_delete="+softDel+"");
			jrList = sessionFactory.getCurrentSession().
					createSQLQuery(queryBldr.toString()).setFirstResult(recordIndex).
					setMaxResults(20).list();
		}
		
		if ((jrList != null) && (jrList.size() > 0)) {
			int size = 0;
			Iterator itr = jrList.iterator();
			while (itr.hasNext()) {
				jrNoBuilder.append("'" + (String) itr.next() + "'");
				if (jrList.size() - 1 > size) {
					jrNoBuilder.append(",");	
				} 
				size = size + 1;
			}
			queryBldr = new StringBuilder("Select apview.jct_as_created_by, " +
					"apview.jct_user_occupation, apview.jct_user_levels, " +
					"apview.jct_as_question_desc, apview.jct_as_question_sub_desc, " +
					"apview.jct_as_answar_desc,apview.jct_first_name,apview.jct_last_name from jct_action_plan_view apview");
			if ((occupation.trim().length() > 0) ) {
				queryBldr.append(" where apview.jct_user_occupation = :fnGrp ");
				queryBldr.append(" and apview.jct_as_created_by in (" + jrNoBuilder.toString() + ") and apview.jct_user_occupation IS NOT NULL and jct_status_id="+statusId+" and jct_action_plan_soft_delete="+softDel+"");
			}  else {
				queryBldr.append(" where apview.jct_as_created_by in (" + jrNoBuilder.toString() + ") and apview.jct_user_occupation IS NOT NULL and jct_status_id="+statusId+" and jct_action_plan_soft_delete="+softDel+"");
			}
					
			if ((null != emailIdList) && (emailIdList.size() > 0)) {
				queryBldr.append(" and jct_as_created_by NOT IN :list ");
			}
			queryBldr.append(" order by apview.jct_as_question_sub_desc");
			
			
			if ((occupation.trim().length() > 0) ) {
				if ((null != emailIdList) && (emailIdList.size() > 0)) {
					returnList = sessionFactory.getCurrentSession().
							createSQLQuery(queryBldr.toString()).setParameter("fnGrp", occupation)
							.setParameter("list", emailIdList)
							.list();
				} else {
					returnList = sessionFactory.getCurrentSession().
							createSQLQuery(queryBldr.toString()).setParameter("fnGrp", occupation)
							.list();
				}
			}  else {
				if ((null != emailIdList) && (emailIdList.size() > 0)) {
					returnList =returnList = sessionFactory.getCurrentSession().
							createSQLQuery(queryBldr.toString())
							.setParameterList("list", emailIdList)
							.list();
				} else {
					returnList = sessionFactory.getCurrentSession().createSQLQuery(queryBldr.toString())
							.list();
				}
				
			}
		} else {
			returnList = new ArrayList<Object>();
		}
		LOGGER.info("<<<< ReportDAOImpl.getActionPlanList");
		return returnList;
	}
	
	@Transactional(propagation=Propagation.REQUIRED)
	public List<Object> getActionPlanListMax(String function, String jobLevel, int recordIndex) throws DAOException {
		LOGGER.info(">>>> ReportDAOImpl.getActionPlanListMax");
		StringBuilder queryBldr = null;
		StringBuilder maxHeaderId = new StringBuilder("");
		List<Object> returnList = null;
		List<String> jrList = null;
		queryBldr = new StringBuilder("Select max(apview.jct_as_header_id) from jct_action_plan_view apview");
		if ((function.trim().length() > 0) && (jobLevel.trim().length() > 0)) {
			queryBldr.append(" where apview.jct_user_function_group = :fnGrp " +
					"and apview.jct_user_levels = :jLevel and jct_status_id=5 and jct_action_plan_soft_delete=1 group by jct_as_created_by");
			jrList = sessionFactory.getCurrentSession().
					createSQLQuery(queryBldr.toString()).setParameter("fnGrp", function).
					setParameter("jLevel", jobLevel).setFirstResult(recordIndex).
					setMaxResults(20).list();
		} else if ((function.trim().length() > 0) && (jobLevel.trim().length() == 0)) {
			queryBldr.append(" where apview.jct_user_function_group = :fnGrp and jct_status_id=5 and jct_action_plan_soft_delete=1 group by jct_as_created_by");
			jrList = sessionFactory.getCurrentSession().
					createSQLQuery(queryBldr.toString()).setParameter("fnGrp", function).
					setFirstResult(recordIndex).setMaxResults(20).list();
		} else if ((function.trim().length() == 0) && (jobLevel.trim().length() > 0)) {
			queryBldr.append(" where apview.jct_user_levels = :jLevel and jct_status_id=5 and jct_action_plan_soft_delete=1 group by jct_as_created_by");
			jrList = sessionFactory.getCurrentSession().
					createSQLQuery(queryBldr.toString()).setParameter("jLevel", jobLevel).
					setFirstResult(recordIndex).setMaxResults(20).list();
		} else {
			queryBldr.append(" where jct_status_id=5 and jct_action_plan_soft_delete=1 group by jct_as_created_by");
			jrList = sessionFactory.getCurrentSession().
					createSQLQuery(queryBldr.toString()).setFirstResult(recordIndex).
					setMaxResults(20).list();
		}
		
		if (jrList.size() > 0) {
			int size = 0;
			Iterator itr = jrList.iterator();
			while (itr.hasNext()) {
				if (jrList.size() == 1) {
					try {
					maxHeaderId.append("'"+(BigInteger) itr.next()+"'");
					if(jrList.size()-1 > size){
						maxHeaderId.append(",");	
					} 
					size = size + 1;
					}catch(Exception e) {
						Iterator itrs = jrList.iterator();
						while (itrs.hasNext()) {
							Object[] obj = (Object[]) itrs.next();
							maxHeaderId.append("'" + (Integer) obj[0] + "'");
							if (jrList.size() - 1 > size) {
								maxHeaderId.append(",");	
							} 
							size = size + 1;
						}
					}
				} else {
					Object[] obj = (Object[]) itr.next();
					maxHeaderId.append("'" + (Integer) obj[0] + "'");
					if (jrList.size() - 1 > size) {
						maxHeaderId.append(",");	
					} 
					size = size + 1;
				}}
			queryBldr = new StringBuilder("Select apview.jct_as_created_by, " +
					"apview.jct_user_function_group, apview.jct_user_levels, " +
					"apview.jct_as_question_desc, apview.jct_as_question_sub_desc, " +
					"apview.jct_as_answar_desc,apview.jct_first_name,apview.jct_last_name from jct_action_plan_view apview");
			if ((function.trim().length() > 0) && (jobLevel.trim().length() > 0)) {
				queryBldr.append(" where apview.jct_user_function_group = :fnGrp " +
						"and apview.jct_user_levels = :jLevel");
				queryBldr.append(" and apview.jct_as_header_id in (" + maxHeaderId.toString() + ") and jct_status_id=5 and jct_action_plan_soft_delete=1");
			} else if ((function.trim().length() > 0) && (jobLevel.trim().length() == 0)) {
				queryBldr.append(" where apview.jct_user_function_group = :fnGrp");
				queryBldr.append(" and apview.jct_as_header_id in (" + maxHeaderId.toString() + ") and jct_status_id=5 and jct_action_plan_soft_delete=1");
			} else if ((function.trim().length() == 0) && (jobLevel.trim().length() > 0)) {
				queryBldr.append(" where apview.jct_user_levels = :jLevel");
				queryBldr.append(" and apview.jct_as_header_id in (" + maxHeaderId.toString() + ") and jct_status_id=5 and jct_action_plan_soft_delete=1");
			} else {
				queryBldr.append(" where apview.jct_as_header_id in (" + maxHeaderId.toString() + ") and jct_status_id=5 and jct_action_plan_soft_delete=1");
			}
			queryBldr.append(" order by apview.jct_as_question_sub_desc");
			
			if ((function.trim().length() > 0) && (jobLevel.trim().length() > 0)) {
				returnList = sessionFactory.getCurrentSession().
						createSQLQuery(queryBldr.toString()).setParameter("fnGrp", function).
						setParameter("jLevel", jobLevel)
						.list();
			} else if ((function.trim().length() > 0) && (jobLevel.trim().length() == 0)) {
				returnList = sessionFactory.getCurrentSession().
						createSQLQuery(queryBldr.toString()).setParameter("fnGrp", function)
						.list();
			} else if ((function.trim().length() == 0) && (jobLevel.trim().length() > 0)) {
				returnList = sessionFactory.getCurrentSession().
						createSQLQuery(queryBldr.toString()).setParameter("jLevel", jobLevel)
						.list();
			} else {
				returnList = sessionFactory.getCurrentSession().createSQLQuery(queryBldr.toString())
						.list();
			}	
		} else {
			returnList = new ArrayList<Object>();
		}
		LOGGER.info("<<<< ReportDAOImpl.getActionPlanListMax");
		return returnList;
	}
	
	@Transactional(propagation=Propagation.REQUIRED)
	public List<Object> getAfterSketchList(List<String> emailIdList, String function, String jobLevel, int recordIndex, int statusId, int softDel) throws DAOException {
		LOGGER.info(">>>> ReportDAOImpl.getAfterSketchList");
		StringBuilder queryBldr = null;
		StringBuilder jrNoBuilder = new StringBuilder("");
		List<Object> returnList = null;
		List<String> jrList = null;
		queryBldr = new StringBuilder("Select distinct(asview.jct_as_created_by) from jct_after_sketch_view asview");
		if ((function.trim().length() > 0) && (jobLevel.trim().length() > 0)) {
			queryBldr.append(" where asview.jct_user_function_group = :fnGrp " +
					"and asview.jct_user_levels = :jLevel and jct_as_status_id = "+statusId+" and jct_as_soft_delete = "+softDel+"");
			jrList = sessionFactory.getCurrentSession().
					createSQLQuery(queryBldr.toString()).setParameter("fnGrp", function).
					setParameter("jLevel", jobLevel).setFirstResult(recordIndex).
					setMaxResults(20).list();
		} else if ((function.trim().length() > 0) && (jobLevel.trim().length() == 0)) {
			queryBldr.append(" where asview.jct_user_function_group = :fnGrp and jct_as_status_id = "+statusId+" and jct_as_soft_delete = "+softDel+"");
			jrList = sessionFactory.getCurrentSession().
					createSQLQuery(queryBldr.toString()).
					setParameter("fnGrp", function).setFirstResult(recordIndex).
					setMaxResults(20).list();
		} else if ((function.trim().length() == 0) && (jobLevel.trim().length() > 0)) {
			queryBldr.append(" where asview.jct_user_levels = :jLevel and jct_as_status_id = "+statusId+" and jct_as_soft_delete = "+softDel+"");
			jrList = sessionFactory.getCurrentSession().
					createSQLQuery(queryBldr.toString()).
					setParameter("jLevel", jobLevel).setFirstResult(recordIndex).
					setMaxResults(20).list();
		} else {
			queryBldr.append(" where jct_as_status_id = "+statusId+" and jct_as_soft_delete = "+softDel+"");
			jrList = sessionFactory.getCurrentSession().
					createSQLQuery(queryBldr.toString()).setFirstResult(recordIndex).
					setMaxResults(20).list();
		}
		
		if (jrList.size() > 0) {
			int size = 0;
			Iterator itr = jrList.iterator();
			while (itr.hasNext()) {
				jrNoBuilder.append("'" + (String) itr.next() + "'");
				if (jrList.size() - 1 > size) {
					jrNoBuilder.append(",");	
				} 
				size = size + 1;
			}
			queryBldr = new StringBuilder("Select asview.jct_as_created_by, " 		//0	
										+ "asview.jct_as_role_desc, "				//1
										+ "asview.jct_as_task_desc, "				//2
										+ "asview.jct_as_task_energy, "				//3
										+ "asview.jct_as_finalpage_time_spent, "	//4
										+ "asview.jct_as_element_desc, " 			//5
										+ "asview.jct_user_occupation , " 		    //6
										+ "asview.jct_user_levels, "				//7
										+ "asview.jct_first_name, "					//8
										+ "asview.jct_last_name "                   //,9
										+ "from jct_after_sketch_view asview");
			if ((function.trim().length() > 0) && (jobLevel.trim().length() > 0)) {
				queryBldr.append(" where asview.jct_user_function_group = :fnGrp " +
						"and asview.jct_user_levels = :jLevel");
				queryBldr.append(" and asview.jct_as_created_by in (" + jrNoBuilder.toString() + ") and jct_as_status_id = "+statusId+" and jct_as_soft_delete = "+softDel+"");
			} else if ((function.trim().length() > 0) && (jobLevel.trim().length() == 0)) {
				queryBldr.append(" where asview.jct_user_function_group = :fnGrp");
				queryBldr.append(" and asview.jct_as_created_by in (" + jrNoBuilder.toString() + ") and jct_as_status_id = "+statusId+" and jct_as_soft_delete = "+softDel+"");
			} else if ((function.trim().length() == 0) && (jobLevel.trim().length() > 0)) {
				queryBldr.append(" where asview.jct_user_levels = :jLevel");
				queryBldr.append(" and asview.jct_as_created_by in (" + jrNoBuilder.toString() + ") and jct_as_status_id = "+statusId+" and jct_as_soft_delete = "+softDel+"");
			} else {
				queryBldr.append(" where asview.jct_as_created_by in (" + jrNoBuilder.toString() + ") and jct_as_status_id = "+statusId+" and jct_as_soft_delete = "+softDel+"");
			}
			if (null != emailIdList) {
				queryBldr.append(" and jct_as_created_by NOT IN :list ");
			}
			queryBldr.append(" order by asview.jct_as_role_desc desc, " +
					"asview.jct_as_created_by, asview.jct_as_task_desc desc");
			
			if ((function.trim().length() > 0) && (jobLevel.trim().length() > 0)) {
						if (null != emailIdList) {
							returnList = sessionFactory.getCurrentSession().
									createSQLQuery(queryBldr.toString()).setParameter("fnGrp", function).
									setParameter("jLevel", jobLevel).
									setParameterList("list", emailIdList)
									.list();
						} else {
							returnList = sessionFactory.getCurrentSession().
									createSQLQuery(queryBldr.toString()).setParameter("fnGrp", function).
									setParameter("jLevel", jobLevel)
									.list();
						}
			} else if ((function.trim().length() > 0) && (jobLevel.trim().length() == 0)) {
				if (null != emailIdList) {
					returnList = sessionFactory.getCurrentSession().
							createSQLQuery(queryBldr.toString()).setParameter("fnGrp", function)
							.setParameterList("list", emailIdList)
							.list();
				} else {
					returnList = sessionFactory.getCurrentSession().
							createSQLQuery(queryBldr.toString()).setParameter("fnGrp", function)
							.list();
				}
				
			} else if ((function.trim().length() == 0) && (jobLevel.trim().length() > 0)) {
				if (null != emailIdList) {
					returnList =returnList = sessionFactory.getCurrentSession().
							createSQLQuery(queryBldr.toString()).setParameter("jLevel", jobLevel)
							.setParameterList("list", emailIdList)
							.list();
				} else {
					returnList = sessionFactory.getCurrentSession().
							createSQLQuery(queryBldr.toString()).setParameter("jLevel", jobLevel)
							.list();
				}
			} else {
				if (null != emailIdList) {
					returnList =returnList = sessionFactory.getCurrentSession().
							createSQLQuery(queryBldr.toString())
							.setParameterList("list", emailIdList)
							.list();
				} else {
					returnList = sessionFactory.getCurrentSession().createSQLQuery(queryBldr.toString())
							.list();
				}
				
			}	
		} else {
			returnList = new ArrayList<Object>();
		}
		LOGGER.info("<<<< ReportDAOImpl.getAfterSketchList");
		return returnList;
	}
	
	@Transactional(propagation=Propagation.REQUIRED)
	public List<Object> getAfterSketchListMax(String function, String jobLevel, int recordIndex) throws DAOException {
		LOGGER.info(">>>> ReportDAOImpl.getAfterSketchListMax");
		StringBuilder queryBldr = null;
		StringBuilder maxHeaderId = new StringBuilder("");
		List<Object> returnList = null;
		List<String> jrList = null;
		queryBldr = new StringBuilder("select max(jct_as_header_id),jct_as_created_by from jct_after_sketch_view asview");
		if ((function.trim().length() > 0) && (jobLevel.trim().length() > 0)) {
			queryBldr.append(" where asview.jct_user_function_group = :fnGrp " +
					"and asview.jct_user_levels = :jLevel and jct_as_status_id = 5 and jct_as_soft_delete = 1 group by jct_as_created_by");
			jrList = sessionFactory.getCurrentSession().
					createSQLQuery(queryBldr.toString()).setParameter("fnGrp", function).
					setParameter("jLevel", jobLevel).setFirstResult(recordIndex).
					setMaxResults(20).list();
		} else if ((function.trim().length() > 0) && (jobLevel.trim().length() == 0)) {
			queryBldr.append(" where asview.jct_user_function_group = :fnGrp and jct_as_status_id = 5 and jct_as_soft_delete = 1 group by jct_as_created_by");
			jrList = sessionFactory.getCurrentSession().
					createSQLQuery(queryBldr.toString()).
					setParameter("fnGrp", function).setFirstResult(recordIndex).
					setMaxResults(20).list();
		} else if ((function.trim().length() == 0) && (jobLevel.trim().length() > 0)) {
			queryBldr.append(" where asview.jct_user_levels = :jLevel and jct_as_status_id = 5 and jct_as_soft_delete = 1 group by jct_as_created_by");
			jrList = sessionFactory.getCurrentSession().
					createSQLQuery(queryBldr.toString()).
					setParameter("jLevel", jobLevel).setFirstResult(recordIndex).
					setMaxResults(20).list();
		} else {
			queryBldr.append(" where jct_as_status_id = 5 and jct_as_soft_delete = 1 group by jct_as_created_by");
			jrList = sessionFactory.getCurrentSession().
					createSQLQuery(queryBldr.toString()).setFirstResult(recordIndex).
					setMaxResults(20).list();
		}
		if (jrList.size() > 0) {
			int size = 0;
			Iterator itr = jrList.iterator();
			while (itr.hasNext()) {
				if (jrList.size() == 1) {
					try {
					maxHeaderId.append("'"+(BigInteger) itr.next()+"'");
					if(jrList.size()-1 > size){
						maxHeaderId.append(",");	
					} 
					size = size + 1;
					}catch(Exception e) {
						Iterator itrs = jrList.iterator();
						while (itrs.hasNext()) {
							Object[] obj = (Object[]) itrs.next();
							maxHeaderId.append("'" + (BigInteger) obj[0] + "'");
							if (jrList.size() - 1 > size) {
								maxHeaderId.append(",");	
							} 
							size = size + 1;
						}
					}
				} else {
					Object[] obj = (Object[]) itr.next();
					maxHeaderId.append("'" + (Integer) obj[0] + "'");
					if (jrList.size() - 1 > size) {
						maxHeaderId.append(",");	
					} 
					size = size + 1;
				}}
			queryBldr = new StringBuilder("Select asview.jct_as_created_by, " 		//0	
										+ "asview.jct_as_role_desc, "				//1
										+ "asview.jct_as_task_desc, "				//2
										+ "asview.jct_as_task_energy, "				//3
										+ "asview.jct_as_finalpage_time_spent, "	//4
										+ "asview.jct_as_element_desc, " 			//5
										+ "asview.jct_user_function_group , " 		//6
										+ "asview.jct_user_levels, "					//7
										+ "asview.jct_first_name, "					//8
										+ "asview.jct_last_name "                   //,9
										+ "from jct_after_sketch_view asview");
			if ((function.trim().length() > 0) && (jobLevel.trim().length() > 0)) {
				queryBldr.append(" where asview.jct_user_function_group = :fnGrp " +
						"and asview.jct_user_levels = :jLevel");
				queryBldr.append(" and asview.jct_as_header_id in (" + maxHeaderId.toString() + ") and jct_as_status_id = 5 and jct_as_soft_delete = 1");
			} else if ((function.trim().length() > 0) && (jobLevel.trim().length() == 0)) {
				queryBldr.append(" where asview.jct_user_function_group = :fnGrp");
				queryBldr.append(" and asview.jct_as_header_id in (" + maxHeaderId.toString() + ") and jct_as_status_id = 5 and jct_as_soft_delete = 1");
			} else if ((function.trim().length() == 0) && (jobLevel.trim().length() > 0)) {
				queryBldr.append(" where asview.jct_user_levels = :jLevel");
				queryBldr.append(" and asview.jct_as_header_id in (" + maxHeaderId.toString() + ") and jct_as_status_id = 5 and jct_as_soft_delete = 1");
			} else {
				queryBldr.append(" where asview.jct_as_header_id in (" + maxHeaderId.toString() + ") and jct_as_status_id = 5 and jct_as_soft_delete = 1");
			}
			queryBldr.append(" order by asview.jct_as_role_desc desc, " +
					"asview.jct_as_created_by, asview.jct_as_task_desc desc");
			
			if ((function.trim().length() > 0) && (jobLevel.trim().length() > 0)) {
				returnList = sessionFactory.getCurrentSession().
						createSQLQuery(queryBldr.toString()).setParameter("fnGrp", function).
						setParameter("jLevel", jobLevel)
						.list();
			} else if ((function.trim().length() > 0) && (jobLevel.trim().length() == 0)) {
				returnList = sessionFactory.getCurrentSession().
						createSQLQuery(queryBldr.toString()).setParameter("fnGrp", function)
						.list();
			} else if ((function.trim().length() == 0) && (jobLevel.trim().length() > 0)) {
				returnList = sessionFactory.getCurrentSession().
						createSQLQuery(queryBldr.toString()).setParameter("jLevel", jobLevel)
						.list();
			} else {
				returnList = sessionFactory.getCurrentSession().createSQLQuery(queryBldr.toString())
						.list();
			}	
		} else {
			returnList = new ArrayList<Object>();
		}
		LOGGER.info("<<<< ReportDAOImpl.getAfterSketchListMax");
		return returnList;
	}
	
	@Transactional(propagation=Propagation.REQUIRED)
	public java.math.BigInteger getTotalTaskCount(String emailId, int statusId, int softDel) throws DAOException {
		LOGGER.info(">>>> ReportDAOImpl.getTotalTaskCount");
		String taskString = "select count( distinct jct_as_task_desc ) from jct_after_sketch_view asview where "
				+ "jct_as_created_by = :email and jct_as_task_desc != '' and jct_as_status_id = "+statusId+" and jct_as_soft_delete = "+softDel+" ";
		LOGGER.info("<<<< ReportDAOImpl.getTotalTaskCount");
		return (java.math.BigInteger)sessionFactory.getCurrentSession().createSQLQuery(taskString).setParameter("email", emailId).uniqueResult();
	}
	
	@Transactional(propagation=Propagation.REQUIRED)
	public java.math.BigInteger getTotalMappingCount(String emailId) throws DAOException {
		LOGGER.info(">>>> ReportDAOImpl.getTotalMappingCount");
		String taskString = "select count( distinct jct_as_element_desc ) from jct_after_sketch_view asview where "
				+ "jct_as_created_by = :email and jct_as_element_desc != '' and jct_as_soft_delete = 0";
		LOGGER.info("<<<< ReportDAOImpl.getTotalMappingCount");
		return (java.math.BigInteger)sessionFactory.getCurrentSession().createSQLQuery(taskString).setParameter("email", emailId).uniqueResult();
	}
	
	@Transactional(propagation=Propagation.REQUIRED)
	public java.math.BigInteger getTotalRoleCount(String emailId, int statusId, int softDel) throws DAOException {
		LOGGER.info(">>>> ReportDAOImpl.getTotalRoleCount");
		StringBuffer sb = new StringBuffer("select count( distinct jct_as_role_desc ) from jct_after_sketch_view asview where " +
				"jct_as_created_by = :email and jct_as_role_desc != '' and jct_as_status_id = "+statusId+" and jct_as_soft_delete = "+softDel+" " +
						"and jct_as_role_desc != 'N/A'");
		LOGGER.info("<<<< ReportDAOImpl.getTotalRoleCount");
		return (java.math.BigInteger)sessionFactory.getCurrentSession().createSQLQuery(sb.toString()).setParameter("email", emailId).uniqueResult();
	}
	
	@Transactional(propagation=Propagation.REQUIRED)
	public List<String> getASTotalCount(List<String> emailIdList, String function, String jobLevel, int statusId, int softDel) throws DAOException {
		LOGGER.info(">>>> ReportDAOImpl.getASTotalCount");
		StringBuilder jrNoBuilder = new StringBuilder("");
		List<String> jrList = null;
		jrNoBuilder = new StringBuilder("Select distinct(asview.jct_as_created_by) from jct_after_sketch_view asview");
		if ((function.trim().length() > 0) && (jobLevel.trim().length() > 0)) {
			if (null != emailIdList) {
				jrNoBuilder.append(" where asview.jct_user_function_group = :fnGrp " +
						"and asview.jct_user_levels = :jLevel and jct_as_status_id="+statusId+" and jct_as_soft_delete="+softDel+" and jct_as_created_by NOT IN :emailList");
				jrList = sessionFactory.getCurrentSession().createSQLQuery(jrNoBuilder.toString()).
						setParameter("fnGrp", function).setParameter("jLevel", jobLevel).
						setParameterList("emailList", emailIdList).list();
			} else {
				jrNoBuilder.append(" where asview.jct_user_function_group = :fnGrp " +
						"and asview.jct_user_levels = :jLevel and jct_as_status_id="+statusId+" and jct_as_soft_delete="+softDel+"");
				jrList = sessionFactory.getCurrentSession().createSQLQuery(jrNoBuilder.toString()).
						setParameter("fnGrp", function).setParameter("jLevel", jobLevel).list();
			}
			
		} else if ((function.trim().length() > 0) && (jobLevel.trim().length() == 0)) {
			if (null != emailIdList) {
				jrNoBuilder.append(" where asview.jct_user_function_group = :fnGrp " +
						"and jct_as_status_id="+statusId+" and jct_as_soft_delete="+softDel+" and jct_as_created_by NOT IN :emailList");
				jrList = sessionFactory.getCurrentSession().createSQLQuery(jrNoBuilder.toString()).
						setParameter("fnGrp", function).
						setParameterList("emailList", emailIdList).list();
			} else {
				jrNoBuilder.append(" where asview.jct_user_function_group = :fnGrp and jct_as_status_id="+statusId+" and jct_as_soft_delete="+softDel+"");
				jrList = sessionFactory.getCurrentSession().createSQLQuery(jrNoBuilder.toString()).
						setParameter("fnGrp", function).list();
			}
		} else if((function.trim().length() == 0) && (jobLevel.trim().length() > 0)){
			if (null != emailIdList) {
				jrNoBuilder.append(" where " +
						"asview.jct_user_levels = :jLevel and jct_as_status_id="+statusId+" and jct_as_soft_delete="+softDel+" and jct_as_created_by NOT IN :emailList");
				jrList = sessionFactory.getCurrentSession().createSQLQuery(jrNoBuilder.toString()).
						setParameter("jLevel", function).
						setParameterList("emailList", emailIdList).list();
			} else {
				jrNoBuilder.append(" where asview.jct_user_levels = :jLevel and jct_as_status_id="+statusId+" and jct_as_soft_delete="+softDel+"");
				jrList = sessionFactory.getCurrentSession().createSQLQuery(jrNoBuilder.toString()).
						setParameter("jLevel", jobLevel).list();
			}
		} else {
			if (null != emailIdList) {
				jrNoBuilder.append(" where jct_as_status_id="+statusId+" and jct_as_soft_delete="+softDel+" and jct_as_created_by NOT IN :emailList");
				jrList = sessionFactory.getCurrentSession().createSQLQuery(jrNoBuilder.toString()).
						setParameterList("emailList", emailIdList).list();
			} else {
				jrNoBuilder.append(" where jct_as_status_id="+statusId+" and jct_as_soft_delete="+softDel+"");
				jrList = sessionFactory.getCurrentSession().createSQLQuery(jrNoBuilder.toString()).list();
			}
		}
		LOGGER.info("<<<< ReportDAOImpl.getASTotalCount");
		return jrList;
	}
	
	@Transactional(propagation=Propagation.REQUIRED)
	public List<String> getASTotalCountMax(String function, String jobLevel) throws DAOException {
		LOGGER.info(">>>> ReportDAOImpl.getASTotalCountMax");
		StringBuilder jrNoBuilder = new StringBuilder("");
		List<String> jrList = null;
		jrNoBuilder = new StringBuilder("Select distinct(asview.jct_as_header_id) from jct_after_sketch_view asview");
		if ((function.trim().length() > 0) && (jobLevel.trim().length() > 0)) {
			jrNoBuilder.append(" where asview.jct_user_function_group = :fnGrp " +
					"and asview.jct_user_levels = :jLevel and jct_as_status_id=5 and jct_as_soft_delete=1");
			jrList = sessionFactory.getCurrentSession().createSQLQuery(jrNoBuilder.toString()).
					setParameter("fnGrp", function).setParameter("jLevel", jobLevel).list();
		} else if ((function.trim().length() > 0) && (jobLevel.trim().length() == 0)) {
			jrNoBuilder.append(" where asview.jct_user_function_group = :fnGrp and jct_as_status_id=5 and jct_as_soft_delete=1");
			jrList = sessionFactory.getCurrentSession().createSQLQuery(jrNoBuilder.toString()).
					setParameter("fnGrp", function).list();
		} else if((function.trim().length() == 0) && (jobLevel.trim().length() > 0)){
			jrNoBuilder.append(" where asview.jct_user_levels = :jLevel and jct_as_status_id=5 and jct_as_soft_delete=1");
			jrList = sessionFactory.getCurrentSession().createSQLQuery(jrNoBuilder.toString()).
					setParameter("jLevel", jobLevel).list();
		} else {
			jrNoBuilder.append(" where jct_as_status_id=5 and jct_as_soft_delete=1");
			jrList = sessionFactory.getCurrentSession().createSQLQuery(jrNoBuilder.toString()).list();
		}
		LOGGER.info("<<<< ReportDAOImpl.getASTotalCountMax");
		return jrList;
	}

	@Transactional(propagation=Propagation.REQUIRED)
	public List<String> getTotalCountActionPlan(List<String> emailIdList, String occupation, int statusId, int softDel) throws DAOException {
		LOGGER.info(">>>> ReportDAOImpl.getTotalCountActionPlan");
		StringBuilder jrNoBuilder = new StringBuilder("");
		List<String> jrList = null;
		jrNoBuilder = new StringBuilder("Select distinct(apview.jct_as_created_by) from jct_action_plan_view apview");
		if((occupation.trim().length() > 0)){
			if (emailIdList == null) {
				jrNoBuilder.append(" where apview.jct_user_occupation = :fnGrp and jct_status_id="+statusId+" and jct_action_plan_soft_delete="+softDel+"");
				jrList = sessionFactory.getCurrentSession().createSQLQuery(jrNoBuilder.toString()).setParameter("fnGrp", occupation).list();
			} else {
				jrNoBuilder.append(" where apview.jct_user_occupation = :fnGrp and jct_status_id="+statusId+" and jct_action_plan_soft_delete="+softDel+" and apview.jct_as_created_by NOT IN :emailIdList");
				jrList = sessionFactory.getCurrentSession().createSQLQuery(jrNoBuilder.toString()).
						setParameter("fnGrp", occupation)
						.setParameterList("emailIdList", emailIdList).list();
			}
			
		}  else {
			if (emailIdList == null) {
				jrNoBuilder.append(" where jct_status_id="+statusId+" and jct_action_plan_soft_delete="+softDel+"");
				jrList = sessionFactory.getCurrentSession().createSQLQuery(jrNoBuilder.toString()).list();
			} else {
				jrNoBuilder.append(" where jct_status_id="+statusId+" and jct_action_plan_soft_delete="+softDel+" and apview.jct_as_created_by NOT IN :emailIdList");
				jrList = sessionFactory.getCurrentSession().createSQLQuery(jrNoBuilder.toString()).
						setParameterList("emailIdList", emailIdList).list();
			}
			
		}
		LOGGER.info("<<<< ReportDAOImpl.getTotalCountActionPlan");
		return jrList;
	}
	
	@Transactional(propagation=Propagation.REQUIRED)
	public List<String> getTotalCountActionPlanMax(String function, String jobLevel) throws DAOException {
		LOGGER.info(">>>> ReportDAOImpl.getTotalCountActionPlanMax");
		StringBuilder jrNoBuilder = new StringBuilder("");
		List<String> jrList = null;
		jrNoBuilder = new StringBuilder("Select distinct(apview.jct_as_created_by) from jct_action_plan_view apview");
		if((function.trim().length() > 0) && (jobLevel.trim().length() > 0)){
			jrNoBuilder.append(" where apview.jct_user_function_group = :fnGrp and apview.jct_user_levels = :jLevel and jct_status_id=5 and jct_action_plan_soft_delete=1");
			jrList = sessionFactory.getCurrentSession().createSQLQuery(jrNoBuilder.toString()).setParameter("fnGrp", function).setParameter("jLevel", jobLevel).list();
		} else if((function.trim().length() > 0) && (jobLevel.trim().length() == 0)){
			jrNoBuilder.append(" where apview.jct_user_function_group = :fnGrp and jct_status_id=5 and jct_action_plan_soft_delete=1");
			jrList = sessionFactory.getCurrentSession().createSQLQuery(jrNoBuilder.toString()).setParameter("fnGrp", function).list();
		} else if((function.trim().length() == 0) && (jobLevel.trim().length() > 0)){
			jrNoBuilder.append(" where apview.jct_user_levels = :jLevel and jct_status_id=5 and jct_action_plan_soft_delete=1");
			jrList = sessionFactory.getCurrentSession().createSQLQuery(jrNoBuilder.toString()).setParameter("jLevel", jobLevel).list();
		} else {
			jrNoBuilder.append(" where jct_status_id=5 and jct_action_plan_soft_delete=1");
			jrList = sessionFactory.getCurrentSession().createSQLQuery(jrNoBuilder.toString()).list();
		}
		LOGGER.info("<<<< ReportDAOImpl.getTotalCountActionPlanMax");
		return jrList;
	}

	@Transactional(propagation=Propagation.REQUIRED)
	public List<Object> getActionPlanListForExcel(List<String> emailIdList, String occupation, int statusId, int softDel) throws DAOException {
		LOGGER.info(">>>> ReportDAOImpl.getActionPlanListForExcel");
		StringBuilder queryBldr = null;
		StringBuilder jrNoBuilder = new StringBuilder("");
		List<Object> returnList = null;
		List<String> jrList = null;
		queryBldr = new StringBuilder("Select distinct(apview.jct_as_created_by) from jct_action_plan_view apview");
		if((occupation.trim().length() > 0)){
			queryBldr.append(" where apview.jct_user_occupation = :fnGrp and jct_status_id="+statusId+" and jct_action_plan_soft_delete="+softDel+"");
			jrList = sessionFactory.getCurrentSession().createSQLQuery(queryBldr.toString()).setParameter("fnGrp", occupation).list();
		} else {
			queryBldr.append(" where jct_status_id="+statusId+" and jct_action_plan_soft_delete="+softDel+"");
			jrList = sessionFactory.getCurrentSession().createSQLQuery(queryBldr.toString()).list();
		}
		
		if((jrList != null) && (jrList.size() > 0)){
			int size = 0;
			Iterator itr = jrList.iterator();
			while(itr.hasNext()){
				jrNoBuilder.append("'"+(String)itr.next()+"'");
				if(jrList.size()-1 > size){
					jrNoBuilder.append(",");	
				} 
				size = size + 1;
			}
			queryBldr = new StringBuilder("Select apview.jct_as_created_by, apview.jct_user_occupation, apview.jct_user_levels, apview.jct_as_question_desc, apview.jct_as_question_sub_desc, apview.jct_as_answar_desc,apview.jct_first_name,apview.jct_last_name from jct_action_plan_view apview");
			if((occupation.trim().length() > 0)){
				queryBldr.append(" where apview.jct_user_occupation = :fnGrp");
				queryBldr.append(" and apview.jct_as_created_by in ("+jrNoBuilder.toString()+") and apview.jct_user_occupation IS NOT NULL and jct_status_id="+statusId+" and jct_action_plan_soft_delete="+softDel+"");
			} else {
				queryBldr.append(" where apview.jct_as_created_by in ("+jrNoBuilder.toString()+") and apview.jct_user_occupation IS NOT NULL and jct_status_id="+statusId+" and jct_action_plan_soft_delete="+softDel+"");
			}
			if (null != emailIdList) {
				queryBldr.append(" and jct_as_created_by NOT IN :list ");
			}
			queryBldr.append(" order by apview.jct_as_question_sub_desc");
			
			/**    		NEW CHANGED         **/
			if ((occupation.trim().length() > 0)) {
				if (null != emailIdList) {
					returnList = sessionFactory.getCurrentSession().
							createSQLQuery(queryBldr.toString()).setParameter("fnGrp", occupation)
							.setParameterList("list", emailIdList)
							.list();
				} else {
					returnList = sessionFactory.getCurrentSession().
							createSQLQuery(queryBldr.toString()).setParameter("fnGrp", occupation)
							.list();
				}
			}else {
				if (null != emailIdList) {
					returnList =returnList = sessionFactory.getCurrentSession().
							createSQLQuery(queryBldr.toString())
							.setParameterList("list", emailIdList)
							.list();
				} else {
					returnList = sessionFactory.getCurrentSession().createSQLQuery(queryBldr.toString())
							.list();
				}
				
			}			
		} else {
			returnList = new ArrayList<Object>();
		}
		LOGGER.info("<<<< ReportDAOImpl.getActionPlanListForExcel");
		return returnList;
	}
	
	@Transactional(propagation=Propagation.REQUIRED)
	public List<Object> getActionPlanListForExcelMax(String function, String jobLevel) throws DAOException {
		LOGGER.info(">>>> ReportDAOImpl.getActionPlanListForExcelMax");
		StringBuilder queryBldr = null;
		StringBuilder maxHeaderId = new StringBuilder("");
		List<Object> returnList = null;
		List<String> jrList = null;
		queryBldr = new StringBuilder("Select max(apview.jct_as_header_id) from jct_action_plan_view apview");
		if((function.trim().length() > 0) && (jobLevel.trim().length() > 0)){
			queryBldr.append(" where apview.jct_user_function_group = :fnGrp and apview.jct_user_levels = :jLevel and jct_status_id=5 and jct_action_plan_soft_delete=1 group by jct_as_created_by");
			jrList = sessionFactory.getCurrentSession().createSQLQuery(queryBldr.toString()).setParameter("fnGrp", function).setParameter("jLevel", jobLevel).list();
		} else if((function.trim().length() > 0) && (jobLevel.trim().length() == 0)){
			queryBldr.append(" where apview.jct_user_function_group = :fnGrp and jct_status_id=5 and jct_action_plan_soft_delete=1 group by jct_as_created_by");
			jrList = sessionFactory.getCurrentSession().createSQLQuery(queryBldr.toString()).setParameter("fnGrp", function).list();
		} else if((function.trim().length() == 0) && (jobLevel.trim().length() > 0)){
			queryBldr.append(" where apview.jct_user_levels = :jLevel and jct_status_id=5 and jct_action_plan_soft_delete=1 group by jct_as_created_by");
			jrList = sessionFactory.getCurrentSession().createSQLQuery(queryBldr.toString()).setParameter("jLevel", jobLevel).list();
		} else {
			queryBldr.append(" where jct_status_id=5 and jct_action_plan_soft_delete=1 group by jct_as_created_by");
			jrList = sessionFactory.getCurrentSession().createSQLQuery(queryBldr.toString()).list();
		}
		
		if(jrList.size() > 0){
			int size = 0;
			Iterator itr = jrList.iterator();
			while(itr.hasNext()){
				if (jrList.size() == 1) {
					try {
					maxHeaderId.append("'"+(BigInteger) itr.next()+"'");
					if(jrList.size()-1 > size){
						maxHeaderId.append(",");	
					} 
					size = size + 1;
					}catch(Exception e) {
						Iterator itrs = jrList.iterator();
						while (itrs.hasNext()) {
							Object[] obj = (Object[]) itrs.next();
							maxHeaderId.append("'" + (Integer) obj[0] + "'");
							if (jrList.size() - 1 > size) {
								maxHeaderId.append(",");	
							} 
							size = size + 1;
						}
					}
				} else {
					Object[] obj = (Object[]) itr.next();
					maxHeaderId.append("'" + (Integer) obj[0] + "'");
					if (jrList.size() - 1 > size) {
						maxHeaderId.append(",");	
					} 
					size = size + 1;
				}}
			queryBldr = new StringBuilder("Select apview.jct_as_created_by, apview.jct_user_function_group, apview.jct_user_levels, apview.jct_as_question_desc, apview.jct_as_question_sub_desc, apview.jct_as_answar_desc,apview.jct_first_name,apview.jct_last_name from jct_action_plan_view apview");
			if((function.trim().length() > 0) && (jobLevel.trim().length() > 0)){
				queryBldr.append(" where apview.jct_user_function_group = :fnGrp and apview.jct_user_levels = :jLevel");
				queryBldr.append(" and apview.jct_as_header_id in ("+maxHeaderId.toString()+") and jct_status_id=5 and jct_action_plan_soft_delete=1");
			} else if((function.trim().length() > 0) && (jobLevel.trim().length() == 0)){
				queryBldr.append(" where apview.jct_user_function_group = :fnGrp");
				queryBldr.append(" and apview.jct_as_header_id in ("+maxHeaderId.toString()+") and jct_status_id=5 and jct_action_plan_soft_delete=1");
			} else if((function.trim().length() == 0) && (jobLevel.trim().length() > 0)){
				queryBldr.append(" where apview.jct_user_levels = :jLevel");
				queryBldr.append(" and apview.jct_as_header_id in ("+maxHeaderId.toString()+") and jct_status_id=5 and jct_action_plan_soft_delete=1");
			} else {
				queryBldr.append(" where apview.jct_as_header_id in ("+maxHeaderId.toString()+") and jct_status_id=5 and jct_action_plan_soft_delete=1");
			}
			queryBldr.append(" order by apview.jct_as_question_sub_desc");
			
			if((function.trim().length() > 0) && (jobLevel.trim().length() > 0)){
				returnList = sessionFactory.getCurrentSession().createSQLQuery(queryBldr.toString()).setParameter("fnGrp", function).setParameter("jLevel", jobLevel)
						.list();
			} else if((function.trim().length() > 0) && (jobLevel.trim().length() == 0)){
				returnList = sessionFactory.getCurrentSession().createSQLQuery(queryBldr.toString()).setParameter("fnGrp", function)
						.list();
			} else if((function.trim().length() == 0) && (jobLevel.trim().length() > 0)){
				returnList = sessionFactory.getCurrentSession().createSQLQuery(queryBldr.toString()).setParameter("jLevel", jobLevel)
						.list();
			} else {
				returnList = sessionFactory.getCurrentSession().createSQLQuery(queryBldr.toString())
						.list();
			}	
		} else {
			returnList = new ArrayList<Object>();
		}
		LOGGER.info("<<<< ReportDAOImpl.getActionPlanListForExcelMax");
		return returnList;
	}
	
	@Transactional(propagation=Propagation.REQUIRED)
	public List<Object> getAfterSketchListForExcel(List<String> emailIdList, String occupation, int statusId, int softDel) throws DAOException {
		LOGGER.info(">>>> ReportDAOImpl.getAfterSketchListForExcel");
		StringBuilder queryBldr = null;
		StringBuilder jrNoBuilder = new StringBuilder("");
		List<Object> returnList = null;
		List<String> jrList = null;
		queryBldr = new StringBuilder("Select distinct(asview.jct_as_created_by) from jct_after_sketch_view asview");
		if((occupation.trim().length() > 0) ){
			queryBldr.append(" where asview.jct_user_occupation = :fnGrp and jct_as_status_id="+statusId+" and jct_as_soft_delete="+softDel+"");
			jrList = sessionFactory.getCurrentSession().createSQLQuery(queryBldr.toString()).setParameter("fnGrp", occupation).list();
		} else {
			queryBldr.append(" where jct_as_status_id="+statusId+" and jct_as_soft_delete="+softDel+"");
			jrList = sessionFactory.getCurrentSession().createSQLQuery(queryBldr.toString()).list();
		}
		
		if(jrList.size() > 0){
			int size = 0;
			Iterator itr = jrList.iterator();
			while(itr.hasNext()){
				jrNoBuilder.append("'"+(String)itr.next()+"'");
				if(jrList.size()-1 > size){
					jrNoBuilder.append(",");	
				} 
				size = size + 1;
			}
			queryBldr = new StringBuilder("Select asview.jct_as_created_by, " 	
										+ "asview.jct_as_role_desc, "
										+ "asview.jct_as_task_desc, "
										+ "asview.jct_as_task_energy, "
										+ "asview.jct_as_finalpage_time_spent, "
										+ "asview.jct_as_element_desc , " 
										+ "asview.jct_user_occupation , " 
										+ "asview.jct_user_levels, "
										+ "asview.jct_first_name, "					//8
										+ "asview.jct_last_name "                   //,9
										+ "from jct_after_sketch_view asview");
			if((occupation.trim().length() > 0) ){
				queryBldr.append(" where asview.jct_user_occupation = :fnGrp");
				queryBldr.append(" and asview.jct_as_created_by in ("+jrNoBuilder.toString()+") and jct_as_status_id="+statusId+" and jct_as_soft_delete="+softDel+"");
			} else {
				queryBldr.append(" where asview.jct_as_created_by in ("+jrNoBuilder.toString()+") and jct_as_status_id="+statusId+" and jct_as_soft_delete="+softDel+"");
			}
			if (null != emailIdList) {
				queryBldr.append(" and asview.jct_as_created_by NOT IN :list");
			}
			
			queryBldr.append(" order by asview.jct_as_role_desc desc,  asview.jct_as_created_by, asview.jct_as_task_desc desc");
			
			if((occupation.trim().length() > 0)){
				if (null != emailIdList) {
					returnList = sessionFactory.getCurrentSession().
							createSQLQuery(queryBldr.toString()).setParameter("fnGrp", occupation).
							setParameterList("list", emailIdList)
							.list();
				} else {
					returnList = sessionFactory.getCurrentSession().
							createSQLQuery(queryBldr.toString()).setParameter("fnGrp", occupation)
							.list();
				}
			} else {
				if (null != emailIdList) {
					returnList = sessionFactory.getCurrentSession().
							createSQLQuery(queryBldr.toString()).
							setParameterList("list", emailIdList)
							.list();
				} else {
					returnList = sessionFactory.getCurrentSession().createSQLQuery(queryBldr.toString())
							.list();
				}
				
			}	
		} else {
			returnList = new ArrayList<Object>();
		}
		LOGGER.info("<<<< ReportDAOImpl.getAfterSketchListForExcel");
		return returnList;
	}
	
	@Transactional(propagation=Propagation.REQUIRED)
	public List<Object> getAfterSketchListForExcelMax(String function,
			String jobLevel) throws DAOException {
		LOGGER.info(">>>> ReportDAOImpl.getAfterSketchListForExcelMax");
		StringBuilder queryBldr = null;
		StringBuilder maxHeaderId = new StringBuilder("");
		List<Object> returnList = null;
		List<String> jrList = null;
		queryBldr = new StringBuilder("Select max(asview.jct_as_header_id) from jct_after_sketch_view asview");
		if((function.trim().length() > 0) && (jobLevel.trim().length() > 0)){
			queryBldr.append(" where asview.jct_user_function_group = :fnGrp and asview.jct_user_levels = :jLevel and jct_as_status_id=5 and jct_as_soft_delete=1 group by jct_as_created_by");
			jrList = sessionFactory.getCurrentSession().createSQLQuery(queryBldr.toString()).setParameter("fnGrp", function).setParameter("jLevel", jobLevel).list();
		} else if((function.trim().length() > 0) && (jobLevel.trim().length() == 0)){
			queryBldr.append(" where asview.jct_user_function_group = :fnGrp and jct_as_status_id=5 and jct_as_soft_delete=1 group by jct_as_created_by");
			jrList = sessionFactory.getCurrentSession().createSQLQuery(queryBldr.toString()).setParameter("fnGrp", function).list();
		} else if((function.trim().length() == 0) && (jobLevel.trim().length() > 0)){
			queryBldr.append(" where asview.jct_user_levels = :jLevel and jct_as_status_id=5 and jct_as_soft_delete=1 group by jct_as_created_by");
			jrList = sessionFactory.getCurrentSession().createSQLQuery(queryBldr.toString()).setParameter("jLevel", jobLevel).list();
		} else {
			queryBldr.append(" where jct_as_status_id=5 and jct_as_soft_delete=1 group by jct_as_created_by");
			jrList = sessionFactory.getCurrentSession().createSQLQuery(queryBldr.toString()).list();
		}
		if(jrList.size() > 0){int size = 0;
		Iterator itr = jrList.iterator();
		while (itr.hasNext()) {
			if (jrList.size() == 1) {
				try {
				maxHeaderId.append("'"+(BigInteger) itr.next()+"'");
				if(jrList.size()-1 > size){
					maxHeaderId.append(",");	
				} 
				size = size + 1;
				}catch(Exception e) {
					Iterator itrs = jrList.iterator();
					while (itrs.hasNext()) {
						Object[] obj = (Object[]) itrs.next();
						maxHeaderId.append("'" + (Integer) obj[0] + "'");
						if (jrList.size() - 1 > size) {
							maxHeaderId.append(",");	
						} 
						size = size + 1;
					}
				}
			} else {
				Object[] obj = (Object[]) itr.next();
				maxHeaderId.append("'" + (Integer) obj[0] + "'");
				if (jrList.size() - 1 > size) {
					maxHeaderId.append(",");	
				} 
				size = size + 1;
			}}
			queryBldr = new StringBuilder("Select asview.jct_as_created_by, " 	
										+ "asview.jct_as_role_desc, "
										+ "asview.jct_as_task_desc, "
										+ "asview.jct_as_task_energy, "
										+ "asview.jct_as_finalpage_time_spent, "
										+ "asview.jct_as_element_desc , " 
										+ "asview.jct_user_function_group , " 
										+ "asview.jct_user_levels, "
										+ "asview.jct_first_name, "					//8
										+ "asview.jct_last_name "                   //,9
										+ "from jct_after_sketch_view asview");
			if((function.trim().length() > 0) && (jobLevel.trim().length() > 0)){
				queryBldr.append(" where asview.jct_user_function_group = :fnGrp and asview.jct_user_levels = :jLevel");
				queryBldr.append(" and asview.jct_as_header_id in ("+maxHeaderId.toString()+") and jct_as_status_id=5 and jct_as_soft_delete=1");
			} else if((function.trim().length() > 0) && (jobLevel.trim().length() == 0)){
				queryBldr.append(" where asview.jct_user_function_group = :fnGrp");
				queryBldr.append(" and asview.jct_as_header_id in ("+maxHeaderId.toString()+") and jct_as_status_id=5 and jct_as_soft_delete=1");
			} else if((function.trim().length() == 0) && (jobLevel.trim().length() > 0)){
				queryBldr.append(" where asview.jct_user_levels = :jLevel");
				queryBldr.append(" and asview.jct_as_header_id in ("+maxHeaderId.toString()+") and jct_as_status_id=5 and jct_as_soft_delete=1");
			} else {
				queryBldr.append(" where asview.jct_as_header_id in ("+maxHeaderId.toString()+") and jct_as_status_id=5 and jct_as_soft_delete=1");
			}
			queryBldr.append(" order by asview.jct_as_role_desc desc,  asview.jct_as_created_by, asview.jct_as_task_desc desc");
			
			if((function.trim().length() > 0) && (jobLevel.trim().length() > 0)){
				returnList = sessionFactory.getCurrentSession().createSQLQuery(queryBldr.toString()).setParameter("fnGrp", function).setParameter("jLevel", jobLevel)
						.list();
			} else if((function.trim().length() > 0) && (jobLevel.trim().length() == 0)){
				returnList = sessionFactory.getCurrentSession().createSQLQuery(queryBldr.toString()).setParameter("fnGrp", function)
						.list();
			} else if((function.trim().length() == 0) && (jobLevel.trim().length() > 0)){
				returnList = sessionFactory.getCurrentSession().createSQLQuery(queryBldr.toString()).setParameter("jLevel", jobLevel)
						.list();
			} else {
				returnList = sessionFactory.getCurrentSession().createSQLQuery(queryBldr.toString())
						.list();
			}	
		} else {
			returnList = new ArrayList<Object>();
		}
		LOGGER.info("<<<< ReportDAOImpl.getAfterSketchListForExcelMax");
		return returnList;
	}
	/***************************** BEFORE SKETCH TO AFTER SKETCH******************************************/	
	@Transactional(propagation=Propagation.REQUIRED)
	public List<Object> getBsToAsList(String occupation, int recordIndex) throws DAOException {
		LOGGER.info(">>>> ReportDAOImpl.getBsToAsList");
		StringBuilder queryBldr = null;
		StringBuilder jrNoBuilder = new StringBuilder("");
		queryBldr = new StringBuilder("Select distinct(bsToAsview.jct_bs_to_as_created_by) from jct_before_sketch_to_after_sketch_view bsToAsview");
		List<Object> returnList = null;
		List<String> jrList = null;
		queryBldr = new StringBuilder("Select distinct(bsToAsview.jct_bs_to_as_created_by) from jct_before_sketch_to_after_sketch_view bsToAsview");
		if((occupation.trim().length() > 0)){
			queryBldr.append(" where bsToAsview.jct_user_occupation = :fnGrp");
			jrList = sessionFactory.getCurrentSession().createSQLQuery(queryBldr.toString()).setParameter("fnGrp", occupation).setFirstResult(recordIndex).setMaxResults(20).list();
		}  else {
			jrList = sessionFactory.getCurrentSession().createSQLQuery(queryBldr.toString()).setFirstResult(recordIndex).setMaxResults(20).list();
		}
		
		if(jrList.size() > 0){
			int size = 0;
			Iterator itr = jrList.iterator();
			while(itr.hasNext()){
				jrNoBuilder.append("'"+(String)itr.next()+"'");
				if(jrList.size()-1 > size){
					jrNoBuilder.append(",");	
				} 
				size = size + 1;
			}
			queryBldr = new StringBuilder("Select bsToAsview.jct_bs_to_as_created_by, bsToAsview.jct_user_occupation, bsToAsview.jct_user_levels, bsToAsview.jct_diff_status, bsToAsview.jct_bs_task_desc, bsToAsview.jct_as_task_desc, bsToAsview.jct_diff_task_desc, bsToAsview.jct_bs_energy, bsToAsview.jct_as_energy, bsToAsview.jct_diff_energy,bsToAsview.jct_first_name, bsToAsview.jct_last_name from jct_before_sketch_to_after_sketch_view bsToAsview");
			if((occupation.trim().length() > 0)){
				queryBldr.append(" where bsToAsview.jct_user_occupation = :fnGrp");
				queryBldr.append(" and bsToAsview.jct_soft_delete=0 and bsToAsview.jct_bs_to_as_created_by in ("+jrNoBuilder.toString()+")");
			} else {
				queryBldr.append(" where bsToAsview.jct_soft_delete=0 and bsToAsview.jct_bs_to_as_created_by in ("+jrNoBuilder.toString()+")");
			}
			queryBldr.append(" order by bsToAsview.jct_bs_to_as_created_by");
			
			if((occupation.trim().length() > 0)){
				returnList = sessionFactory.getCurrentSession().createSQLQuery(queryBldr.toString()).setParameter("fnGrp", occupation).list();
			}  else {
				returnList = sessionFactory.getCurrentSession().createSQLQuery(queryBldr.toString()).list();
			}	
		} else {
			returnList = new ArrayList<Object>();
		}
		LOGGER.info("<<<< ReportDAOImpl.getBsToAsList");
		return returnList;
	}

	@Transactional(propagation=Propagation.REQUIRED)
	public List<String> getTotalCountBsToAs(String function, String jobLevel) throws DAOException {
		LOGGER.info(">>>> ReportDAOImpl.getTotalCountBsToAs");
		StringBuilder jrNoBuilder = new StringBuilder("");
		List<String> jrList = null;
		jrNoBuilder = new StringBuilder("Select distinct(bsToAsview.jct_bs_to_as_created_by) from jct_before_sketch_to_after_sketch_view bsToAsview");
		if((function.trim().length() > 0) && (jobLevel.trim().length() > 0)){
			jrNoBuilder.append(" where bsToAsview.jct_user_occupation = :fnGrp ");
			jrList = sessionFactory.getCurrentSession().createSQLQuery(jrNoBuilder.toString()).setParameter("fnGrp", function).setParameter("jLevel", jobLevel).list();
		} else {
			jrList = sessionFactory.getCurrentSession().createSQLQuery(jrNoBuilder.toString()).list();
		}
		LOGGER.info("<<<< ReportDAOImpl.getTotalCountBsToAs");
		return jrList;
	
	}
	
	
	@Transactional(propagation=Propagation.REQUIRED)
	public List<Object> getBsToAsListForExcel(String occupation) throws DAOException {
		LOGGER.info(">>>> ReportDAOImpl.getBsToAsListForExcel");
		StringBuilder queryBldr = null;
		StringBuilder jrNoBuilder = new StringBuilder("");
		queryBldr = new StringBuilder("Select distinct(bsToAsview.jct_bs_to_as_created_by) from jct_before_sketch_to_after_sketch_view bsToAsview");
		List<Object> returnList = null;
		List<String> jrList = null;
		queryBldr = new StringBuilder("Select distinct(bsToAsview.jct_bs_to_as_created_by) from jct_before_sketch_to_after_sketch_view bsToAsview");
		if((occupation.trim().length() > 0) ){
			queryBldr.append(" where bsToAsview.jct_user_occupation = :fnGrp");
			jrList = sessionFactory.getCurrentSession().createSQLQuery(queryBldr.toString()).setParameter("fnGrp", occupation).list();
		} else {
			jrList = sessionFactory.getCurrentSession().createSQLQuery(queryBldr.toString()).list();
		}
		
		if(jrList.size() > 0){
			int size = 0;
			Iterator itr = jrList.iterator();
			while(itr.hasNext()){
				jrNoBuilder.append("'"+(String)itr.next()+"'");
				if(jrList.size()-1 > size){
					jrNoBuilder.append(",");	
				} 
				size = size + 1;
			}
			queryBldr = new StringBuilder("Select bsToAsview.jct_bs_to_as_created_by, bsToAsview.jct_user_occupation, bsToAsview.jct_user_levels, bsToAsview.jct_diff_status, bsToAsview.jct_bs_task_desc, bsToAsview.jct_as_task_desc, bsToAsview.jct_diff_task_desc, bsToAsview.jct_bs_energy, bsToAsview.jct_as_energy, bsToAsview.jct_diff_energy,bsToAsview.jct_first_name, bsToAsview.jct_last_name from jct_before_sketch_to_after_sketch_view bsToAsview");
			if((occupation.trim().length() > 0)){
				queryBldr.append(" where bsToAsview.jct_user_occupation = :fnGrp");
				queryBldr.append(" and bsToAsview.jct_soft_delete=0 and bsToAsview.jct_bs_to_as_created_by in ("+jrNoBuilder.toString()+")");
			} else {
				queryBldr.append(" where bsToAsview.jct_soft_delete=0 and bsToAsview.jct_bs_to_as_created_by in ("+jrNoBuilder.toString()+")");
			}
			queryBldr.append(" order by bsToAsview.jct_bs_to_as_created_by");
			
			if((occupation.trim().length() > 0) ){
				returnList = sessionFactory.getCurrentSession().createSQLQuery(queryBldr.toString()).setParameter("fnGrp", occupation).list();
			} else {
				returnList = sessionFactory.getCurrentSession().createSQLQuery(queryBldr.toString()).list();
			}	
		} else {
			returnList = new ArrayList<Object>();
		}
		LOGGER.info("<<<< ReportDAOImpl.getBsToAsListForExcel");
		return returnList;
	}	
	
	@Transactional(propagation=Propagation.REQUIRED)
	public List<Object> createMiscHeaderReport() throws DAOException {
		LOGGER.info(">>>> ReportDAOImpl.createMiscHeaderReport");
		StringBuilder qBuilder = new StringBuilder("");
		qBuilder.append("select usr.jct_user_email,  info.jct_jobref_no, "+
						"min(info.jct_start_time), max(info.jct_end_time),"+ 
						"count(info.jct_start_time),usrdetail.jct_user_details_first_name,usrdetail.jct_user_details_last_name from jct_user_login_info info, jct_user usr, jct_user_details usrdetail "+ 
						"where usr.jct_user_id=(select distinct usInfo.jct_user_id  from jct_user_login_info usInfo where usInfo.jct_jobref_no=info.jct_jobref_no) "+ 
						"AND usrdetail.jct_user_id=usr.jct_user_id AND info.jct_jobref_no NOT LIKE '00000000ADPRV%' group by info.jct_user_id;");
		LOGGER.info("<<<< ReportDAOImpl.createMiscHeaderReport");
		return sessionFactory.getCurrentSession().createSQLQuery(qBuilder.toString()).list();
	}
	
	@Transactional(propagation=Propagation.REQUIRED)
	public java.math.BigInteger getSearchableCount(String jrno) throws DAOException {
		LOGGER.info(">>>> ReportDAOImpl.getSearchableCount");
		return (java.math.BigInteger)sessionFactory.getCurrentSession()
				.createSQLQuery("Select count(*) from jct_status_search where jct_jobref_no = :jrNo")
				.setParameter("jrNo", jrno).uniqueResult();
	}
	
	@Transactional(propagation=Propagation.REQUIRED)
	public List<Object> createMiscDetailedReport(String jobRefNo) throws DAOException {
		LOGGER.info(">>>> ReportDAOImpl.createMiscDetailedReport");
		StringBuilder qBuilder = new StringBuilder("");
		qBuilder.append("select info.jct_start_time, "+
						"info.jct_end_time from jct_user_login_info info where "+ 
						"jct_jobref_no = :jobRefNo");
		LOGGER.info("<<<< ReportDAOImpl.createMiscDetailedReport");
		return sessionFactory.getCurrentSession().createSQLQuery(qBuilder.toString()).setParameter("jobRefNo", jobRefNo).list();
	}

	@Transactional(propagation=Propagation.REQUIRED)
	public List<Object> getASDetailedViewForMappings(String emailId, String elementCode, int statusId, int softDelete) throws DAOException {
		LOGGER.info(">>>> ReportDAOImpl.getASDetailedViewForMappings");
		StringBuilder qBuilder = new StringBuilder("");
		//qBuilder.append("select distinct(asview.jct_as_element_desc) from aftersketchview asview where "
		qBuilder.append("select asview.jct_as_element_desc from jct_after_sketch_view asview where "
				+ "asview.jct_as_element_code = :elementCode and asview.jct_as_created_by = :email and jct_as_status_id = "+statusId+" and jct_as_soft_delete = "+softDelete+"");
		LOGGER.info("<<<< ReportDAOImpl.getASDetailedViewForMappings");
		return sessionFactory.getCurrentSession().createSQLQuery(qBuilder.toString())
				.setParameter("elementCode", elementCode)
				.setParameter("email", emailId).list();
	}
	
	@Transactional(propagation=Propagation.REQUIRED)
	public List<Object> getGroupProfileListForExcel(int profile) throws DAOException {
		LOGGER.info(">>>> ReportDAOImpl.getGroupProfileListForExcel");
		List<Object> viewList2 = null;
		List<Object> viewList3 = null;
		List<Object> viewList4 = null;
		List<Object> viewList5 = null;
		List<Object> viewList6 = null;
		List<Object> viewList7 = null;
		List<Object> returnList = new ArrayList<Object>();
		//Map<Integer,Object> map = new HashMap<Integer,Object>();
		StringBuilder queryBldr = null;
		List<Object> profileIdList = null;	
		if( profile == 1){
			queryBldr = new StringBuilder("Select jct_user_profile " +
					"from jct_user_profile where jct_soft_delete = 0 and jct_user_profile != 1");
			profileIdList = sessionFactory.getCurrentSession().
					createSQLQuery(queryBldr.toString()).list();			
			for(int index = 0; index < profileIdList.size(); index++){
				//List<Object> returnList = new ArrayList<Object>();
				viewList2 = getInstructionDescListForExcel((Integer)profileIdList.get(index));
				viewList3 = getRefQtnDescListForExcel((Integer)profileIdList.get(index));
				viewList4 = getActionPlanDescListForExcel((Integer)profileIdList.get(index));
				viewList5 = getValueListForExcel((Integer)profileIdList.get(index));
				viewList6 = getStrengthListForExcel((Integer)profileIdList.get(index));
				viewList7 = getPassionListForExcel((Integer)profileIdList.get(index));
				
				returnList.add(viewList2);
				returnList.add(viewList3);
				returnList.add(viewList4);
				returnList.add(viewList5);
				returnList.add(viewList6);
				returnList.add(viewList7);
				//map.put((Integer)profileIdList.get(index), returnList);
			}
		} else {
			//List<Object> returnList = new ArrayList<Object>();
			//List<Object> viewList1 = getProfileDescListForExcel(profile);
			viewList2 = getInstructionDescListForExcel(profile);
			viewList3 = getRefQtnDescListForExcel(profile);
			viewList4 = getActionPlanDescListForExcel(profile);
			viewList5 = getValueListForExcel(profile);
			viewList6 = getStrengthListForExcel(profile);
			viewList7 = getPassionListForExcel(profile);
			
			//returnList.add(viewList1);
			returnList.add(viewList2);
			returnList.add(viewList3);
			returnList.add(viewList4);
			returnList.add(viewList5);
			returnList.add(viewList6);
			returnList.add(viewList7);
		}
		LOGGER.info("<<<< ReportDAOImpl.getGroupProfileListForExcel");
		return returnList;
	}
	
	@Transactional(propagation=Propagation.REQUIRED)
	public List<Object> getProfileDescListForExcel(int profile) throws DAOException {	
		LOGGER.info(">>>> ReportDAOImpl.getProfileDescListForExcel");
		StringBuilder queryBldr1 = null;
		List<Object> profileIdList = null;	
		
		StringBuilder queryBldr = null;
		List<Object> returnList = null;
		List<Object> finalList = new ArrayList<Object>();	
		if ( profile == 1) {
			queryBldr1 = new StringBuilder("Select jct_user_profile " +
					"from jct_user_profile where jct_soft_delete = 0 and jct_user_profile != 1");
			profileIdList = sessionFactory.getCurrentSession().
					createSQLQuery(queryBldr1.toString()).list();	
			for(int index = 0; index < profileIdList.size(); index++){
				queryBldr = new StringBuilder("Select userProfile.jct_user_profile_desc, " 	
						+ "userProfile.jct_created_ts, "
						+ "count(distinct(userGroup.jct_user_group_desc)) "
						+ "from jct_user_profile userProfile,jct_user_group userGroup " 
						+ "where userProfile.jct_user_profile = "+(Integer)profileIdList.get(index)+" and userProfile.jct_soft_delete = 0 and userProfile.jct_user_profile != 1 "
						+ "and userGroup.jct_user_profile = "+(Integer)profileIdList.get(index)+" and userGroup.jct_soft_delete = 0 ");											
				
				returnList = sessionFactory.getCurrentSession().createSQLQuery(queryBldr.toString()).list();
				finalList.add(returnList);
			}
		
		} else {
			queryBldr = new StringBuilder("Select userProfile.jct_user_profile_desc, " 	
					+ "userProfile.jct_created_ts, "
					+ "count(distinct(userGroup.jct_user_group_desc)) "
					+ "from jct_user_profile userProfile,jct_user_group userGroup " 
					+ "where userProfile.jct_user_profile= "+profile+" and userProfile.jct_soft_delete = 0 "
					+ "and userGroup.jct_user_profile= "+profile+" and userGroup.jct_soft_delete = 0 ");
		
			returnList = sessionFactory.getCurrentSession().
				createSQLQuery(queryBldr.toString()).list();	
			finalList.add(returnList);
		}	
		LOGGER.info("<<<< ReportDAOImpl.getProfileDescListForExcel");
		return finalList;
	}
	
	private List<Object> getInstructionDescListForExcel(int profile) throws DAOException {
		LOGGER.info(">>>> ReportDAOImpl.getInstructionDescListForExcel");
		StringBuilder queryBldr1 = null;
		StringBuilder queryBldr2 = null;
		StringBuilder queryBldr3 = null;
		List<Object> returnList = new ArrayList<Object>();
		List<Object> returnList1 = null;
		List<Object> returnList2 = null;
		List<Object> returnList3 = null;
		queryBldr1 = new StringBuilder("Select instruction.jct_instruction_bar_desc " +
				"from jct_instruction_bar instruction " +
				"where instruction.jct_profile_id = " + profile + " " +
				"and instruction.jct_instruction_bar_soft_delete = 0 " +
				"and instruction.jct_page_details IN ('Creating Before Sketch')");
		
		returnList1 = sessionFactory.getCurrentSession().
				createSQLQuery(queryBldr1.toString()).list();	
		
		
		queryBldr2 = new StringBuilder("Select instruction.jct_instruction_bar_desc " +
				"from jct_instruction_bar instruction " +
				"where instruction.jct_profile_id= " + profile + " " +
				"and instruction.jct_instruction_bar_soft_delete= 0 " +
				"and instruction.jct_page_details IN ('Mapping Yourself')");
		
		returnList2 = sessionFactory.getCurrentSession().
				createSQLQuery(queryBldr2.toString()).list();	
		
		queryBldr3 = new StringBuilder("Select instruction.jct_instruction_bar_desc " +
				"from jct_instruction_bar instruction " +
				"where instruction.jct_profile_id = " + profile + " " +
				"and instruction.jct_instruction_bar_soft_delete = 0 " +
				"and instruction.jct_page_details IN ('Creating After Diagram')");
		
		returnList3 = sessionFactory.getCurrentSession().
				createSQLQuery(queryBldr3.toString()).list();	
		
		returnList.add(returnList1);
		returnList.add(returnList2);
		returnList.add(returnList3);
		LOGGER.info("<<<< ReportDAOImpl.getInstructionDescListForExcel");
		return returnList;
	}
	
	private List<Object> getRefQtnDescListForExcel(int profile) throws DAOException {
		LOGGER.info(">>>> ReportDAOImpl.getRefQtnDescListForExcel");
		StringBuilder queryBldr = null;
		List<Object> returnList = null;	
		queryBldr = new StringBuilder("select question.jct_questions_desc, jct_question_sub_desc " +
				"from jct_questions question " +
				"where question.jct_profile_id = " + profile + " " +
				"and question.jct_question_bsas = 'BS' " +
				"and question.jct_questions_soft_delete = 0");
		returnList = sessionFactory.getCurrentSession().
				createSQLQuery(queryBldr.toString()).list();
		LOGGER.info("<<<< ReportDAOImpl.getRefQtnDescListForExcel");
		return returnList;
	}
	
	private List<Object> getActionPlanDescListForExcel(int profile) throws DAOException {
		LOGGER.info(">>>> ReportDAOImpl.getActionPlanDescListForExcel");
		StringBuilder queryBldr = null;
		List<Object> returnList = null;	
		queryBldr = new StringBuilder("select question.jct_questions_desc, jct_question_sub_desc " +
				"from jct_questions question " +
				"where question.jct_profile_id = " + profile + "" +
				" and question.jct_question_bsas= 'AS' " +
				"and question.jct_questions_soft_delete = 0");
		returnList = sessionFactory.getCurrentSession().
				createSQLQuery(queryBldr.toString()).list();
		LOGGER.info("<<<< ReportDAOImpl.getActionPlanDescListForExcel");
		return returnList;
	}
	
	private List<Object> getValueListForExcel(int profile) throws DAOException {
		LOGGER.info(">>>> ReportDAOImpl.getValueListForExcel");
		StringBuilder queryBldr = null;
		List<Object> returnList = null;	
		queryBldr = new StringBuilder("select jct_job_attribute_name, jct_job_attribute_desc " +
				"from jct_job_attribute " +
				"where jct_job_attribute_code='VAL' " +
				"and jct_profile_id = " + profile + " " +
				"and jct_job_attribute_soft_delete = 0");
		returnList = sessionFactory.getCurrentSession().
				createSQLQuery(queryBldr.toString()).list();
		LOGGER.info("<<<< ReportDAOImpl.getValueListForExcel");
		return returnList;
	}
	
	private List<Object> getStrengthListForExcel(int profile) throws DAOException {
		LOGGER.info(">>>> ReportDAOImpl.getStrengthListForExcel");
		StringBuilder queryBldr = null;
		List<Object> returnList = null;	
		queryBldr = new StringBuilder("select jct_job_attribute_name, jct_job_attribute_desc " +
				"from jct_job_attribute " +
				"where jct_job_attribute_code='STR' " +
				"and jct_profile_id = " + profile + " " +
				"and jct_job_attribute_soft_delete = 0");
		returnList = sessionFactory.getCurrentSession().
				createSQLQuery(queryBldr.toString()).list();
		LOGGER.info("<<<< ReportDAOImpl.getStrengthListForExcel");
		return returnList;
	}
	
	private List<Object> getPassionListForExcel(int profile) throws DAOException {
		LOGGER.info(">>>> ReportDAOImpl.getPassionListForExcel");
		StringBuilder queryBldr = null;
		List<Object> returnList = null;	
		queryBldr = new StringBuilder("select jct_job_attribute_name, jct_job_attribute_desc " +
				"from jct_job_attribute " +
				"where jct_job_attribute_code='PAS' " +
				"and jct_profile_id = " + profile + " " +
				"and jct_job_attribute_soft_delete = 0");
		returnList = sessionFactory.getCurrentSession().
				createSQLQuery(queryBldr.toString()).list();
		LOGGER.info("<<<< ReportDAOImpl.getPassionListForExcel");
		return returnList;
	}
	
	@Transactional(propagation=Propagation.REQUIRED)
	public List<Object> getLoginInfoDetails(String emailId) throws DAOException {
		  LOGGER.info(">>>> ReportDAOImpl.getLoginInfoDetails");
		  StringBuilder queryBldr = null;
		  List<Object> returnList = null; 
		  if (emailId == null) {
		   queryBldr = new StringBuilder("select user.jct_user_email, dtls.jct_user_details_first_name, dtls.jct_user_details_last_name, "
		    + "dtls.jct_user_details_function_group, dtls.jct_user_details_levels, dtls.jct_user_details_group_id  "
		    + "from jct_user user, jct_user_details dtls where user.jct_user_id in (select distinct(jct_user_id) "
		    + "from jct_user_login_info) and user.jct_user_id=dtls.jct_user_id and user.jct_active_yn = 3 and "
		    + "user.jct_user_soft_delete = 0 and user.jct_role_id=1 and dtls.jct_user_onet_occupation NOT IN ('NOT-ACTIVATED')");
		   returnList = sessionFactory.getCurrentSession().
		     createSQLQuery(queryBldr.toString()).list();
		  } else {
		   queryBldr = new StringBuilder("select dtls.jct_user_details_function_group, dtls.jct_user_details_levels, dtls.jct_user_details_group_id  "
		     + "from jct_user user, jct_user_details dtls where user.jct_user_id in (select distinct(jct_user_id) "
		     + "from jct_user_login_info) and user.jct_user_id=dtls.jct_user_id and user.jct_active_yn = 3 and user.jct_user_soft_delete = 0 and "
		     + "user.jct_user_email = :emailId and user.jct_role_id=1 and dtls.jct_user_onet_occupation NOT IN ('NOT-ACTIVATED')");
		   returnList = sessionFactory.getCurrentSession().
		     createSQLQuery(queryBldr.toString())
		     .setParameter("emailId", emailId)
		     .list();
		  }
		  LOGGER.info("<<<< ReportDAOImpl.getLoginInfoDetails");
		  return returnList;
		 }
	@Transactional(propagation=Propagation.REQUIRED)
	public java.sql.Timestamp getGroupCreationDate(Integer groupName) throws DAOException {
		LOGGER.info(">>>> ReportDAOImpl.getGroupCreationDate");
		return (java.sql.Timestamp) sessionFactory.getCurrentSession().
				createSQLQuery("select jct_created_ts from jct_user_group where jct_user_group = :groupName")
				.setParameter("groupName", groupName).uniqueResult();
	}

	@Transactional(propagation=Propagation.REQUIRED)
	public List<Object> getLoggedIndetails(String email) throws DAOException {
		LOGGER.info(">>>> ReportDAOImpl.getLoggedIndetails");
		StringBuilder queryBldr = new StringBuilder("select min(info.jct_start_time), "
				+ "max(info.jct_end_time), count(info.jct_user_id), info.jct_user_id from jct_user_login_info info, "
				+ "jct_user usr where usr.jct_user_id = info.jct_user_id and usr.jct_user_email = :emailId");
		LOGGER.info("<<<< ReportDAOImpl.getLoggedIndetails");
		return sessionFactory.getCurrentSession().
				createSQLQuery(queryBldr.toString())
				.setParameter("emailId", email)
				.list();
	}

	@Transactional(propagation=Propagation.REQUIRED)
	public Integer getTotalTimeSpentOnBs(String emailId, int statusId, int softDel) throws DAOException {
		LOGGER.info(">>>> ReportDAOImpl.getTotalTimeSpentOnBs");
		Integer result = null;
		try {
			result = (Integer) sessionFactory.getCurrentSession().createSQLQuery(
					"select max(jct_bs_time_spent) from jct_before_sketch_header where jct_bs_created_by = :emailId and jct_bs_soft_delete = :sDel and jct_bs_status_id = :stId")
					.setParameter("emailId", emailId)
					.setParameter("sDel", softDel)
					.setParameter("stId", statusId).uniqueResult();
		} catch (Exception e) {
			result = new Integer(0);
		}
		return result;
	}

	@Transactional(propagation=Propagation.REQUIRED)
	public BigInteger getTotalBsTasks(String emailId) throws DAOException {
		LOGGER.info(">>>> ReportDAOImpl.getTotalBsTasks");
		BigInteger totalBsTasks = new BigInteger("0");
		try {
			totalBsTasks = (BigInteger) sessionFactory.getCurrentSession().createSQLQuery(
					"select count(distinct(jct_bs_task_desc)) from jct_before_sketch_details where jct_bs_created_by = :emailId "+ 
					"and jct_bs_soft_delete = 0")
					.setParameter("emailId", emailId).uniqueResult();
		} catch (Exception ex) {
			
		}
		return totalBsTasks;
	}

	@Transactional(propagation=Propagation.REQUIRED)
	public List<Object> getBSDescriptionAndTimeEnergy(String emailId)
			throws DAOException {
		LOGGER.info(">>>> ReportDAOImpl.getBSDescriptionAndTimeEnergy");
		List<Object> returnList = null;
		StringBuilder queryBldr = new StringBuilder("Select distinct(jct_bs_task_desc), jct_bs_energy from jct_before_sketch_details");
		queryBldr.append(" where jct_bs_soft_delete = 0 and jct_bs_created_by = :emailId");
		returnList = sessionFactory.getCurrentSession().createSQLQuery(queryBldr.toString())
				.setParameter("emailId", emailId).list();
		LOGGER.info("<<<< ReportDAOImpl.getBSDescriptionAndTimeEnergy");
		return returnList;
	}
	
	@Transactional(propagation=Propagation.REQUIRED)
	public Integer getTotalTimeSpentOnAs(String emailId, int statusId, int softDel) throws DAOException {
		LOGGER.info(">>>> ReportDAOImpl.getTotalTimeSpentOnAs");
		Integer timeSpent = null;
		try {
			timeSpent = (Integer) sessionFactory.getCurrentSession().createSQLQuery(
					"select max(jct_as_finalpage_time_spent) from jct_after_sketch_header where jct_as_created_by = :emailId and "
							+ "jct_as_soft_delete = :softDel and jct_as_status_id = :stId")
							.setParameter("emailId", emailId)
							.setParameter("softDel", softDel)
							.setParameter("stId", statusId).uniqueResult();
		} catch (Exception ex) {
			timeSpent = new Integer(0);
		}
		return timeSpent;
	}
	
	@Transactional(propagation=Propagation.REQUIRED)
	public List<Object> getTotalAsTasks(String emailId) throws DAOException {
		LOGGER.info(">>>> ReportDAOImpl.getTotalAsTasks");
		return  sessionFactory.getCurrentSession().createSQLQuery(
				"select count(distinct(jct_as_task_desc)), count(distinct(jct_as_role_desc)) from jct_after_sketch_finalpage_details where jct_as_created_by = :emailId "+ 
				"and jct_as_soft_delete = 0 and jct_as_element_code NOT IN ('STR', 'VAL', 'PAS')")
				.setParameter("emailId", emailId).list();
	}

	@Transactional(propagation=Propagation.REQUIRED)
	public List<Object> getRoleFrame(String emailId) throws DAOException {
		LOGGER.info(">>>> ReportDAOImpl.getRoleFrame");
		return  sessionFactory.getCurrentSession().createSQLQuery(
				"select distinct(jct_as_role_desc) from jct_after_sketch_finalpage_details where jct_as_created_by = :emailId "+ 
				"and jct_as_soft_delete = 0")
				.setParameter("emailId", emailId).list();
	}

	@Transactional(propagation=Propagation.REQUIRED)
	public List<Object> getASDescriptionMainPersonAndTimeEnergy(String emailId)
			throws DAOException {
		LOGGER.info(">>>> ReportDAOImpl.getASDescriptionMainPersonAndTimeEnergy");
		/*return  sessionFactory.getCurrentSession().createSQLQuery(
				"select distinct(jct_as_task_desc), jct_as_additional_desc, jct_as_task_energy, jct_as_role_desc "
				+ "from jct_after_sketch_finalpage_details where jct_as_created_by = :emailId "+ 
				"and jct_as_soft_delete = 0 and jct_as_element_code NOT IN ('STR', 'VAL', 'PAS')")
				.setParameter("emailId", emailId).list();*/
		return  sessionFactory.getCurrentSession().createSQLQuery(
				"select distinct(jct_as_task_desc), jct_as_additional_desc, jct_as_task_energy "
				+ "from jct_after_sketch_finalpage_details where jct_as_created_by = :emailId "+ 
				"and jct_as_soft_delete = 0 and jct_as_element_code NOT IN ('STR', 'VAL', 'PAS')")
				.setParameter("emailId", emailId).list();
	}

	@Transactional(propagation=Propagation.REQUIRED)
	public List<Object> getBsToAs(String emailId) throws DAOException {
		LOGGER.info(">>>> ReportDAOImpl.getBsToAs");
		return  sessionFactory.getCurrentSession().createSQLQuery(
				"select jct_diff_energy, jct_diff_status, jct_diff_task_desc from jct_bs_to_as where jct_bs_to_as_created_by = :emailId and jct_soft_delete = 0")
				.setParameter("emailId", emailId).list();
	}

	@Transactional(propagation=Propagation.REQUIRED)
	public List<Object> getStrValPassItems(String emailId) throws DAOException {
		LOGGER.info(">>>> ReportDAOImpl.getStrValPassItems");
		return  sessionFactory.getCurrentSession().createSQLQuery(
				"select jct_as_element_code, jct_as_element_desc, jct_as_role_desc, jct_as_position from jct_after_sketch_finalpage_details where "+
				"jct_as_element_code in ('STR', 'VAL', 'PAS') "+
				" and jct_as_soft_delete = 0 and jct_as_created_by = :emailId order by jct_as_element_code,jct_as_element_desc, jct_as_role_desc, jct_as_position")
				.setParameter("emailId", emailId).list();
	}

	@Transactional(propagation=Propagation.REQUIRED)
	public List<Object> getReflectionQuestions(String emailId, int statusId, int softDel, String diff)
			throws DAOException {
		LOGGER.info(">>>> ReportDAOImpl.getReflectionQuestions");
		List<Object> objs = null;
		//if (diff.equals("N")) {
			objs = sessionFactory.getCurrentSession().createSQLQuery(
					"select distinct (jct_bs_question_desc) from jct_before_sketch_question where "+
					"jct_bs_created_by = :emailId and jct_bs_soft_delete = :softDel and jct_bs_status_id = :statusId and jct_bs_header_id = "
					+ "((select max(jct_bs_header_id) from jct_before_sketch_question where jct_bs_status_id = "+statusId+" and jct_bs_soft_delete="+softDel+" and jct_bs_created_by = :email2))")
					.setParameter("emailId", emailId)
					.setParameter("softDel", softDel)
					.setParameter("statusId", statusId)
					.setParameter("email2", emailId).list();
		/*} else {
			objs = sessionFactory.getCurrentSession().createSQLQuery(
					"select distinct (jct_bs_question_desc) from jct_before_sketch_question where "+
					"jct_bs_created_by = :emailId and jct_bs_soft_delete = :softDel and jct_bs_status_id = :statusId and jct_bs_header_id = "
					+ "(select max(jct_bs_header_id) from jct_before_sketch_question where jct_bs_status_id = 5 and jct_bs_soft_delete=1 and jct_bs_created_by = :email2)")
					.setParameter("emailId", emailId)
					.setParameter("softDel", softDel)
					.setParameter("statusId", statusId)
					.setParameter("email2", emailId).list();
		}*/
		return objs;
	}

	@Transactional(propagation=Propagation.REQUIRED)
	public List<Object> populateActionPlan(String emailId, int statusId, int softDel) throws DAOException {
		LOGGER.info(">>>> ReportDAOImpl.populateActionPlan");
		List<Object> objs = null;
		//if (diff.equals("N")) {
			objs = sessionFactory.getCurrentSession().createSQLQuery(
					"select distinct(jct_as_question_desc) from jct_action_plan where "+
					"jct_action_plan_soft_delete = "+softDel+" and jct_status_id = "+statusId+" and jct_as_created_by = :emailId and jct_as_header_id = "
					+ "(select max(jct_as_header_id) from jct_action_plan where jct_action_plan_soft_delete = "+softDel+" and jct_status_id = "+statusId+" and jct_as_created_by = :emailId2)")
					.setParameter("emailId", emailId)
					.setParameter("emailId2", emailId).list();
		/*} else {
			objs = sessionFactory.getCurrentSession().createSQLQuery(
					"select distinct(jct_as_question_desc) from jct_action_plan where "+
							"jct_action_plan_soft_delete = 1 and jct_status_id = 5 and jct_as_created_by = :emailId and jct_as_header_id = "
							+ "(select max(jct_as_header_id) from jct_action_plan where jct_action_plan_soft_delete = 1 and jct_status_id = 5 and jct_as_created_by = :emailId2)")
							.setParameter("emailId", emailId)
							.setParameter("emailId2", emailId).list();
		}*/
		return objs;
	}

	@Transactional(propagation=Propagation.REQUIRED)
	public List<Object> populateSubQtnActionPlan(String mainQuestion,
			String emailId, String diff, int statusId, int softDel) throws DAOException {
		LOGGER.info(">>>> ReportDAOImpl.populateSubQtnActionPlan");
		List<Object> objs = null;
		//if (diff.equals("N")) {
			objs = sessionFactory.getCurrentSession().createSQLQuery(
					"select distinct(jct_as_question_sub_desc) from jct_action_plan where "+
							"jct_action_plan_soft_delete = "+softDel+" and jct_status_id = "+statusId+" and jct_as_created_by = :emailId and "+
							"jct_as_question_desc = :mainQuestion and jct_as_header_id = "
							+ "(select max(jct_as_header_id) from jct_action_plan where jct_action_plan_soft_delete = "+softDel+" and jct_status_id = "+statusId+" and jct_as_created_by = :emailId2)")
							.setParameter("emailId", emailId)
							.setParameter("mainQuestion", mainQuestion)
							.setParameter("emailId2", emailId).list();
		/*} else {
			objs = sessionFactory.getCurrentSession().createSQLQuery(
					"select distinct(jct_as_question_sub_desc) from jct_action_plan where "+
							"jct_action_plan_soft_delete = 1 and jct_status_id = 5 and jct_as_created_by = :emailId and "+
							"jct_as_question_desc = :mainQuestion and jct_as_header_id = "
							+ "(select max(jct_as_header_id) from jct_action_plan where jct_action_plan_soft_delete = 1 and jct_status_id = 5 and jct_as_created_by = :emailId2)")
							.setParameter("emailId", emailId)
							.setParameter("mainQuestion", mainQuestion)
							.setParameter("emailId2", emailId).list();
		}*/
		return objs;
	}

	@Transactional(propagation=Propagation.REQUIRED)
	public List<Object> getActionPlanAnswers(String mainQuestion,
			String subQuestion, String emailId, String diff, int statusId, int softDel) throws DAOException {
		LOGGER.info(">>>> ReportDAOImpl.getActionPlanAnswers");
		List<Object> objs = null;
		//if (diff.equals("N")) {
			objs = sessionFactory.getCurrentSession().createSQLQuery(
					"select jct_as_answar_desc from jct_action_plan where jct_action_plan_soft_delete = "+softDel+" and " +
					"jct_status_id = "+statusId+" and jct_as_question_desc = :mainQtn and jct_as_question_sub_desc = :subQtn and jct_as_created_by = :emailId and jct_as_header_id = "
					+ "(select max(jct_as_header_id) from jct_action_plan where jct_action_plan_soft_delete = "+softDel+" and jct_status_id = "+statusId+" and jct_as_created_by = :emailId2)")
					.setParameter("mainQtn", mainQuestion)
					.setParameter("subQtn", subQuestion)
					.setParameter("emailId", emailId)
					.setParameter("emailId2", emailId).list();
		/*} else {
			objs = sessionFactory.getCurrentSession().createSQLQuery(
					"select jct_as_answar_desc from jct_action_plan where jct_action_plan_soft_delete = 0 and " +
							"jct_status_id = 5 and jct_as_question_desc = :mainQtn and jct_as_question_sub_desc = :subQtn and jct_as_created_by = :emailId and jct_as_header_id = "
							+ "(select max(jct_as_header_id) from jct_action_plan where jct_action_plan_soft_delete = 1 and jct_status_id = 5 and jct_as_created_by = :emailId2)")
							.setParameter("mainQtn", mainQuestion)
							.setParameter("subQtn", subQuestion)
							.setParameter("emailId", emailId)
							.setParameter("emailId2", emailId).list();
		}*/
		return objs;
	}

	@Transactional(propagation=Propagation.REQUIRED)
	public List<Object> populateASTaskLocation(String emailId)
			throws DAOException {
		LOGGER.info(">>>> ReportDAOImpl.populateASTaskLocation");
		return  sessionFactory.getCurrentSession().createSQLQuery(
				"select jct_as_position from jct_after_sketch_finalpage_details where "+
				"jct_as_element_code not in ('STR', 'VAL', 'PAS') and "+
				"jct_as_created_by = :emailId and jct_as_soft_delete = 0")
				.setParameter("emailId", emailId).list();
	}
	
	@Transactional(propagation=Propagation.REQUIRED)
	public List<Object> populateSubQtnQuestionnaire(String mainQuestion,
			String emailId, String diff, int statusId, int softDel) throws DAOException {
		List<Object> objs = null;
		objs = sessionFactory.getCurrentSession().createSQLQuery(
				"select jct_bs_sub_question from jct_before_sketch_question where "+
						"jct_bs_soft_delete = "+softDel+" and jct_bs_status_id = "+statusId+" and jct_bs_created_by = :emailId and "+
						"jct_bs_question_desc = :mainQuestion and jct_bs_header_id = "
						+ "(select max(jct_bs_header_id) from jct_before_sketch_question where jct_bs_status_id = "+statusId+" and jct_bs_soft_delete="+softDel+" and jct_bs_created_by = :email2)")
						.setParameter("emailId", emailId)
						.setParameter("mainQuestion", mainQuestion)
						.setParameter("email2", emailId).list();
		LOGGER.info(">>>> ReportDAOImpl.populateSubQtnQuestionnaire");
		return objs;
	}
	
	@Transactional(propagation=Propagation.REQUIRED)
	public List<Object> getQuestionnaireAnswers(String mainQuestion,
			String subQuestion, String emailId, String diff, int statusId, int softDel) throws DAOException {
		List<Object> objs = null;
			objs = sessionFactory.getCurrentSession().createSQLQuery(
					"select jct_bs_answar_desc from jct_before_sketch_question where jct_bs_soft_delete = "+softDel+" and jct_bs_status_id = "+statusId+" and " +
					"jct_bs_question_desc = :mainQtn and jct_bs_sub_question = :subQtn and jct_bs_created_by = :emailId and jct_bs_header_id = "
					+ "(select max(jct_bs_header_id) from jct_before_sketch_question where jct_bs_status_id = "+statusId+" and jct_bs_soft_delete="+softDel+" and jct_bs_created_by = :email2)")
					.setParameter("mainQtn", mainQuestion)
					.setParameter("subQtn", subQuestion)
					.setParameter("emailId", emailId)
					.setParameter("email2", emailId).list();
		LOGGER.info(">>>> ReportDAOImpl.getQuestionnaireAnswers");
		return objs;
	}

	@Transactional(propagation=Propagation.REQUIRED)
	public List<String> getRolesForTasks(String emailId, String taskDesc,
			int statusId, int softDel) throws DAOException {
		List<String> objs = null;
		objs = sessionFactory.getCurrentSession().createSQLQuery(
			"select jct_as_role_desc from jct_after_sketch_finalpage_details where "+
					"jct_as_soft_delete = "+softDel+" and jct_as_status_id = "+statusId+" and jct_as_task_desc = :taskDesc and "+
						"jct_as_created_by = :emailId")
						.setParameter("taskDesc", taskDesc)
						.setParameter("emailId", emailId).list();
		return objs;
	}
	
	@Transactional(propagation=Propagation.REQUIRED)
	public List<Object> createSurveyReport(int usrType) throws DAOException {
		LOGGER.info(">>>> ReportDAOImpl.createSurveyReport");
		StringBuilder qBuilder = new StringBuilder("");
		if(usrType==5){
		qBuilder.append("select jct_user.jct_user_name,jct_user.jct_role_id," +
						"jct_user.jct_created_ts,jct_user.jct_account_expiration_date," +
						"jct_survey_questions.jct_survey_main_qtn,jct_survey_questions.jct_survey_answer " +
						"from jct_user,jct_survey_questions " +
						"where jct_user.jct_user_name = jct_survey_questions.jct_survey_taken_by " +
						"and jct_survey_taken_by_user_id = jct_user_id " +
						"and jct_survey_soft_delete ='0'"+
						"and (jct_role_id = '1' or jct_role_id = '3');");
		}else{
		qBuilder.append("select jct_user.jct_user_name,jct_user.jct_role_id," +
						"jct_user.jct_created_ts,jct_user.jct_account_expiration_date," +
						"jct_survey_questions.jct_survey_main_qtn,jct_survey_questions.jct_survey_answer " +
						"from jct_user,jct_survey_questions " +
						"where jct_user.jct_user_name=jct_survey_questions.jct_survey_taken_by " +
						"and jct_survey_taken_by_user_id = jct_user_id " +
						"and jct_survey_soft_delete ='0'"+
						"and jct_role_id = "+usrType+"; ");
		}
		LOGGER.info("<<<< ReportDAOImpl.createSurveyReport");		
		return sessionFactory.getCurrentSession().createSQLQuery(qBuilder.toString()).list();
	}
	
	public Object getSurveyListCommonFieldsForExcel(int answeredUserId) throws DAOException {
		LOGGER.info(">>>> ReportDAOImpl.getSurveyListCommonFieldsForExcel");
		StringBuilder qBuilder = new StringBuilder("");		
		qBuilder.append("Select u.jct_user_name,udt.jct_user_details_profile_name,udt.jct_user_details_group_name,u.jct_user_customer_id,u.jct_role_id,u.jct_account_expiration_date,udt.jct_user_details_first_name,udt.jct_user_details_last_name " +
				"from jct_user u, jct_user_details udt " +
				"where u.jct_user_id = "+answeredUserId+" AND udt.jct_user_id = "+answeredUserId);
		LOGGER.info("<<<< ReportDAOImpl.getSurveyListCommonFieldsForExcel");
		return sessionFactory.getCurrentSession().createSQLQuery(qBuilder.toString()).uniqueResult();
	}
	@Transactional(propagation=Propagation.REQUIRED)
	public List<Object> getSurveyListRedundentFields(int answeredUserId) throws DAOException {
		LOGGER.info(">>>> ReportDAOImpl.getSurveyListRedundentFields");
		StringBuilder qBuilder = new StringBuilder("");	
		qBuilder.append("Select jct_survey_main_qtn, jct_survey_answer_type, jct_survey_answer from jct_survey_questions "+
					"where jct_survey_taken_by_user_id = "+answeredUserId+" AND jct_survey_soft_delete = 0 " +
					"AND jct_survey_answer != 'NOTANSWERED'");
		return sessionFactory.getCurrentSession().createSQLQuery(qBuilder.toString()).list(); 
	}

/*	@Transactional(propagation=Propagation.REQUIRED)
	public List<Object> createUserPaymentReport(int userType) throws DAOException {
		LOGGER.info(">>>>>> ReportDAOImpl.createUserPaymentReport");
		StringBuilder qBuilder = new StringBuilder("");
		List<Object> obj = null;
		if(userType == 1) {
			qBuilder.append("select u.jct_user_name,ud.jct_user_details_profile_name,ud.jct_user_details_group_name,u.jct_user_customer_id,u.jct_role_id," +
					"u.jct_created_by,h.jct_pmt_hdr_total_amt,d.jct_pmt_dtls_pmt_typ_desc,d.jct_pmt_dtls_modified_ts,d.jct_pmt_dtls_cheque_nos,d.jct_pmt_dtls_pmt_trans_nos,ud.jct_user_details_first_name,ud.jct_user_details_last_name " + 
					"from jct_user u,jct_user_details ud,jct_payment_header h,jct_payment_details d " +
					"where u.jct_user_id = ud.jct_user_id and u.jct_user_customer_id = h.jct_user_customer_id and h.jct_pmt_hdr_id = d.jct_pmt_hdr_id "+
					"and u.jct_role_id = "+userType+" and u.jct_user_soft_delete = 0 and jct_user_disabled = 0 and u.jct_user_customer_id like('98%')" +
					"ORDER BY u.jct_user_name");
			obj = sessionFactory.getCurrentSession().createSQLQuery(qBuilder.toString()).list();
		} else if (userType == 3) {
			qBuilder.append("select u.jct_user_name,ud.jct_user_details_profile_name,ud.jct_user_details_group_name,u.jct_user_customer_id,u.jct_role_id," +
					"u.jct_created_by,h.jct_pmt_hdr_total_amt,d.jct_pmt_dtls_pmt_typ_desc,d.jct_pmt_dtls_modified_ts,d.jct_pmt_dtls_cheque_nos,d.jct_pmt_dtls_pmt_trans_nos,fd.jct_fac_type,fd.jct_pmt_hdr_id,ud.jct_user_details_first_name,ud.jct_user_details_last_name " + 
					"from jct_user u,jct_user_details ud,jct_payment_header h,jct_payment_details d,jct_facilitator_details fd " +
					"where u.jct_user_id = ud.jct_user_id and u.jct_user_customer_id = h.jct_user_customer_id and h.jct_pmt_hdr_id = d.jct_pmt_hdr_id " +
					"and fd.jct_pmt_hdr_id = h.jct_pmt_hdr_id and u.jct_role_id = "+ userType +
					" and u.jct_active_yn != 6 and u.jct_user_soft_delete = 0 and jct_user_disabled = 0 " +
					"ORDER BY u.jct_user_name");
			obj = sessionFactory.getCurrentSession().createSQLQuery(qBuilder.toString()).list();
		} else if (userType == 0) {	//	All
			qBuilder.append("select u.jct_user_name,ud.jct_user_details_profile_name,ud.jct_user_details_group_name,u.jct_user_customer_id,u.jct_role_id," +
					"u.jct_created_by,h.jct_pmt_hdr_total_amt,d.jct_pmt_dtls_pmt_typ_desc,d.jct_pmt_dtls_modified_ts,d.jct_pmt_dtls_cheque_nos,d.jct_pmt_dtls_pmt_trans_nos,ud.jct_user_details_first_name,ud.jct_user_details_last_name" + 
					"from jct_user u,jct_user_details ud,jct_payment_header h,jct_payment_details d " +
					"where u.jct_user_id = ud.jct_user_id and u.jct_user_customer_id = h.jct_user_customer_id and h.jct_pmt_hdr_id = d.jct_pmt_hdr_id "+
					"and u.jct_role_id = 1 and u.jct_user_soft_delete = 0 and jct_user_disabled = 0 and u.jct_user_customer_id like('98%')"+
					"ORDER BY u.jct_user_name");
			obj = sessionFactory.getCurrentSession().createSQLQuery(qBuilder.toString()).list();
			qBuilder.setLength(0);	//	reset qBuilder		
			qBuilder.append("select u.jct_user_name,ud.jct_user_details_profile_name,ud.jct_user_details_group_name,u.jct_user_customer_id,u.jct_role_id," +
					"u.jct_created_by,h.jct_pmt_hdr_total_amt,d.jct_pmt_dtls_pmt_typ_desc,d.jct_pmt_dtls_modified_ts,d.jct_pmt_dtls_cheque_nos,d.jct_pmt_dtls_pmt_trans_nos,fd.jct_fac_type,fd.jct_pmt_hdr_id,ud.jct_user_details_first_name,ud.jct_user_details_last_name " + 
					"from jct_user u,jct_user_details ud,jct_payment_header h,jct_payment_details d,jct_facilitator_details fd " +
					"where u.jct_user_id = ud.jct_user_id and u.jct_user_customer_id = h.jct_user_customer_id and h.jct_pmt_hdr_id = d.jct_pmt_hdr_id " +
					"and fd.jct_pmt_hdr_id = h.jct_pmt_hdr_id and u.jct_role_id = 3 " +
					"and u.jct_active_yn != 6 and u.jct_user_soft_delete = 0 and jct_user_disabled = 0 "+
					"ORDER BY u.jct_user_name");
			obj.addAll(sessionFactory.getCurrentSession().createSQLQuery(qBuilder.toString()).list());
			for(int i=0;i<obj.size();i++)
			{
				System.out.println(obj.get(i));
			}
		}
		LOGGER.info("<<<<<< ReportDAOImpl.createUserPaymentReport");

		return obj;
	}
	*/
	@Transactional(propagation=Propagation.REQUIRED)
	 public List<Object> createUserPaymentReport(int userType) throws DAOException {
	  LOGGER.info(">>>>>> ReportDAOImpl.createUserPaymentReport userType:" + userType);
	  StringBuilder qBuilder = new StringBuilder("");
	  List<Object> obj = null;
	  if(userType == 1) {
	   qBuilder.append("select u.jct_user_name,ud.jct_user_details_profile_name,ud.jct_user_details_group_name,u.jct_user_customer_id,u.jct_role_id," +
	     "u.jct_created_by,h.jct_pmt_hdr_total_amt,d.jct_pmt_dtls_pmt_typ_desc,d.jct_pmt_dtls_modified_ts,d.jct_pmt_dtls_cheque_nos,d.jct_pmt_dtls_pmt_trans_nos,ud.jct_user_details_first_name,ud.jct_user_details_last_name " + 
	     "from jct_user u,jct_user_details ud,jct_payment_header h,jct_payment_details d " +
	     "where u.jct_user_id = ud.jct_user_id and u.jct_user_customer_id = h.jct_user_customer_id and h.jct_pmt_hdr_id = d.jct_pmt_hdr_id "+
	     "and u.jct_role_id = "+userType+" and u.jct_user_soft_delete = 0 and jct_user_disabled = 0 and u.jct_user_customer_id like('98%')" +
	     "ORDER BY u.jct_user_name");
	   obj = sessionFactory.getCurrentSession().createSQLQuery(qBuilder.toString()).list();
	  } else if (userType == 3) {
	   qBuilder.append("select u.jct_user_name,ud.jct_user_details_profile_name,ud.jct_user_details_group_name,u.jct_user_customer_id,u.jct_role_id," +
	     "u.jct_created_by,h.jct_pmt_hdr_total_amt,d.jct_pmt_dtls_pmt_typ_desc,d.jct_pmt_dtls_modified_ts,d.jct_pmt_dtls_cheque_nos,d.jct_pmt_dtls_pmt_trans_nos,ud.jct_user_details_first_name,ud.jct_user_details_last_name,fd.jct_fac_type,fd.jct_pmt_hdr_id " + 
	     "from jct_user u,jct_user_details ud,jct_payment_header h,jct_payment_details d,jct_facilitator_details fd " +
	     "where u.jct_user_id = ud.jct_user_id and u.jct_user_customer_id = h.jct_user_customer_id and h.jct_pmt_hdr_id = d.jct_pmt_hdr_id " +
	     "and fd.jct_pmt_hdr_id = h.jct_pmt_hdr_id and u.jct_role_id = "+ userType +
	     " and u.jct_active_yn != 6 and u.jct_user_soft_delete = 0 and jct_user_disabled = 0 " +
	     "ORDER BY u.jct_user_name");
	   obj = sessionFactory.getCurrentSession().createSQLQuery(qBuilder.toString()).list();
	  } else if (userType == 0) { // All
	   qBuilder.append("select u.jct_user_name,ud.jct_user_details_profile_name,ud.jct_user_details_group_name,u.jct_user_customer_id,u.jct_role_id," +
	     "u.jct_created_by,h.jct_pmt_hdr_total_amt,d.jct_pmt_dtls_pmt_typ_desc,d.jct_pmt_dtls_modified_ts,d.jct_pmt_dtls_cheque_nos,d.jct_pmt_dtls_pmt_trans_nos,ud.jct_user_details_first_name,ud.jct_user_details_last_name " + 
	     "from jct_user u,jct_user_details ud,jct_payment_header h,jct_payment_details d " +
	     "where u.jct_user_id = ud.jct_user_id and u.jct_user_customer_id = h.jct_user_customer_id and h.jct_pmt_hdr_id = d.jct_pmt_hdr_id "+
	     "and u.jct_role_id = 1 and u.jct_user_soft_delete = 0 and jct_user_disabled = 0 and u.jct_user_customer_id like('98%')"+
	     "ORDER BY u.jct_user_name");
	   obj = sessionFactory.getCurrentSession().createSQLQuery(qBuilder.toString()).list();
	   qBuilder.setLength(0); // reset qBuilder  
	   qBuilder.append("select u.jct_user_name,ud.jct_user_details_profile_name,ud.jct_user_details_group_name,u.jct_user_customer_id,u.jct_role_id," +
	     "u.jct_created_by,h.jct_pmt_hdr_total_amt,d.jct_pmt_dtls_pmt_typ_desc,d.jct_pmt_dtls_modified_ts,d.jct_pmt_dtls_cheque_nos,d.jct_pmt_dtls_pmt_trans_nos,ud.jct_user_details_first_name,ud.jct_user_details_last_name,fd.jct_fac_type,fd.jct_pmt_hdr_id " + 
	     "from jct_user u,jct_user_details ud,jct_payment_header h,jct_payment_details d,jct_facilitator_details fd " +
	     "where u.jct_user_id = ud.jct_user_id and u.jct_user_customer_id = h.jct_user_customer_id and h.jct_pmt_hdr_id = d.jct_pmt_hdr_id " +
	     "and fd.jct_pmt_hdr_id = h.jct_pmt_hdr_id and u.jct_role_id = 3 " +
	     "and u.jct_active_yn != 6 and u.jct_user_soft_delete = 0 and jct_user_disabled = 0 "+
	     "ORDER BY u.jct_user_name");
	   obj.addAll(sessionFactory.getCurrentSession().createSQLQuery(qBuilder.toString()).list());
	  }
	  LOGGER.info("<<<<<< ReportDAOImpl.createUserPaymentReport");
	  return obj;
	 }
	@Transactional(propagation=Propagation.REQUIRED)
	public int fetchFacilitatorTotalLimit(String custId, int paymentHdrId) throws DAOException {
		LOGGER.info(">>>>>> ReportDAOImpl.fetchFacilitatorTotalLimit");
		StringBuilder qBuilder = new StringBuilder("");		
		qBuilder.append("select jct_fac_total_limit from jct_facilitator_details where jct_user_customer_id = "+custId+" and jct_pmt_hdr_id = "+paymentHdrId);
		LOGGER.info("<<<<<< ReportDAOImpl.fetchFacilitatorTotalLimit");
		return (Integer) sessionFactory.getCurrentSession().createSQLQuery(qBuilder.toString()).uniqueResult();
	}

	@Transactional(propagation=Propagation.REQUIRED)
	public List<SurveyQuestionDTO> fetchUserIdFromSurvey(int userType) throws DAOException{
		LOGGER.info(">>>>>> ReportDAOImpl.fetchUserIdFromSurvey");
		//StringBuilder qBuilder = new StringBuilder("");
		if(userType == 0){
			return sessionFactory.getCurrentSession().getNamedQuery("getSurveyUserIdAndCountForAll").list();
		} else {
			return sessionFactory.getCurrentSession().getNamedQuery("getSurveyUserIdAndCount")
					.setParameter("userType", userType).list();			
		}
			
	}

	@Transactional(propagation=Propagation.REQUIRED)
	public BigInteger getSurveyQuestionCount() throws DAOException {
		LOGGER.info(">>>> ReportDAOImpl.getSurveyQuestionCount");
		String taskString = "Select max(table1.counts) from (select count(distinct(jct_survey_sub_qtn)) as counts from jct_survey_questions where "
				+ "jct_survey_user_type=1 and jct_survey_soft_delete = 0 group by jct_survey_taken_by ) table1";
		LOGGER.info("<<<< ReportDAOImpl.getSurveyQuestionCount");
		return (java.math.BigInteger)sessionFactory.getCurrentSession().createSQLQuery(taskString).uniqueResult();
	}

	@Transactional(propagation=Propagation.REQUIRED)
	public List<Object> getSurveyMainQtns(String emailId) throws DAOException {
		LOGGER.info(">>>> ReportDAOImpl.getSurveyMainQtns");
		return sessionFactory.getCurrentSession().createSQLQuery("select distinct(jct_survey_main_qtn), jct_survey_answer_type from jct_survey_questions where "
				+ "jct_survey_taken_by= :email and jct_survey_user_type=1")
				.setParameter("email", emailId).list();
	}
	
	@Transactional(propagation=Propagation.REQUIRED)
	public List<String> getSurveyAnswers(String emailId, String mainQtn, Integer ansType) throws DAOException {
		LOGGER.info(">>>> ReportDAOImpl.getSurveyAnswers");
		return sessionFactory.getCurrentSession().createSQLQuery("select distinct(jct_survey_answer) from jct_survey_questions where "
				+ "jct_survey_main_qtn = :mainQtn  and jct_survey_taken_by= :email and jct_survey_user_type=1 and jct_survey_answer_type = :ansType")
				.setParameter("mainQtn", mainQtn)
				.setParameter("email", emailId)
				.setParameter("ansType", ansType).list();
	}
	
	
	@Transactional(propagation=Propagation.REQUIRED)
	public List<Object> getBeforeSketchList(String occupationCode, int recordIndex, int status, int softDelete) throws DAOException {
		LOGGER.info(">>>> ReportDAOImpl.getBeforeSketchList");
		StringBuilder queryBldr = null;
		StringBuilder jrNoBuilder = new StringBuilder("");
		List<Object> returnList = null;
		List<String> jrList = null;
		queryBldr = new StringBuilder("Select distinct(bsview.jct_bs_created_by) from jct_before_sketch_view bsview");
		if ((occupationCode.trim().length() > 0)) {
			queryBldr.append(" where bsview.jct_user_occupation = :occupationCode " +
					"					and jct_bs_status_id = "+status+" and jct_bs_soft_delete = "+softDelete+"");
			jrList = sessionFactory.getCurrentSession().
					 createSQLQuery(queryBldr.toString())
					 .setParameter("occupationCode", occupationCode).
					 setFirstResult(recordIndex).setMaxResults(20).list();
		}  else {
			queryBldr.append(" where jct_bs_status_id = "+status+" and jct_bs_soft_delete = "+softDelete+"");
			jrList = sessionFactory.getCurrentSession().
					 createSQLQuery(queryBldr.toString()).
					 setFirstResult(recordIndex).setMaxResults(20).list();
		}
		if (jrList.size() > 0) {
			int size = 0;
			Iterator itr = jrList.iterator();
			while (itr.hasNext()) {
				jrNoBuilder.append("'" + (String) itr.next() + "'");
				if (jrList.size() - 1 > size) {
					jrNoBuilder.append(",");	
				} 
				size = size + 1;
			}
			queryBldr = new StringBuilder ("Select bsview.jct_bs_created_by, " +
					"bsview.jct_user_occupation, " +
					"bsview.jct_user_levels, bsview.jct_bs_task_desc, " +
					"bsview.jct_bs_energy, bsview.jct_bs_time_spent, " +
					"bsview.jct_first_name, bsview.jct_last_name " +
					"from jct_before_sketch_view bsview");
			if ((occupationCode.trim().length() > 0)) {
				queryBldr.append(" where bsview.jct_user_occupation = :occupationCode ");
				queryBldr.append(" and bsview.jct_bs_created_by in (" + jrNoBuilder.toString() + ") and jct_bs_status_id = "+status+" and jct_bs_soft_delete = "+softDelete+"");
				queryBldr.append(" and bsview.jct_bs_task_desc != '' ");
				queryBldr.append(" order by bsview.jct_bs_created_by");
				returnList = sessionFactory.getCurrentSession().
						createSQLQuery(queryBldr.toString()).setParameter("occupationCode", occupationCode).list();
			} else {
				queryBldr.append(" where bsview.jct_bs_created_by in (" + jrNoBuilder.toString() + ") and jct_bs_status_id = "+status+" and jct_bs_soft_delete = "+softDelete+"");
				queryBldr.append(" and bsview.jct_bs_task_desc != '' ");
				queryBldr.append(" order by bsview.jct_bs_created_by");
				returnList = sessionFactory.getCurrentSession().
						createSQLQuery(queryBldr.toString())
						.list();
			}
		} else {
			returnList = new ArrayList<Object>();
		}
		LOGGER.info("<<<< ReportDAOImpl.getBeforeSketchList");
		return returnList;
	}
	
	
	
	
	
	@Transactional(propagation=Propagation.REQUIRED)
	public List<String> getTotalCount(List<String> emailIdList, String occupationCode, int status, int softDel) throws DAOException {
		LOGGER.info(">>>> ReportDAOImpl.getTotalCount");
		StringBuilder jrNoBuilder = new StringBuilder("");
		List<String> jrList = null;
		jrNoBuilder = new StringBuilder("Select distinct(bsview.jct_bs_created_by) " +
				    						"from jct_before_sketch_view bsview");
		if ((occupationCode.trim().length() > 0)) {
			if (emailIdList == null) {
				jrNoBuilder.append(" where bsview.jct_user_occupation = :occupationCode and jct_bs_status_id = "+status+" and jct_bs_soft_delete = "+softDel+"");
				jrList = sessionFactory.getCurrentSession().
						createSQLQuery(jrNoBuilder.toString()).setParameter("occupationCode", occupationCode).list();
			} else {
				jrNoBuilder.append(" where bsview.jct_user_occupation = :occupationCode and jct_bs_status_id = "+status+" and jct_bs_soft_delete = "+softDel+" and bsview.jct_bs_created_by NOT IN :emailIdList");
				jrList = sessionFactory.getCurrentSession().
						createSQLQuery(jrNoBuilder.toString()).setParameter("occupationCode", occupationCode).
						setParameterList("emailIdList", emailIdList).list();
			}
			
		} else {
			if (emailIdList == null) {
				jrNoBuilder.append(" where jct_bs_status_id = "+status+" and jct_bs_soft_delete = "+softDel+"");
				jrList = sessionFactory.getCurrentSession().
						createSQLQuery(jrNoBuilder.toString()).list();
			} else {
				jrNoBuilder.append(" where jct_bs_status_id = "+status+" and jct_bs_soft_delete = "+softDel+" and bsview.jct_bs_created_by NOT IN :emailIdList");
				jrList = sessionFactory.getCurrentSession().
						createSQLQuery(jrNoBuilder.toString()).setParameterList("emailIdList", emailIdList).list();
			}
			
		}
		LOGGER.info("<<<< ReportDAOImpl.getTotalCount");
		return jrList;
	}
	
	
	
	@Transactional(propagation=Propagation.REQUIRED)
	public List<Object> getAfterSketchList(List<String> emailIdList, String occupationCode, int recordIndex, int statusId, int softDel) throws DAOException {
		LOGGER.info(">>>> ReportDAOImpl.getAfterSketchList");
		StringBuilder queryBldr = null;
		StringBuilder jrNoBuilder = new StringBuilder("");
		List<Object> returnList = null;
		List<String> jrList = null;
		queryBldr = new StringBuilder("Select distinct(asview.jct_as_created_by) from jct_after_sketch_view asview");
		if ((occupationCode.trim().length() > 0)) {
			queryBldr.append(" where asview.jct_user_occupation = :fnGrp " +
					"and jct_as_status_id = "+statusId+" and jct_as_soft_delete = "+softDel+"");
			jrList = sessionFactory.getCurrentSession().
					createSQLQuery(queryBldr.toString()).setParameter("fnGrp", occupationCode)
					.setMaxResults(20).list();
		}  else {
			queryBldr.append(" where jct_as_status_id = "+statusId+" and jct_as_soft_delete = "+softDel+"");
			jrList = sessionFactory.getCurrentSession().
					createSQLQuery(queryBldr.toString()).setFirstResult(recordIndex).
					setMaxResults(20).list();
		}
		
		if (jrList.size() > 0) {
			int size = 0;
			Iterator itr = jrList.iterator();
			while (itr.hasNext()) {
				jrNoBuilder.append("'" + (String) itr.next() + "'");
				if (jrList.size() - 1 > size) {
					jrNoBuilder.append(",");	
				} 
				size = size + 1;
			}
			queryBldr = new StringBuilder("Select asview.jct_as_created_by, " 		//0	
										+ "asview.jct_as_role_desc, "				//1
										+ "asview.jct_as_task_desc, "				//2
										+ "asview.jct_as_task_energy, "				//3
										+ "asview.jct_as_finalpage_time_spent, "	//4
										+ "asview.jct_as_element_desc, " 			//5
										+ "asview.jct_user_occupation, "	 		//6
										+ "asview.jct_user_levels, "				//7
										+ "asview.jct_first_name, "					//8
										+ "asview.jct_last_name "                   //,9
										+ "from jct_after_sketch_view asview");
			if ((occupationCode.trim().length() > 0) ) {
				queryBldr.append(" where asview.jct_user_occupation = :fnGrp ");
				queryBldr.append(" and asview.jct_as_created_by in (" + jrNoBuilder.toString() + ") and jct_as_status_id = "+statusId+" and jct_as_soft_delete = "+softDel+"");
			} 
			else {
				queryBldr.append(" where asview.jct_as_created_by in (" + jrNoBuilder.toString() + ") and jct_as_status_id = "+statusId+" and jct_as_soft_delete = "+softDel+"");
			}
			if (null != emailIdList) {
				queryBldr.append(" and jct_as_created_by NOT IN :list ");
			}
			queryBldr.append(" order by asview.jct_as_role_desc desc, " +
					"asview.jct_as_created_by, asview.jct_as_task_desc desc");
			
			if ((occupationCode.trim().length() > 0)) {
						if (null != emailIdList) {
							returnList = sessionFactory.getCurrentSession().
									createSQLQuery(queryBldr.toString()).setParameter("fnGrp", occupationCode)
									.setParameterList("list", emailIdList)
									.list();
						} else {
							returnList = sessionFactory.getCurrentSession().
									createSQLQuery(queryBldr.toString()).setParameter("fnGrp", occupationCode)
									.list();
						}
			}  else {
				if (null != emailIdList) {
					returnList =returnList = sessionFactory.getCurrentSession().
							createSQLQuery(queryBldr.toString())
							.setParameterList("list", emailIdList)
							.list();
				} else {
					returnList = sessionFactory.getCurrentSession().createSQLQuery(queryBldr.toString())
							.list();
				}
				
			}	
		} else {
			returnList = new ArrayList<Object>();
		}
		LOGGER.info("<<<< ReportDAOImpl.getAfterSketchList");
		return returnList;
	}
	
	
	@Transactional(propagation=Propagation.REQUIRED)
	public List<String> getASTotalCount(List<String> emailIdList, String occupationCode, int statusId, int softDel) throws DAOException {
		LOGGER.info(">>>> ReportDAOImpl.getASTotalCount");
		StringBuilder jrNoBuilder = new StringBuilder("");
		List<String> jrList = null;
		jrNoBuilder = new StringBuilder("Select distinct(asview.jct_as_created_by) from jct_after_sketch_view asview");
		if ((occupationCode.trim().length() > 0) ) {
			if (null != emailIdList) {
				jrNoBuilder.append(" where asview.jct_user_occupation = :fnGrp " +
						"and jct_as_status_id="+statusId+" and jct_as_soft_delete="+softDel+" and jct_as_created_by NOT IN :emailList");
				jrList = sessionFactory.getCurrentSession().createSQLQuery(jrNoBuilder.toString()).
						setParameter("fnGrp", occupationCode).
						setParameterList("emailList", emailIdList).list();
			} else {
				jrNoBuilder.append(" where asview.jct_user_occupation = :fnGrp " +
						"and jct_as_status_id="+statusId+" and jct_as_soft_delete="+softDel+"");
				jrList = sessionFactory.getCurrentSession().createSQLQuery(jrNoBuilder.toString()).
						setParameter("fnGrp", occupationCode).list();
			}
			
		}  else {
			if (null != emailIdList) {
				jrNoBuilder.append(" where jct_as_status_id="+statusId+" and jct_as_soft_delete="+softDel+" and jct_as_created_by NOT IN :emailList");
				jrList = sessionFactory.getCurrentSession().createSQLQuery(jrNoBuilder.toString()).
						setParameterList("emailList", emailIdList).list();
			} else {
				jrNoBuilder.append(" where jct_as_status_id="+statusId+" and jct_as_soft_delete="+softDel+"");
				jrList = sessionFactory.getCurrentSession().createSQLQuery(jrNoBuilder.toString()).list();
			}
		}
		LOGGER.info("<<<< ReportDAOImpl.getASTotalCount");
		return jrList;
	}

	@Transactional(propagation=Propagation.REQUIRED)
	public String getJctUserGroupNameById(int groupId) throws DAOException {
		return (String) sessionFactory.getCurrentSession().getNamedQuery("fetchUserGroupNameById")
				.setParameter("userGrpId", groupId).uniqueResult();
	}

}