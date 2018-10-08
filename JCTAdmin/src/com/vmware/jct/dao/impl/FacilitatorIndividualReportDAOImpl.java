package com.vmware.jct.dao.impl;

import java.util.ArrayList;
import java.util.List;

import org.hibernate.SessionFactory;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Repository;
import org.springframework.transaction.annotation.Propagation;
import org.springframework.transaction.annotation.Transactional;

import com.vmware.jct.dao.IFacilitatorIndividualReportDAO;
import com.vmware.jct.exception.DAOException;
import com.vmware.jct.model.JctUserLoginInfo;

/**
 * 
 * <p><b>Class name:</b> FacilitatorIndividualReportDAOImpl.java</p>
 * <p><b>Author:</b> InterraIT</p>
 * <p><b>Purpose:</b> This is a FacilitatorIndividualReportDAOImpl class. This artifact is Persistence Manager layer artifact.
 * FacilitatorIndividualReportDAOImpl implement IFacilitatorIndividualReportDAOImpl interface.
 * </p>
 * <p><b>Description:</b> This class used to perform insert, modify ,fetch data from db and delete operation. Save, update, delete operations are performed in DataAccessObject class </p>
 * </p>
 * The following methods are implemented
 * -createNewProgressReport()
 * -getUserDetails()
 * -createFaciIndiReport()
 */

@Repository
public class FacilitatorIndividualReportDAOImpl implements IFacilitatorIndividualReportDAO {

private static final Logger LOGGER = LoggerFactory.getLogger(FacilitatorIndividualReportDAOImpl.class);
	
	@Autowired
	private SessionFactory sessionFactory;
	
	/*New Progress Report*/
	
	@Transactional(propagation=Propagation.REQUIRED)
	public List<Object> populateAllProgressReports(String jctEmail) throws DAOException{
		LOGGER.info(">>>> FacilitatorIndividualReportDAOImpl.createNewProgressReport");
		List userIdList = null;
		List userLoginInfoList = null;
		List userNotExistsList = null;
		List userNotExistsInfoList = null;
		List userIdExistsList = null;
		List<Object> finalList = new ArrayList<Object>();
		String custID = "select jct_user_customer_id from jct_user where jct_user_email='"+jctEmail+"' and jct_role_id='3' and jct_user_soft_delete='0'";
		String custIDVal = (String) sessionFactory.getCurrentSession().createSQLQuery(custID).uniqueResult();
		String userId = "select jct_user_id from jct_user where jct_user_customer_id =(select jct_user_customer_id from jct_user where jct_user_email='"+jctEmail+"' and jct_role_id='3' and jct_user_soft_delete='0')and jct_role_id='1' and jct_user_soft_delete='0'";
		userIdList = sessionFactory.getCurrentSession().createSQLQuery(userId).list();
		String userIdExists = "SELECT jct_user_id FROM jct_user WHERE jct_user_id IN " +
				"(SELECT jct_user_id FROM jct_user_login_info) " +
				"and jct_user_customer_id = '"+custIDVal+"' " +
				"and jct_role_id = '1' and jct_user_soft_delete = '0'";
		userIdExistsList = sessionFactory.getCurrentSession().createSQLQuery(userIdExists).list();
		int size = userIdList.size();
		StringBuilder qBuilder = null;
		StringBuilder sBuilder = null;
		StringBuilder queryBuilder = null;
		for(int i = 0; i < size; i++){
			sBuilder = new StringBuilder("SELECT jct_user_id FROM jct_user WHERE jct_user_id NOT IN (SELECT jct_user_id FROM jct_user_login_info) " +
					"and jct_user_customer_id = '"+custIDVal+"' and jct_role_id = '1' and jct_user_soft_delete = '0'");
			userNotExistsList = sessionFactory.getCurrentSession().createSQLQuery(sBuilder.toString()).list();
		}
		try{
				for(int j = 0; j < userIdExistsList.size(); j++){
				qBuilder = new StringBuilder("select users.jct_user_name," +
						"users.jct_account_expiration_date," +
						"jct.jct_as_lastmodified_ts ," +
						"sum((TIMESTAMPDIFF(hour, (select min(ob.jct_start_time) from jct_user_login_info ob where ob.jct_jobref_no NOT LIKE '%00000000ADPRV%' and ob.jct_user_id = "+userIdExistsList.get(j)+"), (select max(obs.jct_end_time) from jct_user_login_info obs where obs.jct_jobref_no NOT LIKE '%00000000ADPRV%' and obs.jct_user_id = "+userIdExistsList.get(j)+"))))," +
						"sum(round((TIMESTAMPDIFF(minute, (select min(obg.jct_start_time) from jct_user_login_info obg where obg.jct_jobref_no NOT LIKE '%00000000ADPRV%' and obg.jct_user_id = "+userIdExistsList.get(j)+"), (select max(obss.jct_end_time) from jct_user_login_info obss where obss.jct_jobref_no NOT LIKE '%00000000ADPRV%' and obss.jct_user_id = "+userIdExistsList.get(j)+"))%60)))," +
						"(select max(myObj.jct_page_info) from jct_user_login_info myObj where myObj.jct_jobref_no NOT LIKE '%00000000ADPRV%' and myObj.jct_user_id = "+userIdExistsList.get(j)+" and myObj.jct_user_login_info_id = (select max(otr.jct_user_login_info_id) from jct_user_login_info otr where otr.jct_jobref_no NOT LIKE '%00000000ADPRV%' and otr.jct_user_id = "+userIdExistsList.get(j)+")) as page , dtl.jct_user_details_group_name, " +
						"dtl.jct_user_details_last_name, dtl.jct_user_details_first_name " +
						"from jct_user_login_info jct,jct_user users, jct_user_details dtl " +
						"where jct.jct_jobref_no NOT LIKE '%00000000ADPRV%' and jct.jct_user_id = '"+userIdExistsList.get(j)+"' and users.jct_user_id= '"+userIdExistsList.get(j)+"' " +
						"and users.jct_role_id = '1' and users.jct_user_soft_delete = '0' " +
						"and users.jct_user_id = dtl.jct_user_id and users.jct_user_customer_id = '"+custIDVal+"' " +
						"and jct_start_time IN (select (usr.jct_start_time) from jct_user_login_info usr where jct.jct_user_id = usr.jct_user_id) " +
						"group by dtl.jct_user_details_group_name");
					
				userLoginInfoList = sessionFactory.getCurrentSession().createSQLQuery(qBuilder.toString()).list();
				finalList.add(userLoginInfoList);
			}
			for(int j = 0; j < userNotExistsList.size(); j++){
				queryBuilder = new StringBuilder("select users.jct_user_name, " +
						"users.jct_account_expiration_date, " +
						"dtl.jct_user_details_group_name, " +
						"dtl.jct_user_details_last_name, dtl.jct_user_details_first_name " +
						"from jct_user users, " +
						"jct_user_details dtl " +
						"where users.jct_user_id= '"+userNotExistsList.get(j)+"' " +
								"and users.jct_role_id = '1' " +
								"and users.jct_user_soft_delete = '0' " +
								"and users.jct_user_id = dtl.jct_user_id " +
								"and users.jct_user_customer_id = '"+custIDVal+"' group by dtl.jct_user_details_group_name");
				userNotExistsInfoList = sessionFactory.getCurrentSession().createSQLQuery(queryBuilder.toString()).list();
				finalList.add(userNotExistsInfoList);
				
			}
					}
		catch(Exception e){
			System.out.println(e.getLocalizedMessage());
		}
		LOGGER.info("<<<< FacilitatorIndividualReportDAOImpl.createNewProgressReport");
		return finalList;
	}
	

	
	@Transactional(propagation=Propagation.REQUIRED)
	public List<Object> getUserDetails(int usrType) throws DAOException {	
		LOGGER.info(">>>> FacilitatorIndividualReportDAOImpl.getUserDetails");
		StringBuilder qBuilder = new StringBuilder("");
		if ( usrType != 0 ){
			qBuilder.append("select jct_user_name,jct_user_details_profile_name,jct_user_details_group_name,jct_user_customer_id,jct_created_ts,jct_account_expiration_date,jct_role_id "+
						"from jct_user,jct_user_details,jct_functions "+ 
						"where jct_user.jct_user_id = jct_user_details.jct_user_id and jct_functions.jct_functions_id = jct_user_details.jct_user_details_profile_id and jct_user_soft_delete = '0' and jct_role_id = "+usrType);
		
		} else {	//	= 0 -> all users except ADMIN
			qBuilder.append("select jct_user_name,jct_user_details_profile_name,jct_user_details_group_name,jct_user_customer_id,jct_created_ts,jct_account_expiration_date,jct_role_id "+
						"from jct_user,jct_user_details,jct_functions "+ 
						"where jct_user.jct_user_id = jct_user_details.jct_user_id and jct_functions.jct_functions_id = jct_user_details.jct_user_details_profile_id and jct_user_soft_delete = '0' and jct_role_id NOT IN (2,4)");
		}		
		LOGGER.info("<<<< FacilitatorIndividualReportDAOImpl.getUserDetails");
		return sessionFactory.getCurrentSession().createSQLQuery(qBuilder.toString()).list();
	}
	
	@Transactional(propagation=Propagation.REQUIRED)
	public List<Object> createFaciIndiReport(String userType) throws DAOException {
		LOGGER.info(">>>> FacilitatorIndividualReportDAOImpl.createFaciIndiReport");
		StringBuilder qBuilder = new StringBuilder("");
		if ( !userType.equals("0") ){
			qBuilder.append("select jct_user_name,jct_user_details_profile_name,jct_user_details_group_name,jct_user_customer_id "+
					"from jct_user,jct_user_details,jct_functions "+ 
					"where jct_user.jct_user_id = jct_user_details.jct_user_id and jct_functions.jct_functions_id = jct_user_details.jct_user_details_profile_id and jct_user_soft_delete = '0' and jct_role_id = "+userType);
		} else {	//	= 0 -> all users except ADMIN
			qBuilder.append("select jct_user_name,jct_user_details_profile_name,jct_user_details_group_name,jct_user_customer_id "+
					"from jct_user,jct_user_details,jct_functions "+ 
					"where jct_user.jct_user_id = jct_user_details.jct_user_id and jct_functions.jct_functions_id = jct_user_details.jct_user_details_profile_id and jct_user_soft_delete = '0' and jct_role_id NOT IN (2,4)");
		}
		LOGGER.info("<<<< FacilitatorIndividualReportDAOImpl.createFaciIndiReport");
		return sessionFactory.getCurrentSession().createSQLQuery(qBuilder.toString()).list();
	}
	
	
	@Transactional(propagation=Propagation.REQUIRED)
	public List<Object> populateProgressReports(String jctEmail, String userGroupName) throws DAOException{
		LOGGER.info(">>>> FacilitatorIndividualReportDAOImpl.createNewProgressReport");
	List userIdList = null;
	List userLoginInfoList = null;
	List userNotExistsList = null;
	List userNotExistsInfoList = null;
	List userIdExistsList = null;
	List<Object> finalList = new ArrayList<Object>();
	String custID = "select jct_user_customer_id from jct_user where jct_user_email='"+jctEmail+"' and jct_role_id='3' and jct_user_soft_delete='0'";
	String custIDVal = (String) sessionFactory.getCurrentSession().createSQLQuery(custID).uniqueResult();
	String userId = "select jct_user_id from jct_user where jct_user_customer_id =(select jct_user_customer_id from jct_user where jct_user_email='"+jctEmail+"' and jct_role_id='3' and jct_user_soft_delete='0')and jct_role_id='1' and jct_user_soft_delete='0'";
	userIdList = sessionFactory.getCurrentSession().createSQLQuery(userId).list();	
	String userIdExists = "SELECT usr.jct_user_id " +
			"FROM jct_user usr, jct_user_details userDtl " +
			"WHERE usr.jct_user_id IN (SELECT jct_user_id FROM jct_user_login_info) " +
			"and usr.jct_user_customer_id = '"+custIDVal+"' and usr.jct_role_id = '1' " +
			"and usr.jct_user_soft_delete = '0' and userDtl.jct_user_details_group_name = '"+userGroupName+"' " +
			"and usr.jct_user_id = userDtl.jct_user_id";	
	userIdExistsList = sessionFactory.getCurrentSession().createSQLQuery(userIdExists).list();
	int size = userIdList.size();
	StringBuilder qBuilder = null;
	StringBuilder sBuilder = null;
	StringBuilder queryBuilder = null;
	for(int i = 0; i < size; i++){
		sBuilder = new StringBuilder("SELECT jct_user_id FROM jct_user WHERE jct_user_id NOT IN (SELECT jct_user_id FROM jct_user_login_info) " +
				"and jct_user_customer_id = '"+custIDVal+"' and jct_role_id = '1' and jct_user_soft_delete = '0'");
		userNotExistsList = sessionFactory.getCurrentSession().createSQLQuery(sBuilder.toString()).list();
	}
	try{
			for(int j = 0; j < userIdExistsList.size(); j++){
			qBuilder = new StringBuilder("select users.jct_user_name," +
					"users.jct_account_expiration_date," +
					"jct.jct_as_lastmodified_ts ," +
					"sum((TIMESTAMPDIFF(hour, (select min(ob.jct_start_time) from jct_user_login_info ob where ob.jct_jobref_no NOT LIKE '%00000000ADPRV%' and ob.jct_user_id = "+userIdExistsList.get(j)+"), (select max(obs.jct_end_time) from jct_user_login_info obs where obs.jct_jobref_no NOT LIKE '%00000000ADPRV%' and obs.jct_user_id = "+userIdExistsList.get(j)+"))))," +
					"sum(round((TIMESTAMPDIFF(minute, (select min(obg.jct_start_time) from jct_user_login_info obg where obg.jct_jobref_no NOT LIKE '%00000000ADPRV%' and obg.jct_user_id = "+userIdExistsList.get(j)+"), (select max(obss.jct_end_time) from jct_user_login_info obss where obss.jct_jobref_no NOT LIKE '%00000000ADPRV%' and obss.jct_user_id = "+userIdExistsList.get(j)+"))%60)))," +
					"(select max(myObj.jct_page_info) from jct_user_login_info myObj where myObj.jct_jobref_no NOT LIKE '%00000000ADPRV%' and myObj.jct_user_id = "+userIdExistsList.get(j)+" and myObj.jct_user_login_info_id = (select max(otr.jct_user_login_info_id) from jct_user_login_info otr where otr.jct_jobref_no NOT LIKE '%00000000ADPRV%' and otr.jct_user_id = "+userIdExistsList.get(j)+")) as page , dtl.jct_user_details_group_name, " +
					"dtl.jct_user_details_last_name, dtl.jct_user_details_first_name " +
					"from jct_user_login_info jct,jct_user users, jct_user_details dtl " +
					"where jct.jct_jobref_no NOT LIKE '%00000000ADPRV%' and jct.jct_user_id = '"+userIdExistsList.get(j)+"'" +
					"and users.jct_user_id= '"+userIdExistsList.get(j)+"' " +
					"and users.jct_role_id = '1' " +
					"and users.jct_user_soft_delete = '0'" +
					"and users.jct_user_id = dtl.jct_user_id " +
					"and users.jct_user_customer_id = '"+custIDVal+"' and dtl.jct_user_details_group_name = '"+userGroupName+"'" + 
					"and jct_start_time IN (select (usr.jct_start_time) from jct_user_login_info usr where jct.jct_user_id = usr.jct_user_id) " +
					"order by dtl.jct_user_details_group_name");
			
				userLoginInfoList = sessionFactory.getCurrentSession().createSQLQuery(qBuilder.toString()).list();
				finalList.add(userLoginInfoList);
		}
		for(int j = 0; j < userNotExistsList.size(); j++){
			queryBuilder = new StringBuilder("select users.jct_user_name," +
					"users.jct_account_expiration_date," +
					"dtl.jct_user_details_group_name " +
					"from jct_user users, jct_user_details dtl " +
					"where users.jct_user_id= '"+userNotExistsList.get(j)+"' " +
					"and users.jct_role_id = '1' " +
					"and users.jct_user_soft_delete = '0'" +
					"and users.jct_user_id = dtl.jct_user_id " +
					"and users.jct_user_customer_id = '"+custIDVal+"' and dtl.jct_user_details_group_name = '"+userGroupName+"' order by dtl.jct_user_details_group_name");
			userNotExistsInfoList = sessionFactory.getCurrentSession().createSQLQuery(queryBuilder.toString()).list();
			finalList.add(userNotExistsInfoList);
			}
				}
	catch(Exception e){
		System.out.println(e.getLocalizedMessage());
	}
	LOGGER.info("<<<< FacilitatorIndividualReportDAOImpl.createNewProgressReport");
	System.out.println(finalList);
	return finalList;
	}
	
	public List<JctUserLoginInfo> getTime (String userName) throws DAOException {
		return sessionFactory.getCurrentSession().getNamedQuery("getObjByinfoCrtfBy").setParameter("userName", userName).list();
	}
}