package com.vmware.jct.dao.impl;

import java.util.Date;
import java.util.List;
import java.util.StringTokenizer;

import org.hibernate.SessionFactory;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Repository;
import org.springframework.transaction.annotation.Propagation;
import org.springframework.transaction.annotation.Transactional;

import com.vmware.jct.common.utility.CommonConstants;
import com.vmware.jct.dao.DataAccessObject;
import com.vmware.jct.dao.IAuthenticatorDAO;
import com.vmware.jct.dao.dto.DemographicDTO;
import com.vmware.jct.dao.dto.FunctionDTO;
import com.vmware.jct.dao.dto.LevelDTO;
import com.vmware.jct.dao.dto.OccupationListDTO;
import com.vmware.jct.dao.dto.RegionDTO;
import com.vmware.jct.exception.DAOException;
import com.vmware.jct.exception.DAONullException;
import com.vmware.jct.exception.JCTException;
import com.vmware.jct.model.JctActionPlan;
import com.vmware.jct.model.JctAfterSketchHeader;
import com.vmware.jct.model.JctAfterSketchHeaderTemp;
import com.vmware.jct.model.JctBeforeSketchHeader;
import com.vmware.jct.model.JctBeforeSketchHeaderTemp;
import com.vmware.jct.model.JctBeforeSketchQuestion;
import com.vmware.jct.model.JctInstructionBar;
import com.vmware.jct.model.JctOnetOccupationList;
import com.vmware.jct.model.JctPopupInstruction;
import com.vmware.jct.model.JctStatusSearch;
import com.vmware.jct.model.JctSurveyQuestions;
import com.vmware.jct.model.JctUser;
import com.vmware.jct.model.JctUserInfo;
import com.vmware.jct.model.JctUserLoginInfo;
import com.vmware.jct.model.JctUserProfile;
import com.vmware.jct.service.vo.MyAccountVO;
import com.vmware.jct.service.vo.UserVO;


/**
 * 
 * <p><b>Class name:</b> AuthenticatorDAOImpl.java</p>
 * <p><b>Author:</b> InterraIt</p>
 * <p><b>Purpose:</b> This is a AuthenticatorDAOImpl class. This artifact is Persistence Manager layer artifact.
 * AuthenticatorDAOImpl implement IAuthenticatorDAO interface and override the following  methods.
 * -authenticateUser(UserVO userVO)
 * -checkUser(UserVO userVO)
 * -checkUsers(UserVO userVO)
 * -register(JctUser user)
 * -resetPassword(JctUser user)
 * -merge(JctUser user)</p>
 * <p><b>Description:</b> This class used to perform insert, modify ,fetch data from db and delete operation. Save, update, delete operations are performed in DataAccessObject class </p>
 * <p><b>Revision History:</b>
 * 	<li>InterraIt - 19/Jan/2014 - Implement Exception </li>
 * <li>InterraIt - 31/Jan/2014 - Implement comments, introduce Named query, Transactional attribute </li>
 * </p>
 */

@Repository
public class AuthenticatorDAOImpl extends DataAccessObject implements IAuthenticatorDAO {

	@Autowired
	private SessionFactory sessionFactory;
	
	private static final Logger logger = LoggerFactory.getLogger(AuthenticatorDAOImpl.class);
	
	/**
	 * <p><b>Description:</b>  This method verify user id, password, active status and return user object</p>
	 * @param It will take userVO object as a input parameter
	 * @return It returns list of JctUser data from database
	 * @throws DAOException -it throws DAOException
	 * 
	 */
	@Transactional(propagation=Propagation.REQUIRED)
	public List<JctUser> authenticateUser(UserVO userVO) throws DAOException  {
		
		return sessionFactory.getCurrentSession().getNamedQuery("fetchJctUser").setString("email", userVO.getEmail())
				.setString("password", userVO.getPassword())
				//.setInteger("active", CommonConstants.ACTIVATED)
				.setInteger("complete", CommonConstants.ACTIVATION_COMPLETE)
				.setInteger("active", CommonConstants.ACTIVATED_ONLY_EMAIL)
				.setInteger("inactive", CommonConstants.CREATED)
				.setInteger("chequeProb", CommonConstants.CHECK_PAYMENT_NOT_CLEARED)
				.list();
	}
	
	/**
	 * <p><b>Description:</b>  Bases on user id this method return user object</p>
	 * @param It will take userVO object as a input parameter
	 * @return It returns list of JctUser data from database
	 * @throws DAOException -it throws DAOException
	 * 
	 */
	@Transactional(propagation=Propagation.REQUIRED,readOnly = true)
	public List<JctUser> checkUser(UserVO userVO) throws DAOException {
		return sessionFactory.getCurrentSession().getNamedQuery("fetchJctUserByEmailId").setString("emailId", userVO.getEmail()).list();
	}
	@Transactional(propagation=Propagation.REQUIRED,readOnly = true)
	public List<JctUser> checkFacilitatorUser(UserVO userVO) throws DAOException {
		return sessionFactory.getCurrentSession().getNamedQuery("fetchJctFacilitatorUserByEmailId").setString("emailId", userVO.getEmail()).list();
	}
	@Transactional(propagation=Propagation.REQUIRED,readOnly = true)
	public List<JctUser> checkInactiveUser(UserVO userVO) throws DAOException {
		return sessionFactory.getCurrentSession().getNamedQuery("fetchInactiveJctUserByEmailId").setString("emailId", userVO.getEmail()).list();
	}
	
	@Transactional(propagation=Propagation.REQUIRED,readOnly = true)
	public List<JctUser> checkActiveUser(UserVO userVO) throws DAOException {
		return sessionFactory.getCurrentSession().getNamedQuery("fetchActiveJctUserByEmailId").setString("emailId", userVO.getEmail()).list();
	}
	
	/**
	 * <p><b>Description:</b>  Bases on user id and password this method return user object</p>
	 * @param It will take userVO object as a input parameter
	 * @return It returns list of JctUser data from database
	 * @throws DAOException -it throws DAOException
	 * 
	 */
	@Transactional(propagation=Propagation.REQUIRED,readOnly = true)
	public List<JctUser> checkUsers(UserVO userVO) throws DAOException {
		return sessionFactory.getCurrentSession().getNamedQuery("fetchJctUserByEmailIdPassword")
				.setString("email", userVO.getEmail())
				.setString("password", userVO.getInitialPassword()).list();
	}
	
	
	/**
	 * <p><b>Description:</b> This method is used for storing user data</p>
	 * @param It will take JctUser object as a input parameter
	 * @return It returns string object
	 * @throws DAOException -it throws DAOException
	 * 
	 */
	@Override
	@Transactional(propagation=Propagation.REQUIRED)
	public String register(JctUser user) throws DAONullException, DAOException {
	
	return save(user);
		
	}

	
	/**
	 * <p><b>Description:</b> This method is used for update user data</p>
	 * @param It will take JctUser object as a input parameter
	 * @return It returns string object
	 * @throws DAOException -it throws DAOException
	 * 
	 */
	@Override
	@Transactional(propagation=Propagation.REQUIRED)
	public String updateUser(JctUser user) throws DAOException {
		return update(user);
	}
	
	
	/**
	 * <p><b>Description:</b> This method is used for update user data</p>
	 * @param It will take JctUser object as a input parameter
	 * @return It returns string object
	 * @throws DAOException -it throws DAONullException,DAOException
	 * 
	 */
	@Override
	@Transactional(propagation=Propagation.REQUIRED)
	public String merge(JctUser user) throws DAONullException, DAOException {
					
			
		return update(user);
		
	}

	
	/**
	 * <p><b>Description:</b> This method is used for fetching  user info data based on user id and status</p>
	 * @param It will take JctUser object as a input parameter
	 * @return It returns string object
	 * @throws DAOException -it throws ,DAOException
	 * 
	 */
	@Transactional(propagation=Propagation.REQUIRED,readOnly = true)
	public List<JctUserLoginInfo> getTaskPendingDetails(int userId)
			throws DAOException {
		logger.info(">>>> AuthenticatorDAOImpl.getTaskPendingDetails"); 
		List<JctUserLoginInfo> userList = sessionFactory.getCurrentSession().getNamedQuery("fetchJctUserLogInfo")
										  		.setInteger("userId", userId)
										  		.setInteger("status", CommonConstants.TASK_PENDING).list();
		logger.info("<<<< AuthenticatorDAOImpl.getTaskPendingDetails");
		return userList;
	}

	/**
	 * <p><b>Description:</b> his method is used for storing user info data</p>
	 * @param It will take JctUserLoginInfo object as a input parameter
	 * @return It returns string object
	 * @throws DAOException -it throws ,DAOException
	 * 
	 */
	@Transactional(propagation=Propagation.REQUIRED)
	public String saveLoginInfo(JctUserLoginInfo info) throws DAOException {
		return save(info);
	}
	
	/**
	 * <p><b>Description:</b> This method is used for update user info data</p>
	 * @param It will take jobRefNo ,rowId and lastActivePage  as a input parameter
	 * @return It returns string object
	 * @throws DAOException -it throws ,DAOException
	 * 
	 */
	@Transactional(propagation=Propagation.REQUIRED)
	public String updateLoginInfo(String jobRefNo, int rowId, String lastActivePage) throws DAOException {
		logger.info(">>>> AuthenticatorDAOImpl.updateLoginInfo"); 
		String status = "failure";
		try{
			sessionFactory.getCurrentSession().getNamedQuery("updateUserInfo")
				.setParameter("endTime", new Date())
				.setParameter("lastActivePage",lastActivePage)
				.setParameter("job", jobRefNo)
				.setParameter("rowId", rowId).executeUpdate();
			status = "success";
		}catch(Exception ez){
			logger.error(ez.getLocalizedMessage());
		}
		logger.info("<<<< AuthenticatorDAOImpl.updateLoginInfo"); 
		return status;
	}
	
	/**
	 * <p><b>Description:</b> This method return list of region </p>
	 * @param Inull
	 * @return It returns List of RegionDTO 
	 * @throws DAOException -it throws ,DAOException
	 * 
	 */
	@Transactional(propagation=Propagation.REQUIRED)
	public List<RegionDTO> getJctRegionList() throws DAOException{
		return sessionFactory.getCurrentSession().getNamedQuery("fetchRegion").list();
	}
	
	/**
	 * <p><b>Description:</b> This method return list of function </p>
	 * @param Inull
	 * @return It returns List of RegionDTO 
	 * @throws DAOException -it throws ,DAOException
	 * 
	 */
	@Transactional(propagation=Propagation.REQUIRED)
	public List<FunctionDTO> getJctFunctionList() throws DAOException{
		return sessionFactory.getCurrentSession().getNamedQuery("fetchFunction").list();
	}
	
	/**
	 * <p><b>Description:</b>  This method return list of Job level </p>
	 * @param Null
	 * @return It returns list of  LevelDTO
	 * @throws DAOException -it throws ,DAOException
	 * 
	 */
	@Transactional(propagation=Propagation.REQUIRED)
	public List<LevelDTO> getJctLevelList() throws DAOException{
		return sessionFactory.getCurrentSession().getNamedQuery("fetchJobLevel").list();
	}
	
	/**
	 * <p><b>Description:</b>  This method return list of users </p>
	 * @param It will take userName  as a input parameter
	 * @return It returns list of  JctUser
	 * @throws DAOException -it throws ,DAOException
	 * 
	 */
	@Transactional(propagation=Propagation.REQUIRED)
	public List<JctUser> getUserList(String userName, int roleId) throws DAOException{
		String qry = "fetchUserList";
		List<JctUser> userList = null;
		if (roleId == 0) {
			userList = sessionFactory.getCurrentSession().getNamedQuery(qry)
					.setParameter("email", userName)
					.setParameter("roleId", 3).list();
		} else if (roleId == 1){
			userList = sessionFactory.getCurrentSession().getNamedQuery(qry)
					.setParameter("email", userName)
					.setParameter("roleId", 1).list();
		} else {
			qry = "fetchUserListAll";
			userList = sessionFactory.getCurrentSession().getNamedQuery(qry)
					.setParameter("email", userName).list();
		}
		return userList;
	}
	
	public List<MyAccountVO> getFacilitatorDetails (String custId) throws DAOException {
		return sessionFactory.getCurrentSession().getNamedQuery("getFacilitatorDetails")
				.setParameter("custId", custId).list();
	}
	
	/**
	 * <p><b>Description:</b>  This method return maximum login info PK </p>
	 * @param It will take jobReferenceNo  as a input parameter
	 * @return It returns integer object
	 * @throws DAOException -it throws ,DAOException
	 * 
	 */
	@Transactional(propagation=Propagation.REQUIRED)
	public int getMaxLoginInfoId(String jobReferenceNo) throws DAOException {
		return (Integer)sessionFactory.getCurrentSession().getNamedQuery("selectMaxReslt").setParameter("jobRfNo", jobReferenceNo).uniqueResult();
	}
	
	/**
	 * <p><b>Description:</b>   This method updates the login info for the PK </p>
	 * @param It will take id  as a input parameter
	 * @return It returns string object
	 * @throws DAOException -it throws ,DAOException
	 * 
	 */
	@Transactional(propagation=Propagation.REQUIRED)
	public String updateLoginInfo(int id, String page) throws DAOException {
		logger.info(">>>> AuthenticatorDAOImpl.updateLoginInfo"); 
		String status = "failure";
		try{
			sessionFactory.getCurrentSession().getNamedQuery("updateMaxReslt")
				.setParameter("endTime", new Date())
				//.setParameter("lastActivePage","/user/view/finalPage.jsp")
				.setParameter("lastActivePage", page)
				.setParameter("id", id).executeUpdate();
			status = "success";
		}catch(Exception ez){
			logger.error(ez.getLocalizedMessage());
		}
		logger.info("<<<< AuthenticatorDAOImpl.updateLoginInfo"); 
		return status;
	}
	
	/**
	 * <p><b>Description:</b>   This method checks whether the task is completed or not for the job reference id passed </p>
	 * @param It will take refNo  as a input parameter
	 * @return It returns Long object
	 * @throws DAOException -it throws ,DAOException
	 * 
	 */
	@Transactional(propagation=Propagation.REQUIRED)
	public Long getIsCompleted(String refNo) throws DAOException {
		return (Long)sessionFactory.getCurrentSession().getNamedQuery("fetchCompletedStatusCount")
				.setParameter("jrNo", refNo)
				.setParameter("jrNo2", refNo).uniqueResult();
	}
	
	/**
	 * <p><b>Description:</b>   This method updates the next page val of login info table </p>
	 * @param It will take jobReferenceNo and nextPage  as a input parameter
	 * @return It returns Long object
	 * @throws DAOException -it throws ,DAOException
	 * 
	 */
	@Transactional(propagation=Propagation.REQUIRED)
	public int updateNextPage(String jobReferenceNo, String nextPage)
			throws DAOException {
		return sessionFactory.getCurrentSession().getNamedQuery("updateNextPage")
				.setParameter("nextPage", nextPage)
				.setParameter("jrNo", jobReferenceNo).executeUpdate();
	}
	/**
	 * <p><b>Description:</b>   This method fetch user information based on user id and user name </p>
	 * @param It will take user id and user name   as a input parameter
	 * @return It returns Long object
	 * @throws DAOException -it throws ,DAOException
	 * 
	 */
	@Transactional(propagation=Propagation.REQUIRED)
	public List<JctUserInfo> fetchUserInfo(int userId, String userName)
			throws DAOException {
		return sessionFactory.getCurrentSession().getNamedQuery("fetchFailedInstance")
				.setParameter("userId", userId)
				.setParameter("email", userName).list();
	}
	/**
	 * <p><b>Description:</b>   This method fetch user information  </p>
	 * @param It will take JctUserInfo   as a input parameter
	 * @return It returns Long object
	 * @throws DAOException -it throws ,DAOException
	 * 
	 */
	@Transactional(propagation=Propagation.REQUIRED)
	public String saveUserInfo(JctUserInfo info) throws DAOException {
		return update(info);
	}
	/**
	 * <p><b>Description:</b>   This method fetch email id from user table based on user name</p>
	 * @param It will take userName   as a input parameter
	 * @return It returns Long object
	 * @throws DAOException -it throws ,DAOException
	 * 
	 */
	@Transactional(propagation=Propagation.REQUIRED)
	public List<JctUserInfo> fetchUserInfoByEmail(String userName)
			throws DAOException {
		return sessionFactory.getCurrentSession().getNamedQuery("fetchFailedInstanceByEmail")
				.setParameter("email", userName).list();
	}
	
	/**
	 * <p><b>Description:</b>   This method fetch instruction from instruction table based on profile id and releted page </p>
	 * @param It will take profileId and relatedPage   as a input parameter
	 * @return It returns JctInstructionBar
	 * @throws DAOException -it throws ,DAOException
	 * 
	 */
	@Transactional(propagation = Propagation.REQUIRED)
	public JctInstructionBar populateExistingInstruction(int profileId, String relatedPage)
			throws DAOException {
		logger.info(">>>> AuthenticatorDAOImpl.populateExistingInstruction");
		JctInstructionBar object = null;
		List<JctInstructionBar> listObj = sessionFactory.getCurrentSession()
				.getNamedQuery("fetchInstructionById")
				.setParameter("profileId", profileId)
				.setParameter("relatedPage", relatedPage).list();
		if (listObj.size() > 0) {
			object = (JctInstructionBar) listObj.get(0);
		} 
		logger.info("<<<< AuthenticatorDAOImpl.populateExistingInstruction");
		return object;	
		/*return (JctInstructionBar) sessionFactory.getCurrentSession().
				getNamedQuery("fetchInstructionById")
				.setParameter("profileId", profileId).setParameter("relatedPage", relatedPage).list().get(0);	*/	
	}
	/**
	 * <p><b>Description:</b>   This method fetch profile id from user table based on email id </p>
	 * @param It will take email id as a input parameter
	 * @return It returns profile id - Integer
	 * @throws DAOException -it throws ,DAOException
	 * 
	 */
	@Transactional(propagation = Propagation.REQUIRED)
	public Integer fetchProfileId(String emailId) throws DAOException {
		return (Integer) sessionFactory.getCurrentSession().getNamedQuery("fetchProfileByEmailId")
				.setParameter("email", emailId).uniqueResult();
	}

	/**
	 * <p><b>Description:</b>   This method will fetch list of demographic information</p>
	 * @param It will take email id as a input parameter
	 * @return It returns List<DemographicDTO>
	 * @throws DAOException -it throws ,DAOException
	 * 
	 */
	@Transactional(propagation = Propagation.REQUIRED)
	public List<DemographicDTO> getDemographicInformation(String emailId)
			throws DAOException {
		return sessionFactory.getCurrentSession().getNamedQuery("fetchDemographicByEmailId")
				.setParameter("emailId", emailId).list();
	}
	/**
	 * <p><b>Description:</b>   This method will fetch Email Id</p>
	 * @param It will take rowNo as a input parameter
	 * @return It returns Email Id
	 * @throws DAOException -it throws ,DAOException
	 * 
	 */
	public String getEmailIdByRowNos ( int rowNo ) throws DAOException {
		return (String) sessionFactory.getCurrentSession().getNamedQuery("getEmailId")
				.setParameter("rowId", rowNo).uniqueResult();
	}
	/**
	 * <p><b>Description:</b>   This method will fetch list of before sketch</p>
	 * @param It will take job reference nos as a input parameter
	 * @return It returns List<JctBeforeSketchHeader>
	 * @throws DAOException -it throws ,DAOException
	 * 
	 */
	@Transactional(propagation = Propagation.REQUIRED)
	public List<JctBeforeSketchHeader> getBeforeSketchDummyEntry(
			String jobReferenceNos) throws DAOException {
		return sessionFactory.getCurrentSession().getNamedQuery("getDummyBeforeSketchData")
				.setParameter("jrNo", jobReferenceNos).list();
	}
	/**
	 * 
	 */
	@Transactional(propagation = Propagation.REQUIRED)
	public int deleteBSData (String jobRefNos) throws DAOException {
		return sessionFactory.getCurrentSession().getNamedQuery("deleteDummyBSData")
				.setParameter("jrNo", jobRefNos).executeUpdate();
	}
	
	@Transactional(propagation = Propagation.REQUIRED)
	public int deleteBSChildData (String jobRefNos) throws DAOException {
		return sessionFactory.getCurrentSession().getNamedQuery("deleteDummyBSChildData")
				.setParameter("jrNo", jobRefNos).executeUpdate();
	}
	/**
	 * <p><b>Description:</b>   This method will fetch list of before sketch temp</p>
	 * @param It will take job reference nos as a input parameter
	 * @return It returns List<JctBeforeSketchHeaderTemp>
	 * @throws DAOException -it throws ,DAOException
	 * 
	 */
	@Transactional(propagation = Propagation.REQUIRED)
	public List<JctBeforeSketchHeaderTemp> getBeforeSketchTempDummyEntry(
			String jobReferenceNos) throws DAOException {
		return sessionFactory.getCurrentSession().getNamedQuery("getDummyBeforeSketchTempData")
				.setParameter("jrNo", jobReferenceNos).list();
	}
	/**
	 * 
	 */
	@Transactional(propagation = Propagation.REQUIRED)
	public int deleteBSTempData (String jobRefNos) throws DAOException {
		//deleteDummyBSTempData Delete(header);
		return sessionFactory.getCurrentSession().getNamedQuery("deleteDummyBSTempData")
				.setParameter("jrNo", jobRefNos).executeUpdate();
	}
	/**
	 * 
	 */
	@Transactional(propagation = Propagation.REQUIRED)
	public int deleteBSTempChildData (String jobRefNos) throws DAOException {
		//deleteDummyBSTempData Delete(header);
		return sessionFactory.getCurrentSession().getNamedQuery("deleteDummyBSTempChildData")
				.setParameter("jrNo", jobRefNos).executeUpdate();
	}

	@Transactional(propagation = Propagation.REQUIRED)
	public List<JctBeforeSketchQuestion> getQuestionnaireDummyEntry(
			String jobReferenceNos) throws DAOException {
		return sessionFactory.getCurrentSession().getNamedQuery("getDummyQuestionnaireData")
				.setParameter("jrNo", jobReferenceNos).list();
	}

	@Transactional(propagation = Propagation.REQUIRED)
	public int deleteQuestionnaireTempData(String jobRefNos)
			throws DAOException {
		//Delete(header);	
		return sessionFactory.getCurrentSession().getNamedQuery("deleteDummyQtnData")
				.setParameter("jrNo", jobRefNos).executeUpdate();
	}
	
	@Transactional(propagation = Propagation.REQUIRED)
	public List<JctAfterSketchHeader> getAfterSketchDummyEntry(
			String jobReferenceNos) throws DAOException {
		return sessionFactory.getCurrentSession().getNamedQuery("getAfterSketchDummyData")
				.setParameter("jrNo", jobReferenceNos).list();
	}
	/**
	 * 
	 * @param header
	 * @throws DAOException
	 */
	@Transactional(propagation = Propagation.REQUIRED)
	public int deleteASData (String jobRefNos) throws DAOException {
		return sessionFactory.getCurrentSession().getNamedQuery("deleteAllASData")
				.setParameter("jrNo", jobRefNos).executeUpdate();
	}
	/**
	 * 
	 * @param header
	 * @throws DAOException
	 */
	@Transactional(propagation = Propagation.REQUIRED)
	public int deleteASPageOneData (String jobRefNos) throws DAOException {
		return sessionFactory.getCurrentSession().getNamedQuery("deleteAllASPageOneData")
				.setParameter("jrNo", jobRefNos).executeUpdate();
	}
	/**
	 * 
	 * @param header
	 * @throws DAOException
	 */
	@Transactional(propagation = Propagation.REQUIRED)
	public int deleteASPageTwoData (String jobRefNos) throws DAOException {
		return sessionFactory.getCurrentSession().getNamedQuery("deleteAllASPageTwoData")
				.setParameter("jrNo", jobRefNos).executeUpdate();
	}
	
	
	
	/**
	 * 
	 * @param header
	 * @throws DAOException
	 */
	@Transactional(propagation = Propagation.REQUIRED)
	public int deleteASPageOneTempData (String jobRefNos) throws DAOException {
		return sessionFactory.getCurrentSession().getNamedQuery("deleteAllASPageOneTempData")
				.setParameter("jrNo", jobRefNos).executeUpdate();
	}
	/**
	 * 
	 * @param header
	 * @throws DAOException
	 */
	@Transactional(propagation = Propagation.REQUIRED)
	public int deleteASPageTwoTempData (String jobRefNos) throws DAOException {
		return sessionFactory.getCurrentSession().getNamedQuery("deleteAllASPageTwoTempData")
				.setParameter("jrNo", jobRefNos).executeUpdate();
	}

	@Transactional(propagation = Propagation.REQUIRED)
	public List<JctAfterSketchHeaderTemp> getAfterSketchDummyTempEntry(
			String jobReferenceNos) throws DAOException {
		return sessionFactory.getCurrentSession().getNamedQuery("getAfterSketchDummyTempData")
				.setParameter("jrNo", jobReferenceNos).list();
	}

	@Transactional(propagation = Propagation.REQUIRED)
	public int deleteASTempData(String jobRefNos)
			throws DAOException {
		return sessionFactory.getCurrentSession().getNamedQuery("deleteAllASTempData")
				.setParameter("jrNo", jobRefNos).executeUpdate();
	}
	/**
	 * 
	 */
	@Transactional(propagation = Propagation.REQUIRED)
	public List<JctActionPlan> getActionDummyEntry(String jobReferenceNos)
			throws DAOException {
		return sessionFactory.getCurrentSession().getNamedQuery("getActionPlanDummyData")
				.setParameter("jrNo", jobReferenceNos).list();
	}
	/**
	 * 
	 */
	@Transactional(propagation = Propagation.REQUIRED)
	public int deleteActionPlanTempData(String jobRefNos)
			throws DAOException {
		//Delete(plan);
		return sessionFactory.getCurrentSession().getNamedQuery("deleteAllTempActionPlanData")
				.setParameter("jrNo", jobRefNos).executeUpdate();
	}
	
	/**
	 * 
	 */
	@Transactional(propagation = Propagation.REQUIRED)
	public List<JctStatusSearch> getStatusSearchDummyEntry(String jobReferenceNos)
			throws DAOException {
		return sessionFactory.getCurrentSession().getNamedQuery("getStatusSearchDummyData")
				.setParameter("jrNo", jobReferenceNos).list();
	}
	/**
	 * 
	 */
	@Transactional(propagation = Propagation.REQUIRED)
	public int deleteStatusSearchDummyData(String jobRefNos)
			throws DAOException {
		return sessionFactory.getCurrentSession().getNamedQuery("deleteAllDummyStatusSearchData")
				.setParameter("jrNo", jobRefNos).executeUpdate();
	}

	@Transactional(propagation = Propagation.REQUIRED)
	public int deleteCompletionDummyData(String jobRefNo) throws DAOException {
		return sessionFactory.getCurrentSession().getNamedQuery("deleteAllDummyCompletedData")
				.setParameter("jrNo", jobRefNo).executeUpdate();
	}
	@Transactional(propagation = Propagation.REQUIRED)
	public int deleteDummyLoginInfoData(String jobRefNo) throws DAOException {
		return sessionFactory.getCurrentSession().getNamedQuery("deleteAllDummyLoginInfoData")
				.setParameter("jrNo", jobRefNo).executeUpdate();
	}

	@Transactional(propagation = Propagation.REQUIRED)
	public String getEmailIdByUserId(int userId) throws DAOException {
		return (String) sessionFactory.getCurrentSession().getNamedQuery("getEmailIdByUserId")
				.setParameter("userId", userId).uniqueResult();
	}

	/*@Transactional(propagation = Propagation.REQUIRED)
	public List<String> getTextSurveyQuestion() throws DAOException {
		return sessionFactory.getCurrentSession().getNamedQuery("getTextSurveyQuestion").list();
	}*/
	@Transactional(propagation = Propagation.REQUIRED)
	public List<Object[]> getTextSurveyQuestion() throws DAOException {
		return sessionFactory.getCurrentSession().getNamedQuery("getTextSurveyQuestion").list();
	}
	
	/*@Transactional(propagation = Propagation.REQUIRED)
	public List<String> getSurveyMainQuestion(int ansType) throws DAOException {
		return sessionFactory.getCurrentSession().getNamedQuery("getMainQuestionForMultiple")
				.setParameter("ansType", ansType).list();
	}*/
	@Transactional(propagation = Propagation.REQUIRED)
	public List<Object[]> getSurveyMainQuestion(int ansType) throws DAOException {
		return sessionFactory.getCurrentSession().getNamedQuery("getMainQuestionForMultiple")
				.setParameter("ansType", ansType).list();
	}

	@Transactional(propagation = Propagation.REQUIRED)
	public List<String> getAllSubQtns(int ansType, String mainQtn)
			throws DAOException {
		return sessionFactory.getCurrentSession().getNamedQuery("getSubQuestions")
				.setParameter("ansType", ansType)
				.setParameter("mainQtn", mainQtn).list();
	}

	@Transactional(propagation = Propagation.REQUIRED)
	public int getDaysToExpire(int userId) throws DAOException {
		java.math.BigInteger exp = new java.math.BigInteger("0");
		try {
		exp = (java.math.BigInteger)sessionFactory.getCurrentSession().createSQLQuery("SELECT DATEDIFF("
				+ "(select jct_account_expiration_date from jct_user where jct_user_id="+userId+"), "
				+ "(SELECT SYSDATE())) AS DiffDate").uniqueResult();
		} catch (Exception ex) {
			logger.error("UNABLE TO FETCH THE DAYS LEFT TO ACCOUNT EXPIRATION FOR USER ID: "+userId);
		}
		return exp.intValue(); 
	}

	@Transactional(propagation = Propagation.REQUIRED)
	public List<OccupationListDTO> searchOccupationList(String searchString)
			throws DAOException {
		return sessionFactory.getCurrentSession().getNamedQuery("getAllActiveOccupationList")
				.setParameter("searchString1", searchString)
				.setParameter("searchString2", "%"+searchString)
				.setParameter("searchString3", "%"+searchString+"%")
				.setParameter("searchString4", searchString+"%").list();
	}
	
	/*@Transactional(propagation = Propagation.REQUIRED)
	public List searchOccupationListNew(String searchString)
			throws DAOException {
		StringBuffer qrybuff = new StringBuffer("Select jct_onet_occupation_code, " +
				"jct_onet_occupation_title, jct_onet_occupation_description " +
				"from jct_onet_occupation_list where jct_onet_occupation_title REGEXP ");
		StringBuffer sb = new StringBuffer("");
		if (searchString.contains(",")) {
			String[] tok = searchString.split(",");
			int delimCount = 1;
			for (int index=0; index < tok.length; index++) {
				if (delimCount < tok.length) {
					sb.append((String) tok[index].trim());
					sb.append("|");
				} else {
					sb.append((String) tok[index].trim());
				}
				delimCount = delimCount + 1;
			}
		} else {
			sb.append(searchString);
		}
		String search = "'"+sb.toString()+"'";
		qrybuff.append(search);
		qrybuff.append(" OR jct_onet_occupation_description REGEXP ");
		qrybuff.append(search);
		qrybuff.append(" AND jct_onet_occupation_soft_delete = 0 ");
		
		
		return sessionFactory.getCurrentSession().createSQLQuery(qrybuff.toString()).list();
		}*/
	
	@Transactional(propagation = Propagation.REQUIRED)
	public List searchOccupationListNew(String searchString)
			throws DAOException {
		StringBuffer qrybuff = new StringBuffer("Select jct_onet_occupation_code, " +
				"jct_onet_occupation_title, jct_onet_occupation_description " +
				"from jct_onet_occupation_list where jct_onet_occupation_title REGEXP ");
		StringBuffer sb = new StringBuffer("");
		if (searchString.contains(",")) {
			String[] tok = searchString.split(",");
			int delimCount = 1;
			for (int index=0; index < tok.length; index++) {
				if (delimCount < tok.length) {
					sb.append((String) tok[index].trim());
					sb.append("|");
				} else {
					sb.append((String) tok[index].trim());
				}
				delimCount = delimCount + 1;
			}
		} else {
			sb.append(searchString);
		}
		String search = "'"+sb.toString()+"'";
		qrybuff.append(search);
		qrybuff.append(" OR jct_onet_occupation_description REGEXP ");
		qrybuff.append(search);
		qrybuff.append(" AND jct_onet_occupation_soft_delete = 0 ");
				
		return sessionFactory.getCurrentSession().createSQLQuery(qrybuff.toString()).list();
		}

	@Transactional(propagation = Propagation.REQUIRED)
	public List<String> getOnetDesc(String searchString) throws DAOException {
		return sessionFactory.getCurrentSession().getNamedQuery("getOnetDescList")
				.setParameter("searchString", searchString).list();
	}

	@Transactional(propagation = Propagation.REQUIRED)
	public int getUserId(int row) throws DAOException {
		return (Integer) sessionFactory.getCurrentSession().getNamedQuery("getUserIdForRow")
				.setParameter("rowId", row).uniqueResult();
	}
/*
 * Populate PopUp*/
	@Transactional(propagation = Propagation.REQUIRED)
	public JctPopupInstruction populatePopUp(int profileId,String relatedPage)
			throws DAOException {
		logger.info(">>>> AuthenticatorDAOImpl.populatePopUp");
		JctPopupInstruction object = null;
			String type = (String) sessionFactory.getCurrentSession()
				.getNamedQuery("fetchPopUpType")
				.setParameter("profileId", profileId)
				.setParameter("relatedPage", relatedPage).uniqueResult();
			
		if(type.equalsIgnoreCase("TEXT")){
			object = (JctPopupInstruction)sessionFactory.getCurrentSession()
			.getNamedQuery("fetchTextPopUpInstruction").setParameter("text", type)
			.setParameter("profileId", profileId)
			.setParameter("relatedPage", relatedPage)
			.uniqueResult();
			
		}
		else{
			object = (JctPopupInstruction)sessionFactory.getCurrentSession()
					.getNamedQuery("fetchVideoPopUpInstruction").setParameter("video", type)
					.setParameter("profileId", profileId)
					.setParameter("relatedPage", relatedPage)
					.uniqueResult();
			
		}
		
		logger.info("<<<< AuthenticatorDAOImpl.populateExistingInstruction");
		return object;	
		
	}

	@Transactional(propagation = Propagation.REQUIRED)
	public String getOnetCodeByDesc(String searchString) throws DAOException {
		return (String) sessionFactory.getCurrentSession().getNamedQuery("getOnetCodeByDesc")
				.setParameter("searchString", searchString).uniqueResult();
	}
	
	@Transactional(propagation = Propagation.REQUIRED)
	public String checkInstrList(int profileId,String relatedPage)
			throws DAOException {
		logger.info(">>>> AuthenticatorDAOImpl.populatePopUp");
		String status = "success";
		String type = (String) sessionFactory.getCurrentSession()
				.getNamedQuery("fetchPopUpType")
				.setParameter("profileId", profileId)
				.setParameter("relatedPage", relatedPage).uniqueResult();
		if(type != null){
			status = "success";
		}
		else{
			status = "failure";
		}
		return status;
	}
	/* Jira JCTP-35*/
	@Transactional(propagation = Propagation.REQUIRED)
	public String getMailedPassword(UserVO userVO) 
			throws DAOException{
		logger.info(">>>> AuthenticatorDAOImpl.getMailedPassword");
		return (String) sessionFactory.getCurrentSession().getNamedQuery("getMailedPassword")
				.setParameter("jctEmail", userVO.getEmail()).uniqueResult();
	}
	/* Bugzila JCTP-9045*/
	@Transactional(propagation = Propagation.REQUIRED)
	public String getInitialPassword(UserVO userVO) 
			throws DAOException{
		logger.info(">>>> AuthenticatorDAOImpl.getInitialPassword");
		return (String) sessionFactory.getCurrentSession().getNamedQuery("getInitialPassword")
				.setParameter("jctEmail", userVO.getEmail()).uniqueResult();
	}
	@Transactional(propagation = Propagation.REQUIRED)
	public String getTermsAndConditions(JctUserProfile userProfile, int userType) 
			throws DAOException{
		logger.info(">>>> AuthenticatorDAOImpl.getTermsAndConditions");
		return (String) sessionFactory.getCurrentSession().getNamedQuery("fetchTcByProfileId")
				.setParameter("userProfile", userProfile.getJctUserProfile())
				.setParameter("userType", userType)
				.uniqueResult();
	}

	@Transactional(propagation = Propagation.REQUIRED)
	public JctUserProfile getUserProfileByPk(int userProfileId) throws JCTException {
		
		logger.info(">>>> AuthenticatorDAOImpl.getUserProfileByPk");
		return (JctUserProfile) sessionFactory.getCurrentSession().getNamedQuery("fetchProfileIdByPk")
				.setParameter("userProfileId", userProfileId).uniqueResult();	
	}

	@Transactional(propagation=Propagation.REQUIRED)
	public List<OccupationListDTO> getOccupationListDto(String val) throws DAOException {
		return sessionFactory.getCurrentSession().getNamedQuery("getOnetCodeByTitle")
				.setParameter("searchString", val).list();
	}
	
/*	@Transactional(propagation=Propagation.REQUIRED)
	public List<OccupationListDTO> getOccupationListDto(String searchString) throws DAOException{
		return sessionFactory.getCurrentSession().getNamedQuery("fetchRegion").list();
	}*/
	
	@Transactional(propagation = Propagation.REQUIRED)
	public List<Object[]> getAllSurveyQuestion() throws DAOException {
		return sessionFactory.getCurrentSession().getNamedQuery("getAllSurveyQuestions").list();
	}

	@Transactional(propagation = Propagation.REQUIRED)
	public List<JctSurveyQuestions> alreadyRegisteredOnce(String emailId)
			throws DAOException {
		return sessionFactory.getCurrentSession().getNamedQuery("answerdQtns").setParameter("email", emailId).list();
	}

	@Transactional(propagation = Propagation.REQUIRED)
	public String silentUpdateTime(int rowIdentity) throws DAOException {
		logger.info(">>>> AuthenticatorDAOImpl.silentUpdateTime"); 
		String status = "failure";
		try{
			sessionFactory.getCurrentSession().getNamedQuery("updateUserInfoEndTime")
				.setParameter("endTime", new Date())
				.setParameter("rowId", rowIdentity).executeUpdate();
			status = "success";
		}catch(Exception ez){
			logger.error(ez.getLocalizedMessage());
		}
		logger.info("<<<< AuthenticatorDAOImpl.silentUpdateTime"); 
		return status;
	}
}
