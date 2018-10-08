package com.vmware.jct.service.impl;

import java.util.ArrayList;
import java.util.Date;
import java.util.Iterator;
import java.util.List;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;
import org.springframework.transaction.annotation.Propagation;
import org.springframework.transaction.annotation.Transactional;

import com.vmware.jct.common.utility.CommonConstants;
import com.vmware.jct.dao.IActionPlanDAO;
import com.vmware.jct.dao.IAfterSketchDAO;
import com.vmware.jct.dao.IAuthenticatorDAO;
import com.vmware.jct.dao.IBeforeSketchDAO;
import com.vmware.jct.dao.ICommonDao;
import com.vmware.jct.dao.IConfirmationDAO;
import com.vmware.jct.dao.ISketchSearchDAO;
import com.vmware.jct.dao.dto.DemographicDTO;
import com.vmware.jct.dao.dto.QuestionearDTO;
import com.vmware.jct.exception.DAOException;
import com.vmware.jct.exception.JCTException;
import com.vmware.jct.model.JctActionPlan;
import com.vmware.jct.model.JctAfterSketchFinalpageDetail;
import com.vmware.jct.model.JctAfterSketchFinalpageDetailTemp;
import com.vmware.jct.model.JctAfterSketchHeader;
import com.vmware.jct.model.JctAfterSketchHeaderTemp;
import com.vmware.jct.model.JctAfterSketchPageoneDetail;
import com.vmware.jct.model.JctAfterSketchPageoneDetailTemp;
import com.vmware.jct.model.JctBeforeSketchHeader;
import com.vmware.jct.model.JctBeforeSketchHeaderTemp;
import com.vmware.jct.model.JctCompletionStatus;
import com.vmware.jct.model.JctGlobalProfileHistory;
import com.vmware.jct.model.JctStatus;
import com.vmware.jct.model.JctStatusSearch;
import com.vmware.jct.model.JctUser;
import com.vmware.jct.model.JctUserLoginInfo;
import com.vmware.jct.service.IAfterSketchService;
import com.vmware.jct.service.vo.JctAfterSketchFinalPageVO;
import com.vmware.jct.service.vo.JctAfterSketchHeaderVO;
import com.vmware.jct.service.vo.JctAfterSketchPageOneVO;

/**
 * 
 * <p><b>Class name:</b> AfterSketchServiceImpl.java</p>
 * <p><b>Author:</b> Interra It</p>
 * <p><b>Purpose:</b> This is a AfterSketchServiceImpl class. This artifact is Business layer artifact.
 * AfterSketchServiceImpl implement IAfterSketchService interface and override the following  methods.
 * -saveOrUpdatePageOnde(JctAfterSketchHeaderVO header, List<JctAfterSketchPageOneVO> childList)
 * -fetchCheckedElements(String jobRefNo)
 * -fetchASOneJson(String jobRefNo)
 * </p>
 * <p><b>Description:</b> This class used to perform insert, modify ,fetch data from db and delete operation. Save, update, delete operations are performed in DataAccessObject class </p>
 * <p><b>Revision History:</b>
 * 	<li>Interra It - 19/Jan/2014 - Implement Exception </li>
 * <li>Interra It - 31/Jan/2014 - Implement comments, introduce Named query, Transactional attribute </li>
 * <li>Interra It - 11/Apr/2014 - Implement comments, introduce Named query, Transactional attribute </li>
 * </p>
 */
@Service
public class AfterSketchServiceImpl implements IAfterSketchService {
	
	@Autowired
	private IAfterSketchDAO afterSketchDAO;
	
	@Autowired
	private IBeforeSketchDAO beforeSketchDAO;
	
	@Autowired
	private IActionPlanDAO actionPlanDAO;
	
	@Autowired
	private ICommonDao commonDaoImpl;
	
	@Autowired
	private IConfirmationDAO confirmationDAO;
	
	@Autowired
	private IAuthenticatorDAO authenticatorDAO;
	
	@Autowired
	private ISketchSearchDAO iSketchSearchDAO;
	
	private static final Logger LOGGER = LoggerFactory.getLogger(AfterSketchServiceImpl.class);

	/**
	 * <p><b>Description:</b> Method saves or updates the job attributes (strength, value and passion).</p>
	 * @param JctAfterSketchHeaderVO
	 * @param List<JctAfterSketchPageOneVO>
	 * @return String
	 * @throws JCTException
	 */
	@Transactional(propagation=Propagation.REQUIRED)
	public String saveOrUpdatePageOnde(JctAfterSketchHeaderVO header,
			List<JctAfterSketchPageOneVO> childList, String linkClicked) throws JCTException {
		LOGGER.info(">>>>>> AfterSketchServiceImpl.saveOrUpdatePageOnde");
		/**
		 * PLOT: If the user selects new job attributes then delete all the existing 
		 * 		 job attributes and merge it with the existing ones. Else if for the 
		 * 		 first time then insert them. 		
		 */
		
		Date createdTs = new Date();
		String result = "failure";
		JctAfterSketchHeader parent = null;
		int totalTimeSpent = 0;
		int insertionSeq = 0;
		String jSon = "";
		List<JctAfterSketchPageoneDetail> childEntityList = null;
		try{
			List<JctAfterSketchHeader> list = afterSketchDAO.getList(header.getJobReferenceNo(), linkClicked);
			if(list.size() == 0) { //NEW
				parent = new JctAfterSketchHeader();
				/** THIS IS FOR PUBLIC VERSION **/
				JctUser user = new JctUser();
				user.setJctUserId(header.getJctUserId());
				parent.setJctUser(user);
				jSon = header.getJctPageOneJson();
				// Get the start time
				List<JctUserLoginInfo> infoList = afterSketchDAO.getLoginInfoList(header.getJobReferenceNo());
				JctUserLoginInfo info = (JctUserLoginInfo) infoList.get(infoList.size() - 1); // Get the latest
				Date headerDate = info.getJctAsCreatedTs();
				
				// Get the global profile changes during start time and end time
				List<JctGlobalProfileHistory> hist = afterSketchDAO.getGlobalProfileChangedData(headerDate);
				if (hist.size() > 0) {
					for (int index=0; index<hist.size(); index++) {
						JctGlobalProfileHistory obj = (JctGlobalProfileHistory) hist.get(index);
						String attrOriginalVal = obj.getJctGlobalProfileStrOriginal();
						String attrChanedVal = obj.getJctGlobalProfileStrChanged();
						if (jSon.contains("\"divValue\":\""+attrOriginalVal+"\"")) {
							jSon = jSon.replaceAll("\"divValue\":\""+attrOriginalVal+"\"", "\"divValue\":\""+attrChanedVal+"\"");
						}
					}
					parent.setJctAsPageneJson(jSon);
				} else {
					parent.setJctAsPageneJson(jSon);
				}
				/********************************/
				parent.setJctAsJobrefNo(header.getJobReferenceNo());
				parent.setJctAsStatusId(CommonConstants.TASK_PENDING);
				//parent.setJctAsPageneJson(header.getJctPageOneJson());
				parent.setJctAsCreatedBy(header.getJctCreatedBy());
				parent.setJctAsLastmodifiedTs(createdTs);
				parent.setJctAsLastmodifiedBy(header.getJctCreatedBy());
				parent.setJctAsLastmodifiedTs(createdTs);
				parent.setJctAsUserJobTitle(header.getJctJobTitle());
				parent.setJctAsPageoneCheckedValue(header.getJctPageOneCheckedElements());
				parent.setJctCheckedStrength(header.getCheckedStrength());
				parent.setJctCheckedPassion(header.getCheckedPassion());
				parent.setJctCheckedValue(header.getCheckedValue());
				parent.setJctAsPageoneTimeSpent(header.getJctPageOneTimeSpent());
				parent.setJctAsFinalpageInitialJsonValue(header.getJctFinalPageInitialJSonValue());
				parent.setJctAsInsertionSequence(insertionSeq+1);
				//Iterate and create the list of child
				childEntityList = new ArrayList<JctAfterSketchPageoneDetail>();
				Iterator childIt = childList.iterator();
				while(childIt.hasNext()){
					JctAfterSketchPageOneVO childVO = (JctAfterSketchPageOneVO)childIt.next();
					JctAfterSketchPageoneDetail child = new JctAfterSketchPageoneDetail();
					/*try{
						child.setJctAsPageoneDetailsId(commonDaoImpl.generateKey("jct_after_sketch_page_one_id"));
					} catch (DAOException e) {
						LOGGER.error("----"+e.getLocalizedMessage()+" ----");
						throw new JCTException(e.getMessage());
					}*/
					child.setJctAsCreatedBy(header.getJctCreatedBy());
					child.setJctAsCreatedTs(createdTs);
					child.setJctAsElementCode(childVO.getElementCode());
					child.setJctAsElementId(childVO.getElementId());
					child.setJctAsJobrefNo(childVO.getJobReferenceNo());
					child.setJctAsLastmodifiedBy(header.getJctCreatedBy());
					child.setJctAsLastmodifiedTs(createdTs);
					child.setJctAsPosition(childVO.getPosition());
					child.setJctAsSoftDelete(CommonConstants.ENTRY_NOT_DELETED);
					child.setPassionCount(childVO.getPassionCount());
					child.setValueCount(childVO.getValueCount());
					child.setStrengthCount(childVO.getStrengthCount());
					child.setJctAfterSketchHeader(parent);
					childEntityList.add(child);
				}
			} else { //EXISTING
				//Delete all the existing child entries
				List<JctAfterSketchPageoneDetail> childLst = afterSketchDAO.getASFirstPageList(header.getJobReferenceNo());
				Iterator it = childLst.iterator();
				while(it.hasNext()){
					afterSketchDAO.removePageOneDetails((JctAfterSketchPageoneDetail)it.next());
				}
				parent = (JctAfterSketchHeader)list.get(0);
				//We will not delete the existing row but will just add more values
				//parent.setJctAsPageneJson(header.getJctPageOneJson());
				
				jSon = header.getJctPageOneJson();
				// Get the start time
				List<JctUserLoginInfo> infoList = afterSketchDAO.getLoginInfoList(header.getJobReferenceNo());
				JctUserLoginInfo info = (JctUserLoginInfo) infoList.get(infoList.size() - 1); // Get the latest
				Date headerDate = info.getJctAsCreatedTs();
				
				// Get the global profile changes during start time and end time
				List<JctGlobalProfileHistory> hist = afterSketchDAO.getGlobalProfileChangedData(headerDate);
				if (hist.size() > 0) {
					for (int index=0; index<hist.size(); index++) {
						JctGlobalProfileHistory obj = (JctGlobalProfileHistory) hist.get(index);
						String attrOriginalVal = obj.getJctGlobalProfileStrOriginal();
						String attrChanedVal = obj.getJctGlobalProfileStrChanged();
						if (jSon.contains("\"divValue\":\""+attrOriginalVal+"\"")) {
							jSon = jSon.replaceAll("\"divValue\":\""+attrOriginalVal+"\"", "\"divValue\":\""+attrChanedVal+"\"");
						}
					}
					parent.setJctAsPageneJson(jSon);
				} else {
					parent.setJctAsPageneJson(jSon);
				}
				
				
				parent.setJctAsPageoneCheckedValue(header.getJctPageOneCheckedElements());
				parent.setJctCheckedStrength(header.getCheckedStrength());
				parent.setJctCheckedPassion(header.getCheckedPassion());
				parent.setJctCheckedValue(header.getCheckedValue());
				parent.setJctAsFinalpageInitialJsonValue(parent.getJctAsFinalpageInitialJsonValue());
				totalTimeSpent = header.getJctPageOneTimeSpent() + parent.getJctAsPageoneTimeSpent();
				parent.setJctAsPageoneTimeSpent(totalTimeSpent);
				//Iterate and create the list of child
				childEntityList = new ArrayList<JctAfterSketchPageoneDetail>();
				Iterator childIt = childList.iterator();
				while(childIt.hasNext()){
					JctAfterSketchPageOneVO childVO = (JctAfterSketchPageOneVO)childIt.next();
					JctAfterSketchPageoneDetail child = new JctAfterSketchPageoneDetail();
					/*try{
						child.setJctAsPageoneDetailsId(commonDaoImpl.generateKey("jct_after_sketch_page_one_id"));
					} catch (DAOException e) {
						LOGGER.error("----"+e.getLocalizedMessage()+" ----");
						throw new JCTException(e.getMessage());
					}*/
					child.setJctAsCreatedBy(header.getJctCreatedBy());
					child.setJctAsCreatedTs(createdTs);
					child.setJctAsElementCode(childVO.getElementCode());
					child.setJctAsElementId(childVO.getElementId());
					child.setJctAsJobrefNo(childVO.getJobReferenceNo());
					child.setJctAsLastmodifiedBy(header.getJctCreatedBy());
					child.setJctAsLastmodifiedTs(createdTs);
					child.setJctAsPosition(childVO.getPosition());
					child.setJctAsSoftDelete(CommonConstants.ENTRY_NOT_DELETED);
					child.setPassionCount(childVO.getPassionCount());
					child.setValueCount(childVO.getValueCount());
					child.setStrengthCount(childVO.getStrengthCount());
					child.setJctAfterSketchHeader(parent);
					childEntityList.add(child);
				}
			}
			parent.setJctAfterSketchPageoneDetails(childEntityList);
			result = afterSketchDAO.saveOrUpdatePageOne(parent);
			LOGGER.info("<<<<<< AfterSketchServiceImpl.saveOrUpdatePageOnde");
		} catch (DAOException daoException) {
			LOGGER.error("----"+daoException.getLocalizedMessage()+" ----");
			throw new JCTException(daoException.getMessage());
		}
		return jSon;
	}
	/**
	 * <p><b>Description:</b> Baserd on job reference number this method fetch page one check value from jct after sketch header table.</p>
	 * @param job reference nos
	 * @return String
	 * @throws JCTException
	 */
	@Transactional(propagation=Propagation.REQUIRED)
	public String fetchCheckedElements(String jobRefNo, String linkClicked) throws JCTException {
		LOGGER.info(">>>>>> AfterSketchServiceImpl.fetchCheckedElements");
		String checkedEle = "";
		try {
			checkedEle = afterSketchDAO.fetchCheckedElements(jobRefNo, linkClicked);
		} catch (DAOException daoException) {
			LOGGER.error("----"+daoException.getLocalizedMessage()+" ----");
			throw new JCTException(daoException.getMessage());
		}
		LOGGER.info("<<<<<< AfterSketchServiceImpl.fetchCheckedElements");
		return checkedEle;
	}
	/**
	 * <p><b>Description:</b>  Based on job reference number this method fetch page one Json value from jct after sketch header table</p>
	 * @param It will take jobRefNo as a input parameter
	 * @return It returns string object
	 * @throws JCTException -it throws DAOException
	 * 
	 */
	@Transactional(propagation=Propagation.REQUIRED)
	public String fetchASOneJson(String jobRefNo, String linkClicked) throws JCTException {
		LOGGER.info(">>>>>> AfterSketchServiceImpl.fetchASOneJson");
		String json = "";
		try {
			json = afterSketchDAO.fetchASOneJson(jobRefNo, linkClicked);
			LOGGER.info("<<<<<< AfterSketchServiceImpl.fetchASOneJson");
		} catch (DAOException daoException) {
			LOGGER.error("----"+daoException.getLocalizedMessage()+" ----");
			throw new JCTException(daoException.getMessage());
		}
		return json;
	}
	/**
	 * <p><b>Description:</b>  Based on job reference number this method fetch page one Json value from jct after sketch header table</p>
	 * @param It will take jobRefNo as a input parameter
	 * @return It returns string object
	 * @throws JCTException -it throws DAOException
	 * 
	 */
	@Transactional(propagation=Propagation.REQUIRED)
	public String fetchASOneJsonForEdit(String jobRefNo, String linkClicked) throws JCTException {
		LOGGER.info(">>>>>> AfterSketchServiceImpl.fetchASOneJsonForEdit");
		String json = "";
		try {
			json = afterSketchDAO.fetchASOneJsonForEdit(jobRefNo, linkClicked);
			LOGGER.info("<<<<<< AfterSketchServiceImpl.fetchASOneJsonForEdit");
		} catch (DAOException daoException) {
			LOGGER.error("----"+daoException.getLocalizedMessage()+" ----");
			throw new JCTException(daoException.getMessage());
		}
		return json;
	}
	/**
	 * <p><b>Description:</b>  Based on job reference number this method fetch final page Json value from jct after sketch header table</p>
	 * @param It will take jobRefNo as a input parameter
	 * @return It returns string object
	 * @throws JCTException -it throws DAOException
	 * 
	 */
	@Transactional(propagation=Propagation.REQUIRED)
	public String fetchASFinalPageJson(String jobRefNo, String linkClicked) throws JCTException {
		LOGGER.info(">>>>>> AfterSketchServiceImpl.fetchASFinalPageJson");
		String json = "";
		try {
			json = afterSketchDAO.fetchASFinalPageJson(jobRefNo, linkClicked);
			LOGGER.info("<<<<<< AfterSketchServiceImpl.fetchASFinalPageJson");
		} catch (DAOException daoException) {
			LOGGER.error("----"+daoException.getLocalizedMessage()+" ----");
			throw new JCTException(daoException.getMessage());
		}
		return json;
	}
	/**
	 * <p><b>Description:</b>  Based on job reference number this method fetch final page Json value from jct after sketch header table</p>
	 * @param It will take jobRefNo as a input parameter
	 * @return It returns string object
	 * @throws JCTException -it throws DAOException
	 * 
	 */
	@Transactional(propagation=Propagation.REQUIRED)
	public String fetchASFinalPageJsonForEdit(String jobRefNo, String linkClicked) throws JCTException {
		LOGGER.info(">>>>>> AfterSketchServiceImpl.fetchASFinalPageJsonForEdit");
		String json = "";
		try {
			json = afterSketchDAO.fetchASFinalPageJsonForEdit(jobRefNo, linkClicked);
			LOGGER.info("<<<<<< AfterSketchServiceImpl.fetchASFinalPageJsonForEdit");
		} catch (DAOException daoException) {
			LOGGER.error("----"+daoException.getLocalizedMessage()+" ----");
			throw new JCTException(daoException.getMessage());
		}
		return json;
	}
	/**
	 * <p><b>Description:</b>  Based on job ref no this method fetch after sketch final page's initial jon</p>
	 * @param It will take jobReferenceNos as a input parameter
	 * @return It returns string object
	 * @throws JCTException -it throws DAOException
	 * 
	 */
	@Transactional(propagation=Propagation.REQUIRED)
	public String getFinalPageInitialJson(String jobRefNo,String linkClicked) throws JCTException {
		LOGGER.info(">>>>>> AfterSketchServiceImpl.getFinalPageInitialJson");
		String json = "";
		try {
			json = afterSketchDAO.getFinalPageInitialJson(jobRefNo, linkClicked);
			LOGGER.info("<<<<<< AfterSketchServiceImpl.getFinalPageInitialJson");
		} catch (DAOException daoException) {
			LOGGER.error("----"+daoException.getLocalizedMessage()+" ----");
			throw new JCTException(daoException.getMessage());
		}
		return json;
	}
	/**
	 * <p><b>Description:</b>  Based on job ref no this method fetch after sketch final page's initial jon</p>
	 * @param It will take jobReferenceNos as a input parameter
	 * @return It returns string object
	 * @throws JCTException -it throws DAOException
	 * 
	 */
	@Transactional(propagation=Propagation.REQUIRED)
	public String getFinalPageInitialJsonForEdit(String jobRefNo,String linkClicked) throws JCTException {
		LOGGER.info(">>>>>> AfterSketchServiceImpl.getFinalPageInitialJson");
		String json = "";
		try {
			json = afterSketchDAO.getFinalPageInitialJsonForEdit(jobRefNo, linkClicked);
			LOGGER.info("<<<<<< AfterSketchServiceImpl.getFinalPageInitialJson");
		} catch (DAOException daoException) {
			LOGGER.error("----"+daoException.getLocalizedMessage()+" ----");
			throw new JCTException(daoException.getMessage());
		}
		return json;
	}
	/**
	 * <p><b>Description:</b>  Based on job reference number this method fetches fetches List of passion count</p>
	 * @param It will take jobRefNo as a input parameter
	 * @return It returns Integer object
	 * @throws JCTException -it throws DAOException
	 * 
	 */
	@Transactional(propagation=Propagation.REQUIRED)
	public int getPassionCount(String jobRefNo) throws JCTException{
		LOGGER.info(">>>>>> AfterSketchServiceImpl.getPassionCount");
		int count = 0;
		try {
			List<Integer> countList = afterSketchDAO.getPassionCount(jobRefNo);
			if(countList.size() > 0){
				count = countList.get(0);
			}
			LOGGER.info("<<<<<< AfterSketchServiceImpl.getPassionCount");
		} catch (DAOException daoException) {
			LOGGER.error("----"+daoException.getLocalizedMessage()+" ----");
			throw new JCTException(daoException.getMessage());
		}
		return count;
	}
	/**
	 * <p><b>Description:</b>  Based on job reference number this method fetches List of value count</p>
	 * @param It will take jobRefNo as a input parameter
	 * @return It returns Integer object
	 * @throws JCTException -it throws DAOException
	 * 
	 */
	@Transactional(propagation=Propagation.REQUIRED)
	public int getValueCount(String jobRefNo) throws JCTException{
		LOGGER.info(">>>>>> AfterSketchServiceImpl.getValueCount");
		int count = 0;
		try {
			List<Integer> countList = afterSketchDAO.getValueCount(jobRefNo);
			if(countList.size() > 0){
				count = countList.get(0);
			}
			LOGGER.info("<<<<<< AfterSketchServiceImpl.getValueCount");
		} catch (DAOException daoException) {
			LOGGER.error("----"+daoException.getLocalizedMessage()+" ----");
			throw new JCTException(daoException.getMessage());
		}
		return count;
	}
	/**
	 * <p><b>Description:</b>  Based on job reference number this method fetches List of strength count</p>
	 * @param It will take jobRefNo as a input parameter
	 * @return It returns Integer object
	 * @throws JCTException -it throws DAOException
	 * 
	 */
	@Transactional(propagation=Propagation.REQUIRED)
	public int getStrengthCount(String jobRefNo) throws JCTException{
		LOGGER.info(">>>>>> AfterSketchServiceImpl.getStrengthCount");
		int count = 0;
		try {
			List<Integer> countList = afterSketchDAO.getStrengthCount(jobRefNo);
			if(countList.size() > 0){
				count = countList.get(0);
			}
			LOGGER.info("<<<<<< AfterSketchServiceImpl.getStrengthCount");
		} catch (DAOException daoException) {
			LOGGER.error("----"+daoException.getLocalizedMessage()+" ----");
			throw new JCTException(daoException.getMessage());
		}
		return count;
	}
	/**
	 * <p><b>Description:</b> This method is used for update existing After Sketch header data and insert new 
	 *final page rows</p>
	 * @param It will take JctAfterSketchHeader as a input parameter
	 * @param It will take List<JctAfterSketchFinalPageVO> as a input parameter
	 * @param It will take isCompleted as a input parameter
	 * @return It returns String object
	 * @throws JCTException -it throws JCTException
	 * 
	 */
	@Transactional(propagation=Propagation.REQUIRED)
	public String saveAfterSketchFinalPage(JctAfterSketchHeaderVO header, 
			List<JctAfterSketchFinalPageVO> childList, String isCompleted, 
			String isEdit, String linkClicked) 
			throws JCTException {
		LOGGER.info(">>>>>> AfterSketchServiceImpl.saveAfterSketchFinalPage");
		String result = "";
		try {
			if (isEdit.equals("Y")) {
				result = saveASTemp(header, childList, linkClicked);
			} else {
				JctAfterSketchHeader parent = null;
				JctAfterSketchHeader parentNew = new JctAfterSketchHeader();;
				Date lastTs = new Date();
				//Fetch the header object
				LOGGER.info("ASIMPL----job reference number: "+header.getJobReferenceNo()+"..");
				List<JctAfterSketchHeader> list = afterSketchDAO.getList(header.getJobReferenceNo(), linkClicked);
				LOGGER.info("ASIMPL----SIZE FETCHED: "+list.size()+"..");
				if (list.size() == 0) {
					list = afterSketchDAO.getListForEdit(header.getJobReferenceNo(), linkClicked);
				}
				parent = (JctAfterSketchHeader)list.get(0);
				
				//Update the existing row [parent and childs.]
				//int prevStatusId = parent.getJctAsStatusId();
				int timeSpent = parent.getJctAsFinalpageTimeSpent() + header.getJctFinalPageTimeSpent();
				parent.setJctAsSoftDelete(CommonConstants.ENTRY_SOFT_DELETED);
				parent.setJctAsFinalpageJson(header.getJctFinalPageJson());
				parent.setJctAsFinalpageSnapshot(header.getJctFinalPageSnapshot());
				parent.setJctAsFinalpageSnapshotString(header.getJctFinalPageSnapshotString());
				parent.setJctFinalPageDivHeight(header.getDivHeight());
				parent.setJctFinalPageDivWidth(header.getDivWidth());
				parent.setJctAsTotalCount(header.getAsTotalCount());
				parent.setJctAsFinalpageInitialJsonValue(header.getJctFinalPageInitialJSonValue());
				parent.setJctAsFinalpageTimeSpent(timeSpent);
				
				List<JctAfterSketchFinalpageDetail> details = parent.getJctAfterSketchFinalpageDetails();
				if (details.size() > 0) {
					Iterator detailsItr = details.iterator();
					while (detailsItr.hasNext()) {
						JctAfterSketchFinalpageDetail detailsObj = (JctAfterSketchFinalpageDetail) detailsItr.next();
						detailsObj.setJctAsSoftDelete(CommonConstants.ENTRY_SOFT_DELETED);
						detailsObj.setJctAfterSketchHeader(parent);
					}
				}
				List<JctAfterSketchPageoneDetail> modifiedpageOneList = new ArrayList<JctAfterSketchPageoneDetail> ();
				List<JctAfterSketchPageoneDetail> pageOne = parent.getJctAfterSketchPageoneDetails();
				Iterator<JctAfterSketchPageoneDetail> pageOneItr = pageOne.iterator();
				while (pageOneItr.hasNext()) {
					JctAfterSketchPageoneDetail detailObj = (JctAfterSketchPageoneDetail) pageOneItr.next();
					JctAfterSketchPageoneDetail child = new JctAfterSketchPageoneDetail();
					//child.setJctAsPageoneDetailsId(commonDaoImpl.generateKey("jct_after_sketch_page_one_id"));
					child.setJctAsCreatedBy(detailObj.getJctAsCreatedBy());
					child.setJctAsCreatedTs(lastTs);
					child.setJctAsElementCode(detailObj.getJctAsElementCode());
					child.setJctAsElementId(detailObj.getJctAsElementId());
					child.setJctAsJobrefNo(detailObj.getJctAsJobrefNo());
					child.setJctAsLastmodifiedBy(detailObj.getJctAsLastmodifiedBy());
					child.setJctAsLastmodifiedTs(lastTs);
					child.setJctAsPosition(detailObj.getJctAsPosition());
					child.setJctAsSoftDelete(CommonConstants.ENTRY_NOT_DELETED);
					child.setPassionCount(detailObj.getPassionCount());
					child.setValueCount(detailObj.getValueCount());
					child.setStrengthCount(detailObj.getStrengthCount());
					child.setJctAfterSketchHeader(parentNew);
					modifiedpageOneList.add(child);
					
					detailObj.setJctAsSoftDelete(CommonConstants.ENTRY_SOFT_DELETED);
					detailObj.setJctAfterSketchHeader(parent);
				}
				result = afterSketchDAO.mergeAfterSketchDiag(parent);
				LOGGER.info("RESULT OF INITIAL UPDATE: "+result);
				
				//Insert all new rows..
				
				//parentNew.setJctAsHeaderId(commonDaoImpl.generateKey("jct_after_sketch_header_id"));
				/** THIS IS FOR PUBLIC VERSION **/
				JctUser user = new JctUser();
				user.setJctUserId(header.getJctUserId());
				parentNew.setJctUser(user);
				/********************************/
				parentNew.setJctAsJobrefNo(header.getJobReferenceNo());
				parentNew.setJctAsStatusId(CommonConstants.TASK_PENDING);
				parentNew.setJctAsPageneJson(parent.getJctAsPageneJson());
				parentNew.setJctAsCreatedBy(parent.getJctAsCreatedBy());
				parentNew.setJctAsLastmodifiedTs(lastTs);
				parentNew.setJctAsLastmodifiedBy(parent.getJctAsLastmodifiedBy());
				parentNew.setJctAsLastmodifiedTs(parent.getJctAsLastmodifiedTs());
				parentNew.setJctAsUserJobTitle(parent.getJctAsUserJobTitle());
				parentNew.setJctAsPageoneCheckedValue(parent.getJctAsPageoneCheckedValue());
				parentNew.setJctCheckedStrength(parent.getJctCheckedStrength());
				parentNew.setJctCheckedPassion(parent.getJctCheckedPassion());
				parentNew.setJctCheckedValue(parent.getJctCheckedValue());
				parentNew.setJctAsPageoneTimeSpent(parent.getJctAsPageoneTimeSpent());
				parentNew.setJctAsFinalpageInitialJsonValue(parent.getJctAsFinalpageInitialJsonValue());
				parentNew.setJctAsInsertionSequence(parent.getJctAsInsertionSequence()+1);
				parentNew.setJctAsFinalpageJson(header.getJctFinalPageJson());
				parentNew.setJctAsFinalpageSnapshot(header.getJctFinalPageSnapshot());
				parentNew.setJctAsFinalpageSnapshotString(header.getJctFinalPageSnapshotString());
				parentNew.setJctFinalPageDivHeight(header.getDivHeight());
				parentNew.setJctFinalPageDivWidth(header.getDivWidth());
				parentNew.setJctAsTotalCount(header.getAsTotalCount());
				parentNew.setJctAsFinalpageTimeSpent(timeSpent);
				List<JctAfterSketchFinalpageDetail> modifiedChildList = new ArrayList<JctAfterSketchFinalpageDetail> ();
				Iterator<JctAfterSketchFinalPageVO> childIt = childList.iterator();
				while(childIt.hasNext()){
					JctAfterSketchFinalPageVO childVO = (JctAfterSketchFinalPageVO)childIt.next();
					JctAfterSketchFinalpageDetail child = new JctAfterSketchFinalpageDetail();
					//child.setJctAsFinalpageDetailsId(commonDaoImpl.generateKey("jct_after_sketch_final_page_id"));
					child.setJctAsJobrefNo(childVO.getJobRefNo());
					child.setJctAsElementCode(childVO.getElementCode());
					child.setJctAsElementDesc(childVO.getElementDesc());
					child.setJctAsPosition(childVO.getPosition());
					child.setJctAsRoleDesc(childVO.getRoleDescription());
					child.setJctAsTaskEnergy(childVO.getTaskEnergy());
					child.setJctAsTaskId(childVO.getTaskId());
					child.setJctAsTaskDesc(childVO.getTaskDescription());
					child.setJctAsStatusId(CommonConstants.TASK_PENDING);
					child.setJctAsSoftDelete(CommonConstants.ENTRY_NOT_DELETED);
					child.setJctAsCreatedTs(new Date());
					child.setJctAsCreatedBy(parent.getJctAsCreatedBy());
					child.setJctAsLastmodifiedBy(parent.getJctAsCreatedBy());
					child.setJctAsLastmodifiedTs(new Date());
					child.setJctAsAdditionalDesc(childVO.getAdditionalDesc());
					child.setJctAfterSketchHeader(parentNew);
					modifiedChildList.add(child);
				}
				parentNew.setJctAfterSketchFinalpageDetails(modifiedChildList);
				parentNew.setJctAfterSketchPageoneDetails(modifiedpageOneList);
				result = afterSketchDAO.mergeAfterSketchDiag(parentNew);
				LOGGER.info("RESULT OF FINAL INSERT UPDATE: "+result);
				
				//Call the function to find the difference
				Integer returnVal = afterSketchDAO.calculateDifference(header.getJobReferenceNo());
					
				/*//Update the login info table for changing the completed status to 1 if not logged out
				if(isCompleted.equals("N")){
					afterSketchDAO.changeStatus(header.getJobReferenceNo());
				}*/
				LOGGER.info("<<<<<< AfterSketchServiceImpl.saveAfterSketchFinalPage");
				
			} 
		} catch (DAOException daoException) {
			LOGGER.error("----"+daoException.getLocalizedMessage()+" ----");
			throw new JCTException(daoException.getMessage());
		}
		return result;
	}
	
	private String saveASTemp(JctAfterSketchHeaderVO header,
			List<JctAfterSketchFinalPageVO> childList, String linkClicked) {LOGGER.info(">>>>>> AfterSketchServiceImpl.saveASTemp");
			String result = "failure";
			JctAfterSketchHeader parent = null;
			JctAfterSketchHeaderTemp parentTemp = null;
			List<JctAfterSketchPageoneDetailTemp> pageOneTempList = new ArrayList<JctAfterSketchPageoneDetailTemp> ();
			List<JctAfterSketchFinalpageDetailTemp> finalPageTempList = new ArrayList<JctAfterSketchFinalpageDetailTemp> ();
			try {
				List<JctAfterSketchHeader> list = afterSketchDAO.getListForEdit(header.getJobReferenceNo(), "Y");
				
				//This is because once you click the link, the status becomes 4 , 0
				if (list.size() == 0) {
					list = afterSketchDAO.getList(header.getJobReferenceNo(), "Y");
				}
				
				parent = (JctAfterSketchHeader)list.get(0);
				//Iterate over the parent and child elements, copy them as backup
				parentTemp = new JctAfterSketchHeaderTemp();
				/** THIS IS FOR PUBLIC VERSION **/
				JctUser user = new JctUser();
				user.setJctUserId(header.getJctUserId());
				parentTemp.setJctUser(user);
				/********************************/
				parentTemp.setJctAsCreatedBy(parent.getJctAsCreatedBy());
				parentTemp.setJctAsCreatedTs(parent.getJctAsCreatedTs());
				parentTemp.setJctAsFinalpageJson(header.getJctFinalPageJson());
				parentTemp.setJctAsFinalpageSnapshot(header.getJctFinalPageSnapshot());
				parentTemp.setJctAsFinalpageSnapshotString(header.getJctFinalPageSnapshotString());
				parentTemp.setJctFinalPageDivHeight(header.getDivHeight());
				parentTemp.setJctFinalPageDivWidth(header.getDivWidth());
				parentTemp.setJctAsFinalpageInitialJsonValue(header.getJctFinalPageInitialJSonValue());
				parentTemp.setJctAsFinalpageTimeSpent(parent.getJctAsFinalpageTimeSpent());
				parentTemp.setJctAsJobrefNo(parent.getJctAsJobrefNo());
				parentTemp.setJctAsLastmodifiedBy(parent.getJctAsLastmodifiedBy());
				parentTemp.setJctAsLastmodifiedTs(new Date());
				parentTemp.setJctAsPageneJson(parent.getJctAsPageneJson());
				parentTemp.setJctAsPageoneCheckedValue(parent.getJctAsPageoneCheckedValue());
				parentTemp.setJctAsPageoneSnapshot(parent.getJctAsPageoneSnapshot());
				parentTemp.setJctAsPageoneSnapshotString(parent.getJctAsPageoneSnapshotString());
				parentTemp.setJctAsPageoneTimeSpent(parent.getJctAsPageoneTimeSpent());
				parentTemp.setJctAsSoftDelete(CommonConstants.ENTRY_NOT_DELETED);
				parentTemp.setJctAsStatusId(CommonConstants.TASK_PENDING);
				parentTemp.setJctAsTotalCount(header.getAsTotalCount());
				parentTemp.setJctAsUserJobTitle(parent.getJctAsUserJobTitle());
				parentTemp.setJctCheckedPassion(parent.getJctCheckedPassion());
				parentTemp.setJctCheckedStrength(parent.getJctCheckedStrength());
				parentTemp.setJctCheckedValue(parent.getJctCheckedValue());
				
				Iterator pageOneItr = parent.getJctAfterSketchPageoneDetails().iterator();
				while (pageOneItr.hasNext()) {
					JctAfterSketchPageoneDetail obj = (JctAfterSketchPageoneDetail) pageOneItr.next();
					JctAfterSketchPageoneDetailTemp tempObj = new JctAfterSketchPageoneDetailTemp();
					tempObj.setJctAsCreatedBy(obj.getJctAsCreatedBy());
					tempObj.setJctAsCreatedTs(obj.getJctAsCreatedTs());
					tempObj.setJctAsElementCode(obj.getJctAsElementCode());
					tempObj.setJctAsElementId(obj.getJctAsElementId());
					tempObj.setJctAsJobrefNo(obj.getJctAsJobrefNo());
					tempObj.setJctAsLastmodifiedBy(obj.getJctAsLastmodifiedBy());
					tempObj.setJctAsLastmodifiedTs(new Date());
					tempObj.setJctAsPosition(obj.getJctAsPosition());
					tempObj.setJctAsSoftDelete(CommonConstants.ENTRY_NOT_DELETED);
					tempObj.setPassionCount(obj.getPassionCount());
					tempObj.setStrengthCount(obj.getStrengthCount());
					tempObj.setValueCount(obj.getValueCount());
					tempObj.setJctAfterSketchHeaderTemp(parentTemp);
					pageOneTempList.add(tempObj);
				}
				parentTemp.setJctAfterSketchPageoneDetails(pageOneTempList);
				
				Iterator pageTwoItr = childList.iterator();
				StringBuilder sb = null;
				while (pageTwoItr.hasNext()) {
					sb = new StringBuilder("");
					JctAfterSketchFinalPageVO childVO = (JctAfterSketchFinalPageVO) pageTwoItr.next();
					JctAfterSketchFinalpageDetailTemp child = new JctAfterSketchFinalpageDetailTemp();
					child.setJctAsJobrefNo(childVO.getJobRefNo());
					child.setJctAsElementCode(childVO.getElementCode());
					child.setJctAsElementDesc(childVO.getElementDesc());
					child.setJctAsPosition(childVO.getPosition());
					child.setJctAsRoleDesc(childVO.getRoleDescription());
					child.setJctAsTaskEnergy(childVO.getTaskEnergy());
					child.setJctAsTaskId(childVO.getTaskId());
					child.setJctAsTaskDesc(childVO.getTaskDescription());
					child.setJctAsStatusId(CommonConstants.TASK_PENDING);
					child.setJctAsSoftDelete(CommonConstants.ENTRY_NOT_DELETED);
					child.setJctAsCreatedTs(new Date());
					child.setJctAsCreatedBy(parent.getJctAsCreatedBy());
					child.setJctAsLastmodifiedBy(parent.getJctAsCreatedBy());
					child.setJctAsLastmodifiedTs(new Date());
					child.setJctAsAdditionalDesc(childVO.getAdditionalDesc());
					child.setJctAfterSketchHeaderTemp(parentTemp);
					finalPageTempList.add(child);
					sb.append(parent.getJctAsCreatedBy());
				}
				parentTemp.setJctAfterSketchFinalpageDetails(finalPageTempList);
				
				//Remove all existing in temp tables.
				List<JctAfterSketchHeaderTemp> tempList = afterSketchDAO.getListTemp(header.getJobReferenceNo());
				if (tempList.size() > 0) {
					Iterator<JctAfterSketchHeaderTemp> itr = tempList.iterator();
					while (itr.hasNext()) {
						afterSketchDAO.removeTemp((JctAfterSketchHeaderTemp) itr.next());
					}
				}
				//afterSketchDAO.removeTemp((JctAfterSketchHeaderTemp) afterSketchDAO.getListTemp(header.getJobReferenceNo()).get(0));
				
				result = afterSketchDAO.saveOrUpdateAsHeaderTemp(parentTemp);
				
				// Store it to the search table
				String outcome = storeToSearchTable(parentTemp, sb.toString());
				
				// Store in jct completion status table
				JctCompletionStatus compStatus = new JctCompletionStatus();
				compStatus.setJctJobReferenceNo(parent.getJctAsJobrefNo());
				compStatus.setJctEndDate(new Date());
				compStatus.setJctStartDate(new Date());
				compStatus.setJctOptionSelected(2);
				confirmationDAO.saveOrUpdateBeforeSketch(compStatus);
				LOGGER.info("STORE TO SEARCH TABLE: "+outcome);
				
				LOGGER.info("<<<<<< AfterSketchServiceImpl.saveAfterSketchFinalPage");
			} catch (DAOException daoException) {
				LOGGER.error("----"+daoException.getLocalizedMessage()+" ----");
			}
			return result;
		}
	private String storeToSearchTable(JctAfterSketchHeaderTemp parentTemp, String emailId) {
		String outcome = "failure";
		try {
			// Update all soft delete to 1 for the esisting ones..
			confirmationDAO.updateDeleteStatusForEdit(parentTemp.getJctAsJobrefNo());
			
			JctStatusSearch search = new JctStatusSearch();
			//search.setJctStatusSearchId(commonDaoImpl.generateKey("jct_status_id"));
			search.setJctAsSnapshot(parentTemp.getJctAsFinalpageSnapshot());
			search.setJctJobrefNo(parentTemp.getJctAsJobrefNo());
			search.setJctAsSnapshotString(parentTemp.getJctAsFinalpageSnapshotString());
			JctStatus status = new JctStatus();
			status.setJctStatusId(5);
			search.setJctStatusId(status);
			search.setJctSoftDelete(CommonConstants.ENTRY_NOT_DELETED);
			search.setJctCreatedBy(parentTemp.getJctAsCreatedBy());
			search.setJctCreatedTs(new Date());
			search.setJctLastmodifiedBy(parentTemp.getJctAsCreatedBy());
			search.setJctLastmodifiedTs(new Date());
			search.setJctIsInactive(0);
			search.setJctUserLevels(parentTemp.getJctAsUserJobTitle());
			
			// Fetch the occupation
			List<DemographicDTO> dto = authenticatorDAO.getDemographicInformation(emailId);
			DemographicDTO occupationDto = dto.get(dto.size()-1);
			search.setOccupationCode(occupationDto.getOnetOccupationData());
			
			search.setJctStatusDecision(1); // 1- When shared after edit radio option else 0
			//get snapshot string from temp table
			JctBeforeSketchHeaderTemp bsTemp = beforeSketchDAO.getListTemp(parentTemp.getJctAsJobrefNo()).get(0);
			search.setJctBsSnapshot(bsTemp.getJctBsSnapshot());
			search.setJctBsSnapshotString(bsTemp.getJctBsSnapshotString());
			confirmationDAO.saveToStatusSearch(search);
			
			// Update the flag to 1 in pdf recods table
			confirmationDAO.updateShowPDF(search.getJctJobrefNo());		
			
			
			outcome = "success";
		} catch (DAOException ex) {
			LOGGER.error("------------------------SHOULD HAVE INSERTED VALUES IN STATUS SEARCH DATABASE ----------------- "+ex.getLocalizedMessage());
		}
		return outcome;
	}
	/**
	 * <p><b>Description:</b> This method is used for fetch the action plan questions, sub questions and 
	 * answers if answered based on job reference nos and profile id</p>
	 * @param It will take job ref no as a input parameter
	 * @param It will take job profile id as a input parameter
	 * @return It returns List<List>
	 * @throws JCTException -it throws JCTException
	 * 
	 */
	@Transactional(propagation=Propagation.REQUIRED)
	public List<List> fetActionPlanDynamic (String jobRefNo, int profileId) throws JCTException {
		//This will be returned
		List<List> mainList = null;
		try{
			List<JctBeforeSketchHeader> list = beforeSketchDAO.getList(jobRefNo);
			if (list.size() == 0) {
				list = beforeSketchDAO.getListEdtd(jobRefNo);
			}
			if(list.size()!=0){
				JctBeforeSketchHeader parent = (JctBeforeSketchHeader)list.get(0);
				LOGGER.debug("Test ::"+parent.getJctBsUserJobTitle());
				//Check if new or existing
				List<JctActionPlan> actionPlanlist = actionPlanDAO.getList(jobRefNo);
				if (actionPlanlist.size() == 0) {
					actionPlanlist = actionPlanDAO.getListEdited(jobRefNo);
				}
				if(actionPlanlist.size() > 0){
					mainList = new ArrayList<List>();
					// [[Main Qtn1 subQtn Ans subQtn Ans] [Main Qtn2 subQtn Ans subQtn Ans] .....]
					// Select Distinct Main question List
					List<String> actionPlanMainQtnLst = actionPlanDAO.getDistinctQuestionList(jobRefNo);
					Iterator<String> actionPlanMainQtnLstItr = actionPlanMainQtnLst.iterator();
					while (actionPlanMainQtnLstItr.hasNext()) {
						List<String> innerList = new ArrayList<String>();
						String mainQuestion = (String) actionPlanMainQtnLstItr.next();
						innerList.add(mainQuestion);
						// Select the sub question List 
						List<String> subQtnList = actionPlanDAO.getDistinctSubQuestionListByJrno(mainQuestion, jobRefNo);
						
						Iterator<String> subQtnListItr = subQtnList.iterator();
						while (subQtnListItr.hasNext()) {
							String subQuestion = (String) subQtnListItr.next();
							innerList.add(subQuestion+"~");
							//Fetch answers given
							List<String> ansList = actionPlanDAO.getAnswerListByJrno(mainQuestion, subQuestion, jobRefNo);
							innerList.add((String) ansList.get(0));
							innerList.add((String) ansList.get(1));
							innerList.add((String) ansList.get(2));	
						}
						mainList.add(innerList);
					}
				} else { //new
					// [[Main Qtn1 subQtn Ans subQtn Ans] [Main Qtn2 subQtn Ans subQtn Ans] .....]
					mainList = new ArrayList<List>();
					// Select Distinct Main question List
					List<String> actionPlanMainQtnList = afterSketchDAO.getActionPlanMainQtnList(profileId);
					Iterator<String> actionPlanMainQtnLstItr = actionPlanMainQtnList.iterator();
					while (actionPlanMainQtnLstItr.hasNext()) {
						List<String> innerList = new ArrayList<String>();
						String mainQuestion = (String) actionPlanMainQtnLstItr.next();
						innerList.add(mainQuestion);
						// Select the sub question List
						List<String> subQtnList = afterSketchDAO.getActionPlanSubQtnList(mainQuestion, profileId);
						Iterator<String> subQtnListItr = subQtnList.iterator();
						while (subQtnListItr.hasNext()) {
							innerList.add((String) subQtnListItr.next()+"~");
							innerList.add(""); //This is answer. This is set to blank as this it new
							innerList.add(""); //
							innerList.add(""); //
						}
						mainList.add(innerList);
					}
				}
			} else {
				LOGGER.error("------> SOMETHING IS MISSING IN DATABASE: JctBeforeSketchHeader. NO RECORD FOUND WITH GIVEN JOB REFERENCE NUMBER: "+jobRefNo);
			}
		} catch(DAOException daoException){
			LOGGER.error("----"+daoException.getLocalizedMessage()+" ----");
		}
		return mainList;
	}
	/**
	 * <p><b>Description:</b> This method is used for fetch list of action plan string based on job ref nos 
	 * and profile id</p>
	 * @param It will take job ref nos as a input parameter
	 * @param It will take profile as a input parameter
	 * @return It returns List<String>
	 * @throws JCTException -it throws JCTException
	 * 
	 */
	@Transactional(propagation=Propagation.REQUIRED)
	public List<String> fetchActionPlans(String jobRefNo, int profileId) throws JCTException {
		LOGGER.info(">>>>>> AfterSketchServiceImpl.fetchActionPlans");
		List<JctBeforeSketchHeader> list=null;
		JctBeforeSketchHeader parent = null;
		List<QuestionearDTO> jctQuestionList=null;
		StringBuilder stringBuilder = null;
		List<String> actionPlanList = new ArrayList<String>(); 
		try{
			list = beforeSketchDAO.getList(jobRefNo);
			if(list.size()!=0){
				parent = (JctBeforeSketchHeader)list.get(0);
				LOGGER.debug("Test ::"+parent.getJctBsUserJobTitle());
				//Check if new or existing
				List<JctActionPlan> actionPlanlist = actionPlanDAO.getList(jobRefNo);
				if(actionPlanlist.size() > 0){ //Existing
					Iterator itr = actionPlanlist.iterator();
					int loopIndex = 0;
					while(itr.hasNext()){
						JctActionPlan plan = (JctActionPlan)itr.next();
						stringBuilder = new StringBuilder("");
						if(loopIndex == 0){
							stringBuilder.append(plan.getJctAsQuestionDesc());
							stringBuilder.append("~");
						}
						stringBuilder.append(plan.getJctAsQuestionSubDesc());
						stringBuilder.append("~");
						stringBuilder.append(plan.getJctAsAnswarDesc());
						actionPlanList.add(stringBuilder.toString()+"!||!");
						loopIndex = loopIndex + 1;
					}
					actionPlanList.add("E");
				} else { //new
					jctQuestionList=afterSketchDAO.getFetchQuestion(profileId);
					Iterator itr = jctQuestionList.iterator();
					int loopIndex = 0;
					while(itr.hasNext()){
						QuestionearDTO dto = (QuestionearDTO)itr.next();
						stringBuilder = new StringBuilder("");
						if(loopIndex == 0){
							stringBuilder.append(dto.getJctQuestionsDesc());
							stringBuilder.append("~");
						}
						stringBuilder.append(dto.getJctQuestionSubDesc());
						stringBuilder.append("~");
						stringBuilder.append("");
						actionPlanList.add(stringBuilder.toString()+"!||!");
						loopIndex = loopIndex + 1;
					}
					actionPlanList.add("N");	
				}
			} else {
				LOGGER.error("------> SOMETHING IS MISSING IN DATABASE: JctBeforeSketchHeader. NO RECORD FOUND WITH GIVEN JOB REFERENCE NUMBER: "+jobRefNo);
			}
			LOGGER.info("<<<<<< AfterSketchServiceImpl.fetchActionPlans");
		} catch(DAOException daoException) {
			LOGGER.error("----"+daoException.getLocalizedMessage()+" ----");
		}
		return actionPlanList;
		}
	/**
	 * <p><b>Description:</b> This method is used to fetch div height and width of putting it together page</p>
	 * @param It will take Job reference as a input parameter
	 * @return It returns String object
	 * @throws JCTException -it throws JCTException
	 * 
	 */
	@Transactional(propagation=Propagation.REQUIRED)
	public String getDivHeightAndWidth(String refNo, String linkClicked) throws JCTException{
		LOGGER.info(">>>>>> AfterSketchServiceImpl.getDivHeightAndWidth");
		StringBuilder sb = new StringBuilder("");
		try{
			List<JctAfterSketchHeader> list = afterSketchDAO.getList(refNo, linkClicked);
			if (list.size() == 0) {
				list = afterSketchDAO.getListForEdit(refNo, linkClicked);
			}
			if(list.size()>0){
				JctAfterSketchHeader obj = (JctAfterSketchHeader)list.get(0);
				sb.append(obj.getJctFinalPageDivHeight());
				sb.append("#");
				sb.append(obj.getJctFinalPageDivWidth());
			}
			LOGGER.info("<<<<<< AfterSketchServiceImpl.getDivHeightAndWidth");
		}catch(Exception ex){
			LOGGER.error("----"+ex.getLocalizedMessage()+" ----");
			throw new JCTException();
		}
		return sb.toString();
	}
	/**
	 * <p><b>Description:</b> This method is used to fetch div height and width of putting it together page</p>
	 * @param It will take Job reference as a input parameter
	 * @return It returns String object
	 * @throws JCTException -it throws JCTException
	 * 
	 */
	@Transactional(propagation=Propagation.REQUIRED)
	public String getDivHeightAndWidthForEdit(String refNo, String linkClicked) throws JCTException{
		LOGGER.info(">>>>>> AfterSketchServiceImpl.getDivHeightAndWidthForEdit");
		StringBuilder sb = new StringBuilder("");
		try{
			List<JctAfterSketchHeader> list = afterSketchDAO.getListForEdit(refNo, linkClicked);
			if(list.size()>0){
				JctAfterSketchHeader obj = (JctAfterSketchHeader)list.get(0);
				sb.append(obj.getJctFinalPageDivHeight());
				sb.append("#");
				sb.append(obj.getJctFinalPageDivWidth());
			}
			LOGGER.info("<<<<<< AfterSketchServiceImpl.getDivHeightAndWidthForEdit");
		}catch(Exception ex){
			LOGGER.error("----"+ex.getLocalizedMessage()+" ----");
			throw new JCTException();
		}
		return sb.toString();
	}
	/**
	 * <p><b>Description:</b> This method is used to fetch the job attributes which the user selected 
	 * on mapping attributes page.</p>
	 * @param It will take Job reference as a input parameter
	 * @return It returns String object
	 * @throws JCTException -it throws JCTException
	 * 
	 */
	@Transactional(propagation=Propagation.REQUIRED)
	public String getCheckedItems(String refNo, String linkClicked) throws JCTException{
		LOGGER.info(">>>>>> AfterSketchServiceImpl.getCheckedItems");
		StringBuilder sb = new StringBuilder("");
		try{
			List<JctAfterSketchHeader> list = afterSketchDAO.getList(refNo, linkClicked);
			if(list.size()>0){
				JctAfterSketchHeader obj = (JctAfterSketchHeader)list.get(0);
				sb.append(obj.getJctCheckedStrength());
				sb.append("#");
				sb.append(obj.getJctCheckedPassion());
				sb.append("#");
				sb.append(obj.getJctCheckedValue());
			} else {
				sb.append("");
				sb.append("#");
				sb.append("");
				sb.append("#");
				sb.append("");
			}
			LOGGER.info("<<<<<< AfterSketchServiceImpl.getCheckedItems");
		}catch(Exception ex){
			LOGGER.error("----"+ex.getLocalizedMessage()+" ----");
			throw new JCTException();
		}
		return sb.toString();
	}
	/**
	 * <p><b>Description:</b> This method is used to fetch the base 64 string of putting it together page.</p>
	 * @param It will take Job reference as a input parameter
	 * @return It returns String object
	 * @throws JCTException -it throws JCTException
	 * 
	 */
	@Transactional(propagation=Propagation.REQUIRED)
	public String getFinalPageSnapShot(String refNo, String linkClicked) throws JCTException {
		LOGGER.info(">>>>>> AfterSketchServiceImpl.getFinalPageSnapShot");
		List<JctAfterSketchHeader> list = null;
		StringBuilder sb = new StringBuilder("");
		try{
			list = afterSketchDAO.getList(refNo, linkClicked);
			if(list.size() > 0){
				JctAfterSketchHeader obj = (JctAfterSketchHeader)list.get(0);
				sb.append(obj.getJctAsFinalpageSnapshotString());
			} else {
				list = afterSketchDAO.getListForEdit(refNo, linkClicked);//getListForEdit
				if(list.size() > 0){
					JctAfterSketchHeader obj = (JctAfterSketchHeader)list.get(0);
					sb.append(obj.getJctAsFinalpageSnapshotString());
				}
			}
			LOGGER.info("<<<<<< AfterSketchServiceImpl.getFinalPageSnapShot");
		} catch(Exception ex){
			LOGGER.error("----"+ex.getLocalizedMessage()+" ----");
			throw new JCTException();
		}
		return sb.toString();
	}
	/**
	 *<p><b>Description:</b>  Based on job ref no this method fetch after sketch final diagrams total task element count</p>
	 * @param It will take Job reference as a input parameter
	 * @return It returns Integer object
	 * @throws JCTException -it throws JCTException
	 * 
	 */
	@Transactional(propagation=Propagation.REQUIRED)
	public Integer getAsTotalCount(String jobRefNo, String linkClicked) throws JCTException{
		LOGGER.info(">>>>>> AfterSketchServiceImpl.getAsTotalCount");
		int val = 0;
		try {
			val = afterSketchDAO.getAsTotalCount(jobRefNo, linkClicked);
			LOGGER.info("<<<<<< AfterSketchServiceImpl.getAsTotalCount");
		} 
		catch(NullPointerException npe){
			val = 0;
		}
		catch (DAOException e) {
			LOGGER.error("--------------> COULD NOT FETCH THE TOTAL COUNT ELEMENTS IN AFTER SKETCH DIAGRAM <------------------");
			throw new JCTException();
		}
		return val;
	}
	/**
	 *<p><b>Description:</b>  Based on job ref no this method fetch after sketch final diagrams total task element count</p>
	 * @param It will take Job reference as a input parameter
	 * @return It returns Integer object
	 * @throws JCTException -it throws JCTException
	 * 
	 */
	@Transactional(propagation=Propagation.REQUIRED)
	public Integer getAsTotalCountForEdit(String jobRefNo, String linkClicked) throws JCTException {
		LOGGER.info(">>>>>> AfterSketchServiceImpl.getAsTotalCountForEdit");
		int val = 0;
		try {
			val = afterSketchDAO.getAsTotalCountForEdit(jobRefNo, linkClicked);
			LOGGER.info("<<<<<< AfterSketchServiceImpl.getAsTotalCountForEdit");
		} 
		catch(NullPointerException npe){
			val = 0;
		}
		catch (DAOException e) {
			LOGGER.error("--------------> COULD NOT FETCH THE TOTAL COUNT ELEMENTS IN AFTER SKETCH DIAGRAM <------------------");
			throw new JCTException();
		}
		return val;
	}
	@Transactional(propagation=Propagation.REQUIRED)
	public String getFrozenAttributeById(int attributeId) throws JCTException {
		LOGGER.info(">>>>>>>> AfterSketchServiceImpl.getFrozenAttributeById");
		String attrName = "";
		try {
			attrName = afterSketchDAO.getFrozenAttributeById(attributeId);
		} catch (DAOException dao) {
			throw new JCTException();
		}
		LOGGER.info("<<<<<<<< AfterSketchServiceImpl.getFrozenAttributeById");
		return attrName;
	}
}
