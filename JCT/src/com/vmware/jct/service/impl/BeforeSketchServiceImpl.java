package com.vmware.jct.service.impl;

import java.util.ArrayList;
import java.util.Date;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;
import org.springframework.transaction.annotation.Propagation;
import org.springframework.transaction.annotation.Transactional;

import com.vmware.jct.common.utility.CommonConstants;
import com.vmware.jct.dao.IBeforeSketchDAO;
import com.vmware.jct.dao.ICommonDao;
import com.vmware.jct.dao.ISketchSearchDAO;
import com.vmware.jct.exception.DAOException;
import com.vmware.jct.exception.JCTException;
import com.vmware.jct.model.JctBeforeSketchDetail;
import com.vmware.jct.model.JctBeforeSketchDetailTemp;
import com.vmware.jct.model.JctBeforeSketchHeader;
import com.vmware.jct.model.JctBeforeSketchHeaderTemp;
import com.vmware.jct.model.JctBeforeSketchQuestion;
import com.vmware.jct.model.JctCompletionStatus;
import com.vmware.jct.model.JctQuestion;
import com.vmware.jct.model.JctStatus;
import com.vmware.jct.model.JctUser;
import com.vmware.jct.service.IBeforeSketchService;
import com.vmware.jct.service.vo.JctBeforeSketchDetailsVO;
import com.vmware.jct.service.vo.JctBeforeSketchHeaderVO;

/**
 * 
 * <p><b>Class name:</b> BeforeSketchServiceImpl.java</p>
 * <p><b>Author:</b> InterraIt</p>
 * <p><b>Purpose:</b> This is a BeforeSketchServiceImpl class. This artifact is Business layer artifact.
 * BeforeSketchServiceImpl implement IBeforeSketchService interface and override the following  methods.
 * -saveOrUpdateBeforeSketch(JctBeforeSketchHeaderVO header, List<JctBeforeSketchDetailsVO> childList)
 * -getJson(String jobReferenceNo)
 * -getInitialJson(String jobReferenceNo)
 * -fetchQuestion(String jobRefNo)
 * -fetchQuestionAnswer(String jobRefNo)
 * -fetchBase64String(String jobReferenceNo)</p>
 * <p><b>Description:</b> This class used to perform insert, modify ,fetch data from db and delete operation. Save, update, delete operations are performed in DataAccessObject class </p>
 * <p><b>Revision History:</b>
 * 	<li>InterraIt - 19/Jan/2014 - Implement Exception </li>
 * 	<li>InterraIt - 31/Jan/2014 - Implement comments, introduce Named query, Transactional attribute </li>
 * 	<li>InterraIt - 05/Aug/2014 - Total Time calculation logic changed while saving before sketch diagram </li>
 * </p>
 */
@Service
public class BeforeSketchServiceImpl implements IBeforeSketchService {

	@Autowired
	private IBeforeSketchDAO beforeSketchDAO;
	
	@Autowired
	private ICommonDao commonDaoImpl;
	
	@Autowired
	private ISketchSearchDAO iSketchSearchDAO;
	
	private static final Logger LOGGER = LoggerFactory.getLogger(BeforeSketchServiceImpl.class);
	/**
	 *<p><b>Description:</b> Method is used to save or update the before diagram data.</p>
	 * @param JctBeforeSketchHeaderVO
	 * @param List<JctBeforeSketchDetailsVO>
	 * @return String
	 * @throws JCTException
	 * 
	 */
	@Transactional(propagation=Propagation.REQUIRED)
	public String saveOrUpdateBeforeSketch(JctBeforeSketchHeaderVO header,
			List<JctBeforeSketchDetailsVO> childList, String isEdit) throws JCTException {
		LOGGER.info(">>>> BeforeSketchServiceImpl.saveOrUpdateBeforeSketch");
		String result = "failure";
		if (isEdit.equals("Y")) { // edited diagram
			result = saveToBSTempTable (header, childList);
		} else {
			Date createdTs = new Date();
			JctBeforeSketchHeader parent = null;
			List<JctBeforeSketchDetail> childEntityList = null;
			JctStatus status = new JctStatus();
			status.setJctStatusId(CommonConstants.TASK_PENDING);
			int totalTimeSpent = 0;
			int insertionSequence = 0;
			// Check if record pertaining to same job reference id exists
			try {
				List<JctBeforeSketchHeader> list = beforeSketchDAO.getList(header.getJobReferenceNumber());
				/**
				 * Check if the user has ever pressed the finish button in the excersize (completed the flow once).
				 * If yes:
				 * 		1. make the existing transaction data as soft deleted (parent and children both).
				 * 		2. Insert a new row. [Nothing should be hard deleted in this case]
				 * 
				 * If No:
				 * 		1. Hard delete the data as usual from the parent and child table.
				 * 		2. Insert a new row.
				 * 
				 */
				List<JctCompletionStatus> compList = iSketchSearchDAO.getCompletionStatusObjs(header.getJobReferenceNumber());
				totalTimeSpent = totalTimeSpent + header.getTimeTaken();
				JctBeforeSketchHeader object = null;
				if (list.size() > 0) {
					object = ((JctBeforeSketchHeader)list.get(0));
					totalTimeSpent = object.getJctBsTimeSpent();
					beforeSketchDAO.remove(object);
				}
				/*if (compList.size() > 0) {
					try {
						object = ((JctBeforeSketchHeader) list.get(0));
					} catch(Exception ex) {
						return saveOrUpdateBeforeSketchRestart(header, childList);
					}
					object.setJctBsJson(header.getTotalJson());
					object.setJctBsSnapshot(header.getSnapShot());
					object.setInitialJsonVal(header.getInitialJsonVal());
					object.setJctBsSnapshotString(header.getImageBase64String());
					object.setJctBsTimeSpent(totalTimeSpent + header.getTimeTaken());
					int index = 0;
					HashMap<Integer, String> map = new HashMap<Integer, String>();
					Iterator<JctBeforeSketchDetailsVO> myItr = childList.iterator();
					while (myItr.hasNext()) {
						JctBeforeSketchDetailsVO vo = (JctBeforeSketchDetailsVO) myItr.next();
						map.put(vo.getTaskId(), vo.getTaskDesc());
					}
					
					Iterator<JctBeforeSketchDetail> itr = object.getJctBeforeSketchDetails().iterator();
					while (itr.hasNext()) {
						JctBeforeSketchDetail detail = (JctBeforeSketchDetail) itr.next();
						int taskId = detail.getJctBsTaskId();
						String taskDescFromMap = map.get(taskId);
						detail.setJctBsTaskDesc(taskDescFromMap);
						detail.setJctBeforeSketchHeader(object);
					}
					beforeSketchDAO.saveOrUpdateBeforeSketch(object);
					
				} else { //Delete the existing one*/
					/*if (list.size() > 0) {
						object = ((JctBeforeSketchHeader)list.get(0));
						totalTimeSpent = object.getJctBsTimeSpent();
						beforeSketchDAO.remove(object);
					}*/
					// Now insert the new record
					parent = new JctBeforeSketchHeader();
					//parent.setJctBsHeaderId(commonDaoImpl.generateKey("jct_before_sketch_header_id"));
					
					/** THIS IS FOR PUBLIC VERSION **/
					JctUser user = new JctUser();
					user.setJctUserId(header.getJctUserId());
					parent.setJctUser(user);
					/********************************/
					
					parent.setJctBsJobrefNo(header.getJobReferenceNumber());
					parent.setJctBsJson(header.getTotalJson());
					parent.setJctBsUserJobTitle(header.getJobTitle());
					parent.setJctStatus(status);
					parent.setJctBsLastmodifiedTs(createdTs);
					parent.setJctDsCreatedTs(createdTs);
					parent.setJctBsSnapshot(header.getSnapShot());
					parent.setInitialJsonVal(header.getInitialJsonVal());
					parent.setJctBsCreatedBy(header.getCreatedBy());
					parent.setJctBsSnapshotString(header.getImageBase64String());
					parent.setJctBsTimeSpent(totalTimeSpent+header.getTimeTaken());
					parent.setJctBsInsertionSequence(insertionSequence + 1);
					
					// Iterate and create the list of child
					childEntityList = new ArrayList<JctBeforeSketchDetail>();
					Iterator<JctBeforeSketchDetailsVO> childIt = childList.iterator();
					while (childIt.hasNext()) {
						JctBeforeSketchDetailsVO vo = (JctBeforeSketchDetailsVO) childIt.next();
						JctBeforeSketchDetail child = new JctBeforeSketchDetail();
						/*try {
							child.setJctBsDetailId(commonDaoImpl.generateKey("jct_before_sketch_details_id"));
						} catch (DAOException e) {
							LOGGER.error(e.getLocalizedMessage());
							throw new JCTException(e.getMessage());
						}*/
						child.setJctBsTaskId(vo.getTaskId());
						child.setJctBsEnergy(vo.getEnergy());
						child.setJctBsJobrefNo(vo.getJobReferenceNumber());
						child.setJctBsTaskDesc(vo.getTaskDesc());
						child.setJctBsPosition(vo.getPosition());
						child.setSoftDelete(0); // Active
						child.setJctDsCreatedTs(createdTs);
						child.setJctBsLastmodifiedTs(createdTs);
						child.setJctStatus(status);
						child.setJctBsCreatedBy(vo.getCreatedBy());
						child.setJctBeforeSketchHeader(parent);
						childEntityList.add(child);
					}
					parent.setJctBeforeSketchDetails(childEntityList);
					result = beforeSketchDAO.saveOrUpdateBeforeSketch(parent);
				//}
				
				
				
			} catch (DAOException daoException) {
				LOGGER.error(daoException.getLocalizedMessage());
				throw new JCTException(daoException.getMessage());
			}
		}	
		LOGGER.info("<<<< BeforeSketchServiceImpl.saveOrUpdateBeforeSketch");
		return result;
	}
	
	private String saveOrUpdateBeforeSketchRestart(JctBeforeSketchHeaderVO header,
			List<JctBeforeSketchDetailsVO> childList) throws DAOException {
		Date createdTs = new Date();
		JctBeforeSketchHeader parent = null;
		List<JctBeforeSketchDetail> childEntityList = null;
		JctStatus status = new JctStatus();
		status.setJctStatusId(CommonConstants.TASK_PENDING);
		// Now insert the new record
		parent = new JctBeforeSketchHeader();
		//parent.setJctBsHeaderId(commonDaoImpl.generateKey("jct_before_sketch_header_id"));
		/** THIS IS FOR PUBLIC VERSION **/
		JctUser user = new JctUser();
		user.setJctUserId(header.getJctUserId());
		parent.setJctUser(user);
		/********************************/
		parent.setJctBsJobrefNo(header.getJobReferenceNumber());
		parent.setJctBsJson(header.getTotalJson());
		parent.setJctBsUserJobTitle(header.getJobTitle());
		parent.setJctStatus(status);
		parent.setJctBsLastmodifiedTs(createdTs);
		parent.setJctDsCreatedTs(createdTs);
		parent.setJctBsSnapshot(header.getSnapShot());
		parent.setInitialJsonVal(header.getInitialJsonVal());
		parent.setJctBsCreatedBy(header.getCreatedBy());
		parent.setJctBsSnapshotString(header.getImageBase64String());
		parent.setJctBsTimeSpent(header.getTimeTaken());
		parent.setJctBsInsertionSequence(0);
		// Iterate and create the list of child
		childEntityList = new ArrayList<JctBeforeSketchDetail>();
		Iterator<JctBeforeSketchDetailsVO> childIt = childList.iterator();
		while (childIt.hasNext()) {
			JctBeforeSketchDetailsVO vo = (JctBeforeSketchDetailsVO) childIt.next();
			JctBeforeSketchDetail child = new JctBeforeSketchDetail();
			/*try {
				child.setJctBsDetailId(commonDaoImpl.generateKey("jct_before_sketch_details_id"));
			} catch (DAOException e) {
				LOGGER.error(e.getLocalizedMessage());
				throw new JCTException(e.getMessage());
			}*/
			child.setJctBsTaskId(vo.getTaskId());
			child.setJctBsEnergy(vo.getEnergy());
			child.setJctBsJobrefNo(vo.getJobReferenceNumber());
			child.setJctBsTaskDesc(vo.getTaskDesc());
			child.setJctBsPosition(vo.getPosition());
			child.setSoftDelete(0); // Active
			child.setJctDsCreatedTs(createdTs);
			child.setJctBsLastmodifiedTs(createdTs);
			child.setJctStatus(status);
			child.setJctBsCreatedBy(vo.getCreatedBy());
			child.setJctBeforeSketchHeader(parent);
			childEntityList.add(child);
		}
		parent.setJctBeforeSketchDetails(childEntityList);
		String result = beforeSketchDAO.saveOrUpdateBeforeSketch(parent); 
		return result;
	}
	
	private String saveToBSTempTable(JctBeforeSketchHeaderVO header,
			List<JctBeforeSketchDetailsVO> childList) throws JCTException {
		Date createdTs = new Date();
		String result = "failure";
		JctBeforeSketchHeaderTemp parent = null;
		List<JctBeforeSketchDetailTemp> childEntityList = null;
		JctStatus status = new JctStatus();
		status.setJctStatusId(CommonConstants.TASK_PENDING);
		int totalTimeSpent = 0;
		
		try {
			//Check if record pertaining to same job reference id exists
			List<JctBeforeSketchHeaderTemp> list = beforeSketchDAO.getListTemp(header.getJobReferenceNumber());
			
			//Delete the existing one
			if(list.size() > 0){
				Iterator<JctBeforeSketchHeaderTemp> itr = list.iterator();
				while (itr.hasNext()) {
					JctBeforeSketchHeaderTemp object = (JctBeforeSketchHeaderTemp) itr.next();
					totalTimeSpent = object.getJctBsTimeSpent();
					beforeSketchDAO.removeTemp(object);
				}
				
				//JctBeforeSketchHeaderTemp object = ((JctBeforeSketchHeaderTemp)list.get(0));
				//totalTimeSpent = object.getJctBsTimeSpent();
				//beforeSketchDAO.removeTemp(object);
			}
			totalTimeSpent = totalTimeSpent + header.getTimeTaken();
			// Now insert the new record
			parent = new JctBeforeSketchHeaderTemp();
			//parent.setJctBsHeaderId(commonDaoImpl.generateKey("jct_before_sketch_header_temp_id"));
			/** THIS IS FOR PUBLIC VERSION **/
			JctUser user = new JctUser();
			user.setJctUserId(header.getJctUserId());
			parent.setJctUser(user);
			/********************************/
			parent.setJctBsJobrefNo(header.getJobReferenceNumber());
			parent.setJctBsJson(header.getTotalJson());
			parent.setJctBsUserJobTitle(header.getJobTitle());
			parent.setJctStatus(status);
			parent.setJctBsLastmodifiedTs(createdTs);
			parent.setJctDsCreatedTs(createdTs);
			parent.setJctBsSnapshot(header.getSnapShot());
			parent.setInitialJsonVal(header.getInitialJsonVal());
			parent.setJctBsCreatedBy(header.getCreatedBy());
			parent.setJctBsSnapshotString(header.getImageBase64String());
			parent.setJctBsTimeSpent(totalTimeSpent);
			//Iterate and create the list of child
			childEntityList = new ArrayList<JctBeforeSketchDetailTemp>();
			Iterator<JctBeforeSketchDetailsVO> childIt = childList.iterator();
			while(childIt.hasNext()){
				JctBeforeSketchDetailsVO vo = (JctBeforeSketchDetailsVO)childIt.next();
				JctBeforeSketchDetailTemp child = new JctBeforeSketchDetailTemp();
				/*try {
					child.setJctBsDetailId(commonDaoImpl.generateKey("jct_before_sketch_details_temp_id"));
				} catch (DAOException e) {
					LOGGER.error(e.getLocalizedMessage());
					throw new JCTException(e.getMessage());
				}*/
				child.setJctBsTaskId(vo.getTaskId());
				child.setJctBsEnergy(vo.getEnergy());
				child.setJctBsJobrefNo(vo.getJobReferenceNumber());
				child.setJctBsTaskDesc(vo.getTaskDesc());
				child.setJctBsPosition(vo.getPosition());
				child.setSoftDelete(0); //Active
				child.setJctDsCreatedTs(createdTs);
				child.setJctBsLastmodifiedTs(createdTs);
				child.setJctStatus(status);
				child.setJctBsCreatedBy(vo.getCreatedBy());
				child.setJctBeforeSketchHeader(parent);
				childEntityList.add(child);
			}
			parent.setJctBeforeSketchDetails(childEntityList);
			result = beforeSketchDAO.saveOrUpdateBeforeSketchTemp(parent);
		} catch (DAOException daoException) {
			LOGGER.error(daoException.getLocalizedMessage());
			throw new JCTException(daoException.getMessage());
		}
		return result;
		}
	/**
	 *<p><b>Description:</b> Method is used to fetch before diagram json.</p>
	 * @param Job reference number
	 * @return String
	 * @throws JCTException
	 * 
	 */
	public String getJson(String jobReferenceNo, String linkClicked) throws JCTException {
		LOGGER.info(">>>> BeforeSketchServiceImpl.getJson");
		try{
			LOGGER.info("<<<< BeforeSketchServiceImpl.getJson");
			return beforeSketchDAO.getJson(jobReferenceNo, linkClicked);
		}catch (DAOException daoException) {
			LOGGER.error(daoException.getLocalizedMessage());
			throw new JCTException(daoException.getMessage());
		}
	}
	/**
	 *<p><b>Description:</b> Method is used to fetch before diagram json.</p>
	 * @param Job reference number
	 * @return String
	 * @throws JCTException
	 * 
	 */
	public String getJsonForEdit(String jobReferenceNo, String linkClicked) throws JCTException {
		LOGGER.info(">>>> BeforeSketchServiceImpl.getJsonForEdit");
		try{
			LOGGER.info("<<<< BeforeSketchServiceImpl.getJsonForEdit");
			return beforeSketchDAO.getJsonForEdit(jobReferenceNo, linkClicked);
		}catch (DAOException daoException) {
			LOGGER.error(daoException.getLocalizedMessage());
			throw new JCTException(daoException.getMessage());
		}
	}
	/**
	 *<p><b>Description:</b> Method is used to fetch before diagram json.</p>
	 * @param Job reference number
	 * @return String
	 * @throws JCTException
	 * 
	 */
	public String getJsonFromTemp(String jobReferenceNo, String linkClicked) throws JCTException {
		LOGGER.info(">>>> BeforeSketchServiceImpl.getJsonFromTemp");
		try{
			LOGGER.info("<<<< BeforeSketchServiceImpl.getJsonFromTemp");
			return beforeSketchDAO.getJsonForEditFromTemp(jobReferenceNo, linkClicked);
		}catch (DAOException daoException) {
			LOGGER.error(daoException.getLocalizedMessage());
			throw new JCTException(daoException.getMessage());
		}
	}
	/**
	 *<p><b>Description:</b> Method is used to fetch the initial json of before diagram page.</p>
	 * @param job reference number
	 * @return String
	 * @throws JCTException
	 * 
	 */
	public String getInitialJson(String jobReferenceNo, String linkClicked) throws JCTException {
		LOGGER.info(">>>> BeforeSketchServiceImpl.getInitialJson");
		try{
			LOGGER.info("<<<< BeforeSketchServiceImpl.getInitialJson");
			return beforeSketchDAO.getInitialJson(jobReferenceNo, linkClicked);
		}catch (DAOException daoException) {
			LOGGER.error(daoException.getLocalizedMessage());
			throw new JCTException(daoException.getMessage());
		}
	}
	/**
	 *<p><b>Description:</b> Method is used to fetch the initial json of before diagram page.</p>
	 * @param job reference number
	 * @return String
	 * @throws JCTException
	 * 
	 */
	public String getInitialJsonFromTemp(String jobReferenceNo, String linkClicked) throws JCTException {
		LOGGER.info(">>>> BeforeSketchServiceImpl.getInitialJsonFromTemp");
		try{
			LOGGER.info("<<<< BeforeSketchServiceImpl.getInitialJsonFromTemp");
			return beforeSketchDAO.getInitialJsonFromTemp(jobReferenceNo, linkClicked);
		}catch (DAOException daoException) {
			LOGGER.error(daoException.getLocalizedMessage());
			throw new JCTException(daoException.getMessage());
		}
	}
	/**
	 *<p><b>Description:</b> Method is used to fetch the initial json of before diagram page.</p>
	 * @param job reference number
	 * @return String
	 * @throws JCTException
	 * 
	 */
	public String getInitialJsonForEdit(String jobReferenceNo, String linkClicked) throws JCTException {
		LOGGER.info(">>>> BeforeSketchServiceImpl.getInitialJsonForEdit");
		try{
			LOGGER.info("<<<< BeforeSketchServiceImpl.getInitialJsonForEdit");
			return beforeSketchDAO.getInitialJsonForEdit(jobReferenceNo, linkClicked);
		}catch (DAOException daoException) {
			LOGGER.error(daoException.getLocalizedMessage());
			throw new JCTException(daoException.getMessage());
		}
	}
	/**
	 *<p><b>Description:</b> Method is used to fetch questionnaire.</p>
	 * @param profile id
	 * @return List<JctQuestion>
	 * @throws JCTException
	 * 
	 */
	public List<JctQuestion> fetchQuestion(int profileId) throws JCTException {
		LOGGER.info(">>>> BeforeSketchServiceImpl.fetchQuestion");
		/*List<JctBeforeSketchHeader> list=null;
		JctBeforeSketchHeader parent = null;*/
		List<JctQuestion> jctQuestionList=null;
		try{
			LOGGER.debug("Test profileId::"+profileId);
			jctQuestionList=beforeSketchDAO.getFetchQuestion(profileId);
		} catch(DAOException daoException) {
			LOGGER.error(daoException.getLocalizedMessage());
		}
		LOGGER.info("<<<< BeforeSketchServiceImpl.fetchQuestion");
		return jctQuestionList;
	}
	/**
	 *<p><b>Description:</b> Method is used to fetch question and answers.</p>
	 * @param profile id
	 * @param job reference number
	 * @return List<String>
	 * @throws JCTException
	 * 
	 */
	public List<String> fetchQuestionAnswer(String jobRefNo, int profileId) throws JCTException {
		LOGGER.info(">>>> BeforeSketchServiceImpl.fetchQuestionAnswer");
		List<JctBeforeSketchQuestion> list = null;
		List<String> questionAnswerList = new ArrayList<String>(); 
		StringBuilder stringBuilder = null;
		try {
			list = beforeSketchDAO.fetchQuestionAnswer(jobRefNo);
			if(list.size() > 0){
				Iterator<JctBeforeSketchQuestion> itr = list.iterator();
				while(itr.hasNext()){
					JctBeforeSketchQuestion obj = (JctBeforeSketchQuestion)itr.next();
					stringBuilder = new StringBuilder("");
					stringBuilder.append(obj.getJctBsQuestionDesc());
					stringBuilder.append("~");
					stringBuilder.append(obj.getJctBsAnswarDesc());
					questionAnswerList.add(stringBuilder.toString()+"!||!");
				}
			} else{
				List al = fetchQuestion(profileId);
				Iterator itr = al.iterator();
				while(itr.hasNext()){
					stringBuilder = new StringBuilder("");
					stringBuilder.append((String)itr.next());
					stringBuilder.append("~");
					stringBuilder.append("");
					questionAnswerList.add(stringBuilder.toString()+"!||!");
				}
			}
		} catch(DAOException daoException) {
			LOGGER.error(daoException.getLocalizedMessage());
		}
		LOGGER.info("<<<< BeforeSketchServiceImpl.fetchQuestionAnswer");
		return questionAnswerList;
	}
	/**
	 *<p><b>Description:</b> Method is used to fetch before diagram snap shot.</p>
	 * @param job reference number
	 * @return String
	 * @throws JCTException
	 * 
	 */
	public String fetchBase64String(String jobReferenceNo) throws JCTException {
		LOGGER.info(">>>> BeforeSketchServiceImpl.fetchBase64String");
		List<JctBeforeSketchHeader> list = null;
		String base64 = null;
		try {
			list = beforeSketchDAO.getList(jobReferenceNo);
			if(list.size() != 0){
				JctBeforeSketchHeader obj = (JctBeforeSketchHeader)list.get(0);
				base64 = obj.getJctBsSnapshotString();
			}
		} catch(DAOException daoException) {
			LOGGER.error(daoException.getLocalizedMessage());
		}
		LOGGER.info("<<<< BeforeSketchServiceImpl.fetchBase64String");
		return base64;
	}
}
