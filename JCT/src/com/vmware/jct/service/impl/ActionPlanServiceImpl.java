package com.vmware.jct.service.impl;

import java.util.ArrayList;
import java.util.Date;
import java.util.Iterator;
import java.util.List;
import java.util.StringTokenizer;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;
import org.springframework.transaction.annotation.Propagation;
import org.springframework.transaction.annotation.Transactional;

import com.vmware.jct.common.utility.CommonConstants;
import com.vmware.jct.dao.IActionPlanDAO;
import com.vmware.jct.dao.IAfterSketchDAO;
import com.vmware.jct.dao.ICommonDao;
import com.vmware.jct.dao.IQuestionnaireDAO;
import com.vmware.jct.exception.DAOException;
import com.vmware.jct.exception.JCTException;
import com.vmware.jct.model.JctActionPlan;
import com.vmware.jct.model.JctAfterSketchHeader;
import com.vmware.jct.model.JctGlobalProfileHistory;
import com.vmware.jct.model.JctStatus;
import com.vmware.jct.model.JctUser;
import com.vmware.jct.service.IActionPlanService;
import com.vmware.jct.service.vo.ActionPlanVO;
/**
 * 
 * <p><b>Class name:</b> ActionPlanServiceImpl.java</p>
 * <p><b>Author:</b> InterraIT</p>
 * <p><b>Purpose:</b> This is a ActionPlanServiceImpl class. This artifact is Business layer artifact.
 * ActionPlanServiceImpl implement IActionPlanService interface and override the following  methods.
 * -saveActinoPlans(List<ActionPlanVO> listObj, String jobRefNo) 
 * </p>
 * <p><b>Description:</b> This class used to perform insert, modify ,fetch data from db and delete operation. Save, update, delete operations are performed in DataAccessObject class </p>
 * <p><b>Revision History:</b>
 * 	<li>InterraIT - 19/Jan/2014 - Implement Exception </li>
 * <li>InterraIT - 31/Jan/2014 - Implement comments, introduce Named query, Transactional attribute </li>
 * </p>
 */
@Service
public class ActionPlanServiceImpl implements IActionPlanService{
	
	@Autowired
	private IActionPlanDAO actionPlanDAO;
	
	@Autowired
	private IAfterSketchDAO afterSketchDAO;
	
	@Autowired
	private IQuestionnaireDAO questionnaireDAO;
	
	@Autowired
	private ICommonDao commonDaoImpl;
	
	private static final Logger LOGGER = LoggerFactory.getLogger(ActionPlanServiceImpl.class);
	
	/**
	 * Method saving action plan data.
	 * @param ActionPlan Vo object, job refarance no
	 * @throws JCTException
	 * @return String
	 */
	@Transactional(propagation=Propagation.REQUIRED)
	public String saveActinoPlans(List<ActionPlanVO> listObj, String jobRefNo, String linkClicked) throws JCTException {
		LOGGER.info(">>>>>> ActionPlanServiceImpl.saveActinoPlans");
		String result =  "success";
		int jctAsHeaderId = 0;
		try{
			List<JctActionPlan> actionPlanList = actionPlanDAO.getList(jobRefNo);
			Date lastDate = null;
			if(actionPlanList.size() > 0){
				Iterator<JctActionPlan> itr = actionPlanList.iterator();
				while(itr.hasNext()){
					//JctActionPlan obj = (JctActionPlan) itr.next();
					//obj.setJctActionPlanSoftDelete(CommonConstants.ENTRY_SOFT_DELETED);
					//actionPlanDAO.mergeActinoPlans(obj);
					//Delete existing
					JctActionPlan obj = (JctActionPlan) itr.next();
					lastDate = obj.getJctAsCreatedTs();
					actionPlanDAO.remove(obj);
				}
			}
			//Add new
			Iterator<ActionPlanVO> objItr = listObj.iterator();
			while(objItr.hasNext()){
				Date createdDate = new Date();
				ActionPlanVO vo = (ActionPlanVO)objItr.next();
				JctActionPlan object = new JctActionPlan();
				//object.setJctActionPlanId(commonDaoImpl.generateKey("jct_action_plan"));
				try {
					JctAfterSketchHeader hdr = afterSketchDAO.getList(jobRefNo, linkClicked).get(0);
					jctAsHeaderId = hdr.getJctAsHeaderId();
				} catch (Exception ex) {
					jctAsHeaderId = 0;
				}
				
				//JctAfterSketchHeader hdr = afterSketchDAO.getList(jobRefNo, linkClicked).get(0);
				//jctAsHeaderId = hdr.getJctAsHeaderId();
				object.setJctAsHeaderId(jctAsHeaderId);
				
				/** THIS IS FOR PUBLIC VERSION **/
				JctUser user = new JctUser();
				user.setJctUserId(vo.getJctUserId());
				object.setJctUser(user);
				/********************************/
				
				object.setJctJobrefNo(vo.getJobRefNo());
				object.setJctAsQuestionDesc(vo.getQuestionDesc());
				List<JctGlobalProfileHistory> historyList = questionnaireDAO.getGlobalProfileChangedData(lastDate, vo.getQuestionSubDesc(), vo.getQuestionDesc(), "AP");
				if (historyList.size() > 0) {
					JctGlobalProfileHistory obj = (JctGlobalProfileHistory) historyList.get(historyList.size() - 1); // Get the last one
					object.setJctAsQuestionSubDesc(obj.getJctGlobalProfileApSubQtnChanged());
				} else {
					object.setJctAsQuestionSubDesc(vo.getQuestionSubDesc());
				}
				object.setJctAsAnswarDesc(vo.getAnswerDesc());
				object.setVersion(0);
				object.setJctAsCreatedBy(vo.getUserModified());
				object.setJctAsLastmodifiedBy(vo.getUserModified());
				object.setJctAsCreatedTs(createdDate);
				object.setJctAsLastmodifiedTs(createdDate);
				object.setJctActionPlanSoftDelete(0);
				JctStatus status = new JctStatus();
				status.setJctStatusId(CommonConstants.TASK_PENDING);
				object.setJctStatus(status);
				actionPlanDAO.saveActinoPlans(object);
				LOGGER.info("<<<<<< ActionPlanServiceImpl.saveActinoPlans");
			}
		} catch(DAOException daoEx){
			LOGGER.error("----"+daoEx.getLocalizedMessage()+" ----");
			result = "failure";
			throw new JCTException();
		}
		return result;
	}
	/**
	 * Method check how many times the user has completed
	 */
	public int getNumberOfCompletion (String jobReferenceNos) throws JCTException {
		int completionCount = 0;
		try {
			//completionCount = actionPlanDAO.getNumberOfCompletion(jobReferenceNos).intValue();
			completionCount = actionPlanDAO.getLastCompletionStatus(jobReferenceNos);
		} catch (DAOException ex) {
			LOGGER.error("UNABLE TO DECIDE THE COMPLETION COUNT.... MAKING IT ZERO: "+ex.getLocalizedMessage());
		}
		return completionCount;
	}
	
	/**
	 * Method will freeze the content for questionnaire. This is used only in public version.
	 * @param jrNo
	 * @param userName
	 * @param userId
	 * @param profileId
	 * @return
	 * @throws JCTException
	 */
	@Transactional(propagation=Propagation.REQUIRED)
	public String freezeActionPlanContent (String jrNo, String user, int userId,
			int profileId) throws JCTException {
		String result = "failure";
		try {
		// Select Distinct Main question List
		StringBuilder builder = new StringBuilder("");
		List<String> actionPlanMainQtnList = afterSketchDAO.getActionPlanMainQtnList(profileId);
		Iterator<String> actionPlanMainQtnLstItr = actionPlanMainQtnList.iterator();
		while (actionPlanMainQtnLstItr.hasNext()) {
			//List<String> innerList = new ArrayList<String>();
			String mainQuestion = (String) actionPlanMainQtnLstItr.next();
			builder.append(mainQuestion);
			builder.append("@@@");
			// Select the sub question List
			List<String> subQtnList = afterSketchDAO.getActionPlanSubQtnList(mainQuestion, profileId);
			Iterator<String> subQtnListItr = subQtnList.iterator();
			while (subQtnListItr.hasNext()) {
				builder.append((String) subQtnListItr.next()+"~");
				builder.append(" ");
				builder.append("`");
				builder.append(" ");
				builder.append("`");
				builder.append(" ");
				builder.append("`");
				builder.append("`#");
			}
			builder.append(")(");
		}
		String saveResult = saveFrozenAns(builder.toString(), jrNo, user, userId);
		LOGGER.info("SAVING RESULT: "+saveResult);
		result = "success";
	} catch (DAOException daoEx) {
		LOGGER.error("--> "+daoEx.getLocalizedMessage());
	}
	return result;
	}
	
	private String saveFrozenAns(String ansStr, String jobReferenceNo, String user, int userId) {
		String result = "failure";
		List<ActionPlanVO> actionPlanVOList = new ArrayList<ActionPlanVO>();
		StringTokenizer token = new StringTokenizer(ansStr, ")(");
		while (token.hasMoreTokens()) {
		String totalToken = (String) token.nextToken();
		StringTokenizer innerTokenizer = new StringTokenizer(totalToken, "@@@");
		while (innerTokenizer.hasMoreTokens()) {
			String question = innerTokenizer.nextToken().toString();
			String answerString = innerTokenizer.nextToken().toString();
			StringTokenizer answerToken = new StringTokenizer(answerString, "#");
			while (answerToken.hasMoreTokens()) {
				String subQtnAndAns = (String) answerToken.nextToken();
				StringTokenizer finalToken = new StringTokenizer(subQtnAndAns, "~");
				while (finalToken.hasMoreElements()) {
					String subQtns = (String) finalToken.nextToken();
					try {
						String answers = (String) finalToken.nextToken();
						String[] answerArr = answers.split("`");
						for (int answerIndex = 0; answerIndex < 3; answerIndex++) {
							ActionPlanVO planVO = new ActionPlanVO();
							/** THIS IS FOR PUBLIC VERSION **/
							planVO.setJctUserId(userId);
							/********************************/
							planVO.setJobRefNo(jobReferenceNo);
							planVO.setUserModified(user);
							planVO.setQuestionDesc(question);
							planVO.setAnswerDesc(answerArr[answerIndex].replaceAll("-_-", "#").replaceAll(":=:", "(").replaceAll(";=;", ")").replaceAll(":-:", "@").replaceAll(";_;", "?"));
							planVO.setQuestionSubDesc(subQtns);
							actionPlanVOList.add(planVO);
						}
						saveActinoPlans(actionPlanVOList, jobReferenceNo, "N");
						result = "success";
					} catch (ArrayIndexOutOfBoundsException ex) {
						LOGGER.info("ArrayIndexOutOfBoundsException..."+ex.getMessage());
					}catch (Exception ex) {
						LOGGER.info("Exception..."+ex.getMessage());
					}
				}
			}
		}
	}
		return result;
	}
	/**
	 * Method will return the count.
	 * @param jrNo
	 * @param user
	 * @param userId
	 * @param profileId
	 * @return
	 * @throws JCTException
	 */
	@Transactional(propagation=Propagation.REQUIRED)
	public int getActionPlanCount(String jrNo, String user, int userId,
			int profileId) throws JCTException {
		int count = 0;
		try {
			count = actionPlanDAO.getActionPlanCount(jrNo, user, userId, profileId).intValue();
		} catch (DAOException ex) {
			LOGGER.error("UNABLE TO FIND THE ACTION PLAN COUNT: "+ex.getMessage());
			throw new JCTException();
		}
		return count;
	}
	
}
