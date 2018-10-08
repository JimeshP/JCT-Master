package com.vmware.jct.service.impl;

import java.util.ArrayList;
import java.util.Date;
import java.util.Iterator;
import java.util.List;
import java.util.StringTokenizer;

import org.hibernate.SessionFactory;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;
import org.springframework.transaction.annotation.Propagation;
import org.springframework.transaction.annotation.Transactional;

import com.vmware.jct.common.utility.CommonConstants;
import com.vmware.jct.dao.IAfterSketchDAO;
import com.vmware.jct.dao.IBeforeSketchDAO;
import com.vmware.jct.dao.ICommonDao;
import com.vmware.jct.dao.IQuestionnaireDAO;
import com.vmware.jct.dao.dto.JobAttributeDTO;
import com.vmware.jct.exception.DAOException;
import com.vmware.jct.exception.JCTException;
import com.vmware.jct.model.JctBeforeSketchHeader;
import com.vmware.jct.model.JctBeforeSketchQuestion;
import com.vmware.jct.model.JctGlobalProfileHistory;
import com.vmware.jct.model.JctStatus;
import com.vmware.jct.model.JctUser;
import com.vmware.jct.service.IQuestionnaireService;
import com.vmware.jct.service.vo.QuestionnaireVO;
/**
 * 
 * <p><b>Class name:</b> QuestionnaireServiceImpl.java</p>
 * <p><b>Author:</b> InterreIt</p>
 * <p><b>Purpose:</b> This is a QuestionnaireServiceImpl class. This artifact is Business layer artifact.
 * QuestionnaireServiceImpl implement IQuestionnaireService interface and override the following  methods.
 * -saveOrUpdateAnswers(List<QuestionnaireVO> vo, String jobRefNumber)
 * -fetchJobAttribute()
 * </p>
 * <p><b>Description:</b> This class used to perform insert, modify ,fetch data from db and delete operation. Save, update, delete operations are performed in DataAccessObject class </p>
 * <p><b>Revision History:</b>
 * 	<li>InterreIt - 19/Jan/2014 - Implement Exception </li>
 * <li>InterreIt - InterraIt - 31/Jan/2014 - Implement comments, introduce Named query, Transactional attribute </li>
 * </p>
 */
@Service
public class QuestionnaireServiceImpl implements IQuestionnaireService{

	@Autowired
	private IQuestionnaireDAO questionnaireDAO;
	
	@Autowired
	private IAfterSketchDAO afterSketchDAO;
	
	@Autowired
	private ICommonDao commonDaoImpl;
	
	@Autowired
	private IBeforeSketchDAO beforeSketchDAO;
	
	@Autowired
	private SessionFactory sessionFactory;
	
	private static final Logger LOGGER = LoggerFactory.getLogger(QuestionnaireServiceImpl.class);
	/**
	 *<p><b>Description:</b> This method saves the question and answers the user has provided in questionnaire module.</p>
	 * @param job reference number
	 * @param List<QuestionnaireVO>
	 * @return String
	 * @throws JCTException
	 * 
	 */
	@Transactional(propagation=Propagation.REQUIRED)
	public String saveOrUpdateAnswers(List<QuestionnaireVO> voList, String jobRefNumber, String isCompleted) throws JCTException {
		LOGGER.info(">>>> QuestionnaireServiceImpl.saveOrUpdateAnswers");
		String result =  "success";
		int jctBsHeaderId = 0;
		Date insertionDate = new Date();
		JctStatus status = new JctStatus();
		status.setJctStatusId(CommonConstants.TASK_PENDING);
		try{
			List<JctBeforeSketchQuestion> qtnList = questionnaireDAO.getList(jobRefNumber);
			Date lastDate = null;
			if(qtnList.size() > 0){
				Iterator<JctBeforeSketchQuestion> itr = qtnList.iterator();
				while(itr.hasNext()){
					/*JctBeforeSketchQuestion obj = (JctBeforeSketchQuestion) itr.next();
					if (obj.getJctBsSoftDelete() == 0) {
						obj.setJctBsSoftDelete(CommonConstants.ENTRY_SOFT_DELETED);
						questionnaireDAO.mergeAnswers(obj);
					}*/
					//Delete existing
					JctBeforeSketchQuestion obj = (JctBeforeSketchQuestion) itr.next();
					lastDate = obj.getJctDsCreatedTs();
					questionnaireDAO.remove(obj);
				}
			}
			//Add new
			Iterator<QuestionnaireVO> objItr = voList.iterator();
			while(objItr.hasNext()) {
				QuestionnaireVO vo = (QuestionnaireVO)objItr.next();
				JctBeforeSketchQuestion object = new JctBeforeSketchQuestion();
				try {
					JctBeforeSketchHeader header = beforeSketchDAO.getList(jobRefNumber).get(0);
					jctBsHeaderId = header.getJctBsHeaderId();
				} catch (Exception ex) {
					jctBsHeaderId = 0;
				}
				/** THIS IS FOR PUBLIC VERSION **/
				JctUser user = new JctUser();
				user.setJctUserId(vo.getJctUserId());
				object.setJctUser(user);
				/********************************/
				
				object.setJctBsHeaderId(jctBsHeaderId);
				object.setJctBsJobrefNo(vo.getJobReferenceNumber());
				object.setJctBsQuestionDesc(vo.getQuestionDescription());
				List<JctGlobalProfileHistory> historyList = questionnaireDAO.getGlobalProfileChangedData(lastDate, vo.getQtnSubDesc(), vo.getQuestionDescription(), "RQ");
				if (historyList.size() > 0) {
					JctGlobalProfileHistory obj = (JctGlobalProfileHistory) historyList.get(historyList.size() - 1); // Get the last one
					object.setJctBsSubQuestion(obj.getJctGlobalProfileBsSubQtnChanged());
				} else {
					object.setJctBsSubQuestion(vo.getQtnSubDesc());
				}
				object.setJctBsAnswarDesc(vo.getAnswerDescription());
				object.setJctBsSoftDelete(CommonConstants.ENTRY_NOT_DELETED);
				object.setJctBsCreatedBy(vo.getCreatedBy());
				object.setJctBsLastmodifiedBy(vo.getCreatedBy());
				object.setJctBsLastmodifiedTs(insertionDate);
				object.setJctDsCreatedTs(insertionDate);
				object.setJctStatus(status);
				questionnaireDAO.saveAnswers(object);
				
				//Update the login info table for changing the completed status to 1 if not logged out
				if(isCompleted.equals("Y")){
					afterSketchDAO.changeStatus(jobRefNumber);
				}
			}
		} catch(DAOException daoEx) {
			LOGGER.error("----"+daoEx.getLocalizedMessage()+" ----");
			result = "failure";
			throw new JCTException();
		}
		LOGGER.info("<<<< QuestionnaireServiceImpl.saveOrUpdateAnswers");
		return result;
	}
	/**
	 *<p><b>Description:</b> This method fetches the job attributes which will be shown in mapping page.</p>
	 * @param profile id
	 * @return List<JobAttributeDTO>
	 * @throws JCTException
	 * 
	 */
	 public List<JobAttributeDTO> fetchJobAttribute (int profileId) throws JCTException {
		 LOGGER.info(">>>> QuestionnaireServiceImpl.fetchJobAttribute");
		List<JobAttributeDTO> jctJobAttributeList=null;
		try{					
			jctJobAttributeList=questionnaireDAO.getFetchJobAttribute(profileId);
		}catch(DAOException daoException){
			LOGGER.error(daoException.getLocalizedMessage());
		}
		LOGGER.info("<<<< QuestionnaireServiceImpl.fetchJobAttribute");
		return jctJobAttributeList; 
	 }
	 /**
		 * <p><b>Description:</b> This method is used for fetch questionnaire, sub questions and 
		 * answers if answered based on job reference nos and profile id</p>
		 * @param It will take job ref no as a input parameter
		 * @param It will take job profile id as a input parameter
		 * @return It returns List<List>
		 * @throws JCTException -it throws JCTException
		 * 
		 */
	@Transactional(propagation=Propagation.REQUIRED)
	public List<List> fetQuestionnaireDynamic(String jobRefNo, int profileId)
			throws JCTException {
		//This will be returned
		List<List> mainList = null;
		try{
			List<JctBeforeSketchHeader> list = beforeSketchDAO.getList(jobRefNo);
			if (list.size()!=0) {
				JctBeforeSketchHeader parent = (JctBeforeSketchHeader)list.get(0);
				//Check if new or existing
				List<JctBeforeSketchQuestion> actionPlanlist = questionnaireDAO.getList(jobRefNo);
				if (actionPlanlist.size() > 0) {
					mainList = new ArrayList<List>();
					// [[Main Qtn1 subQtn Ans subQtn Ans] [Main Qtn2 subQtn Ans subQtn Ans] .....]
					// Select Distinct Main question List
					List<String> questionnaireMainQtnLst = questionnaireDAO.getDistinctQuestionList(jobRefNo);
					Iterator<String> questionnaireMainQtnLstItr = questionnaireMainQtnLst.iterator();
					while (questionnaireMainQtnLstItr.hasNext()) {
						List<String> innerList = new ArrayList<String>();
						String mainQuestion = (String) questionnaireMainQtnLstItr.next();
						innerList.add(mainQuestion);
						// Select the sub question List 
						List<String> subQtnList = questionnaireDAO.getDistinctSubQuestionListByJrno(mainQuestion, jobRefNo);
						Iterator<String> subQtnListItr = subQtnList.iterator();
						while (subQtnListItr.hasNext()) {
							String subQuestion = (String) subQtnListItr.next();
							innerList.add(subQuestion+"~");
							//Fetch answers given
							List<String> ansList = questionnaireDAO.getAnswerListByJrno(mainQuestion, subQuestion, jobRefNo);
							innerList.add((String) ansList.get(0));
							/*innerList.add((String) ansList.get(1));
							innerList.add((String) ansList.get(2));*/
						}
						mainList.add(innerList);
					}
				} else { //new
					// [[Main Qtn1 subQtn Ans subQtn Ans] [Main Qtn2 subQtn Ans subQtn Ans] .....]
					mainList = new ArrayList<List>();
					// Select Distinct Main question List
					List<String> questionnaireMainQtnList = questionnaireDAO.getQuestionnaireMainQtnList(profileId);
					Iterator<String> questionnaireMainQtnLstItr = questionnaireMainQtnList.iterator();
					while (questionnaireMainQtnLstItr.hasNext()) {
						List<String> innerList = new ArrayList<String>();
						String mainQuestion = (String) questionnaireMainQtnLstItr.next();
						innerList.add(mainQuestion);
						// Select the sub question List
						List<String> subQtnList = questionnaireDAO.getQuestionnaireSubQtnList(mainQuestion, profileId);
						Iterator<String> subQtnListItr = subQtnList.iterator();
						while (subQtnListItr.hasNext()) {
							innerList.add((String) subQtnListItr.next()+"~");
							innerList.add(""); //This is answer. This is set to blank as this it new
							//innerList.add(""); //
							//innerList.add(""); //
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
	 * Method will freeze the content for questionnaire. This is used only in public version.
	 * @param jrNo
	 * @param userName
	 * @param userId
	 * @param profileId
	 * @return
	 * @throws JCTException
	 */
	@Transactional(propagation=Propagation.REQUIRED)
	public String freezeQuestionnaireContent (String jrNo, String user, int userId,
			int profileId) throws JCTException {
		String result = "failure";
		try {
			StringBuilder builder = new StringBuilder("");
			List<String> questionnaireMainQtnList = questionnaireDAO.getQuestionnaireMainQtnList(profileId);
			Iterator<String> questionnaireMainQtnLstItr = questionnaireMainQtnList.iterator();
			while (questionnaireMainQtnLstItr.hasNext()) {
				String mainQuestion = (String) questionnaireMainQtnLstItr.next();
				builder.append(mainQuestion);
				builder.append("@@@");
				List<String> subQtnList = questionnaireDAO.getQuestionnaireSubQtnList(mainQuestion, profileId);
				Iterator<String> subQtnListItr = subQtnList.iterator();
				while (subQtnListItr.hasNext()) {
					builder.append((String) subQtnListItr.next()+"~");
					builder.append(" ");
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
		List<QuestionnaireVO> qtnrVOList = new ArrayList<QuestionnaireVO>();
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
							for (int answerIndex = 0; answerIndex < 1; answerIndex++) {
								
								QuestionnaireVO answerVO = new QuestionnaireVO();
								/** THIS IS FOR PUBLIC VERSION **/
								answerVO.setJctUserId(userId);
								/********************************/
								answerVO.setJobReferenceNumber(jobReferenceNo);
								answerVO.setQuestionDescription(question);
								answerVO.setAnswerDescription(answerArr[answerIndex].replaceAll("-_-", "#").replaceAll(":=:", "(").replaceAll(";=;", ")").replaceAll(":-:", "@").replaceAll(";_;", "?"));
								answerVO.setQtnSubDesc(subQtns);
								answerVO.setCreatedBy(user);
								qtnrVOList.add(answerVO);
							}
							saveOrUpdateAnswers(qtnrVOList, jobReferenceNo, "N");
							result = "success";
						} catch (ArrayIndexOutOfBoundsException ex) {
							LOGGER.error("ArrayIndexOutOfBoundsException..."+ex.getMessage());
						} catch (Exception ex) {
							LOGGER.error("Exception..."+ex.getMessage());
						}
					}
				}
			}
		}
		return result;
	}
	/**
	 * Method will get the count.
	 * @param jrNo
	 * @param userName
	 * @param userId
	 * @param profileId
	 * @return
	 * @throws JCTException
	 */
	@Transactional(propagation=Propagation.REQUIRED)
	public int getQuestCount(String jrNo, String user, int userId, int profileId)
			throws JCTException {
		int count = 0;
		try {
			count = questionnaireDAO.getQuestCount(jrNo, user, userId, profileId).intValue();
		} catch (DAOException ex) {
			LOGGER.error("UNABLE TO FIND THE QUESTIONNAIRE COUNT: "+ex.getMessage());
			throw new JCTException();
		}
		return count;
	}
	
}