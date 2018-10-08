package com.vmware.jct.service.impl;

import java.util.ArrayList;
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
import com.vmware.jct.dao.IAuthenticatorDAO;
import com.vmware.jct.dao.IQuestionnaireDAO;
import com.vmware.jct.dao.ISketchSearchDAO;
import com.vmware.jct.dao.dto.SearchPdfDTO;
import com.vmware.jct.dao.dto.StatusSearchDTO;
import com.vmware.jct.exception.DAOException;
import com.vmware.jct.exception.JCTException;
import com.vmware.jct.model.JctBeforeSketchQuestion;
import com.vmware.jct.model.JctUser;
import com.vmware.jct.service.ISketchSearchService;
import com.vmware.jct.service.vo.MyAccountVO;
import com.vmware.jct.service.vo.SearchResultVO;
import com.vmware.jct.service.vo.PdfSearchVO;
@Service
public class SketchSearchServiceImpl implements ISketchSearchService{

	@Autowired
	private ISketchSearchDAO dao;
	
	@Autowired
	private IAuthenticatorDAO authenticatorDAO;
	
	@Autowired
	private IQuestionnaireDAO questionnaireDAO;
	
	@Autowired
	private IActionPlanDAO actionPlanDAO;
	
	private static final Logger LOGGER = LoggerFactory.getLogger(SketchSearchServiceImpl.class);
	/**
	 *<p><b>Description:</b> This method fetches the before sketch diagram's base 64 string.</p>
	 * @param job title
	 * @param first result
	 * @param max result
	 * @return List<JobAttributeDTO>
	 * @throws JCTException
	 * 
	 */
	@Transactional(propagation=Propagation.REQUIRED)
	public SearchResultVO getBeforeSketchList(String jobTitle,
			int firstResult, int maxResults) throws JCTException{
		LOGGER.info(">>>> SketchSearchServiceImpl.getBeforeSketchList");
		List<String> returnList = null;
		try{
			List<StatusSearchDTO> list = dao.getBeforeSketchList(jobTitle, firstResult, maxResults);
			returnList = new ArrayList<String>();
			Iterator<StatusSearchDTO> itr = list.iterator();
			while(itr.hasNext()){
				StatusSearchDTO obj = (StatusSearchDTO) itr.next();
				String jrno = obj.getJobReferenceNos();
				String snapShotString = obj.getImageString();
				returnList.add(snapShotString);
				//Fetch the job reference nos for the same
				returnList.add(jrno);
				//Put the availablity
				returnList.add(dao.isASAvailable(jrno, jobTitle));
			}
		}catch(DAOException daoEx){
			LOGGER.error(daoEx.getLocalizedMessage());
			daoEx.getMessage();
			throw new JCTException(CommonConstants.FAILURE);
		}
		SearchResultVO vo = new SearchResultVO();
		vo.setSketchStringList(returnList);
		LOGGER.info("<<<< SketchSearchServiceImpl.getBeforeSketchList");
		return vo;
	}
	/**
	 *<p><b>Description:</b> This method fetches the after sketch diagram's base 64 string.</p>
	 * @param job title
	 * @param first result
	 * @param max result
	 * @return List<JobAttributeDTO>
	 * @throws JCTException
	 * 
	 */
	@Transactional(propagation=Propagation.REQUIRED)
	public SearchResultVO getAfterSketchList(String jobTitle,
			int firstResult, int maxResults) throws JCTException {
		LOGGER.info(">>>> SketchSearchServiceImpl.getAfterSketchList");
		List<String> returnList = null;
		try{
			List<StatusSearchDTO> list = dao.getAfterSketchList(jobTitle, firstResult, maxResults);
			returnList = new ArrayList<String>();
			Iterator<StatusSearchDTO> itr = list.iterator();
			while(itr.hasNext()){
				StatusSearchDTO obj = (StatusSearchDTO) itr.next();
				
				String jrno = obj.getJobReferenceNos();
				String snapShotString = obj.getImageString();
				//returnList.add(itr.next().toString());
				returnList.add(snapShotString);
				//Fetch the job reference nos for the same
				//String jobRefNo = dao.getAfterSketchJrno(snapShotString, jobTitle);
				returnList.add(jrno);
				//Put the availablity
				returnList.add(dao.isBSAvailable(jrno, jobTitle));
			}
		}catch(DAOException daoEx){
			daoEx.getMessage();
			LOGGER.error(daoEx.getLocalizedMessage());
			throw new JCTException(CommonConstants.FAILURE);
		}
		SearchResultVO vo = new SearchResultVO();
		vo.setSketchStringList(returnList);
		LOGGER.info("<<<< SketchSearchServiceImpl.getAfterSketchList");
		return vo;
	}
	/**
	 *<p><b>Description:</b> This method fetches the account details of the user.</p>
	 * @param email id
	 * @return MyAccountVO
	 * @throws JCTException
	 * 
	 */
	@Transactional(propagation=Propagation.REQUIRED)
	public MyAccountVO getAccountDetails(String emailId) throws JCTException {
		LOGGER.info(">>>> SketchSearchServiceImpl.getAccountDetails");
		MyAccountVO vo = new MyAccountVO();
		try {
			List<JctUser> userList = authenticatorDAO.getUserList(emailId, 1);
			JctUser userObj = (JctUser) userList.get(0);
			if (null != userObj.getJctUserDetails().getJctUserDetailsFirstName()) {
				String fName = userObj.getJctUserDetails().getJctUserDetailsFirstName();
				if (fName.trim().length() > 0) {
					vo.setFirstName(fName);
				} else {
					vo.setFirstName("");
				}
			} else {
				vo.setFirstName("");
			}
			
			if (null != userObj.getJctUserDetails().getJctUserDetailsLastName()){
				String lName = userObj.getJctUserDetails().getJctUserDetailsLastName();
				if (lName.trim().length() > 0) {
					vo.setLastName(lName);
				} else {
					vo.setLastName("");
				}
			}else {
				vo.setLastName("");
			}
			
			if (null != userObj.getJctUserDetails().getJctUserDetailsFunctionGroup()) {
				String funcGrp = userObj.getJctUserDetails().getJctUserDetailsFunctionGroup();
				if (funcGrp.trim() != "") {
					vo.setFunctionGroup(funcGrp);
				} else {
					vo.setFunctionGroup("");
				}
			} else {
				vo.setFunctionGroup("");
			}
			
			if (null != userObj.getJctUserDetails().getJctUserDetailsLevels()) {
				String jobLvl = userObj.getJctUserDetails().getJctUserDetailsLevels();
				if (jobLvl.trim() != "") {
					vo.setJobLevel(jobLvl);
				} else {
					vo.setJobLevel("");
				}
			} else {
				vo.setJobLevel("");
			}
			if (null != userObj.getJctUserDetails().getJctUserDetailsRegion()) {
				String region = userObj.getJctUserDetails().getJctUserDetailsRegion();
				if (region.trim().length() > 0) {
					vo.setRegion(region);
				} else {
					vo.setRegion("");
				}
			} else {
				vo.setRegion("");
			}
			String sex = userObj.getJctUserDetails().getJctUserDetailsGender();			
			if (null != sex && sex.equals("F")){
				vo.setSex("Female");
			} else if (null != sex && sex.equals("M")){
				vo.setSex("Male");
			} else {
				vo.setSex("");
			}
			String supervisePpl = userObj.getJctUserDetails().getJctUserDetailsSupervisePeople();
			if (null != supervisePpl && supervisePpl.equals("Y")){
				vo.setSupervisePeople("Yes");
			} else if (null != supervisePpl && supervisePpl.equals("N")){
				vo.setSupervisePeople("No");
			} else {
				vo.setSupervisePeople("");
			}
			if (null != userObj.getJctUserDetails().getJctUserDetailsTenure()) {
				String tenure = userObj.getJctUserDetails().getJctUserDetailsTenure();
				if (tenure.trim().length() > 0) {
					vo.setTenure(tenure);
				} else {
					vo.setTenure("");
				}
			} else {
				vo.setTenure("");
			}			
			vo.setJctAccountExpirationDate(userObj.getJctAccountExpirationDate());
			if (null != userObj.getJctUserCustomerId()) {
				String customerId = userObj.getJctUserCustomerId();
				if (customerId.trim().length() > 0) {
					vo.setCustomerId(customerId);
				} else {
					vo.setCustomerId("");
				}
				if (customerId.substring(0, 2).equals("99")) {
					List<MyAccountVO> facilitatorList = authenticatorDAO.getFacilitatorDetails(customerId);
					MyAccountVO faciVo = facilitatorList.get(0);
					vo.setFaciFirstName(faciVo.getFaciFirstName());
					vo.setFaciLastName(faciVo.getFaciLastName());
					vo.setFaciEmail(faciVo.getFaciEmail());
				} else {
					vo.setFaciFirstName("");
					vo.setFaciLastName("");
					vo.setFaciEmail("");
				}
			} else {
				vo.setCustomerId("");
			}
			String onetDesc = dao.getOccupationDesc(userObj.getJctUserDetails().getJctUserOnetOccupation());
			vo.setOnetDesc(onetDesc);
		} catch (DAOException e) {
			LOGGER.error(e.getLocalizedMessage());
			throw new JCTException(e.getLocalizedMessage());
		}
			
		LOGGER.info("<<<< SketchSearchServiceImpl.getAccountDetails");
		return vo;
	}
	/**
	 *<p><b>Description:</b> This method checks whether the after sketch search to be enabled or not.</p>
	 * @param job reference number
	 * @return SearchResultVO
	 * @throws JCTException
	 * 
	 */
	@Transactional(propagation=Propagation.REQUIRED)
	public SearchResultVO getSearchTabsTobeViewed(String jobReferenceNo)
			throws JCTException {
		LOGGER.info(">>>> SketchSearchServiceImpl.getSearchTabsTobeViewed");
		SearchResultVO searchVO = new SearchResultVO();
		String viewableTab = "beforesketch_searchable";
		try {
			//check if same job reference number exists in
			//1: Questionnaire Table
			List<JctBeforeSketchQuestion> qtnList = questionnaireDAO.getList(jobReferenceNo);
			if (qtnList.size() > 0) { // user have completed the before sketch and hence all is search-able
				viewableTab = "aftersketch_searchable";
			}
			searchVO.setTabToBeShown(viewableTab);
		} catch (DAOException daoEx) {
			daoEx.getMessage();
			LOGGER.error(daoEx.getLocalizedMessage());
			throw new JCTException(CommonConstants.FAILURE);
		}
		LOGGER.info("<<<< SketchSearchServiceImpl.getSearchTabsTobeViewed");
		return searchVO;
	}
	/**
	 *<p><b>Description:</b> Method fetches the base 64 string of the diagram which the user has selected to view large.</p>
	 * @param job reference number
	 * @param diagDistinction - before sketch or after sketch
	 * @return SearchResultVO
	 * @throws JCTException
	 * 
	 */
	@Transactional(propagation=Propagation.REQUIRED)
	public SearchResultVO getSelectedDiagrams(String jobReferenceNo,
			String diagDistinction) throws JCTException {
		LOGGER.info(">>>> SketchSearchServiceImpl.getSelectedDiagrams");
		SearchResultVO searchVO = new SearchResultVO();
		try {
			if (diagDistinction.equals("AS")) { //search after sketch diagram
				searchVO.setSelectedDiagram(dao.getASDiagram(jobReferenceNo));
			} else { //search before sketch diagram
				searchVO.setSelectedDiagram(dao.getBSDiagram(jobReferenceNo));
			}
		} catch (DAOException dao) {
			LOGGER.error(dao.getLocalizedMessage());
		}
		
		LOGGER.info("<<<< SketchSearchServiceImpl.getSelectedDiagrams");
		return searchVO;
	}
	/**
	 *<p><b>Description:</b> Method fetches the base 64 string of the diagram which the user has selected to view large.</p>
	 * @param job reference number
	 * @param diagDistinction - before sketch or after sketch
	 * @return SearchResultVO
	 * @throws JCTException
	 * 
	 */
	@Transactional(propagation=Propagation.REQUIRED)
	public SearchResultVO getPrevDiag(String jobReferenceNo, 
			String distinction) throws JCTException {
		LOGGER.info(">>>> SketchSearchServiceImpl.getPrevDiag");
		SearchResultVO searchVO = new SearchResultVO();
		int softDelt = 0;
		try {
			//Check the number of count
			List<String> list = dao.fetchNosOfTimesShared(jobReferenceNo);
			if (list.size() > 1) {
				softDelt = 1;
			}
			
			if (distinction.equals("AS")) { //search after sketch diagram
				String asSnapshotStr = dao.getPrevASDiagramForView(jobReferenceNo, softDelt);
				if (asSnapshotStr != null) {
					searchVO.setSelectedDiagram(asSnapshotStr);
				} else {
					searchVO.setSelectedDiagram("NA");
				}
			} else { //search before sketch diagram
				String bsSnapshotStr = dao.getPrevBSDiagramForView(jobReferenceNo, softDelt);
				if (bsSnapshotStr != null) {
					searchVO.setSelectedDiagram(bsSnapshotStr);
				} else {
					searchVO.setSelectedDiagram("NA");
				}
			}
		} catch (DAOException dao) {
			LOGGER.error(dao.getLocalizedMessage());
		}
		
		LOGGER.info("<<<< SketchSearchServiceImpl.getPrevDiag");
		return searchVO;
	}
	
	@Transactional(propagation=Propagation.REQUIRED)
	public SearchResultVO getAllDiagram(String distinction, int firstResult,
			int maxResults) throws JCTException {
		LOGGER.info(">>>> SketchSearchServiceImpl.getAllDiagram");
		List<String> returnList = null;
		try{
			List<StatusSearchDTO> list = null;
			if (distinction.trim().equals("BS")) {
				list = dao.getBeforeSketchList(firstResult, maxResults);
			} else {
				list = dao.getAfterSketchList(firstResult, maxResults);
			}
			
			returnList = new ArrayList<String>();
			Iterator<StatusSearchDTO> itr = list.iterator();
			while(itr.hasNext()){
				StatusSearchDTO obj = (StatusSearchDTO) itr.next();
				String jrno = obj.getJobReferenceNos();
				String snapShotString = obj.getImageString();
				returnList.add(snapShotString);
				//Fetch the job reference nos for the same
				returnList.add(jrno);
				String occupationCode =  dao.getOccupationDesc(obj.getOccupation());
				returnList.add(occupationCode);
			}
		}catch(DAOException daoEx){
			LOGGER.error(daoEx.getLocalizedMessage());
			daoEx.getMessage();
			throw new JCTException(CommonConstants.FAILURE);
		}
		SearchResultVO vo = new SearchResultVO();
		vo.setSketchStringList(returnList);
		LOGGER.info("<<<< SketchSearchServiceImpl.getAllDiagram");
		return vo;
	}
	
	@Transactional(propagation = Propagation.REQUIRED)
	public Long fetchAllDiagramCount(String distinction) throws JCTException {
		LOGGER.info(">>>>>>>> RemoveDiagramImpl.fetchAllDiagramCount");
		long count = 0;
		try {
			count = dao.fetchAllDiagramCount(distinction);
		} catch (DAOException dao) {
			LOGGER.error("UNABLE TO FETCH THE BEFORE SKETCH AND AFTER SKETCH DIAGRAM: "+dao.getLocalizedMessage());
		}
		LOGGER.info(">>>>>>>> RemoveDiagramImpl.fetchAllDiagramCount");
		return count;
	}
	
	@Transactional(propagation = Propagation.REQUIRED)
	public Long fetchAllDiagramCount(String distinction, String code) throws JCTException {
		LOGGER.info(">>>>>>>> RemoveDiagramImpl.fetchAllDiagramCount");
		long count = 0;
		try {
			count = dao.fetchAllDiagramCount(distinction, code);
		} catch (DAOException dao) {
			LOGGER.error("UNABLE TO FETCH THE BEFORE SKETCH AND AFTER SKETCH DIAGRAM: "+dao.getLocalizedMessage());
		}
		LOGGER.info(">>>>>>>> RemoveDiagramImpl.fetchAllDiagramCount");
		return count;
	}
	
	@Transactional(propagation=Propagation.REQUIRED)
	public SearchResultVO getAllDiagram(String distinction, int firstResult,
			int maxResults, String code) throws JCTException {
		LOGGER.info(">>>> SketchSearchServiceImpl.getAllDiagram");
		List<String> returnList = null;
		try{
			List<StatusSearchDTO> list = null;
			if (distinction.trim().equals("BS")) {
				list = dao.getBeforeSketchList(firstResult, maxResults, code);
			} else {
				list = dao.getAfterSketchList(firstResult, maxResults, code);
			}
			
			returnList = new ArrayList<String>();
			Iterator<StatusSearchDTO> itr = list.iterator();
			while(itr.hasNext()){
				StatusSearchDTO obj = (StatusSearchDTO) itr.next();
				String jrno = obj.getJobReferenceNos();
				String snapShotString = obj.getImageString();
				returnList.add(snapShotString);
				//Fetch the job reference nos for the same
				returnList.add(jrno);
				String occupationCode =  dao.getOccupationDesc(obj.getOccupation());
				returnList.add(occupationCode);
			}
		}catch(DAOException daoEx){
			LOGGER.error(daoEx.getLocalizedMessage());
			daoEx.getMessage();
			throw new JCTException(CommonConstants.FAILURE);
		}
		SearchResultVO vo = new SearchResultVO();
		vo.setSketchStringList(returnList);
		LOGGER.info("<<<< SketchSearchServiceImpl.getAllDiagram");
		return vo;
	}
	@Override
	@Transactional(propagation=Propagation.REQUIRED)
	public List<SearchPdfDTO> fetchOldPdf(int userId) throws JCTException {
		LOGGER.info(">>>> SketchSearchServiceImpl.fetchOldPdf");
		List<SearchPdfDTO> list = new ArrayList<SearchPdfDTO>();
		
		try{
			list =  dao.fetchOldPdf(userId);
		} catch(DAOException e){
			LOGGER.error(e.getLocalizedMessage());			
		}
		LOGGER.info("<<<< SketchSearchServiceImpl.fetchOldPdf");
		return list;
	}
}
