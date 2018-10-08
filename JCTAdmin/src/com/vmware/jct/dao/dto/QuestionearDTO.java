package com.vmware.jct.dao.dto;

import java.io.Serializable;

/**
 * 
 * <p><b>Class name:</b>QuestionearDTO.java</p>
 * <p><b>Author:</b>  InterrIT</p>
 * <p><b>Purpose:</b>QuestionearDTO provides access to questionnaire data which are fetched into them.</p>
 * <p><b>Description:</b></p>
 * <pre>
 * 		<b>Revision History:</b>
 * 		
 * </pre>
 * <p><b></b></p>
 */
public class QuestionearDTO implements Serializable {
	
	private static final long serialVersionUID = 1L;

	private String jctQuestionSubDesc;
	private String jctQuestionsDesc;
	private String userProfileDesc;
	private int activeStatus;
	private String jctQuestionBsas;
	private int jctQuestionsId;
	private int jctQuestionsSoftDelete;
	private int userProfileId;	
	private int jctQuestionsOrder;
	private int jctSubQuestionsOrder;
	
	
	public QuestionearDTO(){
		
	}
	
	public QuestionearDTO(String jctQuestionsDesc){
		this.jctQuestionsDesc=jctQuestionsDesc;		
	}
	
	public QuestionearDTO(String jctQuestionsDesc,String jctQuestionSubDesc){
		this.jctQuestionsDesc=jctQuestionsDesc;
		this.jctQuestionSubDesc=jctQuestionSubDesc;
		
	}
	
	public QuestionearDTO(String jctQuestionsDesc,String jctQuestionSubDesc, String userProfileDesc, String jctQuestionBsas, 
				Integer jctQuestionsId , Integer jctQuestionsSoftDelete, Integer userProfileId, Integer jctQuestionsOrder, Integer jctSubQuestionsOrder){
		this.jctQuestionsDesc=jctQuestionsDesc;
		this.jctQuestionSubDesc=jctQuestionSubDesc;
		this.userProfileDesc=userProfileDesc;
		this.jctQuestionBsas=jctQuestionBsas;
		this.jctQuestionsId=jctQuestionsId;
		this.jctQuestionsSoftDelete=jctQuestionsSoftDelete;
		this.userProfileId=userProfileId;
		this.jctQuestionsOrder=jctQuestionsOrder;
		this.jctSubQuestionsOrder=jctSubQuestionsOrder;
	}
	
	public QuestionearDTO(String jctQuestionsDesc, Integer jctQuestionsOrder){
		this.jctQuestionsDesc=jctQuestionsDesc;
		this.jctQuestionsOrder=jctQuestionsOrder;
	}
		
	/*public QuestionearDTO(String jctQuestionSubDesc, Integer jctSubQuestionsOrder){
		this.jctQuestionSubDesc=jctQuestionSubDesc;
		this.jctSubQuestionsOrder=jctSubQuestionsOrder;
	}*/
	
	public QuestionearDTO(Integer jctQuestionsId){
		this.jctQuestionsId=jctQuestionsId;		
	}
	
	/**
	 * @return the jctQuestionSubDesc
	 */
	public String getJctQuestionSubDesc() {
		return jctQuestionSubDesc;
	}
	/**
	 * @param jctQuestionSubDesc the jctQuestionSubDesc to set
	 */
	public void setJctQuestionSubDesc(String jctQuestionSubDesc) {
		this.jctQuestionSubDesc = jctQuestionSubDesc;
	}
	/**
	 * @return the jctQuestionsDesc
	 */
	public String getJctQuestionsDesc() {
		return jctQuestionsDesc;
	}
	/**
	 * @param jctQuestionsDesc the jctQuestionsDesc to set
	 */
	public void setJctQuestionsDesc(String jctQuestionsDesc) {
		this.jctQuestionsDesc = jctQuestionsDesc;
	}
	/**
	 * 
	 * @return
	 */
	public String getUserProfileDesc() {
		return userProfileDesc;
	}
	/**
	 * 
	 * @param userProfileDesc
	 */
	public void setUserProfileDesc(String userProfileDesc) {
		this.userProfileDesc = userProfileDesc;
	}
	/**
	 * 
	 * @return
	 */
	public int getActiveStatus() {
		return activeStatus;
	}
	/**
	 * 
	 * @param activeStatus
	 */
	public void setActiveStatus(int activeStatus) {
		this.activeStatus = activeStatus;
	}
	/**
	 * 
	 * @return
	 */
	public String getJctQuestionBsas() {
		return jctQuestionBsas;
	}
	/**
	 * 
	 * @param jctQuestionBsas
	 */
	public void setJctQuestionBsas(String jctQuestionBsas) {
		this.jctQuestionBsas = jctQuestionBsas;
	}
	/**
	 * 
	 * @return
	 */
	public int getJctQuestionsId() {
		return jctQuestionsId;
	}
	/**
	 * 
	 * @param jctQuestionsId
	 */
	public void setJctQuestionsId(int jctQuestionsId) {
		this.jctQuestionsId = jctQuestionsId;
	}
	/**
	 * 
	 * @return
	 */
	public int getJctQuestionsSoftDelete() {
		return jctQuestionsSoftDelete;
	}
	/**
	 * 
	 * @param jctQuestionsSoftDelete
	 */
	public void setJctQuestionsSoftDelete(int jctQuestionsSoftDelete) {
		this.jctQuestionsSoftDelete = jctQuestionsSoftDelete;
	}

	public int getUserProfileId() {
		return userProfileId;
	}

	public void setUserProfileId(int userProfileId) {
		this.userProfileId = userProfileId;
	}

	public int getJctQuestionsOrder() {
		return jctQuestionsOrder;
	}

	public void setJctQuestionsOrder(int jctQuestionsOrder) {
		this.jctQuestionsOrder = jctQuestionsOrder;
	}

	public int getJctSubQuestionsOrder() {
		return jctSubQuestionsOrder;
	}

	public void setJctSubQuestionsOrder(int jctSubQuestionsOrder) {
		this.jctSubQuestionsOrder = jctSubQuestionsOrder;
	}
	
}
