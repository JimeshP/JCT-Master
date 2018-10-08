/**
 * 
 */
package com.vmware.jct.dao.dto;

/**
 * 
 * <p><b>Class name:</b> QuestionearDTO.java</p>
 * <p><b>Author:</b> InterraIT</p>
 * <p><b>Purpose:</b> This class acts as a data transfer object.
 * <p><b>Description:</b> This class acts as a data transfer object.. </p>
 * <p><b>Copyrights:</b> 	All rights reserved by Interra IT and should only be used for its internal application development.</p>
 * <p><b>Created Date:</b> 20/Feb/2014</p>
 * <p><b>Revision History:</b>
 * 	<li>6/May/2014: Changed the saving of action plan to dynamic.</li>
 * </p>
 */
public class QuestionearDTO {
	
	public String jctQuestionSubDesc;

	public String jctQuestionsDesc;
	
	
	public QuestionearDTO(){
		
	}
	public QuestionearDTO(String jctQuestionsDesc,String jctQuestionSubDesc){
		this.jctQuestionsDesc=jctQuestionsDesc;
		this.jctQuestionSubDesc=jctQuestionSubDesc;
		
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

	
	
}
