package com.vmware.jct.dao.dto;

import java.io.Serializable;

public class FacilitatorDetailDTO implements Serializable {
	
	/**
	 * 
	 */
	private static final long serialVersionUID = 1L;
	
	private int jctFacMstrId;
	private int jctFacUserId;
	private String jctFacUserName;
	
	private long jctFacTotalLimit;
	private long jctFacSubscribeLimit;
	
	
	public FacilitatorDetailDTO(){
		
	}
	
	public FacilitatorDetailDTO(int jctFacMstrId, int jctFacUserId, int jctFacTotalLimit, 
			int jctFacSubscribeLimit){
		
		this.jctFacMstrId = jctFacMstrId;
		this.jctFacUserId = jctFacUserId;
		this.jctFacTotalLimit = jctFacTotalLimit;
		this.jctFacSubscribeLimit =jctFacSubscribeLimit;
	}
	
	public FacilitatorDetailDTO(String jctFacUserName, long jctFacTotalLimit, 
			long jctFacSubscribeLimit){
		
		this.jctFacUserName = jctFacUserName;
		this.jctFacTotalLimit = jctFacTotalLimit;
		this.jctFacSubscribeLimit =jctFacSubscribeLimit;
	}

	public FacilitatorDetailDTO( long jctFacTotalLimit, 
			long jctFacSubscribeLimit){
		this.jctFacTotalLimit = jctFacTotalLimit;
		this.jctFacSubscribeLimit =jctFacSubscribeLimit;
	}
	
	
	public int getJctFacMstrId() {
		return jctFacMstrId;
	}
	public void setJctFacMstrId(int jctFacMstrId) {
		this.jctFacMstrId = jctFacMstrId;
	}
	
	public int getJctFacUserId() {
		return jctFacUserId;
	}

	public void setJctFacUserId(int jctFacUserId) {
		this.jctFacUserId = jctFacUserId;
	}
	
	
	public long getJctFacTotalLimit() {
		return jctFacTotalLimit;
	}

	public void setJctFacTotalLimit(long jctFacTotalLimit) {
		this.jctFacTotalLimit = jctFacTotalLimit;
	}

	public long getJctFacSubscribeLimit() {
		return jctFacSubscribeLimit;
	}
	public void setJctFacSubscribeLimit(long jctFacSubscribeLimit) {
		this.jctFacSubscribeLimit = jctFacSubscribeLimit;
	}

	public String getJctFacUserName() {
		return jctFacUserName;
	}

	public void setJctFacUserName(String jctFacUserName) {
		this.jctFacUserName = jctFacUserName;
	}

	
}
