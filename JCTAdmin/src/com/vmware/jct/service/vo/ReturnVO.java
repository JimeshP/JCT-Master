package com.vmware.jct.service.vo;

import java.util.List;

import com.vmware.jct.dao.dto.OccupationListDTO;
/**
 * 
 * <p><b>Class name:</b> ReturnVO.java</p>
 * <p><b>Author:</b> InterraIT</p>
 * <p><b>Purpose:</b> This class acts as a value object. Transfer data from presentation layer to controller and vice versa
 * <p><b>Description:</b>This class acts as a value object. Transfer data from presentation layer to controller and vice versa </p>
 * <p><b>Copyrights:</b> 	All rights reserved by Interra IT and should only be used for its internal application development.</p>
 * <p><b>Created Date:</b> 20/Feb/2014</p>
 * <p><b>Revision History:</b>
 * 	<li>6/May/2014: Changed the saving of action plan to dynamic.</li>
 * </p>
 */
public class ReturnVO {
	private String statusCode;
	private String statusMsg;
	private String statusDesc;
	private String pathFinder;
	private String userName;
	private String email;
	private List list;
	private String bsJson;
	private String initialJson;
	private String existing;
	private String snapShot;
	private byte[] test;
	private int statusCodeInt;
	private List<OccupationListDTO> dtoList;
	/**
	 * 
	 * @return
	 */
	public byte[] getTest() {
		return test;
	}
	/**
	 * 
	 * @param test
	 */
	public void setTest(byte[] test) {
		this.test = test;
	}
	/**
	 * 
	 * @return
	 */
	public String getSnapShot() {
		return snapShot;
	}
	/**
	 * 
	 * @param snapShot
	 */
	public void setSnapShot(String snapShot) {
		this.snapShot = snapShot;
	}
	/**
	 * 
	 * @return
	 */
	public String getExisting() {
		return existing;
	}
	/**
	 * 
	 * @param existing
	 */
	public void setExisting(String existing) {
		this.existing = existing;
	}
	/**
	 * 
	 * @return
	 */
	public String getBsJson() {
		return bsJson;
	}
	/**
	 * 
	 * @param bsJson
	 */
	public void setBsJson(String bsJson) {
		this.bsJson = bsJson;
	}
	/**
	 * 
	 * @return
	 */
	public String getInitialJson() {
		return initialJson;
	}
	/**
	 * 
	 * @param initialJson
	 */
	public void setInitialJson(String initialJson) {
		this.initialJson = initialJson;
	}
	/**
	 * 
	 * @return
	 */
	public List getList() {
		return list;
	}
	/**
	 * 
	 * @param list
	 */
	public void setList(List list) {
		this.list = list;
	}
	/**
	 * 
	 * @return
	 */
	public String getEmail() {
		return email;
	}
	/**
	 * 
	 * @param email
	 */
	public void setEmail(String email) {
		this.email = email;
	}
	/**
	 * 
	 * @return
	 */
	public String getUserName() {
		return userName;
	}
	/**
	 * 
	 * @param userName
	 */
	public void setUserName(String userName) {
		this.userName = userName;
	}
	/**
	 * 
	 * @return
	 */
	public String getStatusCode() {
		return statusCode;
	}
	/**
	 * 
	 * @param statusCode
	 */
	public void setStatusCode(String statusCode) {
		this.statusCode = statusCode;
	}
	/**
	 * 
	 * @return
	 */
	public String getStatusMsg() {
		return statusMsg;
	}
	/**
	 * 
	 * @param statusMsg
	 */
	public void setStatusMsg(String statusMsg) {
		this.statusMsg = statusMsg;
	}
	/**
	 * 
	 * @return
	 */
	public String getStatusDesc() {
		return statusDesc;
	}
	/**
	 * 
	 * @param statusDesc
	 */
	public void setStatusDesc(String statusDesc) {
		this.statusDesc = statusDesc;
	}
	/**
	 * 
	 * @return
	 */
	public String getPathFinder() {
		return pathFinder;
	}
	/**
	 * 
	 * @param pathFinder
	 */
	public void setPathFinder(String pathFinder) {
		this.pathFinder = pathFinder;
	}
	public int getStatusCodeInt() {
		return statusCodeInt;
	}
	public void setStatusCodeInt(int statusCodeInt) {
		this.statusCodeInt = statusCodeInt;
	}
	
	
	public List<OccupationListDTO> getDtoList() {
		return dtoList;
	}
	public void setDtoList(List<OccupationListDTO> dtoList) {
		this.dtoList = dtoList;
	}
	
}
