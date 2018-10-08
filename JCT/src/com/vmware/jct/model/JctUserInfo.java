package com.vmware.jct.model;

import java.io.Serializable;
import java.util.Date;

import javax.persistence.Column;
import javax.persistence.Entity;
import javax.persistence.GeneratedValue;
import javax.persistence.GenerationType;
import javax.persistence.Id;
import javax.persistence.NamedQueries;
import javax.persistence.NamedQuery;
import javax.persistence.Table;
import javax.persistence.Temporal;
import javax.persistence.TemporalType;
/**
 * 
 * <p><b>Class name:</b> JctUserInfo.java</p>
 * <p><b>Author:</b> InterraIT</p>
 * <p><b>Description:</b>  This class basically object reference of jct_user_info table </p>
 * <p><b>Copyrights:</b> 	All rights reserved by Interra IT and should only be used for its internal application development.</p>
 * <p><b>Revision History:</b>
 * 	<li>InterraIT - 31/Jan/2014 - Implement comments, introduce Named query </li>
 * </p>
 */
@NamedQueries({
	@NamedQuery(name = "fetchFailedInstance",
				query = "from JctUserInfo info where jctUserId= :userId and jctUserEmail = :email and jctSoftDelete = 0"),
	@NamedQuery(name = "fetchFailedInstanceByEmail", 
				query = "from JctUserInfo info where jctUserEmail = :email and jctSoftDelete = 0")
})
@Entity
@Table(name="jct_user_info")
public class JctUserInfo implements Serializable{
	private static final long serialVersionUID = 1L;
	
	@Id
	@GeneratedValue(strategy=GenerationType.IDENTITY)
	@Column(name = "jct_user_info_id")
	private int jctUserInfoId;
	
	@Column(name = "jct_user_id")
	private int jctUserId;
	
	@Column(name = "jct_user_email")
	private String jctUserEmail;
	
	@Temporal(TemporalType.TIMESTAMP)
	@Column(name = "jct_created_ts")
	private Date jctCreatedTs;
	
	@Column(name = "jct_no_of_count")
	private int jctNoOfCount;
	
	@Column(name = "jct_soft_delete")
	private int jctSoftDelete;

	/**
	 * @return the jctUserInfoId
	 */
	public int getJctUserInfoId() {
		return jctUserInfoId;
	}

	/**
	 * @param jctUserInfoId the jctUserInfoId to set
	 */
	public void setJctUserInfoId(int jctUserInfoId) {
		this.jctUserInfoId = jctUserInfoId;
	}

	/**
	 * @return the jctUserId
	 */
	public int getJctUserId() {
		return jctUserId;
	}

	/**
	 * @param jctUserId the jctUserId to set
	 */
	public void setJctUserId(int jctUserId) {
		this.jctUserId = jctUserId;
	}

	/**
	 * @return the jctUserEmail
	 */
	public String getJctUserEmail() {
		return jctUserEmail;
	}

	/**
	 * @param jctUserEmail the jctUserEmail to set
	 */
	public void setJctUserEmail(String jctUserEmail) {
		this.jctUserEmail = jctUserEmail;
	}

	/**
	 * @return the jctCreatedTs
	 */
	public Date getJctCreatedTs() {
		return jctCreatedTs;
	}

	/**
	 * @param jctCreatedTs the jctCreatedTs to set
	 */
	public void setJctCreatedTs(Date jctCreatedTs) {
		this.jctCreatedTs = jctCreatedTs;
	}

	/**
	 * @return the jctNoOfCount
	 */
	public int getJctNoOfCount() {
		return jctNoOfCount;
	}

	/**
	 * @param jctNoOfCount the jctNoOfCount to set
	 */
	public void setJctNoOfCount(int jctNoOfCount) {
		this.jctNoOfCount = jctNoOfCount;
	}

	/**
	 * @return the jctSoftDelete
	 */
	public int getJctSoftDelete() {
		return jctSoftDelete;
	}

	/**
	 * @param jctSoftDelete the jctSoftDelete to set
	 */
	public void setJctSoftDelete(int jctSoftDelete) {
		this.jctSoftDelete = jctSoftDelete;
	}
	
	
	
}
