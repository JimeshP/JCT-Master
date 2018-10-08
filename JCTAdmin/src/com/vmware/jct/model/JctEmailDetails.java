package com.vmware.jct.model;

import java.io.Serializable;
import java.util.Date;

import javax.persistence.Column;
import javax.persistence.Entity;
import javax.persistence.Id;
import javax.persistence.NamedQueries;
import javax.persistence.NamedQuery;
import javax.persistence.Table;
import javax.persistence.Temporal;
import javax.persistence.TemporalType;
/**
 * 
 * <p><b>Class name:</b> JctEmailDetails.java</p>
 * <p><b>Author:</b> InterraIT</p>
 * <p><b>Description:</b>  This class basically object reference of jct_email_details table </p>
 * <p><b>Copyrights:</b> 	All rights reserved by InterraIT and should only be used for its internal application development.</p>
 * <p><b>Revision History:</b>
 * 	<li>InterraIT - 31/Jan/2014 - Implement comments, introduce Named query </li>
 * </p>
 */
@NamedQueries({
	@NamedQuery(name="jctEmailObjByEmail",
				query="from JctEmailDetails where jctRegisteredMail = :emailId and jctMailDispatched = 0"),
	@NamedQuery(name="jctEmailsByEmail",
				query="from JctEmailDetails where jctMailDispatched = 0")
})
@Entity
@Table(name="jct_email_details")
public class JctEmailDetails implements Serializable {
	private static final long serialVersionUID = 1L;
	
	@Id
	@Column(name = "jct_registration_id")
	private int jctRegistrationId;
	
	@Column(name = "jct_registered_mail")
	private String jctRegisteredMail;
	
	@Column(name = "jct_registered_password")
	private String jctRegisteredPassword;
	
	@Column(name = "jct_mail_dispatched")
	private int jctMailDispatched;
	
	@Temporal(TemporalType.TIMESTAMP)
	@Column(name = "jct_account_created_date")
	private Date jctAccountCreatedDate;
	
	@Column(name = "jct_account_created_by")
	private String jctAccountCreatedBy;
	
	@Temporal(TemporalType.TIMESTAMP)
	@Column(name = "jct_mail_sent_date")
	private Date jctMailSentDate;
	
	@Column(name = "jct_user_group_desc")
	private String jctUserGroupDesc;
	
	@Column(name = "jct_user_profile_desc")
	private String jctUserProfileDesc;

	/**
	 * @return the jctRegistrationId
	 */
	public int getJctRegistrationId() {
		return jctRegistrationId;
	}

	/**
	 * @param jctRegistrationId the jctRegistrationId to set
	 */
	public void setJctRegistrationId(int jctRegistrationId) {
		this.jctRegistrationId = jctRegistrationId;
	}

	/**
	 * @return the jctRegisteredMail
	 */
	public String getJctRegisteredMail() {
		return jctRegisteredMail;
	}

	/**
	 * @param jctRegisteredMail the jctRegisteredMail to set
	 */
	public void setJctRegisteredMail(String jctRegisteredMail) {
		this.jctRegisteredMail = jctRegisteredMail;
	}

	/**
	 * @return the jctRegisteredPassword
	 */
	public String getJctRegisteredPassword() {
		return jctRegisteredPassword;
	}

	/**
	 * @param jctRegisteredPassword the jctRegisteredPassword to set
	 */
	public void setJctRegisteredPassword(String jctRegisteredPassword) {
		this.jctRegisteredPassword = jctRegisteredPassword;
	}

	/**
	 * @return the jctMailDispatched
	 */
	public int getJctMailDispatched() {
		return jctMailDispatched;
	}

	/**
	 * @param jctMailDispatched the jctMailDispatched to set
	 */
	public void setJctMailDispatched(int jctMailDispatched) {
		this.jctMailDispatched = jctMailDispatched;
	}

	/**
	 * @return the jctAccountCreatedDate
	 */
	public Date getJctAccountCreatedDate() {
		return jctAccountCreatedDate;
	}

	/**
	 * @param jctAccountCreatedDate the jctAccountCreatedDate to set
	 */
	public void setJctAccountCreatedDate(Date jctAccountCreatedDate) {
		this.jctAccountCreatedDate = jctAccountCreatedDate;
	}

	/**
	 * @return the jctAccountCreatedBy
	 */
	public String getJctAccountCreatedBy() {
		return jctAccountCreatedBy;
	}

	/**
	 * @param jctAccountCreatedBy the jctAccountCreatedBy to set
	 */
	public void setJctAccountCreatedBy(String jctAccountCreatedBy) {
		this.jctAccountCreatedBy = jctAccountCreatedBy;
	}

	/**
	 * @return the jctMailSentDate
	 */
	public Date getJctMailSentDate() {
		return jctMailSentDate;
	}

	/**
	 * @param jctMailSentDate the jctMailSentDate to set
	 */
	public void setJctMailSentDate(Date jctMailSentDate) {
		this.jctMailSentDate = jctMailSentDate;
	}

	/**
	 * @return the jctUserGroupDesc
	 */
	public String getJctUserGroupDesc() {
		return jctUserGroupDesc;
	}

	/**
	 * @param jctUserGroupDesc the jctUserGroupDesc to set
	 */
	public void setJctUserGroupDesc(String jctUserGroupDesc) {
		this.jctUserGroupDesc = jctUserGroupDesc;
	}

	/**
	 * @return the jctUserProfileDesc
	 */
	public String getJctUserProfileDesc() {
		return jctUserProfileDesc;
	}

	/**
	 * @param jctUserProfileDesc the jctUserProfileDesc to set
	 */
	public void setJctUserProfileDesc(String jctUserProfileDesc) {
		this.jctUserProfileDesc = jctUserProfileDesc;
	}
	
}
