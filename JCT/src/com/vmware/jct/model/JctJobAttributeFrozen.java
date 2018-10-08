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
 * <p><b>Class name:</b> JctJobAttributeFrozen.java</p>
 * <p><b>Author:</b> InterraIt</p>
 * <p><b>Description:</b>  This class basically object reference of jct_job_attribute_frozen table </p>
 * <p><b>Copyrights:</b> 	All rights reserved by Interra IT and should only be used for its internal application development.</p>
 * <p><b>Revision History:</b>
 * 	<li>InterraIt - 13/Oct/2014 - Implement comments, introduce Named query </li>
 * </p>
 */
@NamedQueries({
	@NamedQuery(name = "fetchFrozenAttrs",
				query = "select new com.vmware.jct.dao.dto.JobAttributeDTO(jctAttributeFrozen.jctJobAttributeId,"
						+ "jctAttributeFrozen.jctJobAttributeCode,"
						+ "jctAttributeFrozen.jctJobAttributeName, "
						+ "jctAttributeFrozen.jctJobAttributeDesc, "
						+ "jctAttributeFrozen.jctJobAttributeProfileDesc, "
						+ "jctAttributeFrozen.jctJobAttributeOrder) " 
						+ "from JctJobAttributeFrozen jctAttributeFrozen where jctAttributeFrozen.jctJobAttributeProfileId =:profileId and "
						+ "jctAttributeFrozen.jctJobAttributeSoftDelete = 0 and "
						+ "jctAttributeFrozen.jctJobAttributeCreatedFor = :createdFor and "
						+ "jctAttributeFrozen.jctJobAttributeUserId = :userId"),
	@NamedQuery(name = "getExtingFrozenJobAttrs", 
				query = "SELECT COUNT(jctJobAttributeCode) FROM JctJobAttributeFrozen frz where "
						+ "frz.jctJobAttributeUserId = :userId AND frz.jctJobAttributeCreatedFor = :createdFor AND "
						+ "frz.jctJobAttributeProfileId = :profileId AND frz.jctJobAttributeSoftDelete = 0"),
	@NamedQuery(name = "getFrozenAttributeById",
				query = "SELECT jctJobAttributeName FROM JctJobAttributeFrozen WHERE jctJobAttributeId = :attributeId")
})
@Entity
@Table(name="jct_job_attribute_frozen")
public class JctJobAttributeFrozen implements Serializable {
	private static final long serialVersionUID = 1L;
	
	@Id
	@GeneratedValue(strategy=GenerationType.IDENTITY)
	@Column(name="jct_job_attribute_id")
	private int jctJobAttributeId;

	@Column(name="jct_job_attribute_code")
	private String jctJobAttributeCode;
	
	@Column(name="jct_job_attribute_name")
	private String jctJobAttributeName;
	
	@Column(name="jct_job_attribute_profile_id")
	private int jctJobAttributeProfileId;
	
	@Column(name="jct_job_attribute_profile_desc")
	private String jctJobAttributeProfileDesc;
	
	@Column(name="jct_job_attribute_desc")
	private String jctJobAttributeDesc;
	
	@Column(name="jct_job_attribute_order")
	private int jctJobAttributeOrder;

	@Column(name="jct_job_attribute_soft_delete")
	private int jctJobAttributeSoftDelete;
	
	@Column(name="jct_job_attribute_created_by")
	private String jctJobAttributeCreatedBy;

	@Temporal(TemporalType.TIMESTAMP)
	@Column(name="jct_job_attribute_created_ts")
	private Date jctJobAttributeCreatedTs;
	
	@Column(name="jct_job_attribute_modified_by")
	private String jctJobAttributemodifiedBy;

	@Temporal(TemporalType.TIMESTAMP)
	@Column(name="jct_job_attribute_modified_ts")
	private Date jctJobAttributemodifiedTs;
	
	@Column(name="jct_job_attribute_created_for")
	private String jctJobAttributeCreatedFor;
	
	@Column(name="jct_job_attribute_user_id")
	private int jctJobAttributeUserId;

	public int getJctJobAttributeId() {
		return jctJobAttributeId;
	}

	public void setJctJobAttributeId(int jctJobAttributeId) {
		this.jctJobAttributeId = jctJobAttributeId;
	}

	public String getJctJobAttributeCode() {
		return jctJobAttributeCode;
	}

	public void setJctJobAttributeCode(String jctJobAttributeCode) {
		this.jctJobAttributeCode = jctJobAttributeCode;
	}

	public String getJctJobAttributeName() {
		return jctJobAttributeName;
	}

	public void setJctJobAttributeName(String jctJobAttributeName) {
		this.jctJobAttributeName = jctJobAttributeName;
	}

	public int getJctJobAttributeProfileId() {
		return jctJobAttributeProfileId;
	}

	public void setJctJobAttributeProfileId(int jctJobAttributeProfileId) {
		this.jctJobAttributeProfileId = jctJobAttributeProfileId;
	}

	public String getJctJobAttributeProfileDesc() {
		return jctJobAttributeProfileDesc;
	}

	public void setJctJobAttributeProfileDesc(String jctJobAttributeProfileDesc) {
		this.jctJobAttributeProfileDesc = jctJobAttributeProfileDesc;
	}

	public String getJctJobAttributeDesc() {
		return jctJobAttributeDesc;
	}

	public void setJctJobAttributeDesc(String jctJobAttributeDesc) {
		this.jctJobAttributeDesc = jctJobAttributeDesc;
	}

	public int getJctJobAttributeOrder() {
		return jctJobAttributeOrder;
	}

	public void setJctJobAttributeOrder(int jctJobAttributeOrder) {
		this.jctJobAttributeOrder = jctJobAttributeOrder;
	}

	public int getJctJobAttributeSoftDelete() {
		return jctJobAttributeSoftDelete;
	}

	public void setJctJobAttributeSoftDelete(int jctJobAttributeSoftDelete) {
		this.jctJobAttributeSoftDelete = jctJobAttributeSoftDelete;
	}

	public String getJctJobAttributeCreatedBy() {
		return jctJobAttributeCreatedBy;
	}

	public void setJctJobAttributeCreatedBy(String jctJobAttributeCreatedBy) {
		this.jctJobAttributeCreatedBy = jctJobAttributeCreatedBy;
	}

	public Date getJctJobAttributeCreatedTs() {
		return jctJobAttributeCreatedTs;
	}

	public void setJctJobAttributeCreatedTs(Date jctJobAttributeCreatedTs) {
		this.jctJobAttributeCreatedTs = jctJobAttributeCreatedTs;
	}

	public String getJctJobAttributemodifiedBy() {
		return jctJobAttributemodifiedBy;
	}

	public void setJctJobAttributemodifiedBy(String jctJobAttributemodifiedBy) {
		this.jctJobAttributemodifiedBy = jctJobAttributemodifiedBy;
	}

	public Date getJctJobAttributemodifiedTs() {
		return jctJobAttributemodifiedTs;
	}

	public void setJctJobAttributemodifiedTs(Date jctJobAttributemodifiedTs) {
		this.jctJobAttributemodifiedTs = jctJobAttributemodifiedTs;
	}

	public String getJctJobAttributeCreatedFor() {
		return jctJobAttributeCreatedFor;
	}

	public void setJctJobAttributeCreatedFor(String jctJobAttributeCreatedFor) {
		this.jctJobAttributeCreatedFor = jctJobAttributeCreatedFor;
	}

	public int getJctJobAttributeUserId() {
		return jctJobAttributeUserId;
	}

	public void setJctJobAttributeUserId(int jctJobAttributeUserId) {
		this.jctJobAttributeUserId = jctJobAttributeUserId;
	}
}
