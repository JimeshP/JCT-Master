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
 * <p><b>Class name:</b> JctOnetOccupationList.java</p>
 * <p><b>Author:</b> InterraIT</p>
 * <p><b>Description:</b>  This class basically object reference of jct_onet_occupation_list table </p>
 * <p><b>Copyrights:</b> 	All rights reserved by InterraIT and should only be used for its internal application development.</p>
 * <p><b>Revision History:</b>
 * 	<li>InterraIT - 10/Oct/2014 - Implement comments, introduce Named query </li>
 * </p>
 */
@NamedQueries({
	@NamedQuery(name = "fetchOnetTitle",
				query = "select onetOccupation.jctOnetOccupationTitle from JctOnetOccupationList onetOccupation " +
						"where onetOccupation.jctOnetOccupationCode = :occupationCode and onetOccupation.jctOnetOccupationSoftDelete = 0"),
	@NamedQuery(name = "fetchOnetDataObj",
				query = "from JctOnetOccupationList onetOccupation where  onetOccupation.jctOnetOccupationCode = :occupationCode " +
						"and onetOccupation.jctOnetOccupationSoftDelete = 0"),
	@NamedQuery(name = "fetchAllOnetData",
				query = "Select new com.vmware.jct.dao.dto.OnetOccupationDTO(jctOnetOccupation.jctOnetOccupationId, " +
						"jctOnetOccupation.jctOnetOccupationCode, jctOnetOccupation.jctOnetOccupationTitle, jctOnetOccupation.jctOnetOccupationDesc) " +
						"from JctOnetOccupationList jctOnetOccupation where jctOnetOccupation.jctOnetOccupationSoftDelete = 0"),
	@NamedQuery(name = "getAllActiveOccupationList",
				query = "Select new com.vmware.jct.dao.dto.OccupationListDTO(jctOnetOccupation.jctOnetOccupationCode, " +
						"jctOnetOccupation.jctOnetOccupationTitle, jctOnetOccupation.jctOnetOccupationDesc) " +
						"from JctOnetOccupationList jctOnetOccupation where "						
						+ "jctOnetOccupation.jctOnetOccupationTitle LIKE :searchString1 "
						+ "OR jctOnetOccupation.jctOnetOccupationTitle LIKE :searchString2 "
						+ "OR jctOnetOccupation.jctOnetOccupationTitle LIKE :searchString3 "
						+ "OR jctOnetOccupation.jctOnetOccupationTitle LIKE :searchString4 " 
						+ "OR jctOnetOccupation.jctOnetOccupationCode LIKE :searchString1 "
						+ "OR jctOnetOccupation.jctOnetOccupationCode LIKE :searchString2 "
						+ "OR jctOnetOccupation.jctOnetOccupationCode LIKE :searchString3 "
						+ "OR jctOnetOccupation.jctOnetOccupationCode LIKE :searchString4  AND jctOnetOccupationSoftDelete = 0")
})

@Entity 
@Table(name="jct_onet_occupation_list")
public class JctOnetOccupationList implements Serializable {
	
	private static final long serialVersionUID = 1L;

	@Id
	@GeneratedValue (strategy = GenerationType.IDENTITY)
	@Column(name="jct_onet_occupation_id")
	private int jctOnetOccupationId;
	
	@Column(name="jct_onet_occupation_code")
	private String jctOnetOccupationCode;
	
	@Column(name="jct_onet_occupation_title")
	private String jctOnetOccupationTitle;
	
	@Column(name="jct_onet_occupation_soft_delete")
	private int jctOnetOccupationSoftDelete;
	
	@Column(name="jct_onet_occupation_created_by")
	private String jctOnetOccupationCreatedBy;

	@Temporal(TemporalType.TIMESTAMP)
	@Column(name="jct_onet_occupation_created_ts")
	private Date jctOnetOccupationCreatedTs;
	
	@Column(name="jct_onet_occupation_description")
	private String jctOnetOccupationDesc;
	
	public JctOnetOccupationList() {
	}

	public int getJctOnetOccupationId() {
		return jctOnetOccupationId;
	}

	public void setJctOnetOccupationId(int jctOnetOccupationId) {
		this.jctOnetOccupationId = jctOnetOccupationId;
	}

	public String getJctOnetOccupationCode() {
		return jctOnetOccupationCode;
	}

	public void setJctOnetOccupationCode(String jctOnetOccupationCode) {
		this.jctOnetOccupationCode = jctOnetOccupationCode;
	}

	public String getJctOnetOccupationTitle() {
		return jctOnetOccupationTitle;
	}

	public void setJctOnetOccupationTitle(String jctOnetOccupationTitle) {
		this.jctOnetOccupationTitle = jctOnetOccupationTitle;
	}

	public int getJctOnetOccupationSoftDelete() {
		return jctOnetOccupationSoftDelete;
	}

	public void setJctOnetOccupationSoftDelete(int jctOnetOccupationSoftDelete) {
		this.jctOnetOccupationSoftDelete = jctOnetOccupationSoftDelete;
	}

	public String getJctOnetOccupationCreatedBy() {
		return jctOnetOccupationCreatedBy;
	}

	public void setJctOnetOccupationCreatedBy(String jctOnetOccupationCreatedBy) {
		this.jctOnetOccupationCreatedBy = jctOnetOccupationCreatedBy;
	}

	public Date getJctOnetOccupationCreatedTs() {
		return jctOnetOccupationCreatedTs;
	}

	public void setJctOnetOccupationCreatedTs(Date jctOnetOccupationCreatedTs) {
		this.jctOnetOccupationCreatedTs = jctOnetOccupationCreatedTs;
	}

	public String getJctOnetOccupationDesc() {
		return jctOnetOccupationDesc;
	}

	public void setJctOnetOccupationDesc(String jctOnetOccupationDesc) {
		this.jctOnetOccupationDesc = jctOnetOccupationDesc;
	}
}
