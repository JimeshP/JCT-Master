package com.vmware.jct.model;

import java.io.Serializable;
import javax.persistence.*;
import java.sql.Timestamp;


/**
 * 
 * <p><b>Class name:</b> JctAfterSketch.java</p>
 * <p><b>Author:</b> InterraIt</p>
 * <p><b>Description:</b>  This class basically object reference of jct_after_sketch table. Store all after sketch information  in this table </p>
 * <p><b>Copyrights:</b> 	All rights reserved by Interra IT and should only be used for its internal application development.</p>
 * <p><b>Revision History:</b>
 * 	<li>InterraIt - 31/Jan/2014 - Implement comments </li>
 * </p>
 */
@Entity
@Table(name="jct_after_sketch")
public class JctAfterSketch implements Serializable {
	private static final long serialVersionUID = 1L;

	@Id
	@Column(name="jct_as_id")
	private int jctAsId;

	@Column(name="jct_as_created_by")
	private String jctAsCreatedBy;

	@Column(name="jct_as_created_ts")
	private Timestamp jctAsCreatedTs;

	@Column(name="jct_as_element_code")
	private String jctAsElementCode;

	@Column(name="jct_as_element_desc")
	private String jctAsElementDesc;

	@Column(name="jct_as_element_energy")
	private int jctAsElementEnergy;

	@Column(name="jct_as_element_time")
	private int jctAsElementTime;

	@Column(name="jct_as_jobref_no")
	private String jctAsJobrefNo;

	@Lob
	@Column(name="jct_as_json")
	private String jctAsJson;

	@Column(name="jct_as_lastmodified_by")
	private String jctAsLastmodifiedBy;

	@Column(name="jct_as_lastmodified_ts")
	private Timestamp jctAsLastmodifiedTs;

	@Column(name="jct_as_position")
	private String jctAsPosition;

	@Column(name="jct_as_role_desc")
	private String jctAsRoleDesc;

	@Column(name="jct_as_role_id")
	private int jctAsRoleId;

	private int version;

	//bi-directional many-to-one association to JctStatus
	@ManyToOne
	@JoinColumn(name="jct_as_status_id")
	private JctStatus jctStatus;

	//bi-directional many-to-one association to JctJobAttribute
	@ManyToOne
	@JoinColumn(name="jct_as_element_id")
	private JctJobAttribute jctJobAttribute;

	public JctAfterSketch() {
	}

	public int getJctAsId() {
		return this.jctAsId;
	}

	public void setJctAsId(int jctAsId) {
		this.jctAsId = jctAsId;
	}

	public String getJctAsCreatedBy() {
		return this.jctAsCreatedBy;
	}

	public void setJctAsCreatedBy(String jctAsCreatedBy) {
		this.jctAsCreatedBy = jctAsCreatedBy;
	}

	public Timestamp getJctAsCreatedTs() {
		return this.jctAsCreatedTs;
	}

	public void setJctAsCreatedTs(Timestamp jctAsCreatedTs) {
		this.jctAsCreatedTs = jctAsCreatedTs;
	}

	public String getJctAsElementCode() {
		return this.jctAsElementCode;
	}

	public void setJctAsElementCode(String jctAsElementCode) {
		this.jctAsElementCode = jctAsElementCode;
	}

	public String getJctAsElementDesc() {
		return this.jctAsElementDesc;
	}

	public void setJctAsElementDesc(String jctAsElementDesc) {
		this.jctAsElementDesc = jctAsElementDesc;
	}

	public int getJctAsElementEnergy() {
		return this.jctAsElementEnergy;
	}

	public void setJctAsElementEnergy(int jctAsElementEnergy) {
		this.jctAsElementEnergy = jctAsElementEnergy;
	}

	public int getJctAsElementTime() {
		return this.jctAsElementTime;
	}

	public void setJctAsElementTime(int jctAsElementTime) {
		this.jctAsElementTime = jctAsElementTime;
	}

	public String getJctAsJobrefNo() {
		return this.jctAsJobrefNo;
	}

	public void setJctAsJobrefNo(String jctAsJobrefNo) {
		this.jctAsJobrefNo = jctAsJobrefNo;
	}

	public String getJctAsJson() {
		return this.jctAsJson;
	}

	public void setJctAsJson(String jctAsJson) {
		this.jctAsJson = jctAsJson;
	}

	public String getJctAsLastmodifiedBy() {
		return this.jctAsLastmodifiedBy;
	}

	public void setJctAsLastmodifiedBy(String jctAsLastmodifiedBy) {
		this.jctAsLastmodifiedBy = jctAsLastmodifiedBy;
	}

	public Timestamp getJctAsLastmodifiedTs() {
		return this.jctAsLastmodifiedTs;
	}

	public void setJctAsLastmodifiedTs(Timestamp jctAsLastmodifiedTs) {
		this.jctAsLastmodifiedTs = jctAsLastmodifiedTs;
	}

	public String getJctAsPosition() {
		return this.jctAsPosition;
	}

	public void setJctAsPosition(String jctAsPosition) {
		this.jctAsPosition = jctAsPosition;
	}

	public String getJctAsRoleDesc() {
		return this.jctAsRoleDesc;
	}

	public void setJctAsRoleDesc(String jctAsRoleDesc) {
		this.jctAsRoleDesc = jctAsRoleDesc;
	}

	public int getJctAsRoleId() {
		return this.jctAsRoleId;
	}

	public void setJctAsRoleId(int jctAsRoleId) {
		this.jctAsRoleId = jctAsRoleId;
	}

	public int getVersion() {
		return this.version;
	}

	public void setVersion(int version) {
		this.version = version;
	}

	public JctStatus getJctStatus() {
		return this.jctStatus;
	}

	public void setJctStatus(JctStatus jctStatus) {
		this.jctStatus = jctStatus;
	}

	public JctJobAttribute getJctJobAttribute() {
		return this.jctJobAttribute;
	}

	public void setJctJobAttribute(JctJobAttribute jctJobAttribute) {
		this.jctJobAttribute = jctJobAttribute;
	}

}