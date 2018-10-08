package com.vmware.jct.model;

import java.io.Serializable;
import java.util.Date;

import javax.persistence.Column;
import javax.persistence.Entity;
import javax.persistence.GeneratedValue;
import javax.persistence.GenerationType;
import javax.persistence.Id;
import javax.persistence.JoinColumn;
import javax.persistence.ManyToOne;
import javax.persistence.NamedQueries;
import javax.persistence.NamedQuery;
import javax.persistence.Table;
import javax.persistence.Temporal;
import javax.persistence.TemporalType;
import javax.persistence.Version;


/**
 * 
 * <p><b>Class name:</b> JctAfterSketchPageoneDetail.java</p>
 * <p><b>Author:</b> InterraIt</p>
 * <p><b>Description:</b>  This class basically object reference of jct_after_sketch_pageone_details table. </p>
 * <p><b>Copyrights:</b> 	All rights reserved by Interra IT and should only be used for its internal application development.</p>
 * <p><b>Revision History:</b>
 * 	<li>InterraIt - 31/Jan/2014 - Implement comments </li>
 * </p>
 */
@NamedQueries({
	@NamedQuery(name="findPassionCount",
				query="Select obj.passionCount from JctAfterSketchPageoneDetail obj " +
						"where obj.jctAsJobrefNo = :jRno and jctAsSoftDelete = 0"),
	@NamedQuery(name="findValueCount",
				query="Select obj.valueCount from JctAfterSketchPageoneDetail obj where " +
						"obj.jctAsJobrefNo = :jRno and jctAsSoftDelete = 0"),
	@NamedQuery(name="findStrengthCount",
				query="Select obj.strengthCount from JctAfterSketchPageoneDetail obj " +
						"where obj.jctAsJobrefNo = :jRno and jctAsSoftDelete = 0"),
	@NamedQuery(name="findASPageOneChildren",
				query="from JctAfterSketchPageoneDetail where jctAsJobrefNo = :refNo and jctAsSoftDelete = 0"),
	@NamedQuery(name="deleteAllASPageOneData", 
				query = "DELETE FROM JctAfterSketchPageoneDetail pon where pon.jctAsJobrefNo = :jrNo")
})

@Entity
@Table(name="jct_after_sketch_pageone_details")
public class JctAfterSketchPageoneDetail implements Serializable {
	private static final long serialVersionUID = 1L;

	@Id
	@GeneratedValue(strategy=GenerationType.IDENTITY)
	@Column(name="jct_as_pageone_details_id")
	private int jctAsPageoneDetailsId;

	@Column(name="jct_as_created_by")
	private String jctAsCreatedBy;

	@Temporal(TemporalType.TIMESTAMP)
	@Column(name="jct_as_created_ts")
	private Date jctAsCreatedTs;

	@Column(name="jct_as_element_code")
	private String jctAsElementCode;

	@Column(name="jct_as_element_id")
	private int jctAsElementId;

	@Column(name="jct_as_jobref_no")
	private String jctAsJobrefNo;

	@Column(name="jct_as_lastmodified_by")
	private String jctAsLastmodifiedBy;

	@Temporal(TemporalType.TIMESTAMP)
	@Column(name="jct_as_lastmodified_ts")
	private Date jctAsLastmodifiedTs;

	@Column(name="jct_as_position")
	private String jctAsPosition;

	@Column(name="jct_as_soft_delete")
	private int jctAsSoftDelete;
	

	@Column(name="jct_as_passion_count")
	private int passionCount;
	
	@Column(name="jct_as_value_count")
	private int valueCount;
	
	@Column(name="jct_as_strength_count")
	private int strengthCount;
	
	public int getPassionCount() {
		return passionCount;
	}

	public void setPassionCount(int passionCount) {
		this.passionCount = passionCount;
	}

	public int getValueCount() {
		return valueCount;
	}

	public void setValueCount(int valueCount) {
		this.valueCount = valueCount;
	}

	public int getStrengthCount() {
		return strengthCount;
	}

	public void setStrengthCount(int strengthCount) {
		this.strengthCount = strengthCount;
	}

	@Version
	private int version;

	//bi-directional many-to-one association to JctAfterSketchHeader
	@ManyToOne
	@JoinColumn(name="jct_as_header_id")
	private JctAfterSketchHeader jctAfterSketchHeader;

	public JctAfterSketchPageoneDetail() {
	}

	public int getJctAsPageoneDetailsId() {
		return this.jctAsPageoneDetailsId;
	}

	public void setJctAsPageoneDetailsId(int jctAsPageoneDetailsId) {
		this.jctAsPageoneDetailsId = jctAsPageoneDetailsId;
	}

	public String getJctAsCreatedBy() {
		return this.jctAsCreatedBy;
	}

	public void setJctAsCreatedBy(String jctAsCreatedBy) {
		this.jctAsCreatedBy = jctAsCreatedBy;
	}

	public Date getJctAsCreatedTs() {
		return this.jctAsCreatedTs;
	}

	public void setJctAsCreatedTs(Date jctAsCreatedTs) {
		this.jctAsCreatedTs = jctAsCreatedTs;
	}

	public String getJctAsElementCode() {
		return this.jctAsElementCode;
	}

	public void setJctAsElementCode(String jctAsElementCode) {
		this.jctAsElementCode = jctAsElementCode;
	}

	public int getJctAsElementId() {
		return this.jctAsElementId;
	}

	public void setJctAsElementId(int jctAsElementId) {
		this.jctAsElementId = jctAsElementId;
	}

	public String getJctAsJobrefNo() {
		return this.jctAsJobrefNo;
	}

	public void setJctAsJobrefNo(String jctAsJobrefNo) {
		this.jctAsJobrefNo = jctAsJobrefNo;
	}

	public String getJctAsLastmodifiedBy() {
		return this.jctAsLastmodifiedBy;
	}

	public void setJctAsLastmodifiedBy(String jctAsLastmodifiedBy) {
		this.jctAsLastmodifiedBy = jctAsLastmodifiedBy;
	}

	public Date getJctAsLastmodifiedTs() {
		return this.jctAsLastmodifiedTs;
	}

	public void setJctAsLastmodifiedTs(Date jctAsLastmodifiedTs) {
		this.jctAsLastmodifiedTs = jctAsLastmodifiedTs;
	}

	public String getJctAsPosition() {
		return this.jctAsPosition;
	}

	public void setJctAsPosition(String jctAsPosition) {
		this.jctAsPosition = jctAsPosition;
	}

	public int getJctAsSoftDelete() {
		return this.jctAsSoftDelete;
	}

	public void setJctAsSoftDelete(int jctAsSoftDelete) {
		this.jctAsSoftDelete = jctAsSoftDelete;
	}

	public int getVersion() {
		return this.version;
	}

	public void setVersion(int version) {
		this.version = version;
	}

	public JctAfterSketchHeader getJctAfterSketchHeader() {
		return this.jctAfterSketchHeader;
	}

	public void setJctAfterSketchHeader(JctAfterSketchHeader jctAfterSketchHeader) {
		this.jctAfterSketchHeader = jctAfterSketchHeader;
	}

}