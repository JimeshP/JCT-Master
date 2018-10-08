package com.vmware.jct.model;

import java.io.Serializable;
import javax.persistence.*;


import java.util.Date;


/**
 * 
 * <p><b>Class name:</b> JctRegion.java</p>
 * <p><b>Author:</b> InterraIt</p>
 * <p><b>Description:</b>  This class basically object reference of jct_region table </p>
 * <p><b>Copyrights:</b> 	All rights reserved by Interra IT and should only be used for its internal application development.</p>
 * <p><b>Revision History:</b>
 * 	<li>InterraIt - 31/Jan/2014 - Implement comments, introduce Named query </li>
 * </p>
 */

@NamedQueries({
	@NamedQuery(name = "fetchRegion",
				query = "Select new com.vmware.jct.dao.dto.RegionDTO(jctRegion.jctRegionId,jctRegion.jctRegionName) " +
						"from JctRegion jctRegion order by jctRegion.jctRegionName")


})
@Entity
@Table(name="jct_region")
public class JctRegion implements Serializable {
	private static final long serialVersionUID = 1L;

	@Id
	@Column(name="jct_region_id")
	private int jctRegionId;

	@Column(name="jct_bs_created_by")
	private String jctBsCreatedBy;

	@Column(name="jct_bs_lastmodified_by")
	private String jctBsLastmodifiedBy;
	
	@Temporal(TemporalType.TIMESTAMP)
	@Column(name="jct_bs_lastmodified_ts")
	private Date jctBsLastmodifiedTs;

	@Temporal(TemporalType.TIMESTAMP)
	@Column(name="jct_ds_created_ts")
	private Date jctDsCreatedTs;

	@Column(name="jct_region_desc")
	private String jctregionDesc;

	@Column(name="jct_region_name")
	private String jctRegionName;

	@Column(name="jct_region_soft_delete")
	private byte jctRegionSoftDelete;

	private int version;

	public JctRegion() {
	}

	

	/**
	 * @return the jctRegionId
	 */
	public int getJctRegionId() {
		return jctRegionId;
	}



	/**
	 * @param jctRegionId the jctRegionId to set
	 */
	public void setJctRegionId(int jctRegionId) {
		this.jctRegionId = jctRegionId;
	}



	/**
	 * @return the jctBsLastmodifiedTs
	 */
	public Date getJctBsLastmodifiedTs() {
		return jctBsLastmodifiedTs;
	}



	/**
	 * @param jctBsLastmodifiedTs the jctBsLastmodifiedTs to set
	 */
	public void setJctBsLastmodifiedTs(Date jctBsLastmodifiedTs) {
		this.jctBsLastmodifiedTs = jctBsLastmodifiedTs;
	}



	/**
	 * @return the jctDsCreatedTs
	 */
	public Date getJctDsCreatedTs() {
		return jctDsCreatedTs;
	}



	/**
	 * @param jctDsCreatedTs the jctDsCreatedTs to set
	 */
	public void setJctDsCreatedTs(Date jctDsCreatedTs) {
		this.jctDsCreatedTs = jctDsCreatedTs;
	}



	/**
	 * @return the jctregionDesc
	 */
	public String getJctregionDesc() {
		return jctregionDesc;
	}



	/**
	 * @param jctregionDesc the jctregionDesc to set
	 */
	public void setJctregionDesc(String jctregionDesc) {
		this.jctregionDesc = jctregionDesc;
	}



	/**
	 * @return the jctRegionName
	 */
	public String getJctRegionName() {
		return jctRegionName;
	}



	/**
	 * @param jctRegionName the jctRegionName to set
	 */
	public void setJctRegionName(String jctRegionName) {
		this.jctRegionName = jctRegionName;
	}



	/**
	 * @return the jctRegionSoftDelete
	 */
	public byte getJctRegionSoftDelete() {
		return jctRegionSoftDelete;
	}



	/**
	 * @param jctRegionSoftDelete the jctRegionSoftDelete to set
	 */
	public void setJctRegionSoftDelete(byte jctRegionSoftDelete) {
		this.jctRegionSoftDelete = jctRegionSoftDelete;
	}



	public String getJctBsCreatedBy() {
		return this.jctBsCreatedBy;
	}

	public void setJctBsCreatedBy(String jctBsCreatedBy) {
		this.jctBsCreatedBy = jctBsCreatedBy;
	}

	public String getJctBsLastmodifiedBy() {
		return this.jctBsLastmodifiedBy;
	}

	public void setJctBsLastmodifiedBy(String jctBsLastmodifiedBy) {
		this.jctBsLastmodifiedBy = jctBsLastmodifiedBy;
	}

	
	public int getVersion() {
		return this.version;
	}

	public void setVersion(int version) {
		this.version = version;
	}

}