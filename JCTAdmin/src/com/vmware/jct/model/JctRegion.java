package com.vmware.jct.model;

import java.io.Serializable;
import javax.persistence.*;


import java.util.Date;


/**
 * 
 * <p><b>Class name:</b> JctRegion.java</p>
 * <p><b>Author:</b> InterraIT</p>
 * <p><b>Description:</b>  This class basically object reference of jct_region table </p>
 * <p><b>Copyrights:</b> 	All rights reserved by InterraIT and should only be used for its internal application development.</p>
 * <p><b>Revision History:</b>
 * 	<li>InterraIT - 31/Jan/2014 - Implement comments, introduce Named query </li>
 * </p>
 */

@NamedQueries({
	@NamedQuery(name = "fetchRegion",
				query = "Select new com.vmware.jct.dao.dto.RegionDTO(jctRegion.jctRegionId,jctRegion.jctRegionName) " +
						"from JctRegion jctRegion where jctRegion.jctRegionSoftDelete = 0 order by jctRegion.jctRegionOrder"),
	@NamedQuery(name = "fetchRegionById",
				query = "from JctRegion jctRegion where  jctRegion.jctRegionId = :pkId and jctRegionSoftDelete = 0"),
	@NamedQuery(name = "fetchRegionByDesc",
				query = "from JctRegion jctRegion where jctRegion.jctRegionSoftDelete = 0 " +
						"and UPPER(jctRegion.jctregionDesc)= :regionName"),
	@NamedQuery(name = "fetchOrderRG",
				query = "select max(jctRegion.jctRegionOrder) from JctRegion jctRegion " +
						"where jctRegion.jctRegionSoftDelete = 0"),
	@NamedQuery(name = "fetchToUpdateRG",
				query = "from JctRegion jctRegion where jctRegion.jctRegionName = :regionName and jctRegion.jctRegionSoftDelete = 0"),
	@NamedQuery(name = "fetchRegionBybOrder",
				query = "select new com.vmware.jct.dao.dto.RegionDTO(jctRegion.jctRegionId) from JctRegion jctRegion " +
						"where jctRegion.jctRegionOrder > :regionOrder and jctRegion.jctRegionSoftDelete = 0 ")


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
	private int jctRegionSoftDelete;
	
	@Column(name="jct_region_order")
	private int jctRegionOrder;

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
	public int getJctRegionSoftDelete() {
		return jctRegionSoftDelete;
	}

	/**
	 * @param jctRegionSoftDelete the jctRegionSoftDelete to set
	 */
	public void setJctRegionSoftDelete(int jctRegionSoftDelete) {
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



	public int getJctRegionOrder() {
		return jctRegionOrder;
	}



	public void setJctRegionOrder(int jctRegionOrder) {
		this.jctRegionOrder = jctRegionOrder;
	}

}