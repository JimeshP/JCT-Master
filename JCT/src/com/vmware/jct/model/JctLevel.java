package com.vmware.jct.model;

import java.io.Serializable;
import javax.persistence.*;

import java.util.Date;
/**
 * 
 * <p><b>Class name:</b> JctLevel.java</p>
 * <p><b>Author:</b> InterraIt</p>
 * <p><b>Description:</b>  This class basically object reference of jct_levels table </p>
 * <p><b>Copyrights:</b> 	All rights reserved by Interra IT and should only be used for its internal application development.</p>
 * <p><b>Revision History:</b>
 * 	<li>InterraIt - 31/Jan/2014 - Implement comments, introduce Named query </li>
 * </p>
 */
 @NamedQueries({
		@NamedQuery(name = "fetchJobLevel",
					query = "Select new com.vmware.jct.dao.dto.LevelDTO(jctLevel.jctLevelsId,jctLevel.jctLevelName) " +
							"from JctLevel jctLevel order by jctLevel.jctLevelName")


	})
@Entity
@Table(name="jct_levels")
public class JctLevel implements Serializable {
	private static final long serialVersionUID = 1L;

	@Id
	@Column(name="jct_levels_id")
	private int jctLevelsId;

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

	@Column(name="jct_levels_desc")
	private String jctLevelsDesc;

	@Column(name="jct_levels_soft_delete")
	private byte jctLevelsSoftDelete;

	private int version;

	
	
	@Column(name="jct_level_name")
	private String jctLevelName;

	/*//bi-directional many-to-one association to JctQuestion
	@OneToMany(mappedBy="jctLevel")
	private List<JctQuestion> jctQuestions;
*/
	public JctLevel() {
	}

	public int getJctLevelsId() {
		return this.jctLevelsId;
	}

	public void setJctLevelsId(int jctLevelsId) {
		this.jctLevelsId = jctLevelsId;
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

	public Date getJctBsLastmodifiedTs() {
		return this.jctBsLastmodifiedTs;
	}

	public void setJctBsLastmodifiedTs(Date jctBsLastmodifiedTs) {
		this.jctBsLastmodifiedTs = jctBsLastmodifiedTs;
	}

	public Date getJctDsCreatedTs() {
		return this.jctDsCreatedTs;
	}

	public void setJctDsCreatedTs(Date jctDsCreatedTs) {
		this.jctDsCreatedTs = jctDsCreatedTs;
	}

	public String getJctLevelsDesc() {
		return this.jctLevelsDesc;
	}

	public void setJctLevelsDesc(String jctLevelsDesc) {
		this.jctLevelsDesc = jctLevelsDesc;
	}

	public byte getJctLevelsSoftDelete() {
		return this.jctLevelsSoftDelete;
	}

	public void setJctLevelsSoftDelete(byte jctLevelsSoftDelete) {
		this.jctLevelsSoftDelete = jctLevelsSoftDelete;
	}

	public int getVersion() {
		return this.version;
	}

	public void setVersion(int version) {
		this.version = version;
	}

	

	/*public List<JctQuestion> getJctQuestions() {
		return this.jctQuestions;
	}

	public void setJctQuestions(List<JctQuestion> jctQuestions) {
		this.jctQuestions = jctQuestions;
	}

	public JctQuestion addJctQuestion(JctQuestion jctQuestion) {
		getJctQuestions().add(jctQuestion);
		jctQuestion.setJctLevel(this);

		return jctQuestion;
	}

	public JctQuestion removeJctQuestion(JctQuestion jctQuestion) {
		getJctQuestions().remove(jctQuestion);
		jctQuestion.setJctLevel(null);

		return jctQuestion;
	}*/

	/**
	 * @return the jctLevelName
	 */
	public String getJctLevelName() {
		return jctLevelName;
	}

	/**
	 * @param jctLevelName the jctLevelName to set
	 */
	public void setJctLevelName(String jctLevelName) {
		this.jctLevelName = jctLevelName;
	}


}
