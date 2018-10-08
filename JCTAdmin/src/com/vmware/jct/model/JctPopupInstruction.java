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

@NamedQueries({
	@NamedQuery(name = "getActiveVideos",
				query = "select obj from JctPopupInstruction obj where obj.jctPopupInstructionProfileId = :profileId "
						+ "AND obj.jctPopupInstructionPage = :insPage"),
	@NamedQuery(name = "getActiveVideosByType",
				query = "from JctPopupInstruction obj where obj.jctPopupInstructionProfileId = :profileId " +
						"AND obj.jctPopupInstructionPage = :page " +
						"AND obj.jctPopupInstructionSoftDelete = 0"),
	@NamedQuery(name = "getActiveInstructionByType",
				query = "select obj from JctPopupInstruction obj where obj.jctPopupInstructionProfileId = :profileId "
						+ "AND obj.jctPopupInstructionPage = :page AND obj.jctPopupInstructionType = 'TEXT' " +
						"AND obj.jctPopupInstructionSoftDelete = 0")
})
@Entity
@Table(name="jct_popup_instruction")
public class JctPopupInstruction implements Serializable {
	private static final long serialVersionUID = 1L;
	
	@Id
	@Column(name="jct_popup_instruction_id")
	private int jctPopupInstructionId;

	@Column(name="jct_popup_instruction_type")
	private String jctPopupInstructionType;

	@Column(name="jct_popup_instruction_page")
	private String jctPopupInstructionPage;
	
	@Column(name="jct_popup_instruction_text_before_video")
	private String jctPopupInstructionTextBeforeVideo;
	
	@Column(name="jct_popup_instruction_soft_delete")
	private int jctPopupInstructionSoftDelete;

	@Temporal(TemporalType.TIMESTAMP)
	@Column(name="jct_popup_instruction_creation_date")
	private Date jctPopupInstructionCreationDate;
	
	@Temporal(TemporalType.TIMESTAMP)
	@Column(name="jct_popup_instruction_last_updation_date")
	private Date jctPopupInstructionLastUpdationDate;
	
	@Column(name="jct_popup_instruction_created_by")
	private String jctPopupInstructionCreatedBy;

	@Column(name="jct_popup_instruction_updated_by")
	private String jctPopupInstructionUpdatedBy;
	
	@Column(name="jct_popup_instruction_video_name")
	private String jctPopupInstructionVideoName;
	
	@Column(name="jct_popup_instruction_profile_id")
	private int jctPopupInstructionProfileId;
	
	@Column(name="jct_popup_instruction_text_after_video")
	private String jctPopupInstructionTextAfterVideo;

	public int getJctPopupInstructionId() {
		return jctPopupInstructionId;
	}

	public void setJctPopupInstructionId(int jctPopupInstructionId) {
		this.jctPopupInstructionId = jctPopupInstructionId;
	}

	public String getJctPopupInstructionType() {
		return jctPopupInstructionType;
	}

	public void setJctPopupInstructionType(String jctPopupInstructionType) {
		this.jctPopupInstructionType = jctPopupInstructionType;
	}

	public String getJctPopupInstructionPage() {
		return jctPopupInstructionPage;
	}

	public void setJctPopupInstructionPage(String jctPopupInstructionPage) {
		this.jctPopupInstructionPage = jctPopupInstructionPage;
	}

	public int getJctPopupInstructionSoftDelete() {
		return jctPopupInstructionSoftDelete;
	}

	public void setJctPopupInstructionSoftDelete(int jctPopupInstructionSoftDelete) {
		this.jctPopupInstructionSoftDelete = jctPopupInstructionSoftDelete;
	}

	public Date getJctPopupInstructionCreationDate() {
		return jctPopupInstructionCreationDate;
	}

	public void setJctPopupInstructionCreationDate(
			Date jctPopupInstructionCreationDate) {
		this.jctPopupInstructionCreationDate = jctPopupInstructionCreationDate;
	}

	public Date getJctPopupInstructionLastUpdationDate() {
		return jctPopupInstructionLastUpdationDate;
	}

	public void setJctPopupInstructionLastUpdationDate(
			Date jctPopupInstructionLastUpdationDate) {
		this.jctPopupInstructionLastUpdationDate = jctPopupInstructionLastUpdationDate;
	}

	public String getJctPopupInstructionCreatedBy() {
		return jctPopupInstructionCreatedBy;
	}

	public void setJctPopupInstructionCreatedBy(String jctPopupInstructionCreatedBy) {
		this.jctPopupInstructionCreatedBy = jctPopupInstructionCreatedBy;
	}

	public String getJctPopupInstructionUpdatedBy() {
		return jctPopupInstructionUpdatedBy;
	}

	public void setJctPopupInstructionUpdatedBy(String jctPopupInstructionUpdatedBy) {
		this.jctPopupInstructionUpdatedBy = jctPopupInstructionUpdatedBy;
	}

	public String getJctPopupInstructionVideoName() {
		return jctPopupInstructionVideoName;
	}

	public void setJctPopupInstructionVideoName(String jctPopupInstructionVideoName) {
		this.jctPopupInstructionVideoName = jctPopupInstructionVideoName;
	}

	public int getJctPopupInstructionProfileId() {
		return jctPopupInstructionProfileId;
	}

	public void setJctPopupInstructionProfileId(int jctPopupInstructionProfileId) {
		this.jctPopupInstructionProfileId = jctPopupInstructionProfileId;
	}

	public String getJctPopupInstructionTextBeforeVideo() {
		return jctPopupInstructionTextBeforeVideo;
	}

	public void setJctPopupInstructionTextBeforeVideo(
			String jctPopupInstructionTextBeforeVideo) {
		this.jctPopupInstructionTextBeforeVideo = jctPopupInstructionTextBeforeVideo;
	}

	public String getJctPopupInstructionTextAfterVideo() {
		return jctPopupInstructionTextAfterVideo;
	}

	public void setJctPopupInstructionTextAfterVideo(
			String jctPopupInstructionTextAfterVideo) {
		this.jctPopupInstructionTextAfterVideo = jctPopupInstructionTextAfterVideo;
	}
}