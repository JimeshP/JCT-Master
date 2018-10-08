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
import javax.persistence.Transient;
@NamedQueries({
	@NamedQuery(name  = "getGlobalRQChangeBySubQtn",
				query = "from JctGlobalProfileHistory where jctGlobalProfileBsSubQtnOriginal = :subQtn and jctGlobalProfileSoftDelete=0 and "
						+ "jctGlobalProfileBsMainQtn = :mainQtn and jctGlobalProfileCreatedTimestamp between :date1 and :date2"),
	@NamedQuery(name  = "getGlobalAPChangeBySubQtn",
				query = "from JctGlobalProfileHistory where jctGlobalProfileApSubQtnOriginal = :subQtn and jctGlobalProfileSoftDelete=0 and "
						+ "jctGlobalProfileApMainQtn = :mainQtn and jctGlobalProfileCreatedTimestamp between :date1 and :date2"),
	@NamedQuery(name  = "getGlobalAttributeChangesPageOneByTimeStamp", 
				query = "from JctGlobalProfileHistory where jctGlobalProfileIsAttribute = 'Y' and jctGlobalProfileAttributeType IN ('STR', 'VAL', 'PAS') "
						+ "AND jctGlobalProfileCreatedTimestamp between :date1 and :date2")
})
@Entity
@Table(name="jct_global_profile_history")
public class JctGlobalProfileHistory implements Serializable {

	private static final long serialVersionUID = 1L;
	
	@Id
	@GeneratedValue(strategy=GenerationType.IDENTITY)
	@Column(name="jct_global_profile_id")
	private int jctGlobalProfileId;

	@Column(name="jct_global_profile_bs_sub_qtn_original")
	private String jctGlobalProfileBsSubQtnOriginal;
	
	@Column(name="jct_global_profile_bs_sub_qtn_changed")
	private String jctGlobalProfileBsSubQtnChanged;
	
	@Column(name="jct_global_profile_ap_sub_qtn_original")
	private String jctGlobalProfileApSubQtnOriginal;
	
	@Column(name="jct_global_profile_ap_sub_qtn_changed")
	private String jctGlobalProfileApSubQtnChanged;
	
	@Column(name="jct_global_profile_str_original")
	private String jctGlobalProfileStrOriginal;
	
	@Column(name="jct_global_profile_str_changed")
	private String jctGlobalProfileStrChanged;
	
	@Column(name="jct_global_profile_str_desc_original")
	private String jctGlobalProfileStrDescOriginal;
	
	@Column(name="jct_global_profile_str_desc_changed")
	private String jctGlobalProfileStrDescChanged;
	
	@Column(name="jct_global_profile_val_original")
	private String jctGlobalProfileValOriginal;
	
	@Column(name="jct_global_profile_val_changed")
	private String jctGlobalProfileValChanged;
	
	@Column(name="jct_global_profile_val_desc_original")
	private String jctGlobalProfileValDescOriginal;
	
	@Column(name="jct_global_profile_val_desc_changed")
	private String jctGlobalProfileValDescChanged;
	
	@Column(name="jct_global_profile_pas_original")
	private String jctGlobalProfilePasOriginal;
	
	@Column(name="jct_global_profile_pas_changed")
	private String jctGlobalProfilePasChanged;
	
	@Column(name="jct_global_profile_pas_desc_original")
	private String jctGlobalProfilePasDescOriginal;
	
	@Column(name="jct_global_profile_pas_desc_changed")
	private String jctGlobalProfilePasDescChanged;

	@Column(name="jct_global_profile_soft_delete")
	private int jctGlobalProfileSoftDelete;

	@Temporal(TemporalType.TIMESTAMP)
	@Column(name="jct_global_profile_created_timestamp")
	private Date jctGlobalProfileCreatedTimestamp;
	
	@Transient
	@Temporal(TemporalType.TIMESTAMP)
	private Date presentTimeStamp;
	
	@Column(name = "jct_global_profile_bs_main_qtn")
	private String jctGlobalProfileBsMainQtn;
	
	@Column(name = "jct_global_profile_ap_main_qtn")
	private String jctGlobalProfileApMainQtn;
	
	@Column(name = "jct_global_profile_is_attribute")
	private String jctGlobalProfileIsAttribute;
	
	@Column(name = "jct_global_profile_attribute_type")
	private String jctGlobalProfileAttributeType;

	public int getJctGlobalProfileId() {
		return jctGlobalProfileId;
	}

	public void setJctGlobalProfileId(int jctGlobalProfileId) {
		this.jctGlobalProfileId = jctGlobalProfileId;
	}

	public String getJctGlobalProfileBsSubQtnOriginal() {
		return jctGlobalProfileBsSubQtnOriginal;
	}

	public void setJctGlobalProfileBsSubQtnOriginal(
			String jctGlobalProfileBsSubQtnOriginal) {
		this.jctGlobalProfileBsSubQtnOriginal = jctGlobalProfileBsSubQtnOriginal;
	}

	public String getJctGlobalProfileBsSubQtnChanged() {
		return jctGlobalProfileBsSubQtnChanged;
	}

	public void setJctGlobalProfileBsSubQtnChanged(
			String jctGlobalProfileBsSubQtnChanged) {
		this.jctGlobalProfileBsSubQtnChanged = jctGlobalProfileBsSubQtnChanged;
	}

	public String getJctGlobalProfileApSubQtnOriginal() {
		return jctGlobalProfileApSubQtnOriginal;
	}

	public void setJctGlobalProfileApSubQtnOriginal(
			String jctGlobalProfileApSubQtnOriginal) {
		this.jctGlobalProfileApSubQtnOriginal = jctGlobalProfileApSubQtnOriginal;
	}

	public String getJctGlobalProfileApSubQtnChanged() {
		return jctGlobalProfileApSubQtnChanged;
	}

	public void setJctGlobalProfileApSubQtnChanged(
			String jctGlobalProfileApSubQtnChanged) {
		this.jctGlobalProfileApSubQtnChanged = jctGlobalProfileApSubQtnChanged;
	}

	public String getJctGlobalProfileStrOriginal() {
		return jctGlobalProfileStrOriginal;
	}

	public void setJctGlobalProfileStrOriginal(String jctGlobalProfileStrOriginal) {
		this.jctGlobalProfileStrOriginal = jctGlobalProfileStrOriginal;
	}

	public String getJctGlobalProfileStrChanged() {
		return jctGlobalProfileStrChanged;
	}

	public void setJctGlobalProfileStrChanged(String jctGlobalProfileStrChanged) {
		this.jctGlobalProfileStrChanged = jctGlobalProfileStrChanged;
	}

	public String getJctGlobalProfileStrDescOriginal() {
		return jctGlobalProfileStrDescOriginal;
	}

	public void setJctGlobalProfileStrDescOriginal(
			String jctGlobalProfileStrDescOriginal) {
		this.jctGlobalProfileStrDescOriginal = jctGlobalProfileStrDescOriginal;
	}

	public String getJctGlobalProfileStrDescChanged() {
		return jctGlobalProfileStrDescChanged;
	}

	public void setJctGlobalProfileStrDescChanged(
			String jctGlobalProfileStrDescChanged) {
		this.jctGlobalProfileStrDescChanged = jctGlobalProfileStrDescChanged;
	}

	public String getJctGlobalProfileValOriginal() {
		return jctGlobalProfileValOriginal;
	}

	public void setJctGlobalProfileValOriginal(String jctGlobalProfileValOriginal) {
		this.jctGlobalProfileValOriginal = jctGlobalProfileValOriginal;
	}

	public String getJctGlobalProfileValChanged() {
		return jctGlobalProfileValChanged;
	}

	public void setJctGlobalProfileValChanged(String jctGlobalProfileValChanged) {
		this.jctGlobalProfileValChanged = jctGlobalProfileValChanged;
	}

	public String getJctGlobalProfileValDescOriginal() {
		return jctGlobalProfileValDescOriginal;
	}

	public void setJctGlobalProfileValDescOriginal(
			String jctGlobalProfileValDescOriginal) {
		this.jctGlobalProfileValDescOriginal = jctGlobalProfileValDescOriginal;
	}

	public String getJctGlobalProfileValDescChanged() {
		return jctGlobalProfileValDescChanged;
	}

	public void setJctGlobalProfileValDescChanged(
			String jctGlobalProfileValDescChanged) {
		this.jctGlobalProfileValDescChanged = jctGlobalProfileValDescChanged;
	}

	public String getJctGlobalProfilePasOriginal() {
		return jctGlobalProfilePasOriginal;
	}

	public void setJctGlobalProfilePasOriginal(String jctGlobalProfilePasOriginal) {
		this.jctGlobalProfilePasOriginal = jctGlobalProfilePasOriginal;
	}

	public String getJctGlobalProfilePasChanged() {
		return jctGlobalProfilePasChanged;
	}

	public void setJctGlobalProfilePasChanged(String jctGlobalProfilePasChanged) {
		this.jctGlobalProfilePasChanged = jctGlobalProfilePasChanged;
	}

	public String getJctGlobalProfilePasDescOriginal() {
		return jctGlobalProfilePasDescOriginal;
	}

	public void setJctGlobalProfilePasDescOriginal(
			String jctGlobalProfilePasDescOriginal) {
		this.jctGlobalProfilePasDescOriginal = jctGlobalProfilePasDescOriginal;
	}

	public String getJctGlobalProfilePasDescChanged() {
		return jctGlobalProfilePasDescChanged;
	}

	public void setJctGlobalProfilePasDescChanged(
			String jctGlobalProfilePasDescChanged) {
		this.jctGlobalProfilePasDescChanged = jctGlobalProfilePasDescChanged;
	}

	public int getJctGlobalProfileSoftDelete() {
		return jctGlobalProfileSoftDelete;
	}

	public void setJctGlobalProfileSoftDelete(int jctGlobalProfileSoftDelete) {
		this.jctGlobalProfileSoftDelete = jctGlobalProfileSoftDelete;
	}

	public Date getJctGlobalProfileCreatedTimestamp() {
		return jctGlobalProfileCreatedTimestamp;
	}

	public void setJctGlobalProfileCreatedTimestamp(
			Date jctGlobalProfileCreatedTimestamp) {
		this.jctGlobalProfileCreatedTimestamp = jctGlobalProfileCreatedTimestamp;
	}

	public String getJctGlobalProfileBsMainQtn() {
		return jctGlobalProfileBsMainQtn;
	}

	public void setJctGlobalProfileBsMainQtn(String jctGlobalProfileBsMainQtn) {
		this.jctGlobalProfileBsMainQtn = jctGlobalProfileBsMainQtn;
	}

	public String getJctGlobalProfileApMainQtn() {
		return jctGlobalProfileApMainQtn;
	}

	public void setJctGlobalProfileApMainQtn(String jctGlobalProfileApMainQtn) {
		this.jctGlobalProfileApMainQtn = jctGlobalProfileApMainQtn;
	}

	public Date getPresentTimeStamp() {
		return presentTimeStamp;
	}

	public void setPresentTimeStamp(Date presentTimeStamp) {
		this.presentTimeStamp = presentTimeStamp;
	}

	public String getJctGlobalProfileIsAttribute() {
		return jctGlobalProfileIsAttribute;
	}

	public void setJctGlobalProfileIsAttribute(String jctGlobalProfileIsAttribute) {
		this.jctGlobalProfileIsAttribute = jctGlobalProfileIsAttribute;
	}

	public String getJctGlobalProfileAttributeType() {
		return jctGlobalProfileAttributeType;
	}

	public void setJctGlobalProfileAttributeType(
			String jctGlobalProfileAttributeType) {
		this.jctGlobalProfileAttributeType = jctGlobalProfileAttributeType;
	}
	
}
