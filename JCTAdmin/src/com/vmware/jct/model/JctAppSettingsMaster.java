package com.vmware.jct.model;

import java.io.Serializable;
import java.util.Date;

import javax.persistence.Column;
import javax.persistence.Entity;
import javax.persistence.GeneratedValue;
import javax.persistence.GenerationType;
import javax.persistence.Id;
import javax.persistence.Lob;
import javax.persistence.NamedQueries;
import javax.persistence.NamedQuery;
import javax.persistence.Table;
import javax.persistence.Temporal;
import javax.persistence.TemporalType;
/**
 * 
 * <p><b>Class name:</b> JctAppSettingsMaster.java</p>
 * <p><b>Author:</b> InterraIT</p>
 * <p><b>Description:</b>  This class basically object reference of jct_action_plan table </p>
 * <p><b>Copyrights:</b> 	All rights reserved by InterraIT and should only be used for its internal application development.</p>
 * <p><b>Revision History:</b>
 * 	<li>InterraIT - 04/Aug/2014 - Implement comments, introduce Named query </li>
 * </p>
 */
@NamedQueries({
	@NamedQuery(name = "fetchLastWhiteLebeling",
				query = "from JctAppSettingsMaster where appSoftDelete = :softDelete")
})
@Entity
@Table(name="jct_app_settings_master")
public class JctAppSettingsMaster implements Serializable {
	private static final long serialVersionUID = 1L;
	
	@Id
	@GeneratedValue(strategy = GenerationType.IDENTITY)
	@Column(name="jct_app_settings_id")
	private int appSettingsId;
	
	@Column(name="jct_app_header_color")
	private String appHeaderColor;
	
	@Column(name="jct_app_footer_color")
	private String appFooterColor;
	
	@Column(name="jct_app_sub_header_color")
	private String appSubHeaderColor;
	
	@Column(name="jct_app_ins_panel_color")
	private String appInsPanelColor;
	
	@Column(name = "jct_app_ins_panel_txt_color")
	private String appInsPanelTxtColor;
	
	@Column(name="jct_app_soft_delete")
	private int appSoftDelete;
	
	@Column(name="jct_app_last_updated_by")
	private String appLastUpdatedBy;
	
	@Temporal(TemporalType.TIMESTAMP)
	@Column(name="jct_app_last_updated_ts")
	private Date appLastUpdatedTs;
	
	@Lob
	@Column(name="jct_app_image")
	private String appImage;

	/**
	 * @return the appSettingsId
	 */
	public int getAppSettingsId() {
		return appSettingsId;
	}

	/**
	 * @param appSettingsId the appSettingsId to set
	 */
	public void setAppSettingsId(int appSettingsId) {
		this.appSettingsId = appSettingsId;
	}

	/**
	 * @return the appHeaderColor
	 */
	public String getAppHeaderColor() {
		return appHeaderColor;
	}

	/**
	 * @param appHeaderColor the appHeaderColor to set
	 */
	public void setAppHeaderColor(String appHeaderColor) {
		this.appHeaderColor = appHeaderColor;
	}

	/**
	 * @return the appFooterColor
	 */
	public String getAppFooterColor() {
		return appFooterColor;
	}

	/**
	 * @param appFooterColor the appFooterColor to set
	 */
	public void setAppFooterColor(String appFooterColor) {
		this.appFooterColor = appFooterColor;
	}

	/**
	 * @return the appSubHeaderColor
	 */
	public String getAppSubHeaderColor() {
		return appSubHeaderColor;
	}

	/**
	 * @param appSubHeaderColor the appSubHeaderColor to set
	 */
	public void setAppSubHeaderColor(String appSubHeaderColor) {
		this.appSubHeaderColor = appSubHeaderColor;
	}

	/**
	 * @return the appInsPanelColor
	 */
	public String getAppInsPanelColor() {
		return appInsPanelColor;
	}

	/**
	 * @param appInsPanelColor the appInsPanelColor to set
	 */
	public void setAppInsPanelColor(String appInsPanelColor) {
		this.appInsPanelColor = appInsPanelColor;
	}

	/**
	 * @return the appSoftDelete
	 */
	public int getAppSoftDelete() {
		return appSoftDelete;
	}

	/**
	 * @param appSoftDelete the appSoftDelete to set
	 */
	public void setAppSoftDelete(int appSoftDelete) {
		this.appSoftDelete = appSoftDelete;
	}

	/**
	 * @return the appLastUpdatedBy
	 */
	public String getAppLastUpdatedBy() {
		return appLastUpdatedBy;
	}

	/**
	 * @param appLastUpdatedBy the appLastUpdatedBy to set
	 */
	public void setAppLastUpdatedBy(String appLastUpdatedBy) {
		this.appLastUpdatedBy = appLastUpdatedBy;
	}

	/**
	 * @return the appLastUpdatedTs
	 */
	public Date getAppLastUpdatedTs() {
		return appLastUpdatedTs;
	}

	/**
	 * @param appLastUpdatedTs the appLastUpdatedTs to set
	 */
	public void setAppLastUpdatedTs(Date appLastUpdatedTs) {
		this.appLastUpdatedTs = appLastUpdatedTs;
	}

	/**
	 * @return the appInsPanelTxtColor
	 */
	public String getAppInsPanelTxtColor() {
		return appInsPanelTxtColor;
	}

	/**
	 * @param appInsPanelTxtColor the appInsPanelTxtColor to set
	 */
	public void setAppInsPanelTxtColor(String appInsPanelTxtColor) {
		this.appInsPanelTxtColor = appInsPanelTxtColor;
	}

	/**
	 * @return the appImage
	 */
	public String getAppImage() {
		return appImage;
	}

	/**
	 * @param appImage the appImage to set
	 */
	public void setAppImage(String appImage) {
		this.appImage = appImage;
	}
	
}
