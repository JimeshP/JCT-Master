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
/**
 * 
 * <p><b>Class name:</b> JctAdminMenu.java</p>
 * <p><b>Author:</b> InterraIT</p>
 * <p><b>Description:</b>  This class basically object reference of jct_admin_menu table </p>
 * <p><b>Copyrights:</b> 	All rights reserved by InterraIT and should only be used for its internal application development.</p>
 * <p><b>Revision History:</b>
 * 	<li>InterraIT - 01/Sep/2014 - Implementation of dynamic menu population</li>
 * </p>
 */
@NamedQueries({
	@NamedQuery(name="getActiveDistinctMenuItems",
				query="select distinct(obj.jctMainMenu) from JctAdminMenu obj where jctRoleId = :roleId " +
						"and jctSoftDelete = 0 order by jctMainMenuOrder"),
	@NamedQuery(name="getActiveMenuDetails",
				query="from JctAdminMenu where jctRoleId = :roleId and jctSoftDelete = 0 and jctMainMenu = :mainMenu " +
						"order by jctSubMenuOrder")
})
@Entity
@Table(name="jct_admin_menu")
public class JctAdminMenu implements Serializable {
	private static final long serialVersionUID = 1L;
	
	@Id
	@Column(name="jct_menu_id")
	private int jctMenuId;
	
	@Column(name="jct_main_menu")
	private String jctMainMenu;
	
	@Column(name="jct_sub_menu")
	private String jctSubMenu;
	
	@Column(name="jct_role_id")
	private int jctRoleId;
	
	@Column(name="jct_soft_delete")
	private int jctSoftDelete;
	
	@Column(name="jct_sub_menu_linked_page")
	private String jctSubMenuLinkedPage;
	
	@Column(name = "jct_created_by")
	private String jctCreatedBy;
	
	@Temporal(TemporalType.TIMESTAMP)
	@Column(name="jct_created_ts")
	private Date jctCreatedTs;
	
	@Column(name="jct_last_modified_by")
	private String jctLastModifiedBy;
	
	@Temporal(TemporalType.TIMESTAMP)
	@Column(name="jct_last_modified_ts")
	private Date jctLastModifiedTs; 
	
	@Column(name="jct_main_menu_order")
	private int jctMainMenuOrder;
	
	@Column(name="jct_sub_menu_order")
	private int jctSubMenuOrder;
	

	public int getJctMenuId() {
		return jctMenuId;
	}

	public void setJctMenuId(int jctMenuId) {
		this.jctMenuId = jctMenuId;
	}

	public String getJctMainMenu() {
		return jctMainMenu;
	}

	public void setJctMainMenu(String jctMainMenu) {
		this.jctMainMenu = jctMainMenu;
	}

	public String getJctSubMenu() {
		return jctSubMenu;
	}

	public void setJctSubMenu(String jctSubMenu) {
		this.jctSubMenu = jctSubMenu;
	}

	public int getJctRoleId() {
		return jctRoleId;
	}

	public void setJctRoleId(int jctRoleId) {
		this.jctRoleId = jctRoleId;
	}

	public int getJctSoftDelete() {
		return jctSoftDelete;
	}

	public void setJctSoftDelete(int jctSoftDelete) {
		this.jctSoftDelete = jctSoftDelete;
	}

	public String getJctSubMenuLinkedPage() {
		return jctSubMenuLinkedPage;
	}

	public void setJctSubMenuLinkedPage(String jctSubMenuLinkedPage) {
		this.jctSubMenuLinkedPage = jctSubMenuLinkedPage;
	}

	public String getJctCreatedBy() {
		return jctCreatedBy;
	}

	public void setJctCreatedBy(String jctCreatedBy) {
		this.jctCreatedBy = jctCreatedBy;
	}

	public Date getJctCreatedTs() {
		return jctCreatedTs;
	}

	public void setJctCreatedTs(Date jctCreatedTs) {
		this.jctCreatedTs = jctCreatedTs;
	}

	public String getJctLastModifiedBy() {
		return jctLastModifiedBy;
	}

	public void setJctLastModifiedBy(String jctLastModifiedBy) {
		this.jctLastModifiedBy = jctLastModifiedBy;
	}

	public Date getJctLastModifiedTs() {
		return jctLastModifiedTs;
	}

	public void setJctLastModifiedTs(Date jctLastModifiedTs) {
		this.jctLastModifiedTs = jctLastModifiedTs;
	}

	public int getJctMainMenuOrder() {
		return jctMainMenuOrder;
	}

	public void setJctMainMenuOrder(int jctMainMenuOrder) {
		this.jctMainMenuOrder = jctMainMenuOrder;
	}

	public int getJctSubMenuOrder() {
		return jctSubMenuOrder;
	}

	public void setJctSubMenuOrder(int jctSubMenuOrder) {
		this.jctSubMenuOrder = jctSubMenuOrder;
	}
}
