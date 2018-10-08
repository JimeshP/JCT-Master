package com.vmware.jct.model;

import java.io.Serializable;
import javax.persistence.*;
/**
 * 
 * <p><b>Class name:</b> JctUserRole.java</p>
 * <p><b>Author:</b> InterraIT</p>
 * <p><b>Description:</b>  This class basically object reference of jct_user_role table. </p>
 * <p><b>Copyrights:</b> 	All rights reserved by Interra IT and should only be used for its internal application development.</p>
 * <p><b>Revision History:</b>
 * 	<li>InterraIT - 31/Jan/2014 - Implement comments, introduce Named query </li>
 * </p>
 */
@Entity
@Table(name="jct_user_role")
public class JctUserRole implements Serializable {
	private static final long serialVersionUID = 1L;

	@Id
	@Column(name="jct_role_id")
	private int jctRoleId;

	@Column(name="jct_role_code")
	private String jctRoleCode;

	@Column(name="jct_role_desc")
	private String jctRoleDesc;

	//bi-directional many-to-one association to JctUser
	//@OneToMany(mappedBy="jctUserRole")
	//@Column()
	//private List<JctUser> jctUsers;

	public JctUserRole() {
	}

	public int getJctRoleId() {
		return this.jctRoleId;
	}

	public void setJctRoleId(int jctRoleId) {
		this.jctRoleId = jctRoleId;
	}

	public String getJctRoleCode() {
		return this.jctRoleCode;
	}

	public void setJctRoleCode(String jctRoleCode) {
		this.jctRoleCode = jctRoleCode;
	}

	public String getJctRoleDesc() {
		return this.jctRoleDesc;
	}

	public void setJctRoleDesc(String jctRoleDesc) {
		this.jctRoleDesc = jctRoleDesc;
	}

	/*public List<JctUser> getJctUsers() {
		return this.jctUsers;
	}

	public void setJctUsers(List<JctUser> jctUsers) {
		this.jctUsers = jctUsers;
	}*/

	/*public JctUser addJctUser(JctUser jctUser) {
		getJctUsers().add(jctUser);
		jctUser.setJctUserRole(this);

		return jctUser;
	}

	public JctUser removeJctUser(JctUser jctUser) {
		getJctUsers().remove(jctUser);
		jctUser.setJctUserRole(null);

		return jctUser;
	}*/

}