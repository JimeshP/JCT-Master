package com.vmware.jct.model;

import java.io.Serializable;
import javax.persistence.*;


/**
 * 
 * <p><b>Class name:</b> IdGenTable.java</p>
 * <p><b>Author:</b> InterraIT</p>
 * <p><b>Description:</b>  This class basically object reference of jct_id_gen_table table. This entity is responsible for generating unique identification number </p>
 * <p><b>Copyrights:</b> 	All rights reserved by InterraIT and should only be used for its internal application development.</p>
 * <p><b>Revision History:</b>
 * 	<li>InterraIT - 31/Jan/2014 - Implement comments, introduce Named query </li>
 * </p>
 */
@NamedQueries({
	@NamedQuery(name = "updateIdGenerationYear", 
			query = "Update JctIdGenTable id set id.jctIdVal = 0 where id.jctIdName = :keyName")
})
@Entity
@Table(name="jct_id_gen_table")
public class JctIdGenTable implements Serializable {
	private static final long serialVersionUID = 1L;

	@Id
	@Column(name="jct_id_gen_id")
	private int jctIdGenId;

	@Column(name="jct_id_name")
	private String jctIdName;

	@Column(name="jct_id_val")
	private int jctIdVal;

	public JctIdGenTable() {
	}

	public int getJctIdGenId() {
		return this.jctIdGenId;
	}

	public void setJctIdGenId(int jctIdGenId) {
		this.jctIdGenId = jctIdGenId;
	}

	public String getJctIdName() {
		return this.jctIdName;
	}

	public void setJctIdName(String jctIdName) {
		this.jctIdName = jctIdName;
	}

	public int getJctIdVal() {
		return this.jctIdVal;
	}

	public void setJctIdVal(int jctIdVal) {
		this.jctIdVal = jctIdVal;
	}

}