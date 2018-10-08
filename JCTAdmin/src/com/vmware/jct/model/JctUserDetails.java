package com.vmware.jct.model;

import java.io.Serializable;
import java.util.Date;

import javax.persistence.Column;
import javax.persistence.Entity;
import javax.persistence.GeneratedValue;
import javax.persistence.GenerationType;
import javax.persistence.Id;
import javax.persistence.JoinColumn;
import javax.persistence.Lob;
import javax.persistence.NamedQueries;
import javax.persistence.NamedQuery;
import javax.persistence.OneToOne;
import javax.persistence.Table;
import javax.persistence.Temporal;
import javax.persistence.TemporalType;


@NamedQueries({
		@NamedQuery(name = "disableUserDtlsForRefundRequestByUserId",
				query = "update JctUserDetails set jctUserDetailsSoftDelete = 1 where jctUser.jctUserId = :userId"),
		@NamedQuery(name = "checkUsersCreatedByFacilitator",
			query = "from JctUserDetails where jctUserDetailsFacilitatorId = :userId"),
		@NamedQuery(name = "fetchUserFirstName",
			query = "select jctUserDetailsFirstName from JctUserDetails where jctUser.jctUserName = :userName and jctUser.jctUserRole.jctRoleId = :role")
})
				
@Entity
@Table(name="jct_user_details")
public class JctUserDetails implements Serializable {
	private static final long serialVersionUID = 1L;
		
	@Id
	@GeneratedValue(strategy=GenerationType.IDENTITY)
	@Column(name="jct_user_details_id")
	private int jctUserDetailsId;
	
	@Column(name="jct_user_details_first_name")
	private String jctUserDetailsFirstName;
	
	@Column(name="jct_user_details_last_name")
	private String jctUserDetailsLastName;
	
	@Column(name="jct_user_details_phone")
	private String jctUserDetailsPhone;
	
	@Column(name="jct_user_details_gender")
	private String jctUserDetailsGender;
	
	@Temporal(TemporalType.DATE)
	@Column(name="jct_user_details_date_of_birth")
	private Date jctUserDetailsDateOfBirth;
	
	@Lob
	@Column(name="jct_user_details_profile_picture")
	private byte[] jctUserDetailsProfilePicture;
	
	@Temporal(TemporalType.DATE)
	@Column(name="jct_user_details_created_ts")
	private Date jctUserDetailsCreatedTs;
	
	@Column(name="jct_user_details_lastmodified_by")
	private String jctUserDetailsLastmodifiedBy;
	
	@Column(name="jct_user_details_created_by")
	private String jctUserDetailsCreatedBy;
	
	@Temporal(TemporalType.DATE)
	@Column(name="jct_user_details_lastmodified_ts")
	private Date jctUserDetailsLastmodifiedTs;
	
	@Column(name="jct_user_details_base_string")
	private String jctUserDetailsBaseString;
	
	@Column(name="jct_user_details_region")
	private String jctUserDetailsRegion;
	
	@Column(name="jct_user_details_levels")
	private String jctUserDetailsLevels;
	
	@Column(name="jct_user_details_tenure")
	private String jctUserDetailsTenure;
	
	@Column(name="jct_user_details_supervise_people")
	private String jctUserDetailsSupervisePeople;
	
	@Column(name="jct_user_details_function_group")
	private String jctUserDetailsFunctionGroup;
	
	@Column(name="jct_user_details_group_id")
	private int jctUserDetailsGroupId;
	
	@Column(name="jct_user_details_group_name")
	private String jctUserDetailsGroupName;
	
	@Column(name="jct_user_details_profile_id")
	private int jctUserDetailsProfileId;
	
	@Column(name = "jct_user_details_profile_name")
	private String jctUserDetailsProfileName;
	
	@Column(name="jct_user_details_facilitator_id")
	private int jctUserDetailsFacilitatorId;
	
	@Column(name="jct_user_details_admin_id")
	private int jctUserDetailsAdminId;
	
	@OneToOne
    @JoinColumn(name="jct_user_id")
	private JctUser jctUser;
	
	@Column(name="jct_user_details_soft_delete")
	private int jctUserDetailsSoftDelete;
	

	public int getJctUserDetailsId() {
		return jctUserDetailsId;
	}

	public void setJctUserDetailsId(int jctUserDetailsId) {
		this.jctUserDetailsId = jctUserDetailsId;
	}

	public String getJctUserDetailsFirstName() {
		return jctUserDetailsFirstName;
	}

	public void setJctUserDetailsFirstName(String jctUserDetailsFirstName) {
		this.jctUserDetailsFirstName = jctUserDetailsFirstName;
	}

	public String getJctUserDetailsLastName() {
		return jctUserDetailsLastName;
	}

	public void setJctUserDetailsLastName(String jctUserDetailsLastName) {
		this.jctUserDetailsLastName = jctUserDetailsLastName;
	}

	public String getJctUserDetailsPhone() {
		return jctUserDetailsPhone;
	}

	public void setJctUserDetailsPhone(String jctUserDetailsPhone) {
		this.jctUserDetailsPhone = jctUserDetailsPhone;
	}

	public String getJctUserDetailsGender() {
		return jctUserDetailsGender;
	}

	public void setJctUserDetailsGender(String jctUserDetailsGender) {
		this.jctUserDetailsGender = jctUserDetailsGender;
	}

	public Date getJctUserDetailsDateOfBirth() {
		return jctUserDetailsDateOfBirth;
	}

	public void setJctUserDetailsDateOfBirth(Date jctUserDetailsDateOfBirth) {
		this.jctUserDetailsDateOfBirth = jctUserDetailsDateOfBirth;
	}

	public byte[] getJctUserDetailsProfilePicture() {
		return jctUserDetailsProfilePicture;
	}

	public void setJctUserDetailsProfilePicture(byte[] jctUserDetailsProfilePicture) {
		this.jctUserDetailsProfilePicture = jctUserDetailsProfilePicture;
	}

	public Date getJctUserDetailsCreatedTs() {
		return jctUserDetailsCreatedTs;
	}

	public void setJctUserDetailsCreatedTs(Date jctUserDetailsCreatedTs) {
		this.jctUserDetailsCreatedTs = jctUserDetailsCreatedTs;
	}

	public String getJctUserDetailsLastmodifiedBy() {
		return jctUserDetailsLastmodifiedBy;
	}

	public void setJctUserDetailsLastmodifiedBy(String jctUserDetailsLastmodifiedBy) {
		this.jctUserDetailsLastmodifiedBy = jctUserDetailsLastmodifiedBy;
	}

	public String getJctUserDetailsCreatedBy() {
		return jctUserDetailsCreatedBy;
	}

	public void setJctUserDetailsCreatedBy(String jctUserDetailsCreatedBy) {
		this.jctUserDetailsCreatedBy = jctUserDetailsCreatedBy;
	}

	public Date getJctUserDetailsLastmodifiedTs() {
		return jctUserDetailsLastmodifiedTs;
	}

	public void setJctUserDetailsLastmodifiedTs(Date jctUserDetailsLastmodifiedTs) {
		this.jctUserDetailsLastmodifiedTs = jctUserDetailsLastmodifiedTs;
	}

	public String getJctUserDetailsBaseString() {
		return jctUserDetailsBaseString;
	}

	public void setJctUserDetailsBaseString(String jctUserDetailsBaseString) {
		this.jctUserDetailsBaseString = jctUserDetailsBaseString;
	}

	public String getJctUserDetailsRegion() {
		return jctUserDetailsRegion;
	}

	public void setJctUserDetailsRegion(String jctUserDetailsRegion) {
		this.jctUserDetailsRegion = jctUserDetailsRegion;
	}

	public String getJctUserDetailsLevels() {
		return jctUserDetailsLevels;
	}

	public void setJctUserDetailsLevels(String jctUserDetailsLevels) {
		this.jctUserDetailsLevels = jctUserDetailsLevels;
	}

	public String getJctUserDetailsTenure() {
		return jctUserDetailsTenure;
	}

	public void setJctUserDetailsTenure(String jctUserDetailsTenure) {
		this.jctUserDetailsTenure = jctUserDetailsTenure;
	}

	public String getJctUserDetailsSupervisePeople() {
		return jctUserDetailsSupervisePeople;
	}

	public void setJctUserDetailsSupervisePeople(
			String jctUserDetailsSupervisePeople) {
		this.jctUserDetailsSupervisePeople = jctUserDetailsSupervisePeople;
	}

	public String getJctUserDetailsFunctionGroup() {
		return jctUserDetailsFunctionGroup;
	}

	public void setJctUserDetailsFunctionGroup(String jctUserDetailsFunctionGroup) {
		this.jctUserDetailsFunctionGroup = jctUserDetailsFunctionGroup;
	}

	public int getJctUserDetailsGroupId() {
		return jctUserDetailsGroupId;
	}

	public void setJctUserDetailsGroupId(int jctUserDetailsGroupId) {
		this.jctUserDetailsGroupId = jctUserDetailsGroupId;
	}

	public String getJctUserDetailsGroupName() {
		return jctUserDetailsGroupName;
	}

	public void setJctUserDetailsGroupName(String jctUserDetailsGroupName) {
		this.jctUserDetailsGroupName = jctUserDetailsGroupName;
	}

	public int getJctUserDetailsProfileId() {
		return jctUserDetailsProfileId;
	}

	public void setJctUserDetailsProfileId(int jctUserDetailsProfileId) {
		this.jctUserDetailsProfileId = jctUserDetailsProfileId;
	}

	public String getJctUserDetailsProfileName() {
		return jctUserDetailsProfileName;
	}

	public void setJctUserDetailsProfileName(String jctUserDetailsProfileName) {
		this.jctUserDetailsProfileName = jctUserDetailsProfileName;
	}

	public int getJctUserDetailsFacilitatorId() {
		return jctUserDetailsFacilitatorId;
	}

	public void setJctUserDetailsFacilitatorId(int jctUserDetailsFacilitatorId) {
		this.jctUserDetailsFacilitatorId = jctUserDetailsFacilitatorId;
	}

	public int getJctUserDetailsAdminId() {
		return jctUserDetailsAdminId;
	}

	public void setJctUserDetailsAdminId(int jctUserDetailsAdminId) {
		this.jctUserDetailsAdminId = jctUserDetailsAdminId;
	}

	public JctUser getJctUser() {
		return jctUser;
	}

	public void setJctUser(JctUser jctUser) {
		this.jctUser = jctUser;
	}

	public int getJctUserDetailsSoftDelete() {
		return jctUserDetailsSoftDelete;
	}

	public void setJctUserDetailsSoftDelete(int jctUserDetailsSoftDelete) {
		this.jctUserDetailsSoftDelete = jctUserDetailsSoftDelete;
	}
}
