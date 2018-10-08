package com.vmware.jct.dao.dto;

import java.util.Date;

public class PaymentDetailsDTO {
	private int id;
	private Date date;
	private String chequeNos;
	private String userName;
	private int typeId;
	private String typeDesc;
	private String transNos;
	private String bankName;	
	private String createdBy;
	private Date createdTs;
	private String modifiedBy;
	private Date dtlModifiedTs;
	private int softDelete;
	private String customerId;
	private int roleId;
	private int headerId;
	private String jctPmtdtlsPmtTransNos;
	private String password;
	
	public PaymentDetailsDTO(){
		
	}
	
	public PaymentDetailsDTO(String userName, String password, String typeDesc, int roleId) {
		this.userName = userName;
		this.password = password;
		this.typeDesc = typeDesc;
		this.roleId = roleId;
	}

	public PaymentDetailsDTO(String userName, String customerId, int roleId, int headerId, int id){
		this.userName = userName;
		this.customerId = customerId;
		this.roleId = roleId;
		this.headerId = headerId;
		this.id = id;
	}
	
	public PaymentDetailsDTO(int id,Date date,String chequeNos,
			String userName, String customerId, int roleId, int headerId){
		this.id=id;
		this.date = date;
		this.chequeNos = chequeNos;
		this.userName = userName;
		this.customerId = customerId;
		this.roleId = roleId;
		this.headerId = headerId;	
	}
	
	public PaymentDetailsDTO(String chequeNos, Date date, String jctPmtdtlsPmtTransNos){
		this.chequeNos = chequeNos;
		this.date = date;	
		this.jctPmtdtlsPmtTransNos = jctPmtdtlsPmtTransNos;
	}
	
	public int getId() {
		return id;
	}
	public void setId(int id) {
		this.id = id;
	}
	public Date getDate() {
		return date;
	}
	public void setDate(Date date) {
		this.date = date;
	}
	public int getTypeId() {
		return typeId;
	}
	public void setTypeId(int typeId) {
		this.typeId = typeId;
	}
	public String getTypeDesc() {
		return typeDesc;
	}
	public void setTypeDesc(String typeDesc) {
		this.typeDesc = typeDesc;
	}
	public String getTransNos() {
		return transNos;
	}
	public void setTransNos(String transNos) {
		this.transNos = transNos;
	}
	public String getBankName() {
		return bankName;
	}
	public void setBankName(String bankName) {
		this.bankName = bankName;
	}
	public String getChequeNos() {
		return chequeNos;
	}
	public void setChequeNos(String chequeNos) {
		this.chequeNos = chequeNos;
	}
	public String getCreatedBy() {
		return createdBy;
	}
	public void setCreatedBy(String createdBy) {
		this.createdBy = createdBy;
	}
	public Date getCreatedTs() {
		return createdTs;
	}
	public void setCreatedTs(Date createdTs) {
		this.createdTs = createdTs;
	}
	public String getModifiedBy() {
		return modifiedBy;
	}
	public void setModifiedBy(String modifiedBy) {
		this.modifiedBy = modifiedBy;
	}
	public Date getDtlModifiedTs() {
		return dtlModifiedTs;
	}
	public void setDtlModifiedTs(Date dtlModifiedTs) {
		this.dtlModifiedTs = dtlModifiedTs;
	}
	public int getSoftDelete() {
		return softDelete;
	}
	public void setSoftDelete(int softDelete) {
		this.softDelete = softDelete;
	}
	public String getUserName() {
		return userName;
	}
	public void setUserName(String userName) {
		this.userName = userName;
	}
	public String getCustomerId() {
		return customerId;
	}
	public void setCustomerId(String customerId) {
		this.customerId = customerId;
	}
	public int getRoleId() {
		return roleId;
	}
	public void setRoleId(int roleId) {
		this.roleId = roleId;
	}
	public int getHeaderId() {
		return headerId;
	}
	public void setHeaderId(int headerId) {
		this.headerId = headerId;
	}
	public String getJctPmtdtlsPmtTransNos() {
		return jctPmtdtlsPmtTransNos;
	}
	public void setJctPmtdtlsPmtTransNos(String jctPmtdtlsPmtTransNos) {
		this.jctPmtdtlsPmtTransNos = jctPmtdtlsPmtTransNos;
	}

	public String getPassword() {
		return password;
	}

	public void setPassword(String password) {
		this.password = password;
	}
	
}