package com.vmware.jct.dao.dto;

import java.util.Date;

public class RemoveDiagramDTO {
	private int rowId;
	private String beforeSketchBaseString;
	private String aftereSketchBaseString;
	private String jctJobrefNo;
	private Date jctCreatedTs;
		
	

	public RemoveDiagramDTO (int rowId, String beforeSketchBaseString, 
			String aftereSketchBaseString) {
		this.rowId = rowId;
		this.beforeSketchBaseString = beforeSketchBaseString;
		this.aftereSketchBaseString = aftereSketchBaseString;
	}
	
	public RemoveDiagramDTO (int rowId, String beforeSketchBaseString, 
			String aftereSketchBaseString, String jctJobrefNo) {
		this.rowId = rowId;
		this.beforeSketchBaseString = beforeSketchBaseString;
		this.aftereSketchBaseString = aftereSketchBaseString;
		this.jctJobrefNo = jctJobrefNo;
	}
	
	public RemoveDiagramDTO (int rowId, String beforeSketchBaseString, 
			String aftereSketchBaseString, String jctJobrefNo,
			Date jctCreatedTs) {
		this.rowId = rowId;
		this.beforeSketchBaseString = beforeSketchBaseString;
		this.aftereSketchBaseString = aftereSketchBaseString;
		this.jctJobrefNo = jctJobrefNo;
		this.jctCreatedTs = jctCreatedTs;
	}

	public int getRowId() {
		return rowId;
	}

	public void setRowId(int rowId) {
		this.rowId = rowId;
	}

	public String getBeforeSketchBaseString() {
		return beforeSketchBaseString;
	}

	public void setBeforeSketchBaseString(String beforeSketchBaseString) {
		this.beforeSketchBaseString = beforeSketchBaseString;
	}

	public String getAftereSketchBaseString() {
		return aftereSketchBaseString;
	}

	public void setAftereSketchBaseString(String aftereSketchBaseString) {
		this.aftereSketchBaseString = aftereSketchBaseString;
	}

	public String getJctJobrefNo() {
		return jctJobrefNo;
	}

	public void setJctJobrefNo(String jctJobrefNo) {
		this.jctJobrefNo = jctJobrefNo;
	}	
	
	public Date getJctCreatedTs() {
		return jctCreatedTs;
	}

	public void setJctCreatedTs(Date jctCreatedTs) {
		this.jctCreatedTs = jctCreatedTs;
	}
	
}
