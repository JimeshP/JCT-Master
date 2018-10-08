package com.vmware.jct.service.vo;

import java.util.List;
/**
 * 
 * <p><b>Class name:</b> SearchResultVO.java</p>
 * <p><b>Author:</b> InterraIT</p>
 * <p><b>Purpose:</b> This class acts as a data transfer object.
 * <p><b>Description:</b> This class acts as a data transfer object.. </p>
 * <p><b>Copyrights:</b> 	All rights reserved by Interra IT and should only be used for its internal application development.</p>
 * <p><b>Created Date:</b> 20/Feb/2014</p>
 * <p><b>Revision History:</b>
 * 	<li>6/May/2014: Changed the saving of action plan to dynamic.</li>
 * </p>
 */
public class SearchResultVO {
	private List<String> sketchStringList;
	//FOR STATUS
	private String statusCode;
	private String statusMsg;
	private String tabToBeShown;
	//For selected diagrams
	private String selectedDiagram;
	private long totalDiagramCount;

	public String getStatusCode() {
		return statusCode;
	}

	public void setStatusCode(String statusCode) {
		this.statusCode = statusCode;
	}

	public String getStatusMsg() {
		return statusMsg;
	}

	public void setStatusMsg(String statusMsg) {
		this.statusMsg = statusMsg;
	}

	public List<String> getSketchStringList() {
		return sketchStringList;
	}

	public void setSketchStringList(List<String> sketchStringList) {
		this.sketchStringList = sketchStringList;
	}

	public String getTabToBeShown() {
		return tabToBeShown;
	}

	public void setTabToBeShown(String tabToBeShown) {
		this.tabToBeShown = tabToBeShown;
	}

	public String getSelectedDiagram() {
		return selectedDiagram;
	}

	public void setSelectedDiagram(String selectedDiagram) {
		this.selectedDiagram = selectedDiagram;
	}

	public long getTotalDiagramCount() {
		return totalDiagramCount;
	}

	public void setTotalDiagramCount(long totalDiagramCount) {
		this.totalDiagramCount = totalDiagramCount;
	}
	
}
