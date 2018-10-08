package com.vmware.jct.dao.dto;

import java.io.Serializable;

public class DemographicDTO implements Serializable {
	
	private static final long serialVersionUID = 1L;
	public String region;
	public String functionGroup;
	public String jobLevel;
	public String onetOccupationData;
	
	public DemographicDTO (String region, String functionGroup, String jobLevel) {
		this.region = region;
		this.functionGroup = functionGroup;
		this.jobLevel = jobLevel;
	}
	
	public DemographicDTO (String onetOccupationData) {
		this.onetOccupationData = onetOccupationData;		
	}

	public String getRegion() {
		return region;
	}

	public void setRegion(String region) {
		this.region = region;
	}

	public String getFunctionGroup() {
		return functionGroup;
	}

	public void setFunctionGroup(String functionGroup) {
		this.functionGroup = functionGroup;
	}

	public String getJobLevel() {
		return jobLevel;
	}

	public void setJobLevel(String jobLevel) {
		this.jobLevel = jobLevel;
	}

	public String getOnetOccupationData() {
		return onetOccupationData;
	}

	public void setOnetOccupationData(String onetOccupationData) {
		this.onetOccupationData = onetOccupationData;
	}
	
}
