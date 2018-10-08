package com.vmware.jct.service;

import com.vmware.jct.exception.JCTException;

public interface IJCTUtilService {

	/**
	 * 	
	 * @param keyName
	 * @return
	 * @throws WorkFlowException
	 */
	public Integer generateKey(String keyName) throws JCTException;
	
}
