package com.vmware.jct.service;

import java.io.InputStream;

public interface IStorageService {
	
	public void store(String key, InputStream in, String contnentType);

}
