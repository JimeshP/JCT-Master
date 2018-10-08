/**
 * 
 */
package com.vmware.jct.exception;

import com.vmware.jct.common.utility.ExceptionCode;

/**
 * @author InterraIt
 *
 */
public class DAOException extends JCTBaseException {
	
	/**
	 * 
	 */
	private static final long serialVersionUID = -5363602445654413021L;

	public DAOException() {
		super();
	}

	public DAOException(Exception ex) {
		super(ex);
	}

	public DAOException(String message) {
		super(message);
	}

	public DAOException(String message, Exception ex) {
		super(message, ex);
	}

	public DAOException(ExceptionCode exceptionCode) {
		this();
		this.setExceptionCode(exceptionCode);
	}

	public DAOException(Exception ex, ExceptionCode exceptionCode) {
		this(ex);
		this.setExceptionCode(exceptionCode);
	}

}
