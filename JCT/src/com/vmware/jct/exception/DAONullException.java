/**
 * 
 */
package com.vmware.jct.exception;

import com.vmware.jct.common.utility.ExceptionCode;

/**
 * @author InterraIt
 *
 */
public class DAONullException extends DAOException {

	/**
	 * 
	 */
	private static final long serialVersionUID = 140730517806217919L;
	public DAONullException() {
		super();
	}

	public DAONullException(Exception ex) {
		super(ex);
	}

	public DAONullException(String message) {
		super(message);
	}

	public DAONullException(String message, Exception ex) {
		super(message, ex);
	}

	public DAONullException(ExceptionCode exceptionCode) {
		this();
		this.setExceptionCode(exceptionCode);
	}

	public DAONullException(Exception ex, ExceptionCode exceptionCode) {
		this(ex);
		this.setExceptionCode(exceptionCode);
	}
}
