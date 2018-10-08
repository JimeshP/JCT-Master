/**
 * 
 */
package com.vmware.jct.exception;

import com.vmware.jct.common.utility.ExceptionCode;

/**
 * @author InterraIT
 *
 */
public class DAOEntityNotFoundException extends DAOException {

	/**
	 * 
	 */
	private static final long serialVersionUID = -6122874565698549779L;
	public DAOEntityNotFoundException() {
		super();
	}

	public DAOEntityNotFoundException(Exception ex) {
		super(ex);
	}

	public DAOEntityNotFoundException(String message) {
		super(message);
	}

	public DAOEntityNotFoundException(String message, Exception ex) {
		super(message, ex);
	}

	public DAOEntityNotFoundException(ExceptionCode exceptionCode) {
		this();
		this.setExceptionCode(exceptionCode);
	}

	public DAOEntityNotFoundException(Exception ex, ExceptionCode exceptionCode) {
		this(ex);
		this.setExceptionCode(exceptionCode);
	}

}
