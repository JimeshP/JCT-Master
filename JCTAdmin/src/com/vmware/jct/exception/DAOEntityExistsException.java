/**
 * 
 */
package com.vmware.jct.exception;

import com.vmware.jct.common.utility.ExceptionCode;

/**
 * @author InterraIT
 *
 */
public class DAOEntityExistsException extends DAOException {

	/**
	 * 
	 */
	private static final long serialVersionUID = 650204195411084486L;
	public DAOEntityExistsException() {
		super();
	}

	public DAOEntityExistsException(Exception ex) {
		super(ex);
	}

	public DAOEntityExistsException(String message) {
		super(message);
	}

	public DAOEntityExistsException(String message, Exception ex) {
		super(message, ex);
	}

	public DAOEntityExistsException(ExceptionCode exceptionCode) {
		this();
		this.setExceptionCode(exceptionCode);
	}

	public DAOEntityExistsException(Exception ex, ExceptionCode exceptionCode) {
		this(ex);
		this.setExceptionCode(exceptionCode);
	}
}
