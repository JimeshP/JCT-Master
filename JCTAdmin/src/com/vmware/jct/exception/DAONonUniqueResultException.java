package com.vmware.jct.exception;

import com.vmware.jct.common.utility.ExceptionCode;

/**
 * @author InterraIT
 *
 */
public class DAONonUniqueResultException extends DAOException {

	private static final long serialVersionUID = -4635213382743942313L;
	public DAONonUniqueResultException() {
		super();
	}

	public DAONonUniqueResultException(Exception ex) {
		super(ex);
	}

	public DAONonUniqueResultException(String message) {
		super(message);
	}

	public DAONonUniqueResultException(String message, Exception ex) {
		super(message, ex);
	}

	public DAONonUniqueResultException(ExceptionCode exceptionCode) {
		this();
		this.setExceptionCode(exceptionCode);
	}

	public DAONonUniqueResultException(Exception ex, ExceptionCode exceptionCode) {
		this(ex);
		this.setExceptionCode(exceptionCode);
	}
}
