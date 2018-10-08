package com.vmware.jct.exception;

import com.vmware.jct.common.utility.ExceptionCode;

/**
 * @author InterraIT
 *
 */
public class DAONoResultException extends DAOException {

	private static final long serialVersionUID = -7773063237128614642L;
	public DAONoResultException() {
		super();
	}

	public DAONoResultException(Exception ex) {
		super(ex);
	}

	public DAONoResultException(String message) {
		super(message);
	}

	public DAONoResultException(String message, Exception ex) {
		super(message, ex);
	}

	public DAONoResultException(ExceptionCode exceptionCode) {
		this();
		this.setExceptionCode(exceptionCode);
	}

	public DAONoResultException(Exception ex, ExceptionCode exceptionCode) {
		this(ex);
		this.setExceptionCode(exceptionCode);
	}
}
