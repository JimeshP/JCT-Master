/**
 * 
 */
package com.vmware.jct.exception;

import com.vmware.jct.common.utility.ExceptionCode;

/**
 * @author InterraIt
 *
 */
public class JCTBaseException extends Exception {

	
	/**
	 * 
	 */
	private static final long serialVersionUID = -5912951665093813552L;
	
	protected ExceptionCode exceptionCode;

	public JCTBaseException() {
		super();
		// TODO Auto-generated constructor stub
	}
	public JCTBaseException(Exception ex) {
		super(ex);
	}

	public JCTBaseException(String message, Throwable cause,
			boolean enableSuppression, boolean writableStackTrace) {
		//super(message, cause, enableSuppression, writableStackTrace);
		// TODO Auto-generated constructor stub
	}

	public JCTBaseException(String message, Throwable cause) {
	//	super();
		super(message, cause);
		// TODO Auto-generated constructor stub
	}

	public JCTBaseException(String message) {
		super(message);
		// TODO Auto-generated constructor stub
	}

	public JCTBaseException(ExceptionCode exceptionCode) {
		this();
		this.setExceptionCode(exceptionCode);
	}
	
	public JCTBaseException(Exception ex, ExceptionCode exceptionCode) {
		this(ex);
		this.setExceptionCode(exceptionCode);
	}
	
	/**
	 * @see #Throwable.getMessage
	 */
	public String getMessage() {
		String message = null;
		if (this.getExceptionCode() != null) {
			message = this.getExceptionCode().getExceptionMessage();
		} else {
			message = super.getMessage();
		}
		return message;
	}

	/**
	 * @see #Throwable.getLocalizedMessage
	 */
	public String getLocalizedMessage() {
		String localizedMessage = null;
		if (this.getExceptionCode() != null) {
			localizedMessage = this.getExceptionCode()
					.getLocalizedExceptionMessage();
		} else {
			localizedMessage = super.getLocalizedMessage();
		}
		return localizedMessage;
	}

	/**
	 * @return the exceptionCode
	 */
	public ExceptionCode getExceptionCode() {
		return this.exceptionCode;
	}

	/**
	 * @param exceptionCode
	 *            the exceptionCode to set
	 */
	public void setExceptionCode(ExceptionCode exceptionCode) {
		this.exceptionCode = exceptionCode;
	}
}
