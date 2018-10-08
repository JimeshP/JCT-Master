package com.vmware.jct.exception;

/**
 * @author InterraIT
 *
 */
public class MailingException extends JCTBaseException {

	private static final long serialVersionUID = 1L;

	public MailingException() {
		super();
	}

	public MailingException(String message, Throwable cause,
			boolean enableSuppression, boolean writableStackTrace) {
		super(message, cause, enableSuppression, writableStackTrace);
	}

	public MailingException(String message, Throwable cause) {
		super(message, cause);
	}

	public MailingException(String message) {
		super(message);
	}

	/*public MailingException(Throwable cause) {
		super(cause);
		// TODO Auto-generated constructor stub
	}*/
}
