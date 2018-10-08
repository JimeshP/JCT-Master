package com.vmware.jct.exception;

/**
 * @author InterraIT
 *
 */
public class JCTException extends JCTBaseException {

	private static final long serialVersionUID = 1L;

	public JCTException() {
		super();
	}

	public JCTException(String message, Throwable cause,
			boolean enableSuppression, boolean writableStackTrace) {
		super(message, cause, enableSuppression, writableStackTrace);
	}

	public JCTException(String message, Throwable cause) {
		super(message, cause);
	}

	public JCTException(String message) {
		super(message);
	}

	/*public JCTException(Throwable cause) {
		super(cause);
		// TODO Auto-generated constructor stub
	}*/
}
