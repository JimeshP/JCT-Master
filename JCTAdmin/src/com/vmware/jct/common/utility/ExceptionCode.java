package com.vmware.jct.common.utility;

import java.io.Serializable;

/**
 * 
 * <p><b>Class name:</b> ExceptionCode.java</p>
 * <p><b>Author:</b> InterraIT</p>
 * <p><b>Purpose:</b> This class acts for exception.
 * <p><b>Copyrights:</b>All rights reserved by Interra IT and should only be used for its internal application development.</p>
 * <p><b>Created Date:</b> 10/Apr/2014</p>
 * <p><b>Revision History:</b>
 * 	<li></li>
 * </p>
 */
public class ExceptionCode implements Serializable {

	private static final long serialVersionUID = 1327072126348432757L;

    public static final String DEFAULT_EXCEPTION_CODE = "exception.default";
    
    public static final String DEFAULT_EXCEPTION_MESSAGE = "An exception has occured.";
    
    public static final String DEFAULT_PROPERTY_KEY = DEFAULT_EXCEPTION_CODE;
    
    public static final String DEFAULT_PROPERTY_VALUE = DEFAULT_EXCEPTION_MESSAGE;
    
    public static final String DEFAULT_LOCALIZED_PROPERTY_VALUE = DEFAULT_EXCEPTION_MESSAGE;
    
	/**
     * The exception code is the key for the property that we want. 
     */
    protected String exceptionCode;
    
    /**
     * Holds the description for this exception message.
     */
    protected String exceptionMessage;
    
    /**
     * The localized exception message, defaults to  exception message.
     */
    protected String localizedExceptionMessage;
    
    protected ExceptionCode () {
    	this(DEFAULT_PROPERTY_KEY);
    }
    
    protected ExceptionCode (String exceptionCode) {
    	this(exceptionCode, DEFAULT_PROPERTY_VALUE);
    }
    
    protected ExceptionCode (String exceptionCode, String exceptionMessage) {
    	this(exceptionCode, exceptionMessage, DEFAULT_LOCALIZED_PROPERTY_VALUE);
    }
    
    protected ExceptionCode (String exceptionCode, String exceptionMessage, String localizedExceptionMessage) {
    	 this.setExceptionCode(exceptionCode);
    	 this.setExceptionMessage(exceptionMessage);
    	 this.setLocalizedExceptionMessage(localizedExceptionMessage);
    }
    
    /**
     * @return the exceptionCode
     */
    public String getExceptionCode () {
        return this.exceptionCode;
    }

    /**
     * @param exceptionCode the exceptionCode to set
     */
    public void setExceptionCode (String exceptionCode) {
        this.exceptionCode = exceptionCode;
    }

    /**
     * @return the exceptionMessage
     */
    public String getExceptionMessage () {
        return this.exceptionMessage;
    }

    /**
     * @param exceptionMessage the exceptionMessage to set
     */
    public void setExceptionMessage (String exceptionMessage) {
        this.exceptionMessage = exceptionMessage;
    }

    /**
     * @return the localizedExceptionMessage
     */
    public String getLocalizedExceptionMessage () {
        return this.localizedExceptionMessage;
    }

    /**
     * @param localizedExceptionMessage the localizedExceptionMessage to set
     */
    public void setLocalizedExceptionMessage (String localizedExceptionMessage) {
        this.localizedExceptionMessage = localizedExceptionMessage;
    }
}
