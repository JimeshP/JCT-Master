package com.vmware.jct.security;
import java.io.IOException;

import javax.servlet.Filter;
import javax.servlet.FilterChain;
import javax.servlet.FilterConfig;
import javax.servlet.ServletException;
import javax.servlet.ServletRequest;
import javax.servlet.ServletResponse;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import javax.servlet.http.HttpSession;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;


/**
 * 
 * <p><b>Class name:</b> LoginPageFilter.java</p>
 * <p><b>Author:</b> InterraIT</p>
 * <p><b>Description:</b>  This class basically acts as Login filter </p>
 * <p><b>Copyrights:</b> 	All rights reserved by InterraIT and should only be used for its internal application development.</p>
 * <p><b>Revision History:</b>
 * 	<li>InterraIT - 31/Jan/2014 - Implement comments, introduce Named query </li>
 * </p>
 */
public class LoginPageFilter implements Filter
{
	private static final Logger logger = LoggerFactory.getLogger(LoginPageFilter.class);
	   public void init(FilterConfig filterConfig) throws ServletException
	   {

	   }

	   public void doFilter(ServletRequest servletRequest, ServletResponse servletResponse,   FilterChain filterChain) throws IOException, ServletException
	   {
		   logger.info(">>>>  LoginPageFilter.doFilter");
	       HttpServletRequest request = (HttpServletRequest) servletRequest;
	       HttpServletResponse response = (HttpServletResponse) servletResponse;
	      request.getRequestURL();
	       HttpSession httpSession = request.getSession();
	       Boolean isLoggedIn = (Boolean)httpSession.getAttribute("isLoggedIn");
	       
	       if(isLoggedIn != null && isLoggedIn==true){ //If user is already authenticated
	    	   logger.debug("Inside Filter if");
	    	   filterChain.doFilter(servletRequest, servletResponse);
	       } 
	       else if("/admin/".equalsIgnoreCase(request.getRequestURI()) || "/admin/signup-page.jsp".equalsIgnoreCase(request.getRequestURI()) || "/admin/view/activateFacilitator.jsp".equalsIgnoreCase(request.getRequestURI())){
	    	   logger.debug("sign up page called");
	    	   filterChain.doFilter(servletRequest, servletResponse);
	       }
	       else{
	    	   logger.debug("Inside Filter else");
	    	   response.sendRedirect("/admin/");// or, forward using RequestDispatcher	          
	       }
	       logger.info("<<<<  LoginPageFilter.doFilter");
	   }

	   public void destroy()
	   {

	   }
	}

