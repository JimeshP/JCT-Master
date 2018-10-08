package com.vmware.jct.dao;

import java.util.List;

import com.vmware.jct.dao.dto.PaymentDetailsDTO;
import com.vmware.jct.exception.DAOException;
import com.vmware.jct.model.JctCheckPaymentUserDetails;
import com.vmware.jct.model.JctPaymentDetails;
/**
 * 
 * <p><b>Class name:</b>IPaymentDAO.java</p>
 * <p><b>@author:</b> InterraIT </p>
 * <p><b>Purpose:</b> 		This DAO interface represents a IPaymentDAO interface
 * <p><b>Description:</b> 	This will have all methods mapped to the database and fetching data from database to be used for Manage user service. </p>
 * <p><b>Note:</b> 			This methods implemented by DAOImpl classes and execute the mapped queries and get the records.</p>  
 * <p><b>Copyrights:</b> 	All rights reserved by Interra IT and should only be used for its internal application development.</p>
 *  <pre>
 * <b>Revision History:</b>
 * <li> </li>
 * </pre>
 */
public interface IPaymentDAO {
	/**
	 * Method saves payment details
	 * @param details
	 * @return
	 * @throws DAOException
	 */
	public String savePaymentDetails (JctPaymentDetails details) throws DAOException;
	/**
	 * Method fetches all existing check payments users details which are qnique for all users 
	 * @param fetchType
	 * @param chequeNum
	 * @param userName
	 * @return
	 * @throws DAOException
	 */
	public List<PaymentDetailsDTO> fetchExistingChequeUsersUniqueFields(String fetchType, String chequeNum, String userName) throws DAOException;
	/**
	 * Method fetches by transaction Id
	 * @return
	 * @throws DAOException
	 */
	public List<JctCheckPaymentUserDetails> searchByTranId(String tranId) throws DAOException;
	/**
	 * Method fetches all existing cheque payment users details which are not unique
	 * @return
	 * @throws DAOException
	 */
	public List<PaymentDetailsDTO> fetchuserRemainingFields(String transId) throws DAOException;
	/**
	 * method to fetch check number using user name
	 * @param userName
	 * @return String 
	 * @throws DAOException
	 * */	
	public String getCHkNoByUserName(String userName)throws DAOException;
}
