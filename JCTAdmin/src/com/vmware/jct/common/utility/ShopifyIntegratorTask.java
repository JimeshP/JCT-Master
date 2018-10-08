package com.vmware.jct.common.utility;

import java.io.BufferedReader;
import java.io.IOException;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.io.OutputStreamWriter;
import java.net.MalformedURLException;
import java.net.URL;
import java.net.URLConnection;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.context.MessageSource;
import org.springframework.stereotype.Service;

import com.google.gson.Gson;
import com.vmware.jct.dao.IAuthenticatorDAO;
import com.vmware.jct.dao.ICommonDao;
import com.vmware.jct.dao.IManageServiceDAO;
import com.vmware.jct.exception.DAOException;
import com.vmware.jct.exception.JCTException;
import com.vmware.jct.exception.MailingException;
import com.vmware.jct.model.JctUser;
import com.vmware.jct.service.IManageUserService;
import com.vmware.jct.service.vo.FacilitatorAccountVO;
import com.vmware.jct.service.vo.GeneralUserAccountVO;
import com.vmware.jct.service.vo.UserVO;

@Service("shopifyIntegratorTask")
public class ShopifyIntegratorTask {
	private static final Logger LOGGER = LoggerFactory.getLogger(ShopifyIntegratorTask.class);
	
	@Autowired  
	private MessageSource messageSource;
	
	@Autowired
	private ICommonDao commonDaoImpl;
	
	@Autowired
	private IManageServiceDAO serviceDAO;
	
	@Autowired
	private IAuthenticatorDAO authenticatorDAO;
	
	@Autowired
	private IManageUserService service;
	
	@Autowired
	private MailNotificationHelper mailer;
	
	public void pingShopify() {
		LOGGER.info(">>>>>>>> ShopifyIntegratorTask.pingShopify");
		String apiHostName = this.messageSource.getMessage("shopify.api.hostname",null, null);
		String apiTemplate = this.messageSource.getMessage("shopify.api.template",null, null);
		String apiKey = this.messageSource.getMessage("shopify.api.key",null, null);
		String apiPwd = this.messageSource.getMessage("shopify.api.pwd",null, null);
		try {
			String dataFromShopify = connectAndFetchData(apiHostName, apiTemplate, apiKey, apiPwd);
			if (dataFromShopify.length() > 30) {
				processShopifyData(dataFromShopify);
			} else {
				LOGGER.info("NO ORDERS FOUND FROM ECOMMERCE: "+new Date());
			}
		} catch (MailingException mex) {
			LOGGER.error("UNABLE TO SEND MAIL..."+mex.getLocalizedMessage());
		} catch (DAOException dex) {
			dex.printStackTrace();
		} catch (MalformedURLException ex) {
			LOGGER.error("PLEASE CHECK THE URL..."+ex.getLocalizedMessage());
		} catch (IOException e) {
			e.printStackTrace();
		}
	}
	/**
	 * Method connects to the shopify and fetches the order data
	 * @param apiHostName
	 * @param apiTemplate
	 * @param apiKey
	 * @param apiPwd
	 * @return
	 * @throws MalformedURLException
	 * @throws IOException
	 */
	private String connectAndFetchData (final String apiHostName, final String apiTemplate, 
			final String apiKey, final String apiPwd) throws IOException {
		URL url = new URL(String.format(apiTemplate, apiKey, apiPwd, apiHostName));
		URLConnection conn = url.openConnection();
		String userpass = apiKey + ":" + apiPwd;
		String basicAuth = "Basic " + javax.xml.bind.DatatypeConverter.printBase64Binary(userpass.getBytes());
		conn.setRequestProperty ("Authorization", basicAuth);
		InputStream input = conn.getInputStream();
		BufferedReader br = null;
		StringBuilder sb = new StringBuilder();
		String line;
		br = new BufferedReader(new InputStreamReader(input));
		while ((line = br.readLine()) != null) {
			sb.append(line);
		}
		return sb.toString();
	}
	/**
	 * Method process and stores data in database
	 * @param shopifyData
	 * @throws DAOException
	 * @throws MailingException
	 * @throws IOException 
	 */
	private void processShopifyData(String shopifyData) throws DAOException, MailingException, IOException {
		Gson gson = new Gson();
		ShopifyOrderModel model = gson.fromJson(shopifyData, ShopifyOrderModel.class);
		List<ShopifyModel> listOfOrders = model.getOrders();
		// Iterate over all the orders
		for (int orderIndex = 0; orderIndex < listOfOrders.size(); orderIndex++) {
			ShopifyModel order = (ShopifyModel) listOfOrders.get(orderIndex);
			List<ShopifyNodeAttributeModel> noteAttributes = order.getNote_attributes();
			// First array customer type
			ShopifyNodeAttributeModel obj1 = (ShopifyNodeAttributeModel) noteAttributes.get(0);
			String custType = obj1.getValue();
			
			// Second array Subscription type
			ShopifyNodeAttributeModel obj2 = (ShopifyNodeAttributeModel) noteAttributes.get(1);
			String subscriptionType = obj2.getValue();
			
			// Third array emailId
			ShopifyNodeAttributeModel obj3 = (ShopifyNodeAttributeModel) noteAttributes.get(2);
			String email = obj3.getValue();
			
			if (custType.equals(CommonConstants.SHOPIFY_INDIVIDUAL_CUSTOMER)) {
				if (subscriptionType.equals(CommonConstants.SHOPIFY_NEW_SUBSCRIPTION)) {
					createNewIndividualUser(email, order);
				} else if (subscriptionType.equals(CommonConstants.SHOPIFY_RENEW_SUBSCRIPTION)) {
					renewIndividualCustomer(email, order);
				} else {
					LOGGER.error("INVALID SUBSCRIPTION TYPE FOUND: "+subscriptionType);
				}
			} else if (custType.equals(CommonConstants.SHOPIFY_FACILITATOR_CUSTOMER)) {
				if (subscriptionType.equals(CommonConstants.SHOPIFY_NEW_SUBSCRIPTION)) {
					// Fourth array emailId
					ShopifyNodeAttributeModel obj4 = (ShopifyNodeAttributeModel) noteAttributes.get(3);
					int numberOfUsers = Integer.parseInt(obj4.getValue());
					
					// Facilitator to have tool access
					ShopifyNodeAttributeModel obj5 = (ShopifyNodeAttributeModel) noteAttributes.get(4);
					String toolAccess = obj5.getValue();
					if (toolAccess.equals("Y")) {
						numberOfUsers = numberOfUsers - 1;
					}
					createNewFacilitator(email, order, numberOfUsers, toolAccess);
				} else if (subscriptionType.equals(CommonConstants.SHOPIFY_RENEW_SUBSCRIPTION)) {
					// Fourth array emailId
					ShopifyNodeAttributeModel obj4 = (ShopifyNodeAttributeModel) noteAttributes.get(3);
					String facilitatorId = obj4.getValue();
					
					/*ShopifyNodeAttributeModel obj5 = (ShopifyNodeAttributeModel) noteAttributes.get(4);
					int numberOfUsers = Integer.parseInt(obj5.getValue());*/
					
					// Get fulfillable quanity as user has the option to change the quantity
					List<ShopifyLineItemsModel> lineItemsList = order.getLine_items();
					ShopifyLineItemsModel lineItems = (ShopifyLineItemsModel) lineItemsList.get(0);
					int numberOfUsers = Integer.parseInt(lineItems.getQuantity());
					
					renewSubscriptionCount(email, order, facilitatorId, numberOfUsers);
				} else {
					LOGGER.error("INVALID SUBSCRIPTION TYPE FOUND: "+subscriptionType);
				}
			} else {
				LOGGER.error("INVALID CUSTOMER TYPE FOUND: "+custType);
			}
		}
	}
	/**
	 * Method creates new individual account
	 * @param email
	 * @param order
	 * @return
	 * @throws DAOException
	 * @throws MailingException
	 * @throws IOException
	 */
	private boolean createNewIndividualUser(String email, ShopifyModel order) throws DAOException, MailingException, IOException {
		boolean isSuccess = false;
		List<String> validEmailIdList = new ArrayList<String>();
		validEmailIdList.add(email + "#`#" +"FirstName" + "#`#" +"LastName" + "~");
		UserVO userVO = new UserVO();		
		GeneralUserAccountVO accountVO = new GeneralUserAccountVO(order.getEmail(),
				CommonConstants.SHOPIFY_ECOMMERCE, "6", "1!Default User Group",
				validEmailIdList, new ArrayList<String>(), getPresentFormattedDate(), "1",
				"", order.getTotal_price());
		try {
			if(authenticatorDAO.checkUserByEmailAndRole(userVO,1).size() != 0 ){
				isSuccess = false;	
				cancelOrder(order.getId());
			} else {
				service.saveGeneralUserViaChequePayment(accountVO, validEmailIdList.size(), CommonConstants.SHOPIFY_ECOMMERCE);
				isSuccess = true;
				// Once the user is created, the order will be closed.
				closeOrder(order.getId());				
			}		
			
			
		} catch (JCTException ex) {
			LOGGER.error("UNABLE TO SAVE GENERAL USER BY MANUAL: "+ex.getLocalizedMessage());
		}
		return isSuccess;
	}
	/**
	 * Method renews the individual user account
	 * @param email
	 * @param order
	 * @return
	 * @throws DAOException
	 * @throws MailingException
	 * @throws IOException
	 */
	private boolean renewIndividualCustomer(String email, ShopifyModel order)
			throws DAOException, MailingException, IOException {
		boolean isSuccess = false;
		UserVO userVO = new UserVO();
		userVO.setEmail(email);
		List<JctUser> userLst = authenticatorDAO.checkUserByEmailAndRole(userVO, CommonConstants.GENERAL_USER);
		if (userLst.size() == 0) {
			// Email the users and admin about failure (user exists)
			LOGGER.info("USER: " + email + " IS NOT REGISTERED....");
		} else {
			List<String> emailWithExpDateList = new ArrayList<String>();
			emailWithExpDateList.add(email+"#`#"+"6");
			//emailWithExpDateList.add("6~");
			GeneralUserAccountVO accountVO = new GeneralUserAccountVO(order.getEmail(),
					CommonConstants.SHOPIFY_ECOMMERCE, emailWithExpDateList, getPresentFormattedDate(), "1",
					"", order.getTotal_price());
			try {
				service.renewGeneralUserPaymentManual(accountVO, 1, CommonConstants.SHOPIFY_ECOMMERCE);
				isSuccess = true;
				// Once the user is created, the order will be closed.
				closeOrder(order.getId());
				
			} catch (JCTException e) {
				LOGGER.error("UNABLE TO RENEW INDIVIDUAL CUSTOMER FOR USER: "+email);
			}
		}
		return isSuccess;
	}
	/**
	 * Method renews subscription count for facilitator
	 * @param email
	 * @param order
	 * @param facilitatorId
	 * @param numberOfUsers
	 * @return
	 * @throws DAOException
	 * @throws MailingException
	 * @throws IOException
	 */
	private boolean renewSubscriptionCount(String email, ShopifyModel order,
			String facilitatorId, int numberOfUsers) throws DAOException,
			MailingException, IOException {
		boolean isSuccess = false;
		UserVO userVO = new UserVO();
		userVO.setEmail(email);
		List<JctUser> userLst = authenticatorDAO.checkUserByEmailAndRole(userVO, CommonConstants.FACILITATOR_USER);
		if (userLst.size() == 0) {
			// Email the users and admin about failure (user exists)
			LOGGER.info("USER: " + email + " IS NOT REGISTERED....");
		} else {
			FacilitatorAccountVO facilitatorChequeVO = new FacilitatorAccountVO(order.getEmail(), 
					facilitatorId, email, getPresentFormattedDate(), numberOfUsers+"",
					"", order.getTotal_price(), CommonConstants.SHOPIFY_ECOMMERCE);					
			try {
				service.renewSubscribeViaChequePayment(facilitatorChequeVO,"AD");
				isSuccess = true;
				// Once the user is created, the order will be closed.
				closeOrder(order.getId());
				
			} catch (JCTException e) {
				LOGGER.error("UNABLE TO RENEW SUBSCRIPTION COUNT FOR USER: "+email);
			}
		}
		return isSuccess;
	}
	/**
	 * Method creates new Facilitator
	 * @param email
	 * @param order
	 * @param numberOfUsers
	 * @param toolAccess
	 * @return
	 * @throws DAOException
	 * @throws MailingException
	 * @throws IOException 
	 */
	private boolean createNewFacilitator(String email, ShopifyModel order,
			int numberOfUsers, String toolAccess) throws DAOException,
			MailingException, IOException {
		boolean isSuccess = false;
		UserVO userVO = new UserVO();
		userVO.setEmail(email);
		List<JctUser> userLst = authenticatorDAO.checkUserByEmailAndRole(userVO, CommonConstants.FACILITATOR_USER);
		if (userLst.size() == 0) {
			FacilitatorAccountVO facilitatorChequeVO = new FacilitatorAccountVO("1!Default User Group", email, getPresentFormattedDate(), 
					numberOfUsers+"", "", order.getTotal_price(), toolAccess, order.getEmail(), CommonConstants.SHOPIFY_ECOMMERCE);
			try {
				if(authenticatorDAO.checkUserByEmailAndRole(userVO,3).size() != 0 ){
					isSuccess = false;
					cancelOrder(order.getId());
				} else {
					service.saveFacilitatorViaChequePayment(facilitatorChequeVO, CommonConstants.SHOPIFY_ECOMMERCE);
					isSuccess = true;
					// Once the user is created, the order will be closed.
					closeOrder(order.getId());
				}
			} catch (JCTException e) {
				LOGGER.error("UNABLE TO CREATE USER: "+email);
			}
		} else {
			// Email the users and admin about failure (user exists)
			LOGGER.info("USER: " + email + " IS ALREADY REGISTERED...");
			mailer.mailIndUserDuplicateAccountEcomerce(email);
		}
		return isSuccess;
	}
	/**
	 * Method closes the order id in shopify
	 * @param id
	 * @throws IOException
	 */
	private void closeOrder(String id) throws IOException {
		String apiHostName = this.messageSource.getMessage("shopify.api.hostname",null, null);
		String apiTemplate = this.messageSource.getMessage("shopify.api.close.order.template",null, null);
		String apiKey = this.messageSource.getMessage("shopify.api.key",null, null);
		String apiPwd = this.messageSource.getMessage("shopify.api.pwd",null, null);
		URL url = new URL(String.format(apiTemplate+id+"/close.json", apiKey, apiPwd, apiHostName));
		URLConnection conn = url.openConnection();
		String userpass = apiKey + ":" + apiPwd;
		String basicAuth = "Basic "
				+ javax.xml.bind.DatatypeConverter.printBase64Binary(userpass
						.getBytes());
		conn.setRequestProperty("Authorization", basicAuth);
		conn.setDoOutput(true);
		conn.setRequestProperty("content-type",
				"application/x-www-form-urlencoded");
		OutputStreamWriter out = new OutputStreamWriter(conn.getOutputStream());
		out.write("message=Hello world");
		out.flush();
		out.close();
		InputStream inputStream = conn.getInputStream();
	    String encoding = conn.getContentEncoding();
	}
	/**
	 * Method closes the order id in shopify
	 * @param id
	 * @throws IOException
	 */
	private void cancelOrder(String id) throws IOException {
		String apiHostName = this.messageSource.getMessage("shopify.api.hostname",null, null);
		String apiTemplate = this.messageSource.getMessage("shopify.api.close.order.template",null, null);
		String apiKey = this.messageSource.getMessage("shopify.api.key",null, null);
		String apiPwd = this.messageSource.getMessage("shopify.api.pwd",null, null);
		URL url = new URL(String.format(apiTemplate+id+"/cancel.json", apiKey, apiPwd, apiHostName));
		URLConnection conn = url.openConnection();
		String userpass = apiKey + ":" + apiPwd;
		String basicAuth = "Basic "
				+ javax.xml.bind.DatatypeConverter.printBase64Binary(userpass
						.getBytes());
		conn.setRequestProperty("Authorization", basicAuth);
		conn.setDoOutput(true);
		conn.setRequestProperty("content-type",
				"application/x-www-form-urlencoded");
		OutputStreamWriter out = new OutputStreamWriter(conn.getOutputStream());
		out.write("message=Hello world");
		out.flush();
		out.close();
		InputStream inputStream = conn.getInputStream();
	    String encoding = conn.getContentEncoding();
	}
	/**
	 * Method returns present date in MM/dd/yyyy format 
	 * @return
	 */
	private String getPresentFormattedDate() {
		SimpleDateFormat sdf = new SimpleDateFormat("MM/dd/yyyy");
		return sdf.format(new Date());
	}
}
