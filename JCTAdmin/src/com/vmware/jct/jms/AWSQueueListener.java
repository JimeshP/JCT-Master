package com.vmware.jct.jms;

import java.io.IOException;

import javax.jms.JMSException;
import javax.jms.Message;
import javax.jms.MessageListener;
import javax.jms.TextMessage;
import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.parsers.ParserConfigurationException;

import org.apache.commons.io.IOUtils;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.beans.factory.annotation.Autowired;
import org.w3c.dom.Document;
import org.w3c.dom.Element;
import org.w3c.dom.Node;
import org.w3c.dom.NodeList;
import org.xml.sax.SAXException;

import com.vmware.jct.common.utility.CommonUtility;
import com.vmware.jct.common.utility.StatusConstants;
import com.vmware.jct.exception.JCTException;
import com.vmware.jct.service.IManageUserService;
import com.vmware.jct.service.vo.FacilitatorAccountVO;
/**
 * 
 * <p><b>Class name:</b> AWSQueueListener.java</p>
 * <p><b>Author:</b> InterraIT</p>
 * <p><b>Purpose:</b> This class acts as a listener to the message queue. It listens to AWS message queue for new registration.
 * AWSQueueListener implements MessageListener and has following  methods.
 * -onMessage(Message message)
 * <p><b>Description:</b> This class acts as a listener to the message queue. It listens to AWS message queue for new registration. </p>
 * <p><b>Copyrights:</b> All rights reserved by InterraIT and should only be used for its internal application development.</p>
 * <p><b>Created Date:</b> 03/Dec/2014</p>
 * <p><b>Revision History:</b>
 * 	<li></li>
 * </p>
 */
public class AWSQueueListener implements MessageListener {
	
	@Autowired
	private IManageUserService service;
	private static final Logger LOGGER = LoggerFactory.getLogger(AWSQueueListener.class);
	
	public void onMessage(Message message) {
		LOGGER.info(">>>>>>>> AWSQueueListener.onMessage");
		if (message instanceof TextMessage) {
			try {
				String msgText = ((TextMessage) message).getText();
				LOGGER.info("MESSAGE FROM QUEUE: -> "+msgText);
				if (msgText.contains("aws-facilitator")) {
					FacilitatorAccountVO faciVo = getFacilitatorAccountVO(msgText);
					if (null != faciVo) {
						int status = service.saveFacilitatorViaChequePayment(faciVo, "CHK");
						if (status == StatusConstants.STATUS_SUCCESS) {
							LOGGER.info("---- QUEUE-LISTENER: USER WAS SUCCUSSFULLY REGISTERED");
						} else if (status == StatusConstants.USER_REGISTERED_MAIL_SEND_ERROR) {
							LOGGER.info("---- QUEUE-LISTENER: USER WAS SUCCUSSFULLY REGISTERED BUT MAIL WAS NOT SENT");
						} else if (status == StatusConstants.TABLE_OR_DATA_RELATED_ERROR) {
							LOGGER.info("---- QUEUE-LISTENER: THERE ARE SOME TABLE OR DATA RELATED ERROR HENCE THE DATA IS NOT SAVED");
						} else {
							LOGGER.info("---- QUEUE-LISTENER: DATA IS NOT SAVED. FAILED.");
						}
					}
				} else {
					//TODO: FOR GENERAL USER
				}				
			} catch (JMSException e) {
				e.printStackTrace();
			} catch (ParserConfigurationException e) {
				e.printStackTrace();
			} catch (IOException e) {
				e.printStackTrace();
			} catch (SAXException e) {
				e.printStackTrace();
			} catch (JCTException e) {
				e.printStackTrace();
			}
		} else {
			LOGGER.error("THE MESSAGE IS NOT OF TEXT TYPE....");
			throw new RuntimeException();
		}
		LOGGER.info("<<<<<<<< AWSQueueListener.onMessage");
	}
	
	private FacilitatorAccountVO getFacilitatorAccountVO (String xml) throws ParserConfigurationException, IOException, SAXException {
		FacilitatorAccountVO vo = null;
		DocumentBuilderFactory factory = DocumentBuilderFactory.newInstance();
		DocumentBuilder builder = factory.newDocumentBuilder();
		Document document = builder.parse(IOUtils.toInputStream(xml, "UTF-8"));
		document.getDocumentElement().normalize();
		NodeList nList = document.getElementsByTagName("aws-facilitator");
		for (int temp = 0; temp < nList.getLength(); temp++) {
			Node nNode = nList.item(temp);
			if (nNode.getNodeType() == Node.ELEMENT_NODE) {
	 			Element eElement = (Element) nNode;
	 			Boolean isValid = CommonUtility.isEmailValid(eElement.getElementsByTagName("facilitator-email").item(0).getTextContent());
	 			if (isValid) {
	 				vo = new FacilitatorAccountVO(eElement.getElementsByTagName("user-group").item(0).getTextContent(), 
		 					//eElement.getElementsByTagName("bank-details").item(0).getTextContent(), 
		 					eElement.getElementsByTagName("facilitator-email").item(0).getTextContent(), 
		 					eElement.getElementsByTagName("payment-date").item(0).getTextContent(), 
		 					eElement.getElementsByTagName("number-of-users").item(0).getTextContent(), 
		 					eElement.getElementsByTagName("cheque-nos").item(0).getTextContent(), 
		 					eElement.getElementsByTagName("total-amount-to-be-paid").item(0).getTextContent(), 
		 					eElement.getElementsByTagName("facilitator-to-have-account").item(0).getTextContent(), 
		 					eElement.getElementsByTagName("created-by").item(0).getTextContent(), 
		 					eElement.getElementsByTagName("payment-type").item(0).getTextContent());
	 			} else {
	 				LOGGER.error("---- QUEUE-LISTENER: Invalid Email Id... Cannot construct the FacilitatorAccountVO");
	 			}
			}
		}
		return vo;
	}
}
