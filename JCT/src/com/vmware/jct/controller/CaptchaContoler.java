package com.vmware.jct.controller;

import java.io.IOException;

import javax.servlet.http.HttpServletRequest;

import net.tanesha.recaptcha.ReCaptchaImpl;
import net.tanesha.recaptcha.ReCaptchaResponse;

import org.codehaus.jackson.JsonNode;
import org.codehaus.jackson.JsonProcessingException;
import org.codehaus.jackson.map.ObjectMapper;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.RequestBody;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestMethod;
import org.springframework.web.bind.annotation.ResponseBody;

@Controller
public class CaptchaContoler {

	@RequestMapping(value = "/validateCaptcha", method = RequestMethod.POST)
	@ResponseBody()
	public String validateCaptcha(@RequestBody String body, HttpServletRequest request){
		ObjectMapper mapper = new ObjectMapper();
		JsonNode node = null;
		try {
			node = mapper.readTree(body);
		} catch (JsonProcessingException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		}
		
		String remoteAddr = request.getRemoteAddr();
		ReCaptchaImpl reCaptcha = new ReCaptchaImpl();
		reCaptcha.setPrivateKey(node.get("challenges").toString().replaceAll("\"", ""));
		ReCaptchaResponse reCaptchaResponse =
			    reCaptcha.checkAnswer(remoteAddr, node.get("challenges").toString().replaceAll("\"", ""), node.get("responses").toString().replaceAll("\"", ""));
		
		return reCaptchaResponse.getErrorMessage();
	}

}
