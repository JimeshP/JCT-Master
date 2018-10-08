var imgByte = null;

/**
 * function handles the login functionality
 */
$(document).ready(function() {	
	callMe();
    $(".login").click(function () {
    	$('input[type=text], textarea').placeholder();
    $(".login_area_collapse").slideToggle();
    $(".password_area_collapse").slideUp("slow");   
    document.getElementById('forgotEmail').value="";
    });

    $(".forgot-pass").click(function () {
    	$('input[type=text], textarea').placeholder();
    $(".password_area_collapse").slideToggle();
    $(".login_area_collapse").slideUp("slow");
    document.getElementById('inputEmails').value="";
    document.getElementById('pwd').value="";
    }); 
    
    $("#loginBtn").click(function () { 
    	var myReset = sessionStorage.getItem("isNextReset");
    	//Clear all the session items if any
    	if (!sessionStorage.getItem("isReset")) {
    		sessionStorage.clear();
    	}
    	
    	
    	var myEmail = $('#inputEmails').val();
    	var password = $('#pwd').val();
    	//If the validation passes, proceed further
    	if(validateLoginCredential(myEmail, password) == true) { 
    		
    		//Get the screen resolution
        	var screenResolution = getCurrentScreenResolution();
        	
        	//Get the user os
        	var opertngSys = getOSDetails();
    		
    	$(".loader_bg").fadeIn();
    	$(".loader").fadeIn();
    	//Create a model
		var Contact = Spine.Model.sub();
		Contact.configure("/user/auth/authenticate", "emailId", "password", "userScreenResolution", "userOS");
		Contact.extend( Spine.Model.Ajax );

		//Populate the model with data to transfer
		var modelPopulator = new Contact({  
			emailId: myEmail, 
			password: encode(password+password),
			userScreenResolution: screenResolution,
			userOS: opertngSys
		});
		modelPopulator.save(); //POST
		
		Contact.bind("ajaxError", function(record, xhr, settings, error) {
			//var jsonStr = JSON.stringify(xhr);
			alertify.alert("<img src='img/server-na-icon.png'><br /><p>Unable to connect to the server.</p>");
			$(".loader_bg").fadeOut();
		    $(".loader").hide();
		});
		
		Contact.bind("ajaxSuccess", function(record, xhr, settings, error) {
			var jsonStr = JSON.stringify(xhr);
			var obj = jQuery.parseJSON(jsonStr);
			var statusCode = obj.statusCode;
			var url = obj.url;
			if((url == "/user/view/resultPage.jsp") || (url == "/user/view/resultPageFinal.jsp")) {
				sessionStorage.setItem("infoDispReq", "No");
			}
			
			/** THIS IS FOR PUBLIC VERSION **/
			if (obj.jctUserId) {
				sessionStorage.setItem("jctUserId", obj.jctUserId);
			}
			/********************************/
			
			if(statusCode == 200){ //If user activated and login is success
				if (obj.accountExpWarning) {
					sessionStorage.setItem("accountExp", obj.accountExpWarning);
					sessionStorage.setItem('notificationIconColor','red');
				}
				sessionStorage.setItem("restartExcersize",0);
				sessionStorage.setItem("linkClicked", "N");
				sessionStorage.setItem("name", obj.firstName);
				sessionStorage.setItem("jctEmail", myEmail);
				sessionStorage.setItem("pageSequence", obj.pageSequence);
				sessionStorage.setItem("jobReferenceNo", obj.jobRefNo); //Job reference
				sessionStorage.setItem("next_page", obj.lastPage);
				sessionStorage.setItem("bsView", "D");
				sessionStorage.setItem("as1View", "D");
				sessionStorage.setItem("as2View", "D");
				sessionStorage.setItem("profileId", obj.profileId);
				sessionStorage.setItem("groupId", obj.groupId);
				sessionStorage.setItem("isCompleted", obj.isCompleted);
				sessionStorage.setItem("myPage", "BS");
				
				if (obj.source) {
					sessionStorage.setItem("src", obj.source);
				}
				if (obj.demographicPopupMsg) {
					sessionStorage.setItem("demoInfoMsg", obj.demographicPopupMsg);
				}
				if (obj.correct) {
					sessionStorage.setItem("correctInfo", obj.correct);
				}
				if (obj.wrong) {
					sessionStorage.setItem("changeInfo", obj.wrong);
				}
				if((obj.bsJson != "none") && (null != obj.bsJson)) {
					sessionStorage.setItem("total_json_add_task", obj.bsJson); //Job reference
					var myJsonObj = JSON.parse(obj.bsJson);
					sessionStorage.setItem("total_count", myJsonObj.length+1);
					sessionStorage.setItem("div_intial_value", obj.initialJson);
					sessionStorage.setItem("myPage", "BS");
				}
				
				sessionStorage.setItem("totalTime",obj.totalTime);
				sessionStorage.setItem("rowIdentity", obj.identifier);
				sessionStorage.setItem("jobTitle", obj.jobTitle);
				sessionStorage.setItem("isLogout", "N");
				if(obj.pageSequence == 2){
					//sessionStorage.setItem("questions", obj.list);
					sessionStorage.setItem("snapShotURLS", obj.jctBaseString);
					sessionStorage.setItem("myPage", "BS");
				}
				if(obj.pageSequence == 3){
					sessionStorage.setItem("back_to_previous", 1);
					sessionStorage.setItem("strength",JSON.stringify(obj.strengthMap));
					sessionStorage.setItem("passion",JSON.stringify(obj.passionMap));
					sessionStorage.setItem("value",JSON.stringify(obj.valueMap));
					sessionStorage.setItem("checked_element", obj.afterSketchCheckedEle);
					sessionStorage.setItem("total_json_obj", obj.afterSkPageOneTotalJson);
					sessionStorage.setItem("strength_count", obj.strengthCount);
					sessionStorage.setItem("passion_count", obj.passionCount);
					sessionStorage.setItem("value_count", obj.valueCount);
					
					sessionStorage.setItem("checked_passion", obj.checkedPassion);
					sessionStorage.setItem("checked_strength", obj.checkedStrength);
					sessionStorage.setItem("checked_value", obj.checkedValue);
					sessionStorage.setItem("myPage", "AS");
				} if(obj.pageSequence == 4){
					sessionStorage.setItem("total_json_obj", obj.pageOneJson);
					sessionStorage.setItem("div_intial_value", obj.divInitialValue);
					sessionStorage.setItem("total_json_after_sketch_final", obj.afterSketch2TotalJsonObj);
					//sessionStorage.setItem("total_count", JSON.parse(obj.totalJsonAddTask).length+1);
					sessionStorage.setItem("total_count", obj.asTotalCount);
					sessionStorage.setItem("divHeight", obj.divHeight);
					sessionStorage.setItem("divWidth", obj.divWidth);
					sessionStorage.setItem("myPage", "AS");
				} if(obj.pageSequence == 5){
					//sessionStorage.setItem("actionPlanQtnList", obj.list);
					sessionStorage.setItem("existing", obj.existing);
					sessionStorage.setItem("snapShotURLS", obj.snapShot);
					sessionStorage.setItem("myPage", "AS");
				} if(obj.pageSequence == 6) {
					sessionStorage.setItem("myPage", "AS");
				}
				sessionStorage.setItem("url", url);
				if (obj.excCompletionCount > 0) {
					if (url == "/user/view/finalPage.jsp") {
						sessionStorage.setItem("pageSequence", 8);
					}
					window.location = url;
				} else {
					if ((myReset != null) && (myReset == "Y")){
						window.location = "landing-page.jsp";
					} else if (null == sessionStorage.getItem("isReset")) {
						window.location = "view/landing-page.jsp";
					} else if ((sessionStorage.getItem("isReset") != "Y")){
						window.location = "landing-page.jsp";
					} /*else if ((myReset != null) && (myReset == "Y")){
						window.location = "landing-page.jsp";
					}*/
					else{
						window.location = "view/landing-page.jsp";
					}
				}
				
				
			} else if(statusCode == 201) {//If user is still to be activated
				document.getElementById("pwd").value="";
				sessionStorage.setItem("statusMsg", obj.statusDesc);
				sessionStorage.setItem("jctEmail", myEmail);
				window.location.replace(url);
			} else if(statusCode == 404) {//User not found
				document.getElementById("pwd").value="";
				alertify.alert("<img src='img/failed-icon.png'><br /><p>"+obj.statusDesc+"<p>");
				$(".loader_bg").fadeOut();
			    $(".loader").hide();
			} else if(statusCode == 224) {//User found but password incorrect
				document.getElementById("pwd").value="";
				alertify.alert("<img src='img/failed-icon.png'><br /><p>"+obj.statusDesc+"<p>");
				$(".loader_bg").fadeOut();
			    $(".loader").hide();
			} else if(statusCode == 987) {//Admin deactivated the user
				document.getElementById("pwd").value="";
				alertify.alert("<img src='img/failed-icon.png'><br /><p>"+obj.statusDesc+"<p>");
				$(".loader_bg").fadeOut();
			    $(".loader").hide();
			} else if(statusCode == 990) {//User account expired
				//document.getElementById("pwd").value="";
				//alertify.alert("<img src='img/failed-icon.png'><br /><p>"+obj.statusDesc+"<p>");
				//$(".loader_bg").fadeOut();
			    //$(".loader").hide();
				var expDate = obj.expirationDate;
				sessionStorage.setItem("expDate", expDate);
				sessionStorage.setItem("jctEmail", myEmail);
				window.location = "view/accountRenew.jsp";
				
				
			} else if(statusCode == 776) {//User account expired
				/*document.getElementById("pwd").value="";
				alertify.alert("<img src='img/failed-icon.png'><br /><p>"+obj.statusDesc+"<p>");
				$(".loader_bg").fadeOut();
			    $(".loader").hide();*/				
				var expDate = obj.expirationDate;
				sessionStorage.setItem("expDate", expDate);
				sessionStorage.setItem("jctEmail", myEmail);
				window.location = "view/accountRenew.jsp";
			}
			else {	// Check with other status codes... Hence it is copy paste
				document.getElementById("pwd").value="";
				alertify.alert("<img src='img/failed-icon.png'><br /><p>"+obj.statusDesc+"<p>");
				$(".loader_bg").fadeOut();
			    $(".loader").hide();
			}
		});
    }
});
    
    function getDivCount(jSon){
    	
    }
    
    function getCurrentScreenResolution() {
    	var sHeight = "Not-Found";
    	var sWidth = "Not-Found";
    	try {
    		sHeight = screen.height;
        	sWidth = screen.width;
	    } catch(error) {}
    	return sWidth+"/"+sHeight;
    }
    
    function getOSDetails() {
    	var OSName="Unknown OS";
    	if (navigator.appVersion.indexOf("Win")!=-1) OSName="Windows";
    	if (navigator.appVersion.indexOf("Mac")!=-1) OSName="MacOS";
    	if (navigator.appVersion.indexOf("X11")!=-1) OSName="UNIX";
    	if (navigator.appVersion.indexOf("Linux")!=-1) OSName="Linux";
    	return OSName;
    }
/**
 * function checks whether the selected username 
 * is free to use or not. Email Id is treated as a user name.
 */
$("#inputEmail").blur(function () {
    	var selectedUserName = $('#inputEmail').val();
    	if((selectedUserName != "") && (validateEmail(selectedUserName) == true)){
    	//Create a model
		var Contact = Spine.Model.sub();
		Contact.configure("/user/auth/checkUser", "emailId");
		Contact.extend( Spine.Model.Ajax );

		//Populate the model with data to transder
		var getHTML = new Contact({  
			emailId: selectedUserName
		});
		
		var userDetails = new Object();
		userDetails.name = selectedUserName;
		getHTML.save(); //POST
		
		Contact.bind("ajaxError", function(record, xhr, settings, error) {
			//var jsonStr = JSON.stringify(xhr);
			alertify.alert("<img src='img/server-na-icon.png'><br /><p>Unable to connect to the server.</p>");
			$(".loader_bg").fadeOut();
		    $(".loader").hide();
		});
		
		Contact.bind("ajaxSuccess", function(record, xhr, settings, error) {
		var jsonStr = JSON.stringify(xhr);
		//var obj = jQuery.parseJSON(jsonStr);
		var status = jsonStr.indexOf("userExists");
		if(status !=  -1){
			//document.getElementById('signUpBtn').
			$('#signUpBtn').prop('disabled', true).addClass('ui-disabled');
			document.getElementById('userInfo').innerHTML = "<img src='img/wrong.png' /> User already exists.";
		}else{
			//$("#signUpBtn")
			$('#signUpBtn').prop('disabled', false).addClass('ui-enabled');
			document.getElementById('userInfo').innerHTML = "<img src='img/right.gif' /> Available.";
		}
		});
    	}
    });
});

/**
 * function to validate the email id
 */
function validateEmail(email){
	var emailReg = /^\w+([-+.']\w+)*@\w+([-.]\w+)*\.\w+([-.]\w+)*$/;
	if (!emailReg.test(email)){
		alertify.alert("<img src='img/alert-icon.png'><br /><p>Please enter valid E-mail ID.</p>");
		return false;
	} else {
		return true;
	}
}

/**
 * function validates the login credentials
 */
function validateLoginCredential(emailId, password){
	//validation - username
	//var illegalChars = /\W/; // allow letters, numbers, and underscores
	//var illegalChars = /^(?=.*\d)(?=.*[a-z])(?=.*[A-Z])(?=.*[^a-zA-Z0-9])(?!.*\s).{8,16}$/;
	var emailReg = /^\w+([-+.']\w+)*@\w+([-.]\w+)*\.\w+([-.]\w+)*$/;
	if (!emailReg.test(emailId)){
		alertify.alert("<img src='img/alert-icon.png'><br /><p>Please enter valid E-mail ID.</p>");
		return false;
	} else if (password == "") {
    	alertify.alert("<img src='img/alert-icon.png'><br /><p>Please enter password.</p>");
        return false;
    } /*else if ((password.length < 8) || (password.length > 16)) {
    	alertify.alert('Length of password should be between 8 - 16.');
        return false;
    } else if (!illegalChars.test(password)) {
    	alertify.alert('Password should contain atleast one lowercase letter, one uppercase letter, one numeric digit, and one special character.');
        return false;
    } */
	return true;
}
/**
 * function validates the signup form data
 * @param firstName
 * @param lastName
 * @param location
 * @param jobTitle
 * @param emailId
 * @param dateOfBirth
 * @param answer
 * @returns {Boolean}
 */
function validateSignupForm(emailId, functionGrp, jobLevel){
	//var illegalChars = /\W/; // allow letters, numbers, and underscores
	var emailReg = /^\w+([-+.']\w+)*@\w+([-.]\w+)*\.\w+([-.]\w+)*$/;
	if (!emailReg.test(emailId)){
		alertify.alert('Please enter valid E-mail ID.');
		return false;
	}
	if (functionGrp == "") {
		alertify.alert('Please select Function Group.');
        return false;
    }
	if (jobLevel == "") {
		alertify.alert('Please select Job Level.');
        return false;
    }
	return true;
}

/**
 * function is used to register the new user
 */
$("#signUpBtn").click(function () {
	var firstName = $('#inputFN').val();
	var lastName = $('#inputLN').val();
	var location = document.getElementById("inputLocation").options[document.getElementById("inputLocation").selectedIndex].value;
	var emailId = $('#inputEmail').val();
	var gender = null;
	var supervisePeople = null;
	var tenure = document.getElementById("inputTenure").options[document.getElementById("inputTenure").selectedIndex].value;
	var functionGrp = document.getElementById("inputFunctionGroup").options[document.getElementById("inputFunctionGroup").selectedIndex].value;
	var jobLevel = document.getElementById("inputJobLevel").options[document.getElementById("inputJobLevel").selectedIndex].value;;
			
	if(validateSignupForm(emailId, functionGrp, jobLevel) == true)	{
		if(document.getElementById("male").checked){
			gender = document.getElementById("male").value;
		} else if(document.getElementById("female").checked){
			gender = document.getElementById("female").value;
		} else {
			gender = "U";
		}
		if(document.getElementById("superviseYes").checked){
			supervisePeople = document.getElementById("superviseYes").value;
		} else if(document.getElementById("superviseNo").checked){
			supervisePeople = document.getElementById("superviseNo").value;
		} else {
			supervisePeople = "U";
		}
		$(".loader_bg").fadeIn();
	    $(".loader").fadeIn();
		    
		//Create a model
		var userProfile = Spine.Model.sub();
		userProfile.configure("/user/auth/register", "firstNames", "lastNames", "location", "emailIds",  "genders", "supervisePeoples", "tenures", "functionGrps", "jobLevels","profileId");
		userProfile.extend( Spine.Model.Ajax );
			
		//Populate the model with data to transfer
		var modelPopulation = new userProfile({  
			firstNames: firstName,
			lastNames: lastName,
			location: location,
			emailIds: emailId,
			genders: gender,
			supervisePeoples: supervisePeople,
			tenures: tenure,
			functionGrps: functionGrp,
			jobLevels: jobLevel,
			profileId: sessionStorage.getItem("profileId")
		});
		modelPopulation.save(); //POST the data
		userProfile.bind("ajaxError", function(record, xhr, settings, error) {
			alertify.alert("Unable to connect to the server.");
		});
			
		userProfile.bind("ajaxSuccess", function(record, xhr, settings, error) {
			var jsonStr = JSON.stringify(xhr);
			var obj = jQuery.parseJSON(jsonStr);
			var statusCode = obj.statusCode;
			if((statusCode == 200 || statusCode == 128)){
				sessionStorage.setItem("jctEmail", obj.email);
				sessionStorage.setItem("statusMsg", obj.statusDesc);
				window.location = "view/resetPassword.jsp";
			} 
			else if(statusCode == 888){ // mail sent error
				alertify.alert("<img src='img/server-na-icon.png'><br /><p>"+obj.statusDesc+"</p>");
				$(".loader_bg").fadeOut();
			    $(".loader").hide();
			}
		});
	} 
});
/**
 * function encode the login password which 
 * is transmitted over the network.
 * @param text
 * @returns {String}
 */
function encode(text) { 
    //var ref = "1&(9%2^*37@86^54&&^%&^0CKj=mo(%np(^qs$YTuQ%^rxyd=ahg)&ei=lBfWzC@*KJM*&%ioON=PQajpSTUVR~XYDA*&^*HGE=ILBFWZ";
	//var ref = "Q%1wS~(RxGzPA$By&N!Oa_2#^CM>D)<bFWLE@3c4*Kd5J+IVe-6Hf7g8=h9iT0jU";
	//var ref = "P!QR#S$T%U&VW(X)Y*Z+[,-].^/_0`1a2b3c4d5e6f7g8h9i:j;k<l=m>n?o@pAqBrCsDtEuFvGwHxIyJzK{L|M}N~O";
	var ref = "P!QR#S$T%U&VW(X)Y*Z+[,-].^/_0`1a2b3c4d5e6f7g8h9i:j;k<l=m>n?o@pAqBrCsDtEuFvGwHxIyJzK{L|M}N~O";
	var result="";
	for (var count=0; count<text.length; count++) {
	    var char=text.substring (count, count+1); 
	    var num=ref.indexOf (char);
	    var encodeChar=ref.substring(num+1, num+2);
	    result += encodeChar;
	}
	return result;
}


/**
 * Function handles the forgot password functionality
 */
function forgotPwd() {
	var toSend = $('#forgotEmail').val();
	var emailReg = /^\w+([-+.']\w+)*@\w+([-.]\w+)*\.\w+([-.]\w+)*$/;
	if (!emailReg.test(toSend)){
		alertify.alert("<img src='img/alert-icon.png'><br /><p>Please enter valid E-mail ID.</p>");
		return false;
	}
	$(".loader_bg").fadeIn();
	$(".loader").fadeIn();
	//Create a model
	var forgotPwd = Spine.Model.sub();
	forgotPwd.configure("/user/auth/forgotPassword", "myEmailId");
	forgotPwd.extend( Spine.Model.Ajax );
	
	//Populate the model
	var populateModel = new forgotPwd({  
		myEmailId: toSend
	});
	populateModel.save(); //POST
	
	forgotPwd.bind("ajaxError", function(record, xhr, settings, error) {
		alertify.alert("<img src='img/right.gif'><br /><p>Your password has been reset.<br />Please check your inbox for the new password.</p>");
		//sessionStorage.setItem("jctEmail", toSend);
		//sessionStorage.setItem("statusMsg", "Please reset your password");
		//window.location = "view/resetPassword.jsp";
		$(".loader_bg").fadeOut();
	    $(".loader").hide();
	    document.getElementById("closeId").click();
	});
	
	forgotPwd.bind("ajaxSuccess", function(record, xhr, settings, error) {
		var jsonStr = JSON.stringify(xhr);
		var obj = jQuery.parseJSON(jsonStr);
		var statusCode = obj.statusCode;
		if((statusCode == 200)){ //updated properly
			//store the username and the status message in session storage
			//sessionStorage.setItem("jctEmail", obj.email);
			//sessionStorage.setItem("statusMsg", obj.statusDesc);
			//window.location = "view/resetPassword.jsp";
			alertify.alert("<img src='img/right.gif'><br /><p>Your password has been reset.<br />Please check your inbox for the new password.</p>");
			$(".loader_bg").fadeOut();
		    $(".loader").hide();
		    document.getElementById("closeId").click();
		} else if (statusCode = 874) { //Password reset is not complete
			alertify.alert("<img src='img/failed-icon.png'><br /><p>"+obj.statusDesc+"</p>");
			$(".loader_bg").fadeOut();
		    $(".loader").hide();
		} else if (statusCode = 989) { //Password reset is not complete
			alertify.alert("<img src='img/failed-icon.png'><br /><p>"+obj.statusDesc+"</p>");
			$(".loader_bg").fadeOut();
		    $(".loader").hide();
		}else {
			alertify.alert("<img src='img/failed-icon.png'><br /><p>"+obj.statusDesc+"</p>");
			$(".loader_bg").fadeOut();
		    $(".loader").hide();
			return false;
		}
	});
}

/**
 * function used by the landing page to move to 
 * the last saved page by the user.
 */
function start(){
	window.location = sessionStorage.getItem("url");
}

/**
 * function to call login against Enter key.
 */
function searchKeyPress(e)
{
    // look for window.event in case event isn't passed in
    if (typeof e == 'undefined' && window.event) { e = window.event; }
    if (e.keyCode == 13)
    {
        document.getElementById('loginBtn').click();
    }
}

/**
 * function to call forgot password against Enter key.
 */
function onKeyPress(e)
{
    // look for window.event in case event isn't passed in
    if (typeof e == 'undefined' && window.event) { e = window.event; }
    if (e.keyCode == 13)
    {
        document.getElementById('fgtMyPwd').click();
    }
}


/**
 * Function add to allow only alphabets as a input
 * in first name and last name field
 * @param evt
 */
function alphaOnly(event) {	
	 var c = event.which || event.keyCode;
	  if((c > 32 && c < 65) || (c > 90 && c < 97) ||
	     (c > 122 && c !== 127 ))
	    return false;
	  
}

function callMe() {
	var parameter = window.location.search.replace( "?", "" ); // will return the GET parameter 
	var values = parameter.split("=");
	if (values[1] == "804") {
		alertify.alert("<img src='img/info-icon.png'><br /><p>You are not associated to any session. <br /> Please log in.</p>");
	}
}