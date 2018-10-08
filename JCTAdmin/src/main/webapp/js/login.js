
var imgByte = null;

/**
 * function handles the login functionality
 */
$(document).ready(function() {
  
    $("#loginBtn").click(function () { 
    	var myEmail = $('#inputEmails').val();
    	var password = $('#pwd').val();
    	//If the validation passes, proceed further
    	if(validateLoginCredential(myEmail, password) == true) { 
    		
    	$(".loader_bg").fadeIn();
    	$(".loader").fadeIn();
    	//Create a model
		var Contact = Spine.Model.sub();
		Contact.configure("/admin/authAdmin/authenticateAdmin", "emailId", "password");
		Contact.extend( Spine.Model.Ajax );

		//Populate the model with data to transfer
		var modelPopulator = new Contact({  
			emailId: myEmail, 
			password: encode(password+password) 
		});
		
		modelPopulator.save(); //POST
		
	    
		Contact.bind("ajaxError", function(record, xhr, settings, error) {
			//var jsonStr = JSON.stringify(xhr);
			alertify.alert("Unable to connect to the server.");
			$(".loader_bg").fadeOut();
		    $(".loader").hide();
		});
		
		Contact.bind("ajaxSuccess", function(record, xhr, settings, error) {
			var jsonStr = JSON.stringify(xhr);
			var obj = jQuery.parseJSON(jsonStr);
			var statusCode = obj.statusCode;
			var url = obj.url;	
			/** THIS IS FOR PUBLIC VERSION **/
			if (obj.jctUserId) {
				sessionStorage.setItem("jctUserId", obj.jctUserId);
			}
			if(statusCode == 200){ //If user activated and login is success
				sessionStorage.setItem("name", obj.firstName);
				sessionStorage.setItem("jctEmail", myEmail);
				sessionStorage.setItem("pageSequence", obj.pageSequence);
				sessionStorage.setItem("jobTitle", obj.jobTitle);
				sessionStorage.setItem("isLogout", "N");
				sessionStorage.setItem("menuItems", obj.menu);
				sessionStorage.setItem("userRole", obj.roleId);
				sessionStorage.setItem("authMsg", obj.unauthMsg);
				sessionStorage.setItem("interfaceTitleStr", obj.interfaceTitleStr);
				sessionStorage.setItem("customerId", obj.customerId);
				sessionStorage.setItem("profileId", obj.profileId);
				if (obj.roleId == 2) {  // Admin
					window.location = "view/userProfile.jsp";
				} else {				// Facilitaror
					if(obj.activeYn == 2) {
						window.location = "view/createUserFacilitator.jsp";
					} else {
						window.location = "view/activateAccount.jsp";
					}		
				}				
			} else if(statusCode == 201) {//If user is still to be activated
				sessionStorage.setItem("statusMsg", obj.statusDesc);
				sessionStorage.setItem("jctEmail", myEmail);
				window.location.replace(url);
			} else if(statusCode == 404) {//User not found
				//alertify.alert("Unable to connect to the server.");
				alertify.alert(obj.statusDesc);
				$(".loader_bg").fadeOut();
			    $(".loader").hide();
			} else if(statusCode == 990) {//User not found
				//alertify.alert("Unable to connect to the server.");
				alertify.alert(obj.statusDesc);
				$(".loader_bg").fadeOut();
			    $(".loader").hide();
			} else if(statusCode == 10 || statusCode == 11) { // Facilitator Created
				sessionStorage.setItem("jctEmail", myEmail);
				window.location = "/admin/view/activateFacilitator.jsp?emailId="+myEmail;
				$(".loader_bg").fadeOut();
			    $(".loader").hide();
			} else {	// Check with other status codes... Hence it is copy paste
				alertify.alert(obj.statusDesc);
				$(".loader_bg").fadeOut();
			    $(".loader").hide();
			}
		});
    }
});
    
    function getDivCount(jSon){
    	
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
		Contact.configure("/admin/auth/checkUser", "emailId");
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
			alertify.alert("Unable to connect to the server.");
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



	$("#click").click(function(){
		$('#fgtMyPwd').removeAttr('data-dismiss');	
		$("#forgotEmail").val('');	
	});

});

/**
 * function validates the email id
 * @param null
 */
function validateEmail(email){
	var emailReg = /^\w+([-+.']\w+)*@\w+([-.]\w+)*\.\w+([-.]\w+)*$/;
	if (!emailReg.test(email)){
		alertify.alert('Please enter valid email id.');
		return false;
	} else {
		return true;
	}
}

/**
 * function validates the login credentials
 * @param emailId
 * @param password
 * @returns {Boolean}
 */
function validateLoginCredential(emailId, password){
	//validation - username
	//var illegalChars = /^(?=.*\d)(?=.*[a-z])(?=.*[A-Z])(?=.*[^a-zA-Z0-9])(?!.*\s).{8,16}$/;
	var emailReg = /^\w+([-+.']\w+)*@\w+([-.]\w+)*\.\w+([-.]\w+)*$/;
	if (!emailReg.test(emailId)){
		alertify.alert('Please enter valid email id.');
		return false;
	} else if (password == "") {
    	alertify.alert('Please enter password.');
        return false;
    } else if ((password.length < 8) || (password.length > 16)) {
    	alertify.alert('Length of password should be between 8 - 16.');
        return false;
    }
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
function validateSignupForm(firstName, lastName, location, emailId, dateOfBirth, tenure, functionGrp, jobLevel){
	//var illegalChars = /\W/; // allow letters, numbers, and underscores
	var emailReg = /^\w+([-+.']\w+)*@\w+([-.]\w+)*\.\w+([-.]\w+)*$/;
	if (firstName == "") {
		alertify.alert('Please enter your first name.');
        return false;
    } else if ((firstName.length < 3)) {
    	alertify.alert('First Name should contain minimum three character.');
        return false;
    }
	if (lastName == "") {
		alertify.alert('Please enter your last name.');
        return false;
    } else if ((lastName.length < 3)) {
    	alertify.alert('Last Name should contain minimum three character.');
        return false;
    }
	if (location == "") {
		alertify.alert('Please enter Region.');
        return false;
    } 
	
	if (!emailReg.test(emailId)){
		alertify.alert('Please enter valid email id.');
		return false;
	}
	if (dateOfBirth == "") {
		alertify.alert('Please enter your date of birth.');
        return false;
    }
	var today = new Date();
	var selectedDate = new Date(document.getElementById('dpd1').value);
	var age = today.getFullYear() - selectedDate.getFullYear();
	var m = today.getMonth() - selectedDate.getMonth();
	if (m < 0 || (m === 0 && today.getDate() < selectedDate.getDate())) {
        age--;
    }
	if (selectedDate.valueOf() > today.valueOf()){
        document.getElementById('dpd1').value="";
        alertify.alert('Date of Birth cannot be a future date');
        return false;
    } else if(age < 18){
    	document.getElementById('dpd1').value="";
    	alertify.alert('Age should be atleast 18');
    	return false;
    }
	/*if(imgByte == null){
		alertify.alert('Please select your profile picture.');
        return false;
	}*/
	/*if (answer == "") {
		alertify.alert('Please enter the captcha answer.');
        return false;
    }*/
	if (tenure == "") {
		alertify.alert('Please select your Tenure.');
        return false;
    }
	if (functionGrp == "") {
		alertify.alert('Please select your Function Group.');
        return false;
    }
	if (jobLevel == "") {
		alertify.alert('Please select your Job Level.');
        return false;
    }
	return true;
}


/**
 * function encode the login password which 
 * is transmitted over the network.
 * @param text
 * @returns {String}
 */
function encode(text) { 
// /   var ref = "1&(9%2^*37@86^54&&^%&^0CKj=mo(%np(^qs$YTuQ%^rxyd=ahg)&ei=lBfWzC@*KJM*&%ioON=PQajpSTUVR~XYDA*&^*HGE=ILBFWZ";
    var ref = "P!QR#S$T%U&VW(X)Y*Z+[,-].^/_0`1a2b3c4d5e6f7g8h9i:j;k<l=m>n?o@pAqBrCsDtEuFvGwHxIyJzK{L|M}N~O";
	var result="";
	for (var count=0; count<text.length; count++) {
	    var char=text.substring (count, count+1); 
	    var num=ref.indexOf (char);
	    var encodeChar=ref.substring(num+1, num+2);
	    result += encodeChar;
	}
	//console.log(result);
	return result;
}


/**
 * function to call login against Enter key.
 * @param e
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
 * @param e
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
 * @param event
 * @returns {Boolean}
 */
function alphaOnly(event) {	
	 var c = event.which || event.keyCode;
	  if((c > 32 && c < 65) || (c > 90 && c < 97) ||
	     (c > 122 && c !== 127 ))
	    return false;
	  
}
/**
 * Function handles the forgot password functionality
 */
$("#fgtMyPwd").click(function(){

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
		alertify.alert("<img src='img/server-na-icon.png'><br /><p>Unable to connect to the server.</p>");
	});
	
	forgotPwd.bind("ajaxSuccess", function(record, xhr, settings, error) {
		var jsonStr = JSON.stringify(xhr);
		var obj = jQuery.parseJSON(jsonStr);
		var statusCode = obj.statusCode;
		if((statusCode == 200)){ //updated properly
			//store the username and the status message in session storage
			sessionStorage.setItem("jctEmail", obj.email);
			sessionStorage.setItem("statusMsg", obj.statusDesc);
			window.location = "view/resetPassword.jsp";
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

	
	
});


function forgotPwd() {
	$('#fgtMyPwd').removeAttr('data-dismiss');	
	var toSend = $('#forgotEmail').val();
	var emailReg = /^\w+([-+.']\w+)*@\w+([-.]\w+)*\.\w+([-.]\w+)*$/;
	if (!emailReg.test(toSend)){
		alertify.alert("<img src='img/alert-icon.png'><br /><p>Please enter valid E-mail ID.</p>");
		return false;
	} else {
		//$('#fgtMyPwd').attr('data-dismiss',"modal");	//	close the modal	
		$(".loader_bg").fadeIn();
		$(".loader").fadeIn();
		//Create a model
		var forgotPwd = Spine.Model.sub();
		forgotPwd.configure("/admin/manageuser/forgotPassword", "emailId");
		forgotPwd.extend( Spine.Model.Ajax );
		//Populate the model
		var populateModel = new forgotPwd({  
			emailId: toSend
		});
		populateModel.save(); //POST
		
		forgotPwd.bind("ajaxError", function(record, xhr, settings, error) {
			alertify.alert("<img src='img/server-na-icon.png'><br /><p>Unable to connect to the server.</p>");
		});
		
		forgotPwd.bind("ajaxSuccess", function(record, xhr, settings, error) {
			var jsonStr = JSON.stringify(xhr);
			var obj = jQuery.parseJSON(jsonStr);
			var statusCode = obj.statusCode;
			if((statusCode == 200)){ // success
				alertify.alert("<img src='img/right.gif'><br /><p>"+obj.statusDesc+"</p>");
				$('#fgtMyPwd').attr('data-dismiss',"modal");
				$("#forgotEmail").val('');
				$(".loader_bg").fadeOut();
			    $(".loader").hide();	
			} else if (statusCode == 203) {
				$('#fgtMyPwd').attr('data-dismiss',"modal");
				$("#forgotEmail").val('');				
				alertify.alert("<img src='img/alert-icon.png'><br /><p>Password reset. But failed to send mail</p>");
				$(".loader_bg").fadeOut();
			    $(".loader").hide();
			} else if (statusCode == 404) {
				alertify.alert("<img src='img/failed-icon.png'><br /><p>Facilitator doesn't exist.</p>");
				$(".loader_bg").fadeOut();
			    $(".loader").hide();
			} else {
					alertify.alert("<img src='img/failed-icon.png'><br /><p>Unnable to reset your password.<br /> Try again later.</p>");
				$(".loader_bg").fadeOut();
			    $(".loader").hide();
				return false;
			}
		});
	}
}