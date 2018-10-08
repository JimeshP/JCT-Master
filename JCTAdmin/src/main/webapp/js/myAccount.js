/**
 * Function calls when the page is loaded 
 * to populate user email id, 
 * password, first name and last name
 */
$(document).ready(function() {
	populateUserData();
});

/**
 * Function to populate the user details
 * i.e, email id, password, first name and last name
 * @param null
 */
function populateUserData() {
	var emailId = sessionStorage.getItem("jctEmail");
	var userProf = Spine.Model.sub();
	userProf.configure("/admin/myAccount/populateExistingUserData", "none");
	userProf.extend( Spine.Model.Ajax );
	//Populate the model with data to transfer
	var modelPopulator = new userProf({  
		none: ""
	});
	modelPopulator.save(); //POST
	userProf.bind("ajaxError", function(record, xhr, settings, error) {
		alertify.alert("Unable to connect to the server.");
	});
	userProf.bind("ajaxSuccess", function(record, xhr, settings, error) {
		var jsonStr = JSON.stringify(xhr);
		var obj = jQuery.parseJSON(jsonStr);
		var statusCode = obj.statusCode;
		if(statusCode == 200) {			
			showUserData(obj.existingUserDataList);
		}  else if(statusCode == 500) {
			//Show error message
		}
	});
}


/**
 * Function to show the user details
 * @param userDataList
 */
function showUserData(userDataList) {	
	var splitVal = userDataList.split('~');	
	var firstName = "";
	var lastName = "";
	if (null == splitVal[0] || splitVal[0] == ""){
		firstName = "N/A";
	} else {
		firstName = splitVal[0];
	}	
	if (null == splitVal[1] || splitVal[1] == ""){
		lastName = "N/A";
	} else {
		lastName = splitVal[1];
	}	
	var passwordLength = (splitVal[2].length)/2;
	var stringBuilder = "";
	for(var i=0; i<passwordLength; i++){
		stringBuilder = stringBuilder + "*";
	}
	var password=splitVal[2];
	document.getElementById("firstNameInptId").innerHTML = firstName+"&nbsp;&nbsp;<a href='#' onClick='editFirstName(\""+splitVal[4]+"\")' class= 'hyperlink_text' >edit</a>";
	document.getElementById("lastNameInptId").innerHTML = lastName+"&nbsp;&nbsp;<a href='#' onClick='editLastName(\""+splitVal[4]+"\")' class= 'hyperlink_text' >edit</a>";
	document.getElementById("chgPwdInptId").innerHTML = stringBuilder+"&nbsp;&nbsp;<a href='#' onClick='editPassword(\""+splitVal[4]+"\", \""+splitVal[2]+"\")' class= 'hyperlink_text' >change password</a>";
	document.getElementById("emailIdInptId").innerHTML = splitVal[3]+"&nbsp;&nbsp;<a href='#' onClick='editEmailId(\""+splitVal[4]+"\")' class= 'hyperlink_text' >edit</a>";
}

/**
 * Function to edit the user first name
 * @param tablePkId
 */
function editFirstName(tablePkId){
	var valueChgDiv = document.getElementById("firstNameChgMainDiv");
	valueChgDiv.setAttribute("style", "display:block");	
	var mainDiv = document.getElementById("divGeneralId");
	mainDiv.setAttribute("style", "display:block");			
	document.getElementById("firstNameChgInptId").value="";
	document.getElementById("firstNameChgOkBtnId").setAttribute("onclick", "updateFirstName("+tablePkId+")");
	document.getElementById("firstNameChgCancelBtnId").setAttribute("onclick", "CancelFirstName()");
}


/**
 * Function to update the user first name
 * @param tablePkId
 */
function updateFirstName(tablePkId){
	var valueDesc = document.getElementById("firstNameChgInptId").value;
	if(validateInput(valueDesc)) {
		var newUserProf = Spine.Model.sub();
		newUserProf.configure("/admin/myAccount/updateFirstName", "valueDesc", "tablePkId");
		newUserProf.extend( Spine.Model.Ajax );
		//Populate the model with data to transfer
		var modelPopulator = new newUserProf({  
			valueDesc: valueDesc,
			tablePkId: tablePkId
		});
		modelPopulator.save(); //POST
		newUserProf.bind("ajaxError", function(record, xhr, settings, error) {
			alertify.alert("Unable to connect to the server.");
		});
		newUserProf.bind("ajaxSuccess", function(record, xhr, settings, error) {
			var jsonStr = JSON.stringify(xhr);
			var obj = jQuery.parseJSON(jsonStr);
			var statusCode = obj.statusCode;
			if(statusCode == 200) {		
				populateUserData();
				alertify.alert("First name has been updated successfully.");				
				document.getElementById("firstNameInptId").innerHTML = obj.userFirstName+"&nbsp;&nbsp;<a href='#' onClick='editFirstName("+obj.primaryKey+")' class= 'hyperlink_text' >edit</a>";
				var valueChgDiv = document.getElementById("firstNameChgMainDiv");
				valueChgDiv.setAttribute("style", "display:none");	
				var mainDiv = document.getElementById("divGeneralId");
				mainDiv.setAttribute("style", "display:none");	
			} else if(statusCode == 500) {
				// Error Message
			}
		});
	}
}

/**
 * Function to cancel the first name div
 * while click on cancel button
 * @param null
 */
function CancelFirstName() {
	var valueChgDiv = document.getElementById("firstNameChgMainDiv");
	valueChgDiv.setAttribute("style", "display:none");	
	var mainDiv = document.getElementById("divGeneralId");
	mainDiv.setAttribute("style", "display:none");	
}



/**
 * Function to edit the user last name
 * @param tablePkId
 */
function editLastName(tablePkId){
	var valueChgDiv = document.getElementById("lastNameChgMainDiv");
	valueChgDiv.setAttribute("style", "display:block");
	document.getElementById('lastNameChgInptId').value = "";
	var mainDiv = document.getElementById("divGeneralId");
	mainDiv.setAttribute("style", "display:block");	
	document.getElementById("lastNameChgInptId").value = "";	
	document.getElementById("lastNameChgOkBtnId").setAttribute("onclick", "updateLastName("+tablePkId+")");
	document.getElementById("lastNameChgCancelBtnId").setAttribute("onclick", "CancelLastName()");
}


/**
 * Function to update the user last name
 * @param tablePkId
 */
function updateLastName(tablePkId){
	var valueDesc = document.getElementById("lastNameChgInptId").value;
	if(validateInput(valueDesc)) {
		var newUserProf = Spine.Model.sub();
		newUserProf.configure("/admin/myAccount/updateLastName", "valueDesc", "tablePkId");
		newUserProf.extend( Spine.Model.Ajax );
		//Populate the model with data to transfer
		var modelPopulator = new newUserProf({  
			valueDesc: valueDesc,
			tablePkId: tablePkId
		});
		modelPopulator.save(); //POST
		newUserProf.bind("ajaxError", function(record, xhr, settings, error) {
			alertify.alert("Unable to connect to the server.");
		});
		newUserProf.bind("ajaxSuccess", function(record, xhr, settings, error) {
			var jsonStr = JSON.stringify(xhr);
			var obj = jQuery.parseJSON(jsonStr);
			var statusCode = obj.statusCode;
			if(statusCode == 200) {
				//alertify.set({ delay: 2700 });
				populateUserData();
				alertify.alert("Last name has been updated successfully.");
				document.getElementById("lastNameInptId").innerHTML = obj.userLastName+"&nbsp;&nbsp;<a href='#' onClick='editLastName("+obj.primaryKey+")' class= 'hyperlink_text' >edit</a>";
				var valueChgDiv = document.getElementById("lastNameChgMainDiv");
				valueChgDiv.setAttribute("style", "display:none");
				var mainDiv = document.getElementById("divGeneralId");
				mainDiv.setAttribute("style", "display:none");	
				
			} else if(statusCode == 500) {
				// Error Message
			}
		});
	}
}


/**
 * Function to cancel the last name div
 * while click on cancel button
 * @param null
 */
function CancelLastName() {
	var valueChgDiv = document.getElementById("lastNameChgMainDiv");
	valueChgDiv.setAttribute("style", "display:none");	
	var mainDiv = document.getElementById("divGeneralId");
	mainDiv.setAttribute("style", "display:none");	
}


/**
 * Function to edit the user last name
 * @param tablePkId
 */
function editEmailId(tablePkId){
	var valueChgDiv = document.getElementById("emailIdChgMainDiv");
	valueChgDiv.setAttribute("style", "display:block");
	var mainDiv = document.getElementById("divGeneralId");
	mainDiv.setAttribute("style", "display:block");	
	document.getElementById("emailIdChgInptId").value = "";
	document.getElementById("emailIdChgOkBtnId").setAttribute("onclick", "updateEmailId("+tablePkId+")");
	document.getElementById("emailIdChgCancelBtnId").setAttribute("onclick", "CancelEmailId()");
}


/**
 * Function to update the user email id
 * @param tablePkId
 */
function updateEmailId(tablePkId){
	var valueDesc = document.getElementById("emailIdChgInptId").value;
	if(validateEmailId(valueDesc)) {
		var newUserProf = Spine.Model.sub();
		newUserProf.configure("/admin/myAccount/updateEmailId", "valueDesc", "tablePkId");
		newUserProf.extend( Spine.Model.Ajax );
		//Populate the model with data to transfer
		var modelPopulator = new newUserProf({  
			valueDesc: valueDesc,
			tablePkId: tablePkId
		});
		modelPopulator.save(); //POST
		newUserProf.bind("ajaxError", function(record, xhr, settings, error) {
			alertify.alert("Unable to connect to the server.");
		});
		newUserProf.bind("ajaxSuccess", function(record, xhr, settings, error) {
			var jsonStr = JSON.stringify(xhr);
			var obj = jQuery.parseJSON(jsonStr);
			var statusCode = obj.statusCode;
			if(statusCode == 200) {
				populateUserData();
				alertify.alert("Email Id has been updated successfully.");
				document.getElementById("emailIdInptId").innerHTML = obj.userEmailId+"&nbsp;&nbsp;<a href='#' onClick='editEmailId("+obj.primaryKey+")' class= 'hyperlink_text' >edit</a>";
				var valueChgDiv = document.getElementById("emailIdChgMainDiv");
				valueChgDiv.setAttribute("style", "display:none");	
				var mainDiv = document.getElementById("divGeneralId");
				mainDiv.setAttribute("style", "display:none");	
				
			} else if(statusCode == 500) {
				// Error Message
			}
		});
	}
}


/**
 * Function to cancel the email id div
 * while click on cancel button
 * @param null
 */
function CancelEmailId() {
	var valueChgDiv = document.getElementById("emailIdChgMainDiv");
	valueChgDiv.setAttribute("style", "display:none");	
	var mainDiv = document.getElementById("divGeneralId");
	mainDiv.setAttribute("style", "display:none");	
}


/**
 * Function to edit password
 * @param tablePkId
 * @param prevValue
 */
function editPassword(tablePkId, prevValue){
	var valueChgDiv = document.getElementById("pwdChgMainDiv");
	valueChgDiv.setAttribute("style", "display:block");	
	var mainDiv = document.getElementById("divGeneralId");
	mainDiv.setAttribute("style", "display:block");	
	document.getElementById('oldPwdChgInptId').value = "";
	document.getElementById("newPwdChgInptId").value = "";	
	document.getElementById("confirmPwdChgInptId").value = "";	
	document.getElementById("pwdChgOkBtnId").setAttribute("onclick", "updatePassword("+tablePkId+", '"+prevValue+"')");
	document.getElementById("pwdChgCancelBtnId").setAttribute("onclick", "CancelPassword()");
}


/**
 * Function to update the password
 * @param tablePkId
 * @param prevPassword
 */
function updatePassword(tablePkId, prevPassword){
	var oldPassword = document.getElementById("oldPwdChgInptId").value;
	var newPassword = document.getElementById("newPwdChgInptId").value;
	var confirmPassword = document.getElementById("confirmPwdChgInptId").value;
	if(validatePasswordEntries(oldPassword, newPassword, confirmPassword, prevPassword) == true) {
		var newUserProf = Spine.Model.sub();
		newUserProf.configure("/admin/myAccount/updatePassword", "valueDesc", "tablePkId");
		newUserProf.extend( Spine.Model.Ajax );
		//Populate the model with data to transfer
		var modelPopulator = new newUserProf({  
			valueDesc: encode(confirmPassword+confirmPassword),
			tablePkId: tablePkId
		});
		modelPopulator.save(); //POST
		newUserProf.bind("ajaxError", function(record, xhr, settings, error) {
			alertify.alert("Unable to connect to the server.");
		});
		newUserProf.bind("ajaxSuccess", function(record, xhr, settings, error) {
			var jsonStr = JSON.stringify(xhr);
			var obj = jQuery.parseJSON(jsonStr);
			var statusCode = obj.statusCode;
			if(statusCode == 200) {
				populateUserData();
				alertify.alert("Password has been updated successfully.");
				var passwordLength = ((obj.userPassword).length)/2;
				var stringBuilder = "";
				for(var i=0; i<passwordLength; i++){
					stringBuilder = stringBuilder + "*";
				}
				
				document.getElementById("chgPwdInptId").innerHTML = stringBuilder+"&nbsp;&nbsp;<a href='#' onClick='editPassword("+obj.primaryKey+", "+obj.userPassword+")' class= 'hyperlink_text' >edit</a>";
				var valueChgDiv = document.getElementById("pwdChgMainDiv");
				valueChgDiv.setAttribute("style", "display:none");	
				var mainDiv = document.getElementById("divGeneralId");
				mainDiv.setAttribute("style", "display:none");	
			} else if(statusCode == 500) {
				// Error Message
			}
		});
	}
}

/**
 * Function to cancel the password div
 * while click on cancel button
 * @param null
 */
function CancelPassword() {
	var valueChgDiv = document.getElementById("pwdChgMainDiv");
	valueChgDiv.setAttribute("style", "display:none");	
	var mainDiv = document.getElementById("divGeneralId");
	mainDiv.setAttribute("style", "display:none");	
}


/**
 * Function to validate the
 * firts name , last name  text
 * @param value
 */
function validateInput(value){
	var re=/^\s*$/;
	if (re.test(value)){
		alertify.alert("Please enter value.");
		return false;
	}	
	return true;
}


/**
 * function validates email id
 * @param emailId
 */
function validateEmailId(emailId){
	var emailReg = /^\w+([-+.']\w+)*@\w+([-.]\w+)*\.\w+([-.]\w+)*$/;
	if (!emailReg.test(emailId)){
		alertify.alert('Please enter valid email id.');
		return false;
	} 
	return true;
}


/**
 * function validates the reset password entries
 * which are provided by the user.
 * @param mailedPassword
 * @param newPassword
 * @param confirmPassword
 * @param prevPassword
 * @returns {Boolean}
 */
function validatePasswordEntries(oldPassword, newPassword, confirmPassword, prevPassword){
	//var illegalChars = /\W/; // allow letters, numbers, and underscores
	var illegalChars = /^(?=.*\d)(?=.*[a-z])(?=.*[A-Z])(?=.*[^a-zA-Z0-9])(?!.*\s).{8,16}$/;
	var newValue = encode(oldPassword+oldPassword);
	if (oldPassword == "") {
		alertify.alert('Please enter your old password.');
        return false;
    } else if ((oldPassword.length < 8) || (oldPassword.length > 16)) {
    	alertify.alert('Length of password should be between 8 - 16.');
        return false;
    } else if (newValue != prevPassword){
    	alertify.alert('Old Password is different. Please correct it.');
        return false;
    } /*else if (illegalChars.test(oldPassword)) {
    	alertify.alert('Password can contain only letters, numbers, and underscores.');
        return false;
    }*/ else if (newPassword == "") {
    	alertify.alert('Please enter the New Password.');
        return false;
    } else if ((newPassword.length < 8) || (newPassword.length > 16)) {
    	alertify.alert('Length of new password should be between 8 - 16.');
        return false;
    } /*else if (illegalChars.test(newPassword)) {
    	alertify.alert('New Password can contain only letters, numbers, and underscores.');
        return false;
    }*/ else if (confirmPassword == "") {	
    	alertify.alert('Please enter the Confirm Password.');
        return false;
    } else if ((confirmPassword.length < 8) || (confirmPassword.length > 16)) {
    	alertify.alert('Length of confirm password should be between 8 - 16.');
        return false;
    } /*else if (illegalChars.test(confirmPassword)) {
    	alertify.alert('Confirm Password can contain only letters, numbers, and underscores.');
        return false;
    }*/ else if (newPassword != confirmPassword){
    	alertify.alert('New Password is different from Confirm Password. Please correct it.');
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
   // var ref = "1&(9%2^*37@86^54&&^%&^0CKj=mo(%np(^qs$YTuQ%^rxyd=ahg)&ei=lBfWzC@*KJM*&%ioON=PQajpSTUVR~XYDA*&^*HGE=ILBFWZ";
	//var ref = "P!QR#S$T%U&V'W(X)Y*Z+[,-].^/_0`1a2b3c4d5e6f7g8h9i:j;k<l=m>n?o@pAqBrCsDtEuFvGwHxIyJzK{L|M}N~O";
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

