if(sessionStorage.getItem("isNextReset")) {
	document.getElementById('passwordResetForm').style.display = "none";
	document.getElementById('loginForm').style.display = "block";
}

/**
 * function to reset the password.
 */
$("#resetMyPassword").click(function () {
		var selectedUserName = sessionStorage.getItem("jctEmail");
    	var mailedPwd = $('#mailedPassword').val();
    	var newPassword = $('#inputNPwd').val();
    	var confirmPassword = $('#inputCPwd').val();
    	if(validatePasswordEntries(mailedPwd, newPassword, confirmPassword) == true) {
    		$(".loader_bg").fadeIn();
		    $(".loader").fadeIn();		    
    		//Create a model
    		var passwordResetModel = Spine.Model.sub();
    		passwordResetModel.configure("/user/auth/resetPassword", "forUser", "initialPassword", "userPassword", "status", "clickStatus", "dist", "mailedPwd");
    		passwordResetModel.extend( Spine.Model.Ajax );   		
    		//Populate the model
    		var modelPopulator = new passwordResetModel({  
    			forUser: selectedUserName,
    			initialPassword: encode(mailedPwd+mailedPwd),
    			userPassword: encode(newPassword+newPassword),
    			status: "3",
    			clickStatus: "reset",
    			dist: 2,
    			mailedPwd: mailedPwd
    		});
    		//POST
    		modelPopulator.save();
    		
    		passwordResetModel.bind("ajaxSuccess", function(record, xhr, settings, error) {
    			var jsonStr = JSON.stringify(xhr);
    			var obj = jQuery.parseJSON(jsonStr);
    			var statusCode = obj.statusCode;
    			if(statusCode == 200 || statusCode == 128){ //success
    				sessionStorage.setItem("statusMsg", obj.statusDesc);
    				document.getElementById('resetPwdId').innerHTML = "Please Login";
    				document.getElementById('passwordResetForm').style.display = "none";
    				document.getElementById('loginForm').style.display = "block";
    				//sessionStorage.setItem("isReset","Y");
    				sessionStorage.setItem("isReset","N");
    				$(".loader_bg").fadeOut();
				    $(".loader").hide();
    				showDescOnLoad();
    			} else if(statusCode == 404){
    				sessionStorage.setItem("statusMsg", obj.statusDesc);
    				$(".loader_bg").fadeOut();
				    $(".loader").hide();
    				showDescOnLoad();
    			} else if (statusCode = 874) { //Password reset is not complete
    				alertify.alert(obj.statusDesc);
    				$(".loader_bg").fadeOut();
				    $(".loader").hide();
    			}
    		});
    		
    		passwordResetModel.bind("ajaxError", function(record, xhr, settings, error) {
    			//var jsonStr = JSON.stringify(xhr);
    			alertify.alert("Unable to connect to the server.");
    			$(".loader_bg").fadeOut();
			    $(".loader").hide();
    		});
    		
    	}
});
/**
 * function validates the reset password entries
 * which are provided by the user.
 * @param mailedPassword
 * @param newPassword
 * @param confirmPassword
 * @returns {Boolean}
 */
function validatePasswordEntries(mailedPassword, newPassword, confirmPassword){
	var illegalChars = /^(?=.*\d)(?=.*[a-z])(?=.*[A-Z])(?=.*[^a-zA-Z0-9])(?!.*\s).{8,16}$/;
	if (mailedPassword == "") {
		alertify.alert('Please enter the Password which we mailed you.');
        return false;
    } else if ((mailedPassword.length < 5) || (mailedPassword.length > 16)) {
    	alertify.alert('Password should be between 5 - 16.');
        return false;
    } else if (newPassword == "") {
    	alertify.alert('Please enter the New Password.');
        return false;
    } else if ((newPassword.length < 8) || (newPassword.length > 16)) {
    	alertify.alert('New Password should be between 8 - 16 characters.');
        return false;
    } else if (!illegalChars.test(newPassword)) {
    	alertify.alert('New Password should contain atleast one lowercase letter, one uppercase letter, one numeric digit, and one special character.');
        return false;
    } else if (confirmPassword == "") {	//JIRA Defect Id: VMJCT-31
    	alertify.alert('Please enter the Confirm Password.');
        return false;
    } else if ((confirmPassword.length < 8) || (confirmPassword.length > 16)) {
    	alertify.alert('Confirm Password should be between 8 - 16 characters.');
        return false;
    } /*else if (!illegalChars.test(confirmPassword)) { //JIRA Defect Id: VMJCT-227
    	alertify.alert('Confirm Password should contain atleast one lowercase letter, one uppercase letter, one numeric digit, and one special character.');
        return false;
    }*/ 
    else if (newPassword != confirmPassword){
    	alertify.alert('New Password is different from Confirm Password. Please correct it.');
        return false;
    } else if (mailedPassword == newPassword) {
    	alertify.alert('Your password should be different from the one we sent you.');
        return false;
    }
	return true;
	
}
/**
 * function encodes the new password which is 
 * then sent over the network
 * @param text
 * @returns {String}
 */
function encode(text) { 
    //var ref = "1&(9%2^*37@86^54&&^%&^0CKj=mo(%np(^qs$YTuQ%^rxyd=ahg)&ei=lBfWzC@*KJM*&%ioON=PQajpSTUVR~XYDA*&^*HGE=ILBFWZ";
	//var ref = "Q%1wS~(RxGzPA$By&N!Oa_2#^CM>D)<bFWLE@3c4*Kd5J+IVe-6Hf7g8=h9iT0jU";
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
/**
 * function shows the status messages.
 */
function showDescOnLoad(){
	var output = document.getElementById("dispalyMessage");
	/*if(null != output){
		output.innerHTML = sessionStorage.getItem("statusMsg");
	}else{
		document.getElementById("dispalyMessageLogin").innerHTML = sessionStorage.getItem("statusMsg");
	}*/
}

/**
 * function to call forgot password against Enter key.
 */
function searchKeyEvent(e)
{
    // look for window.event in case event isn't passed in
    if (typeof e == 'undefined' && window.event) { e = window.event; }
    if (e.keyCode == 13)
    {
        document.getElementById('resetMyPassword').click();
    }
}