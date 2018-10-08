$(document).ready( function() {
	storeParamValue();	
});

function storeParamValue() {
	sessionStorage.clear();
	var parameter = window.location.search.replace( "?", "" ); // will return the GET parameter 
	var values = parameter.split("=");
	sessionStorage.setItem("jctEmail", values[1].split("&")[0]);
	
	// decide div to be populated
	$(".loader_bg").fadeIn();
    $(".loader").fadeIn();
  
	var model = Spine.Model.sub();
	model.configure("/admin/authAdmin/decideDiv", "email");
	model.extend( Spine.Model.Ajax );
	
	var populator = new model({  
		email: sessionStorage.getItem("jctEmail")
	});
	//POST
	populator.save();
	model.bind("ajaxError", function(record, xhr, settings, error) {
		alertify.alert("Unable to connect to the server.");
		$(".loader_bg").fadeOut();
	    $(".loader").hide();
	});
	model.bind("ajaxSuccess", function(record, xhr, settings, error) {
		var jsonStr = JSON.stringify(xhr);
		var obj = jQuery.parseJSON(jsonStr);
		var statusCode = obj.statusCode;
		if (statusCode == 200) {
			document.getElementById("passwordResetForm").style.display = "block";
			document.getElementById("userValidatedSection").style.display = "none";
		} else {
			document.getElementById("userValidatedSection").style.display = "block";
			document.getElementById("passwordResetForm").style.display = "none";
		}
		$(".loader_bg").fadeOut();
	    $(".loader").hide();
	});
	
}

function activateAccount() {
	var selectedUserName = sessionStorage.getItem("jctEmail");
	var mailedPwd = $('#mailedPassword').val();
	var newPassword = $('#inputNPwd').val();
	var confirmPassword = $('#inputCPwd').val();
	if(validatePasswordEntries(mailedPwd, newPassword, confirmPassword) == true) {
		$(".loader_bg").fadeIn();
	    $(".loader").fadeIn();
	    
		//Create a model
		var passwordResetModel = Spine.Model.sub();
		passwordResetModel.configure("/admin/authAdmin/resetFacilitatorInitialPwd", "forUser", "initialPassword", "userPassword", "mailedPwd");
		passwordResetModel.extend( Spine.Model.Ajax );
		
		//Populate the model
		var modelPopulator = new passwordResetModel({  
			forUser: selectedUserName,
			initialPassword: encode(mailedPwd+mailedPwd),
			userPassword: encode(newPassword+newPassword),
			mailedPwd : mailedPwd
		});
		//POST
		modelPopulator.save();
		
		passwordResetModel.bind("ajaxSuccess", function(record, xhr, settings, error) {
			var jsonStr = JSON.stringify(xhr);
			var obj = jQuery.parseJSON(jsonStr);
			var statusCode = obj.statusCode;
			if(statusCode == 200){ //success
				sessionStorage.setItem("statusMsg", obj.statusDesc);
				$(".loader_bg").fadeOut();
			    $(".loader").hide();
			    if (obj.jctUserId) {
					sessionStorage.setItem("jctUserId", obj.jctUserId);
				}
			   // console.log(obj);
				sessionStorage.setItem("name", obj.firstName);
				sessionStorage.setItem("jctEmail", selectedUserName);
				sessionStorage.setItem("pageSequence", obj.pageSequence);
				sessionStorage.setItem("jobTitle", obj.jobTitle);
				sessionStorage.setItem("isLogout", "N");
				sessionStorage.setItem("menuItems", obj.menu);
				sessionStorage.setItem("userRole", obj.roleId);
				sessionStorage.setItem("authMsg", obj.unauthMsg);
				sessionStorage.setItem("interfaceTitleStr", obj.interfaceTitleStr);
				sessionStorage.setItem("customerId", obj.customerId);
				sessionStorage.setItem("profileId", obj.profileId);
				if(obj.statusCodeInt == 10) {
					window.location = "/admin/view/activateAccount.jsp";
				} else {
					window.location = "/admin/view/createUserFacilitator.jsp";
				}			
			} else {//User not found
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
}

function validatePasswordEntries(mailedPassword, newPassword, confirmPassword){
	
	var illegalChars = /^(?=.*\d)(?=.*[a-z])(?=.*[A-Z])(?=.*[^a-zA-Z0-9])(?!.*\s).{8,16}$/;
	if (mailedPassword == "") {
		alertify.alert('Please enter the Password which we mailed you.');
        return false;
    }   else if (newPassword == "") {
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
    } else if (!illegalChars.test(confirmPassword)) { //JIRA Defect Id: VMJCT-227
    	alertify.alert('Confirm Password should contain atleast one lowercase letter, one uppercase letter, one numeric digit, and one special character.');
        return false;
    } else if (newPassword != confirmPassword){
    	alertify.alert('New Password is different from Confirm Password. Please correct it.');
        return false;
    } else if (mailedPassword == newPassword) {
    	alertify.alert('Your password should be different from the one we sent you.');
        return false;
    }
	return true;
}
function encode(text) { 
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