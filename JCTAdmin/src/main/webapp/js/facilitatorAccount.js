/**
 * Function calls when the page is loaded 
 * to populate user email id, 
 * password, first name and last name
 */
$(document).ready(function() {
	fetchUserDropdown();
	//fetchExistingFacilitator();
	
	fetchExistingUsersManageUser();
	
	populateUserData();
	document.getElementById("existing_data_div").style.display = "none";
	$('.tabs .nav-tabs a').on('click', function(e)  {
        var currentAttrValue = $(this).attr('href');
        if(currentAttrValue == '#tab2'){
        	//$("#existingChangeRoleList").show();  
        	//$("#existing_data_div").css("display", "block !important");
        	if (document.getElementById("existing_data_div")) {
    			document.getElementById("existing_data_div").style.display = "block";
    		}
        	//document.getElementById("existing_data_div").style.display = "block";
        } else {
        	if (document.getElementById("existing_data_div")) {
    			document.getElementById("existing_data_div").style.display = "none";
    		}
        }
        // Show/Hide Tabs
        $('.tabs ' + currentAttrValue).show().siblings().hide(); 
        // Change/remove current tab to active
        $(this).parent('li').addClass('active').siblings().removeClass('active'); 
        e.preventDefault();
    });
});

/**
 * Function to populate the user details
 * i.e, email id, password, first name and last name
 * @param null
 */
function populateUserData() {
	var emailId = sessionStorage.getItem("jctEmail");
	var userProf = Spine.Model.sub();
	userProf.configure("/admin/facilitatorAccount/populateExistingFacilitatorData", "facilitatorEmail");
	userProf.extend( Spine.Model.Ajax );
	//Populate the model with data to transfer
	var modelPopulator = new userProf({  
		facilitatorEmail: sessionStorage.getItem("jctEmail")
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
	var passwordLength = (splitVal[1].length)/2;
	var stringBuilder = "";
	for(var i=0; i<passwordLength; i++){
		stringBuilder = stringBuilder + "*";
	}
	
	document.getElementById("firstNameInptId").innerHTML = splitVal[0]+"&nbsp;&nbsp;<a href='#' onClick='editUserName(\""+splitVal[3]+"\")' class= 'hyperlink_text' >edit</a>";
	//document.getElementById("lastNameInptId").innerHTML = lastName+"&nbsp;&nbsp;<a href='#' onClick='editLastName(\""+splitVal[4]+"\")' class= 'hyperlink_text' >edit</a>";
	document.getElementById("chgPwdInptId").innerHTML = stringBuilder+"&nbsp;&nbsp;<a href='#' onClick='editPassword(\""+splitVal[3]+"\", \""+splitVal[1]+"\")' class= 'hyperlink_text' >change password</a>";
	document.getElementById("emailIdInptId").innerHTML = splitVal[2]+"&nbsp;&nbsp;<a href='#' onClick='editEmailId(\""+splitVal[3]+"\")' class= 'hyperlink_text' >edit</a>";
	//document.getElementById("facCustomerId").value = splitVal[4];
	document.getElementById("facCustomerId").innerHTML = splitVal[4];
}

/**
 * Function to edit the user first name
 * @param tablePkId
 */
function editUserName(tablePkId){
	var valueChgDiv = document.getElementById("firstNameChgMainDiv");
	valueChgDiv.setAttribute("style", "display:block");	
	var mainDiv = document.getElementById("divGeneralId");
	mainDiv.setAttribute("style", "display:block");			
	document.getElementById("firstNameChgInptId").value="";
	document.getElementById("firstNameChgOkBtnId").setAttribute("onclick", "updateUserName("+tablePkId+")");
	document.getElementById("firstNameChgCancelBtnId").setAttribute("onclick", "CancelUserName()");
}


/**
 * Function to update the user first name
 * @param tablePkId
 */
function updateUserName(tablePkId){
	var valueDesc = document.getElementById("firstNameChgInptId").value;
	if(validateEmailId(valueDesc)) {
		var newUserProf = Spine.Model.sub();
		newUserProf.configure("/admin/facilitatorAccount/updateUserNameFacilitator", "valueDesc", "tablePkId", "facilitatorEmail");
		newUserProf.extend( Spine.Model.Ajax );
		//Populate the model with data to transfer
		var modelPopulator = new newUserProf({  
			valueDesc: valueDesc,
			tablePkId: tablePkId,
			facilitatorEmail: sessionStorage.getItem("jctEmail")
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
				sessionStorage.setItem("jctEmail", valueDesc);
				populateUserData();
				//alertify.alert("User name has been updated successfully.");	
				alertify.alert(obj.statusDesc);	
				document.getElementById("firstNameInptId").innerHTML = obj.userFirstName+"&nbsp;&nbsp;<a href='#' onClick='editUserName("+obj.primaryKey+")' class= 'hyperlink_text' >edit</a>";
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
function CancelUserName() {
	var valueChgDiv = document.getElementById("firstNameChgMainDiv");
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
		newUserProf.configure("/admin/facilitatorAccount/updateEmailIdFacilitator", "valueDesc", "tablePkId", "facilitatorEmail");
		newUserProf.extend( Spine.Model.Ajax );
		//Populate the model with data to transfer
		var modelPopulator = new newUserProf({  
			valueDesc: valueDesc,
			tablePkId: tablePkId,
			facilitatorEmail: sessionStorage.getItem("jctEmail")
		});
		modelPopulator.save(); //POST
		newUserProf.bind("ajaxError", function(record, xhr, settings, error) {
			alertify.alert("Unable to connect to the server.");
		});
		newUserProf.bind("ajaxSuccess", function(record, xhr, settings, error) {
			var jsonStr = JSON.stringify(xhr);
			var obj = jQuery.parseJSON(jsonStr);
			var statusCode = obj.statusCode;		
				showUserData(obj.existingUserDataList);
				//alertify.alert("Email Id has been updated successfully.");
				alertify.alert(obj.statusDesc);
				if(statusCode == 200) {
					$('#alertify-ok').click(function() {	
			    	sessionStorage.clear();
			    	window.location = "../signup-page.jsp";
			    });								
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
		newUserProf.configure("/admin/facilitatorAccount/updatePasswordFacilitator", "valueDesc", "tablePkId", "facilitatorEmail");
		newUserProf.extend( Spine.Model.Ajax );
		//Populate the model with data to transfer
		var modelPopulator = new newUserProf({  
			valueDesc: encode(confirmPassword+confirmPassword),
			tablePkId: tablePkId,
			facilitatorEmail: sessionStorage.getItem("jctEmail")
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
				showUserData(obj.existingUserDataList);
				alertify.alert("Password has been updated successfully.");
				$('#alertify-ok').click(function() {	
			    	sessionStorage.clear();
			    	window.location = "../signup-page.jsp";
			    });
				/*populateUserData();
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
				mainDiv.setAttribute("style", "display:none");	*/
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
		alertify.alert("Please enter user name.");
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
    } else if(oldPassword == newPassword){
    	alertify.alert('Old and and new password cannot be same.');
        return false;
    } else if (!illegalChars.test(newPassword)) {
    	alertify.alert('Password should contain atleast one lowercase letter, one uppercase letter, one numeric digit, and one special character.');
        return false;
    }  else if (confirmPassword == "") {	
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
/**	************************************** CODE FOR #tab2 ****************************************** **/
/**
 * Function calls when the page is loaded to populate 
 * 1. The drop-down for User
 * 2. The Total Subscribed User and total Registered User
 */
/*$(document).ready(function() {
	fetchUserDropdown();
	fetchExistingFacilitator();
});*/
/**
 * Method populate existing user to populate drop down list
 */
function fetchUserDropdown() {
	var userGrp = Spine.Model.sub();
	userGrp.configure("/admin/manageuserFacilitator/populateUserDropDown", "facilitatorEmail", "customerId");
	userGrp.extend( Spine.Model.Ajax );
	//Populate the model with data to transfer
	var modelPopulator = new userGrp({  
		facilitatorEmail: sessionStorage.getItem("jctEmail"),
		customerId : sessionStorage.getItem("customerId")
	});
	modelPopulator.save(); //POST
	userGrp.bind("ajaxError", function(record, xhr, settings, error) {
		alertify.alert("Unable to connect to the server.");
	});
	userGrp.bind("ajaxSuccess", function(record, xhr, settings, error) {
		var jsonStr = JSON.stringify(xhr);
		var obj = jQuery.parseJSON(jsonStr);
		var statusCode = obj.statusCode;
		if(statusCode == 200) {
			populateDropDown(obj.userNameMap);
		}  else {
			//Show error message
			alertify.alert("Some thing went wrong.");
		}
	});
}


/**
 * Function to populate the drop down for user
 */
function populateDropDown(userNameMap) {
	document.getElementById("existingUserInputId").innerHTML = "";	
	var userNameSelect = document.getElementById("existingUserInputId");
	var defaultOpn = document.createElement("option");
	defaultOpn.text = "Select User";
	defaultOpn.value = "";
	defaultOpn.className = "form-control-general";
	userNameSelect.add(defaultOpn, null);	
	for (var key in userNameMap) {				
		var option = document.createElement("option");
		option.text = userNameMap[key];
	    option.value = key+"!"+userNameMap[key];
	    option.className = "form-control-general";
	    try {
	    	userNameSelect.add(option, null); //Standard 
	    }catch(error) {
	    	//regionSelect.add(option); // IE only
	    }
	}
}


/**
 * Function to populate the existing user fields
 * while user click on existing user radio button
 * @param selection
 */
function toogleSelectionDiv(selection) {
	if (selection == "E") {
		document.getElementById("newUserDiv").style.display = "none";
		document.getElementById("newUserAddDiv").style.display = "none";	
		document.getElementById("existingUserDiv").style.display = "block";	
		document.getElementById("existingUserAddDiv").style.display = "block";	
		document.getElementById("existingUserInputId").value = "";
	} else {
		document.getElementById("existingUserDiv").style.display = "none";	
		document.getElementById("existingUserAddDiv").style.display = "none";	
		document.getElementById("newUserDiv").style.display = "block";		
		document.getElementById("newUserAddDiv").style.display = "block";	
	}
}

/**
 * Function to populate the existing activate and deactivate facilitator
 */
function fetchExistingFacilitator() {
	var userGrp = Spine.Model.sub();
	userGrp.configure("/admin/manageuserFacilitator/populateExistingFacilitator", "customerId");
	userGrp.extend( Spine.Model.Ajax );
	//Populate the model with data to transfer
	var modelPopulator = new userGrp({  
		customerId : sessionStorage.getItem("customerId")
	});
	modelPopulator.save(); //POST
	userGrp.bind("ajaxError", function(record, xhr, settings, error) {
		alertify.alert("Unable to connect to the server.");
	});
	userGrp.bind("ajaxSuccess", function(record, xhr, settings, error) {
		var jsonStr = JSON.stringify(xhr);
		var obj = jQuery.parseJSON(jsonStr);
		var statusCode = obj.statusCode;
		if(statusCode == 200) {
			$(".loader_bg").fadeOut();
		    $(".loader").hide();
			populateTable(obj.existingFacilitators);
		}  else {
			//Show error message
			alertify.alert("Some thing went wrong.");
		}
	});
}

/**
 * Function to show the existing facilitator in table view
 * @param existingUsers
 */
function populateTable(existingUsers) {
	if(existingUsers.length > 0) {
		document.getElementById("existingChangeRoleList").style.display="block";
		var userList = document.getElementById("existingUsersTableId");
		var tableStr = "<table width='94%' border='1' bordercolor='#78C0D3' id='myTable' align='center'><thead class='tab_header'><tr><th>SL. No.</th><th>Facilitator Email</th><th>User Group Name</th><th width='16%'>Status</th></tr></thead><tbody>";
		var counter = 1;
		for (var index=0; index<existingUsers.length; index++) {
			var trColor = "";
			if(index % 2 == 0) {
				trColor = "#D2EAF0";
			} else {
				trColor = "#F1F1F1";
			}
						
			tableStr = tableStr + "<tr class='user_list_row_width' bgcolor='"+trColor+"'><td align='center'>"+counter+".</td><td align='center'>"+existingUsers[index].email+"</td><td align='center'>"+existingUsers[index].userGroup+"</td>";
			tableStr = tableStr + "<td class='user_list_row_width' >" +
					"<table width='100%' border='0'>" +
					"<tr>";
			if(existingUsers[index].softDelete == 0){
				tableStr = tableStr + "<td align='center'>Active</td></td></tr></table></td></tr>";
			} else {
				tableStr = tableStr + "<td align='center'>Inactive</td></td></tr></table></td></tr>";
			}
						
			counter = counter + 1;
		}
		tableStr = tableStr + "</tbody></table>";
		userList.innerHTML = tableStr;
		new SortableTable(document.getElementById('myTable'), 1);
		
		if (document.getElementById("existing_data_div")) {
			document.getElementById("existing_data_div").style.display = "none";
		}
	} else {
		document.getElementById("existingChangeRoleList").style.display="none";
		document.getElementById("existingUsersTableId").innerHTML = "<div align='center'><br /><br /><br /><img src='../img/no-record.png'><br /><div class='textStyleNoExist'>No Existing Facilitators</div></div>";
	}
}

/**
 * Function to update any the existing user to facilitator
 * @param null
 */
function updateUserToFacilitator() {
	var userVal = document.getElementById("existingUserInputId").options[document.getElementById("existingUserInputId").selectedIndex].value;
	if(validateUser(userVal)) {
	alertify.set({ buttonReverse: true });
	alertify.confirm("Are you sure you would like to replace your account with this new facilitator account?", function (e) {
		if (e) {		
			$(".loader_bg").fadeIn();
			$(".loader").fadeIn();
			var addFcltr = Spine.Model.sub();
			addFcltr.configure("/admin/manageuserFacilitator/updateUserToFacilitator", "userVal", "facilitatorEmail" ,"customerId");
			addFcltr.extend( Spine.Model.Ajax );
			//Populate the model with data to transfer
			var modelPopulator = new addFcltr({  
				userVal: userVal,
				facilitatorEmail: sessionStorage.getItem("jctEmail"),
				customerId : sessionStorage.getItem("customerId")
			});
			modelPopulator.save(); //POST
			addFcltr.bind("ajaxError", function(record, xhr, settings, error) {
				alertify.alert("Unable to connect to the server.");
			});
			addFcltr.bind("ajaxSuccess", function(record, xhr, settings, error) {
				var jsonStr = JSON.stringify(xhr);
				var obj = jQuery.parseJSON(jsonStr);
				var statusCode = obj.statusCode;
				if(statusCode == 200) {
					alertify.alert(obj.statusDesc);
					fetchUserDropdown();
					fetchExistingFacilitator();
					$(".loader_bg").fadeOut();
				    $(".loader").hide();			    
				    $('#alertify-ok').click(function() {	
				    	sessionStorage.clear();
				    	window.location = "../signup-page.jsp";
				    });
				} else if(statusCode == 999) {
					alertify.alert(obj.statusDesc);
					$(".loader_bg").fadeOut();
				    $(".loader").hide();			    
				}  else {
					alertify.alert("Some thing went wrong.");
				}
			});
			}		
		});	
	}
}


/**
 * Function to update any the existing user to facilitator
 * @param null
 */
function addNewFacilitator() {
	var newFacilitatorVal = document.getElementById("newUserInputId").value;
	if(validateEmailId(newFacilitatorVal)) {
	alertify.set({ buttonReverse: true });
	alertify.confirm("Are you sure you would like to replace your account with this new facilitator account?", function (e) {
		if (e) {	
			$(".loader_bg").fadeIn();
			$(".loader").fadeIn();
			var addFcltr = Spine.Model.sub();
			addFcltr.configure("/admin/manageuserFacilitator/addNewFacilitator", "newFacilitatorVal", "facilitatorEmail", "customerId");
			addFcltr.extend( Spine.Model.Ajax );
			//Populate the model with data to transfer
			var modelPopulator = new addFcltr({  
				newFacilitatorVal: newFacilitatorVal,
				facilitatorEmail: sessionStorage.getItem("jctEmail"),
				customerId : sessionStorage.getItem("customerId")
			});
			modelPopulator.save(); //POST
			addFcltr.bind("ajaxError", function(record, xhr, settings, error) {
				alertify.alert("Unable to connect to the server.");
				$(".loader_bg").fadeOut();
			    $(".loader").hide();
			});
			addFcltr.bind("ajaxSuccess", function(record, xhr, settings, error) {
				var jsonStr = JSON.stringify(xhr);
				var obj = jQuery.parseJSON(jsonStr);
				alertify.alert(obj.message);
				fetchUserDropdown();
				fetchExistingFacilitator();
				$(".loader_bg").fadeOut();
			    $(".loader").hide();
			    if(obj.statusCode != 999){
			    	$('#alertify-ok').click(function() {	
				    	sessionStorage.clear();
				    	window.location = "../signup-page.jsp";
				    });
			    }    
			});
		}	
	});
	}
}

/**
 * Function validates the user name.
 * @param userVal
 */
function validateUser(userVal) {
	if(userVal == "") {
		alertify.alert("Please select new facilitator.");
		return false;
	}
	return true;
}

/**
 * Function validates the email id.
 * @param emailId
 */
function validateEmailId(emailId) {	
	var filter = /^([a-zA-Z0-9_\.\-])+\@(([a-zA-Z0-9\-])+\.)+([a-zA-Z0-9]{2,4})+$/;
	if(emailId == "") {
		alertify.alert("Please provide new facilitator email.");
		return false;
	} else if (!filter.test(emailId)) {
		alertify.alert("Please provide a valid email address.");
    	return false;
    }
    return true;
}

/**
 * Function to populate the existing users
 * under the facilitator 
 */
function fetchExistingUsersManageUser() {
	var userGrp = Spine.Model.sub();
	userGrp.configure("/admin/manageuserFacilitator/populateUsersByFacilitatorId", "emailId", "customerId");
	userGrp.extend( Spine.Model.Ajax );
	//Populate the model with data to transfer
	var modelPopulator = new userGrp({  
		emailId: sessionStorage.getItem("jctEmail"),
		customerId : sessionStorage.getItem("customerId")
	});
	modelPopulator.save(); //POST
	userGrp.bind("ajaxError", function(record, xhr, settings, error) {
		alertify.alert("Unable to connect to the server.");
	});
	userGrp.bind("ajaxSuccess", function(record, xhr, settings, error) {
		var jsonStr = JSON.stringify(xhr);
		var obj = jQuery.parseJSON(jsonStr);
		var statusCode = obj.statusCode;
		if(statusCode == 200) {
			$(".loader_bg").fadeOut();
		    $(".loader").hide();
		    populateExistingUserTable(obj.existingUsers);
		}  else {
			//Show error message
			alertify.alert("Some thing went wrong.");
		}
	});
}




/**
 * Function to show the existing users in table view
 * @param existingUsers
 */
function populateExistingUserTable(existingUsers) {	
	if(existingUsers.length > 0) {
		document.getElementById("existingChangeRoleList").style.display="block";
		var userList = document.getElementById("existingUsersTableId");
		var tableStr = "<table width='94%' border='1' bordercolor='#78C0D3' id='myTable' align='center' class='tablesorter'><thead class='tab_header'><tr><th width='4%'>#</th><th>Name (Last, First)</th><th>Email (User ID)</th><th>User Group</th><th>Expiration Date</th></tr></thead><tbody>";
		var counter = 1;
		var fName= "";
		var lName = "";
		for (var index=0; index<existingUsers.length; index++) {
			var trColor = "";
			if(index % 2 == 0) {
				trColor = "#D2EAF0";
			} else {
				trColor = "#F1F1F1";
			}
			var expiryDate = new Date(existingUsers[index].jctAccountExpirationDate).toDateString();
			var userExpryDate = dateformat(new Date (expiryDate));		
		
			if(null != existingUsers[index].lastName) {
				lName = existingUsers[index].lastName;
			} else {
				lName = "N/A";
			}
			if(null != existingUsers[index].firstName) {
				fName = existingUsers[index].firstName;
			} else {
				fName = "N/A";
			}
			
			tableStr = tableStr + "<tr class='user_list_row_width' bgcolor='"+trColor+"'><td align='center'>"+counter+"</td><td align='center'>"+lName+", "+fName+"</td><td align='center'>"+existingUsers[index].email+"</td><td align='center'>"+existingUsers[index].userGroup+"</td><td align='center'>"+userExpryDate+"</td></tr>";	
			counter = counter + 1;
		}
		tableStr = tableStr + "</tbody></table>";
		userList.innerHTML = tableStr;
		
		$(function(){
			  $("table").tablesorter({
			    headers: {
			      0: { sorter: false }      // disable first column			     
			    }
			  });
			});
		
		if (document.getElementById("existing_data_div")) {
			document.getElementById("existing_data_div").style.display = "none";
		}
	} else {
		document.getElementById("existingChangeRoleList").style.display="none";
		document.getElementById("existingUsersTableId").innerHTML = "<div align='center'><br /><br /><br /><img src='../img/no-record.png'><br /><div class='textStyleNoExist'>No Existing Facilitators</div></div>";
	}
}