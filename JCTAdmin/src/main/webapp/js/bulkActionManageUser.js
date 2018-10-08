var totalUsers = 0;
var map = new Object();
var checkedEmailString = ""; 

function performAction(type) {
	//var checkedEmailString = "";
	var userToRenew = document.getElementById("subscribedUserId").value;
	var checkBoxes=document.getElementsByName("email");
	var totalUsers=$('input[name="email"]:checked').length;
	for (var i=0; i<checkBoxes.length; i++) {
		 if (checkBoxes[i].checked) {
			 if (checkBoxes.length-1 == i) {
				 checkedEmailString = checkedEmailString + checkBoxes[i].value;
			 } else {
				 checkedEmailString = checkedEmailString + checkBoxes[i].value + "~";
			 }
		 }
	}	
	if ($('input[name="email"]:checked').length == 0) {
		alertify.alert("Please select at least one user.");
	} else {
		if(type == "passwordId") {
			//resetPasswordInBatch(checkedEmailString, totalUsers);
			populateResetPasswordModal(checkedEmailString, totalUsers);
		} else if(type == "groupId") {
			//changeGroupInBatch(checkedEmailString, totalUsers);
			populateAssignGroupModal(checkedEmailString, totalUsers);
		} else if(type == "renewId") {
			if(totalUsers > userToRenew) {
				alertify.alert("Please purchase subscription to renew "+totalUsers+" user(s).");
			} else {
				populateSuccessValidation(checkedEmailString, totalUsers);
			}
		}
	}
}


function populateResetPasswordModal(checkedEmailString, totalUsers) {
	$('#resetPasswordModal').modal('show');
	document.getElementById("resetPassword").innerHTML = "You have selected "+totalUsers+" user(s) listed below. Would you like to <br>Reset Password for these users?<br>";
	var msg = "";
	msg = msg + "<div align='center'> <table style='border: 1px solid black; width: 90%;'><tr class='set-backgrnd-color'><td align='center'><b>User(s):</b></td></tr>";
	var checkBoxes=document.getElementsByName("email");
	for (var i=0; i<checkBoxes.length; i++) {
		 if (checkBoxes[i].checked) {	
			//var expiryDate = new Date(map[checkBoxes[i].value]);
		//	var expiryDate1= expiryDate.setMonth(expiryDate.getMonth() + 6);
		//	var newExpiryDate = dateformat(new Date (expiryDate1));		
			msg = msg + "<tr><td align='center'>"+checkBoxes[i].value+"</td></tr>";
		 }
	}
	msg = msg + "</table></div>";	
	document.getElementById("tableDataResetPassword").innerHTML = msg;
}

/**
 * Function resets the password for the email ids in email id string
 * @param checkedEmailString
 */
function resetPasswordInBatch() {
	$('#resetPasswordModal').modal('hide');
	$(".loader_bg").fadeIn();
	$(".loader").fadeIn();
	var bulkPwdRst = Spine.Model.sub();
	bulkPwdRst.configure("/admin/manageuserFacilitator/resetPasswordInBatchFacilitator", "emailIdString", "updatedBy", "customerId");
	bulkPwdRst.extend( Spine.Model.Ajax );
	//Populate the model with data to transfer
	var modelPopulator = new bulkPwdRst({  
		emailIdString: checkedEmailString,
		updatedBy: sessionStorage.getItem("jctEmail"),
		customerId : sessionStorage.getItem("customerId")
	});
	modelPopulator.save(); //POST
	bulkPwdRst.bind("ajaxError", function(record, xhr, settings, error) {
		alertify.alert("Unable to connect to the server.");
	});
	bulkPwdRst.bind("ajaxSuccess", function(record, xhr, settings, error) {
		var jsonStr = JSON.stringify(xhr);
		var obj = jQuery.parseJSON(jsonStr);
		var statusCode = obj.statusCode;
		if (statusCode == 201) {
			alertify.alert('Password reset was not successful.');
			$(".loader_bg").fadeOut();
		    $(".loader").hide();
		}  else if (statusCode == 222) {
			alertify.alert(obj.statusDesc);
			$(".loader_bg").fadeOut();
		    $(".loader").hide();
		} else {
			alertify.alert('Password has been reset successfully.');
			$(".loader_bg").fadeOut();
		    $(".loader").hide();
		}
	});
}


function populateSuccessValidation(emailIdString,totalUsers) {
	// Prepare message
	$('#bulkRenewModal').modal('show');
	document.getElementById("renewUser").innerHTML = "You have selected "+totalUsers+" user(s) below. Would you like to <br>renew these users?<br>";
	var msg = "";
	msg = msg + "<div align='center'> <table style='border: 1px solid black; width: 90%;'><tr class='set-backgrnd-color'><td align='center'><b>User(s):</b></td><td align='center'><b>New Expiration Date(s):</b></td></tr>";
	var checkBoxes=document.getElementsByName("email");
	for (var i=0; i<checkBoxes.length; i++) {
		 if (checkBoxes[i].checked) {	
			var expiryDate = new Date(map[checkBoxes[i].value]);
			var expiryDate1= expiryDate.setMonth(expiryDate.getMonth() + 6);
			var newExpiryDate = dateformat(new Date (expiryDate1));		
			msg = msg + "<tr><td align='center'>"+checkBoxes[i].value+"</td><td align='center'>"+newExpiryDate+"</td></tr>";
		 }
	}
	msg = msg + "</table></div>";	
	document.getElementById("tableDataRenew").innerHTML = msg;
}


function renewBulkUserAfterSuccess() {
	$('#bulkRenewModal').modal('hide');
	$(".loader_bg").fadeIn();
	$(".loader").fadeIn();
	var emailIdString = "";
	var totalUsers=$('input[name="email"]:checked').length;
	var checkBoxes=document.getElementsByName("email");
	for (var i=0; i<checkBoxes.length; i++) {
		 if (checkBoxes[i].checked) {
			 if (checkBoxes.length-1 == i) {
				 emailIdString = emailIdString + checkBoxes[i].value;
			 } else {
				 emailIdString = emailIdString + checkBoxes[i].value + "~";
			 }
		 }
	}
	var model = Spine.Model.sub();
	model.configure("/admin/manageuserFacilitator/renewBulkUserByFacilitator", "emailIdString", "totalUsers", "createdBy", "customerId");
	model.extend( Spine.Model.Ajax );
	//Populate the model with data to transfer
	var modelPopulator = new model({  			
		emailIdString: emailIdString,		
		totalUsers: totalUsers,
		createdBy: sessionStorage.getItem("jctEmail"),
		customerId : sessionStorage.getItem("customerId")
	});
	modelPopulator.save(); //POST
	model.bind("ajaxError", function(record, xhr, settings, error) {
		alertify.alert("Unable to connect to the server.");
		$(".loader_bg").fadeOut();
	    $(".loader").hide();
	});
	model.bind("ajaxSuccess", function(record, xhr, settings, error) {
		var jsonStr = JSON.stringify(xhr);
		var obj = jQuery.parseJSON(jsonStr);		
		$(".loader_bg").fadeOut();
	    $(".loader").hide();
	    if(obj.statusCode == 200) {
	    	alertify.alert((totalUsers)+" user(s) has been renewed successfully.");
	    	fetchSubscribedUserFacilitator();
	    	fetchExistingUsersManageUser();
	    } else {
	    	//error
	    }
	});
}

function populateAssignGroupModal(checkedEmailString, totalUsers) {
	$('#assignNewGroupModal').modal('show');
	document.getElementById("userGroupToAssignId").disabled = false;
	document.getElementById("createAssignUserGroupDiv").style.display = "none";
	document.getElementById("assignNewGroup").innerHTML = "You have selected "+totalUsers+" user(s) listed below. <br>Choose a group for these users?<br>";
	var msg = "";
	msg = msg + "<div align='center'> <table style='border: 1px solid black; width: 90%;'><tr class='set-backgrnd-color'><td align='center'><b>User(s):</b></td></tr>";
	var checkBoxes=document.getElementsByName("email");
	for (var i=0; i<checkBoxes.length; i++) {
		 if (checkBoxes[i].checked) {	
			//var expiryDate = new Date(map[checkBoxes[i].value]);
		//	var expiryDate1= expiryDate.setMonth(expiryDate.getMonth() + 6);
		//	var newExpiryDate = dateformat(new Date (expiryDate1));		
			msg = msg + "<tr><td align='center'>"+checkBoxes[i].value+"</td></tr>";
		 }
	}
	msg = msg + "</table></div>";	
	fetchUserGroupToAssignGrp();
	document.getElementById("tableDataAassignNewGroup").innerHTML = msg;
}

/**
 * Function to populate the drop-down for User Group 
 */
function fetchUserGroupToAssignGrp() {
	var userGrp = Spine.Model.sub();
	userGrp.configure("/admin/manageuserFacilitator/populateUserGroupDroDwnFacilitator", "customerId");
	userGrp.extend( Spine.Model.Ajax );
	//Populate the model with data to transfer
	var modelPopulator = new userGrp({  
		customerId: sessionStorage.getItem("customerId")
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
			populateAssignGroupDropDown(obj.userGroupMap);
		}  else {
			//Show error message
			alertify.alert("Some thing went wrong.");
		}
	});
}

/**
 * Function to populate the dropdown for user group
 */
function populateAssignGroupDropDown(userGroupMap) {
	document.getElementById("userGroupToAssignId").innerHTML = "";	
	var userGroupSelect = document.getElementById("userGroupToAssignId");
	var defaultOpn = document.createElement("option");
	defaultOpn.text = "None";
	defaultOpn.value = "2!None";
	defaultOpn.className = "form-control-general";
	userGroupSelect.add(defaultOpn, null);
	
	for (var key in userGroupMap) {				
		var option = document.createElement("option");
		option.text = userGroupMap[key];
	    option.value = key+"!"+userGroupMap[key];
	    option.className = "form-control-general";
	    try {
	    	userGroupSelect.add(option, null); //Standard 
	    }catch(error) {
	    	//regionSelect.add(option); // IE only
	    }
	}
}

function populateAssignToUserGrpDiv() {
	document.getElementById("createAssignUserGroupDiv").style.display = "block";
	document.getElementById("userGroupToAssignId").selectedIndex = "0";
	document.getElementById("userGroupToAssignId").disabled = true;
}

function validateUserGroupToAssignGrp() {
	var newUserGrpDiv = $("#createAssignUserGroupDiv").attr('style');
	if(newUserGrpDiv.search('block') != -1) {
		var userGroup = document.getElementById("newAssignUserGroupInptId").value;
		if(userGroup.trim().length == 0){
			alertify.alert("Please enter new user group name");
			return false;
		 } 
	 } /*else if (newUserGrpDiv.search('none') != -1) {
		var userGrpVal = document.getElementById("userGroupToAssignId").options[document.getElementById("userGroupToAssignId").selectedIndex].value;
		if(userGrpVal.trim().length == 0){
			alertify.alert("Please select user group");
			return false;
		 } 
	} */
	return true;
}

function assignNewGroupAfterSuccess() {
	if(validateUserGroupToAssignGrp()) {
		$('#assignNewGroupModal').modal('hide');
		$(".loader_bg").fadeIn();
		$(".loader").fadeIn();
		var emailIdString = "";
		var totalUsers=$('input[name="email"]:checked').length;
		var checkBoxes=document.getElementsByName("email");
		for (var i=0; i<checkBoxes.length; i++) {
			 if (checkBoxes[i].checked) {
				 if (checkBoxes.length-1 == i) {
					 emailIdString = emailIdString + checkBoxes[i].value;
				 } else {
					 emailIdString = emailIdString + checkBoxes[i].value + "~";
				 }
			 }
		}
		var userGrpVal = "";
		var newUserGroup= "";
		var newUserGrpDiv = $("#createAssignUserGroupDiv").attr('style');
		if(newUserGrpDiv.search('block') != -1) {
			userGrpVal = document.getElementById("newAssignUserGroupInptId").value;
			newUserGroup = "YES";
		} else {
			userGrpVal = document.getElementById("userGroupToAssignId").options[document.getElementById("userGroupToAssignId").selectedIndex].value;
			if(userGrpVal.trim().length == 0){
				userGrpVal = "2!None";
			}
			newUserGroup = "NO";
		}
		var model = Spine.Model.sub();
		model.configure("/admin/manageuserFacilitator/assignGroupInBatchFacilitator", "emailIdString", "totalUsers", "createdBy", "customerId", "userGroup", "newUserGroup");
		model.extend( Spine.Model.Ajax );
		//Populate the model with data to transfer
		var modelPopulator = new model({  			
			emailIdString: emailIdString,		
			totalUsers: totalUsers,
			createdBy: sessionStorage.getItem("jctEmail"),
			customerId : sessionStorage.getItem("customerId"),
			userGroup: userGrpVal,
			newUserGroup : newUserGroup
		});
		modelPopulator.save(); //POST
		model.bind("ajaxError", function(record, xhr, settings, error) {
			alertify.alert("Unable to connect to the server.");
			$(".loader_bg").fadeOut();
		    $(".loader").hide();
		});
		model.bind("ajaxSuccess", function(record, xhr, settings, error) {
			var jsonStr = JSON.stringify(xhr);
			var obj = jQuery.parseJSON(jsonStr);		
			$(".loader_bg").fadeOut();
		    $(".loader").hide();
		    if(obj.statusCode == 200) {
		    	alertify.alert((totalUsers)+" user(s) has been assigned to new group.");
		    	fetchSubscribedUserFacilitator();
		    	fetchExistingUsersManageUser();
		    	//fetchUserGroupToAssignGrp();
		    	fetchUserGroupForFacilitator();
		    } else {
		    	//error
		    }
		});
	}
}

function closeMyRenewModel() {
	$('#bulkRenewModal').modal('hide');
}

function closeResetPasswordModel() {
	$('#resetPasswordModal').modal('hide');
}

function closeAssignGroupModal() {
	$('#assignNewGroupModal').modal('hide');
}