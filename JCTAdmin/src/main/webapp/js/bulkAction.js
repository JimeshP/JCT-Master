var emailIdCSVString = "";
var usrType = sessionStorage.getItem("userType");

$(document).ready(function() {
	fetchUserGroup();
	if(usrType == 1){
		document.getElementById("inputUserTypeInpt").selectedIndex = "0";
		$("#inputUserTypeInpt").change();
	} else {
		document.getElementById("inputUserTypeInpt").selectedIndex = "1";
		setTimeout(function(){disableUserGroup();},700);		
	}
});


function disableUserGroup() {
	document.getElementById("inputUserGroupInpt").selectedIndex = "1";
	document.getElementById("inputUserGroupInpt").disabled = true;
}
/**
 * Function fetches the user group
 */
function fetchUserGroup() {
	var userGrp = Spine.Model.sub();
	userGrp.configure("/admin/manageuser/populateUserGroup", "none");
	userGrp.extend( Spine.Model.Ajax );
	//Populate the model with data to transfer
	var modelPopulator = new userGrp({  
		none: "" 
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
			populateDropDown(obj.userGroupMap);
		}  else {
			//Show error message
			alertify.alert("Some thing went wrong.");
		}
	});
}
/**
 * Function populates the user group
 * @param userGroupMap
 */
function populateDropDown(userGroupMap) {
	var userGroupSelect = document.getElementById("inputUserGroupInpt");
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
/**
 * Function opens the div which contains the radio buttons 
 * after selecting the user group 
 */
function openRadioDiv() {
	var userGrpVal = document.getElementById("inputUserGroupInpt").options[document.getElementById("inputUserGroupInpt").selectedIndex].value;
	//  clean & reset existing data & div
	$("#existing_data_div_id").hide();
	$("#deleteId").prop("checked", false);
    $("#activeId").prop("checked", false);
    $("#inctiveId").prop("checked", false);
    $("#passwordId").prop("checked", false);

	if (userGrpVal != "") {
		document.getElementById('radioDrpDwnList').style.display = 'block';
	} else {
		document.getElementById('radioDrpDwnList').style.display = 'none';
		document.getElementById("buttonDivId").style.display = "none";
		document.getElementById('mainPopulationDiv').style.display = 'none';
		document.getElementById('existing_data_div_id').style.display = 'none';		
		//remove the checked items...
		/*if(document.getElementById('deleteId').checked) {
			document.getElementById('deleteId').checked = false;
		} else if(document.getElementById('activeId').checked) {
			document.getElementById('activeId').checked = false;
		} else if(document.getElementById('inctiveId').checked) {
			document.getElementById('inctiveId').checked = false;
		} else if(document.getElementById('passwordId').checked) {
			document.getElementById('passwordId').checked = false;
		}*/
	}
}

function enableDisableSelection() {
	var userTypeVal = document.getElementById("inputUserTypeInpt").options[document.getElementById("inputUserTypeInpt").selectedIndex].value;
	if (document.getElementById('radioDrpDwnList')) {
		document.getElementById('radioDrpDwnList').style.display = "none";
	}
	if (userTypeVal == "F") {
		 document.getElementById("inputUserGroupInpt").selectedIndex = "1";
		 document.getElementById("inputUserGroupInpt").disabled = true;
		 document.getElementById('radioDrpDwnList').style.display = "block";
		 document.getElementById("activeIdTxt").style.display = "none";
		 document.getElementById("inctiveIdTxt").style.display = "none";
		 document.getElementById('passwordDivId').style.left="-62%";
		 
	} else {
		 document.getElementById("inputUserGroupInpt").selectedIndex = "0";
		 document.getElementById("inputUserGroupInpt").disabled = false;
		 document.getElementById("activeIdTxt").style.display = "block";
		 document.getElementById("inctiveIdTxt").style.display = "block";
		 document.getElementById('passwordDivId').style.left="0%";
	}
	 if (document.getElementById("mainPopulationDiv")) {
		 document.getElementById("mainPopulationDiv").innerHTML = "";
	 }
	 if (document.getElementById("theButton")) {
		 document.getElementById("theButton").style.display = "none";
	 }
	 if (document.getElementById("existing_data_div_id")) {
		 document.getElementById("existing_data_div_id").style.display = "none";
	 }
	 
	 //remove the checked items...
	 if(document.getElementById('activeId').checked) {
		 document.getElementById('activeId').checked = false;
	 } else if(document.getElementById('inctiveId').checked) {
	 	 document.getElementById('inctiveId').checked = false;
	 } else if(document.getElementById('passwordId').checked) {
		 document.getElementById('passwordId').checked = false;
	 }
}
/**
 * Function decides which div to open and plot fetched data in it
 * @param id
 */
function decidePlotting(id) {
	alert("decidePlotting: 142");
	var displayDiv = document.getElementById("mainPopulationDiv");
	displayDiv.innerHTML = "<div align='center'><img src='../img/dataLoading.GIF' /><p> Loading Existing Users</p></div>";
	var userGrpVal = document.getElementById("inputUserGroupInpt").options[document.getElementById("inputUserGroupInpt").selectedIndex].value;
	var userTypeVal = document.getElementById("inputUserTypeInpt").options[document.getElementById("inputUserTypeInpt").selectedIndex].value;
	//fetch the username list containing the id and display in table
	var userGrp = Spine.Model.sub();
	var modelPopulator = null;
	if (document.getElementById('activeId').checked || document.getElementById('inctiveId').checked) {
		var status = 0;
		if(document.getElementById('activeId').checked) {
			status = 1;
		}
		userGrp.configure("/admin/manageuser/specificUserListForAdmin", "userGroup", "statusToBe", "emailId", "userType");
		userGrp.extend( Spine.Model.Ajax );
		//Populate the model with data to transfer
		modelPopulator = new userGrp({  
			userGroup: userGrpVal,
			statusToBe: status,
			emailId: sessionStorage.getItem("jctEmail"),
			userType: userTypeVal
		});
	} else {
		userGrp.configure("/admin/manageuser/fetchExistingUserNameForAdmin", "userGroup", "emailId", "userType");
		userGrp.extend( Spine.Model.Ajax );
		//Populate the model with data to transfer
		modelPopulator = new userGrp({  
			userGroup: userGrpVal,
			emailId: sessionStorage.getItem("jctEmail"),
			userType: userTypeVal
		});
	}
	
	modelPopulator.save(); //POST
	userGrp.bind("ajaxError", function(record, xhr, settings, error) {
		alertify.alert("Unable to connect to the server.");
	});
	userGrp.bind("ajaxSuccess", function(record, xhr, settings, error) {
		var jsonStr = JSON.stringify(xhr);
		var obj = jQuery.parseJSON(jsonStr);
		var statusCode = obj.statusCode;
		if(statusCode == 200) {
			populateTable(obj.existingUsers, id, obj.userType);
		}  else {
			//Show error message
			alertify.alert("Some thing went wrong.");
		}
	});
}
/**
 * Function populates and displays the fetched data in table.
 * @param existingUsers
 * @param id
 */
function populateTable(existingUsers, id, userType) {
	document.getElementById("existing_data_div_id").style.display = "block";
	document.getElementById("mainPopulationDiv").style.display = "block";
	if (existingUsers.length > 0) {
		var buttonText = "";
		if(id == "activeId") {
			buttonText = "Activate";
		} else if(id == "inctiveId") {
			buttonText = "Deactivate";
		} else if(id == "passwordId") {
			buttonText = "Reset";
		}
		var userList = document.getElementById("mainPopulationDiv");
		var tableStr = "<table width='94%' border='1' bordercolor='#78C0D3' id='myTable' align='center' class='tablesorter'><thead class='tab_header'><tr><td width='8%'></td><th>User Name</th></tr></thead><tbody>";
		var counter = 1;
		for (var index=0; index<existingUsers.length; index++) {
			var trColor = "";
			if(index % 2 == 0) {
				trColor = "#D2EAF0";
			} else {
				trColor = "#F1F1F1";
			}
			var activeStatus = existingUsers[index].activeInactiveStatus;
			var emailToSend = existingUsers[index].email+"`!!`"+userType;
			if (activeStatus == 3) {
				tableStr = tableStr + "<tr class='user_list_row_width' bgcolor='"+trColor+"'><td align='center'> <input type='checkbox' name='email' value='"+emailToSend+"' disabled> </td><td align='center'>"+existingUsers[index].email+"</td>";
			} else {
				tableStr = tableStr + "<tr class='user_list_row_width' bgcolor='"+trColor+"'><td align='center'> <input type='checkbox' name='email' value='"+emailToSend+"' checked> </td><td align='center'>"+existingUsers[index].email+"</td>";
			}
			
			counter = counter + 1;
		}
		tableStr = tableStr + "</tbody></table>";
		userList.innerHTML = tableStr;
		$("table").tablesorter(); 
		
		document.getElementById("buttonDivId").style.display = "block";
		document.getElementById("buttonDivId").innerHTML = "<table align='center'><tr><td style='padding:10px 0 0 0;'><input class='btn_admin btn_admin-primary btn_admin-sm search_btn' type='button' id='theButton' value='"+buttonText+"' onclick='performGrpAction(\""+id+"\")'></td></tr></table>";
	} else {
		document.getElementById("buttonDivId").style.display = "none";
		if(id == "activeId") {
			document.getElementById("mainPopulationDiv").innerHTML = "<div align='center'> All the users have been activated.</div>";
		} else if(id == "inctiveId") {
			document.getElementById("mainPopulationDiv").innerHTML = "<div align='center'> All the users have been deactivated.</div>";
		} else {
			document.getElementById("mainPopulationDiv").innerHTML = "<div align='center'> No user to reset the password.</div>";
		}
		
	}
}
/**
 * Function prepares the email id string separated by ~ and depending 
 * upon the type further actions will be performed
 * @param type
 */
function performGrpAction(type) {
	var checkedEmailString = "";
	var checkBoxes=document.getElementsByName("email");
	for (var i=0; i<checkBoxes.length; i++) {
		 if (checkBoxes[i].checked) {
			 if (checkBoxes.length-1 == i) {
				 checkedEmailString = checkedEmailString + checkBoxes[i].value;
			 } else {
				 checkedEmailString = checkedEmailString + checkBoxes[i].value + "~";
			 }
		 }
	}
	
	if (checkedEmailString == "") {
		alertify.alert("Please select at least one user.");
	} else {
		$(".loader_bg").fadeIn();
		$(".loader").fadeIn();
		if(type == "activeId") {
			changeStatusInBatch(checkedEmailString, 0);
		} else if(type == "inctiveId") {
			changeStatusInBatch(checkedEmailString, 1);
		} else if(type == "passwordId") {
			resetPasswordInBatch(checkedEmailString);
		}
	}
}
/**
 * Function changes the status from active to inactive for the email ids in email id string
 * @param checkedEmailString
 * @param status
 */
function changeStatusInBatch(checkedEmailString, status) {
	var userGrpVal = document.getElementById("inputUserGroupInpt").options[document.getElementById("inputUserGroupInpt").selectedIndex].value;
	var userTypeVal = document.getElementById("inputUserTypeInpt").options[document.getElementById("inputUserTypeInpt").selectedIndex].value;
	var userType = 1;
	if (userTypeVal == "F") {
		userType = 3;
	}
	var userGrp = Spine.Model.sub();
	userGrp.configure("/admin/manageuser/changeUserStatusInBatch", "emailIdString", "statusToBe", "userGrpVal", "userType");
	userGrp.extend( Spine.Model.Ajax );
	//Populate the model with data to transfer
	var modelPopulator = new userGrp({  
		emailIdString: checkedEmailString,
		statusToBe: status,
		userGrpVal: userGrpVal,
		userType: userType
	});
	modelPopulator.save(); //POST
	userGrp.bind("ajaxError", function(record, xhr, settings, error) {
		$(".loader_bg").fadeOut();
	    $(".loader").hide();
		alertify.alert("Unable to connect to the server.");
	});
	userGrp.bind("ajaxSuccess", function(record, xhr, settings, error) {
		var jsonStr = JSON.stringify(xhr);
		var obj = jQuery.parseJSON(jsonStr);
		var statusCode = obj.statusCode;
		if(statusCode == 200) {
			var activeStatus = "";
			if (status == 0) {
				activeStatus = "activeId";
			} else {
				activeStatus = "inctiveId";
			}
			$(".loader_bg").fadeOut();
		    $(".loader").hide();
			populateTable(obj.existingUsers, activeStatus);
		}  else {
			//Show error message
			alertify.alert("Some thing went wrong.");
		}
	});
}
/**
 * Function resets the password for the email ids in email id string
 * @param checkedEmailString
 */
function resetPasswordInBatch(checkedEmailString) {
	var userGrpVal = document.getElementById("inputUserGroupInpt").options[document.getElementById("inputUserGroupInpt").selectedIndex].value;
	var bulkPwdRst = Spine.Model.sub();
	bulkPwdRst.configure("/admin/manageuser/resetPasswordInBatch", "emailIdString", "updatedBy", "userGroup");
	bulkPwdRst.extend( Spine.Model.Ajax );
	//Populate the model with data to transfer
	var modelPopulator = new bulkPwdRst({  
		emailIdString: checkedEmailString,
		updatedBy: sessionStorage.getItem("jctEmail"),
		userGroup: userGrpVal
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
			$(".loader_bg").fadeOut();
		    $(".loader").hide();
			alertify.alert('Password reset not done successfully.');
		}  else if (statusCode == 222) {
			$(".loader_bg").fadeOut();
		    $(".loader").hide();
			alertify.alert(obj.statusDesc);
		} else {
			$(".loader_bg").fadeOut();
		    $(".loader").hide();
			alertify.alert('Password has been reset successfully.');
		}
	});
}


/**
 * USER WILL NOT BE DELETED. ONLY ACTIVATED OR INACTIVATED
 */
/*function performDeleteAction(emailString) {
	//alert('these will be deleted'+ emailString);
	var userGrp = Spine.Model.sub();
	userGrp.configure("/admin/manageuser/deleteBatchUsers", "none");
	userGrp.extend( Spine.Model.Ajax );
	//Populate the model with data to transfer
	var modelPopulator = new userGrp({  
		none: "" 
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
			populateTable(obj.existingUsers);
		}  else {
			//Show error message
			alertify.alert("Some thing went wrong.");
		}
	});

}*/